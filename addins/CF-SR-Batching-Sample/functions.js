"use strict";
/**
 * Adds two numbers without using batching
 * @CustomFunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function addNoBatch(first, second) {
    return first + second;
}
/**
 * Divides two numbers using batching
 * @CustomFunction
 * @param dividend The number being divided
 * @param divisor The number to divide the dividend with
 * @returns The sum of the two numbers
 */
function div2(dividend, divisor) {
    return _pushOperation("div2", [dividend, divisor]);
}
/**
 * Multiplies two numbers together using batching
 * @CustomFunction
 * @param first First number to multiply
 * @param second Second number to multiply
 * @returns The product of the two numbers
 */
function mul2(first, second) {
    return _pushOperation("mul2", [first, second]);
}
/**
 * Defines the implementation of the custom functions
 * for the function id defined in the metadata file (functions.json).
 */
CustomFunctions.associate("ADDNOBATCH", addNoBatch);
CustomFunctions.associate("DIV2", div2);
CustomFunctions.associate("MUL2", mul2);
const _batch = [];
let _isBatchedRequestScheduled = false;
// This function encloses your custom functions as individual entries,
// which have some additional properties so you can keep track of whether or not
// a request has been resolved or rejected.
function _pushOperation(op, args) {
    // Create an entry for your custom function.
    var invocationEntry = {
        operation: op,
        args: args,
        resolve: undefined,
        reject: undefined,
    };
    // Create a unique promise for this invocation,
    // and save its resolve and reject functions into the invocation entry.
    var promise = new Promise(function (resolve, reject) {
        invocationEntry.resolve = resolve;
        invocationEntry.reject = reject;
    });
    // Push the invocation entry into the next batch.
    _batch.push(invocationEntry);
    // If a remote request hasn't been scheduled yet,
    // schedule it after a certain timeout, e.g. 100 ms.
    if (!_isBatchedRequestScheduled) {
        _isBatchedRequestScheduled = true;
        setTimeout(_makeRemoteRequest, 100);
    }
    // Return the promise for this invocation.
    return promise;
}
// This is a private helper function, used only within your custom function add-in.
// You wouldn't call _makeRemoteRequest in Excel, for example.
// This function makes a request for remote processing of the whole batch,
// and matches the response batch to the request batch.
function _makeRemoteRequest() {
    // Copy the shared batch and allow the building of a new batch while you are waiting for a response.
    // Note the use of "splice" rather than "slice", which will modify the original _batch array
    // to empty it out.
    var batchCopy = _batch.splice(0, _batch.length);
    _isBatchedRequestScheduled = false;
    // Build a simpler request batch that only contains the arguments for each invocation.
    var requestBatch = batchCopy.map(function(item) {
        return { operation: item.operation, args: item.args };
    });
    // Make the remote request.
    _fetchFromRemoteService(requestBatch)
        .then(function (responseBatch) {
        // Match each value from the response batch to its corresponding invocation entry from the request batch,
        // and resolve the invocation promise with its corresponding response value.
        responseBatch.forEach(function(response, index) {
            if (response.error) {
                batchCopy[index].reject(new Error(response.error));
            }
            else {
                console.log(response);
                batchCopy[index].resolve(response.result);
            }
        });
    });
}
// --------------------- A public API ------------------------------
// This function simulates the work of a remote service. Because each service
// differs, you will need to modify this function appropriately to work with the service you are using. 
// This function takes a batch of argument sets and returns a [promise of] batch of values.
// NOTE: When implementing this function on a server, also apply an appropriate authentication mechanism
//       to ensure only the correct callers can access it.
function _fetchFromRemoteService(requestBatch) {
    // Simulate a slow network request to the server;
    return pause(7000).then(function(){
        return requestBatch.map(function(request) {
            // const { operation, args } = request;
            try {
                if (request.operation === "div2") {
                    // Divide the first argument by the second argument.
                    return {
                        result: request.args[0] / request.args[1]
                    };
                }
                else if (request.operation === "mul2") {
                    // Multiply the arguments for the given entry.
                    var myresult = request.args[0] * request.args[1];
                    console.log(myresult);
                    return {
                        result: myresult
                    };
                }
                else {
                    return {
                        error: 'Operation not supported: '+ request.operation
                    };
                }
            }
            catch (error) {
                return {
                    error: 'Operation failed: ' + request.operation
                };
            }
        });
    });
}
function pause(ms) {
    return new Promise(function (resolve) { setTimeout(resolve, ms) } );
}