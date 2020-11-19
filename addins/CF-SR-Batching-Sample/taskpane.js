/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize function must be run each time a new page is loaded
Office.initialize = function () {
  document.getElementById('BtnInvoke').onclick = buttonClick;
};

function buttonClick() {
  var context = new Excel.RequestContext();
  var range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:M1344");
   // Read the range address
   range.load("address");
   range.formulas = "=CONTOSO.MUL2(2,1000)";
   return context.sync().then (function() {
      console.log('Executed 18000 CFs');
   });
}

function action(event) {
  // Your code goes here
  buttonClick();
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
