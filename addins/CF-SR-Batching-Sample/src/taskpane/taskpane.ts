/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize function must be run each time a new page is loaded
Office.initialize = () => {
  document.getElementById('BtnInvoke').onclick = buttonClick;
};

export function buttonClick() {
  run().then( function() {
    console.log("Completed run");
  });
}

async function run() {
  try {
    await Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A1:M1344");

      // Read the range address
      range.load("address");

      // using the MUL2 formula
      let arrayOfFormulas = [
        ["=CONTOSO.MUL2(2,1000)"]
      ];

      range.formulas = <any>"=CONTOSO.MUL2(2,1000)";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
