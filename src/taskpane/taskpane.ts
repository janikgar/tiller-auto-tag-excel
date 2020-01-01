/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    // document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = getAutoTagDataRange;
  }
});

async function getAutoTagDataRange() {
  await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");

    await context.sync()
      .then(() => {
        for (let sheet of sheets.items) {
          if (sheet.name === "AutoTag") {
            getAutoTagData(sheet.id);    
          }
        }
      }).catch((reason) => {
        console.log(reason);
      })
  });
}

async function getAutoTagData(sheetId: string) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(sheetId).getUsedRange(true);
    sheet.load();

    await context.sync()
      .then(() => {
        let autoTagData = []
        let values = sheet.text;
        let keys = values.shift();
        for (let row of values) {
          let autoTagRow = {}
          for (let i = 0; i < keys.length; i++) {
            autoTagRow[keys[i]] = row[i]
          }
          autoTagData.push(autoTagRow)
        }
        console.log(autoTagData);
      })
      .catch((err) => {
        console.log(err);
      })
  })
}