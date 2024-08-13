/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("refresh").onclick = refreshAllPivotTables;
  }
	console.log(info.host);
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem("Issue Tracking Dev");
      const issuTable = sheet.tables.getItem("IssueTrackDev");
      issuTable.rows.load("count");
      await context.sync();

      const firstBlankRow = issuTable.rows.count + 2;
      const rangeName = `A${firstBlankRow}:A${firstBlankRow + 30}`;
      const newRange = sheet.getRange(rangeName);
      const usedNewRange = newRange.getUsedRange();

      usedNewRange.load("rowCount");
      await context.sync();
      issuTable.resize(`A1:AE${firstBlankRow + usedNewRange.rowCount - 1}`);
      await context.sync();
	document.getElementById("result").innerHTML = "Update Successful";
	document.getElementById("result").style.color = "#00ee00";

      console.log("Update done");
    });
  } catch (error) {
	document.getElementById("result").innerHTML = "Update Failed";
	document.getElementById("result").style.color = "#ee0000";
    console.error(error);
  }
}

    async function refreshAllPivotTables() {
	try{
      await Excel.run(async (context) => {
        const workbook = context.workbook;
        const pivotTables = workbook.pivotTables;
        pivotTables.refreshAll();
        await context.sync();
	document.getElementById("result").innerHTML = "Refresh Successful";
	document.getElementById("result").style.color = "#00ee00";
        console.log("Update done");
      });
}	catch(error){
	document.getElementById("result").innerHTML = "Refresh Failed";
document.getElementById("result").style.color = "#ee0000";
    console.error(error);
}
    }

    async function tryCatch(callback) {
      try {
        await callback();
      } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
      }
    }
