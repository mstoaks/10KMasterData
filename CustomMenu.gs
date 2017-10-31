/*
File: CustomMenu.gs
Author: Max Stoaks
Purpose: Creates the custom menu that invokes the population script.
TODO:
*
*/



/*******************************************************************************************
// This method adds a custom menu item to run the script
********************************************************************************************/
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("10000ft",
             [{ name: "Get Project Data From 10K", functionName: "get10KProjectData" },
              { name: "Clear Data", functionName: "clearData" }
             ])
}
