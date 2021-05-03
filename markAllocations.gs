/*
goal of this function is to identify and mark what materials are allocated by SKU
This should only need to run if the ledgering system isn't in proper order, and
at least once before I update the ledgering function to consider the checkbox
*/

var ss = SpreadsheetApp.getActiveSpreadsheet();
  
var ledger = ss.getSheetByName("Ledger");
var lastRow = sheet.getLastRow(); // final ledger entry
var lastCol = sheet.getLastColumn(); // final ledger column
var ui = SpreadsheetApp.getUi();

function detectAllocatedMaterials() {
  var boxesChecked = 0;
  // check for empty lot field
  for (var row = 2; row <= lastRow; row++){
    
    var lotCell = ledger.getRange(row, LOT_I+1);
    if (lotCell.isBlank()){
      Logger.log("is blank");
      // check for negative qty
      var currentQty = ledger.getRange(row, QTY_I+1).getValue();
      if (currentQty < 0){
        Logger.log("is negative");
      // check PO/Job/Order for "*(C/L)****"
      var jobNumber = ledger.getRange(row, POJOB_I+1).getValue();
      if (jobNumber.toString().length === 6){
        Logger.log("length is 6");
        if (jobNumber[1] === 'C' || jobNumber[1] === 'L'){
          Logger.log("second entry is L or C");
          // check box
          ledger.getRange(row, CHECK_I+1).check();
          boxesChecked++;

        }
      }
    }
  }
  Logger.log("skip");
  continue;
  }
  ui.alert("Marked " + boxesChecked + " items as Allocated. Check Allocation Pivot for specifics.");
}
