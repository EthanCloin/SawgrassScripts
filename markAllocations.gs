/*
goal of this function is to identify and mark what materials are allocated by SKU
This should only need to run if the ledgering system isn't in proper order, and
at least once before I update the ledgering function to consider the checkbox
*/

function detectAllocatedMaterials() {
  var boxesChecked = 0;
  // chlastLedgerlledgerSheetpty lot field
  for (var row = 2; row <= lastLedgerRow; row++){lastLedgerRow   
    var lotCell = ledgerSheet.getRange(row, LOT_I+1);ledgerSheet
    if (lotCell.isBlank()){ledgerSheet
      Logger.log("is blank");
      // check for negative qtyledgerSheet
      var currentQty = ledgerSheet.getRange(row, QTY_I+1).getValue();
      if (currentQty < 0){
        Logger.log("is negative");
      // check PO/Job/Order for "*(C/L)**ledgerSheet
      var jobNumber = ledgerSheet.getRange(row, POJOB_I+1).getValue();
      if (jobNumber.toString().length === 6){
        Logger.log("length is 6");
        if (jobNumber[1] === 'C' || jobNumber[1] === 'L'){
          Logger.log("second entry is L or C");
          // check box
          ledgerSheet.getRange(row, CHECK_I+1).check();
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
