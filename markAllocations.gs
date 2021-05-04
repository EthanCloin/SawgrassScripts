/*
goal of this function is to identify and mark what materials are allocated by SKU
This should only need to run if the ledgering system isn't in proper order, and
at least once before I update the ledgering function to consider the checkbox
*/

function detectAllocatedMaterials() {
  Logger.log("start function");
  
  var boxesChecked = 0;
  ledgerSheet.sort(LOT_I+1); // blanks sorted to bottom
  var activeLedgerRange = ledgerSheet.getActiveRange();
  var currentRow = ledgerSheet.getRange(activeLedgerRange.getLastRow(), 1, 1, lastLedgerCol);
  var currentLot = currentRow.getCell(1, LOT_I+1);

  while (currentLot.isBlank()){
   
    var row = currentRow.getRow();
     Logger.log("on row: " + row);
    // check for negative qty
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
          if (ledgerSheet.getRange(row, CHECK_I+1).isChecked()) {
            Logger.log("already checked");
            // move up a row
            currentRow = ledgerSheet.getRange(row-1, 1, 1, lastLedgerCol);
            currentLot = currentRow.getCell(1, LOT_I+1);
            continue;
          }
          else{
          ledgerSheet.getRange(row, CHECK_I+1).check();
          boxesChecked++;
          // move up a row
          currentRow = ledgerSheet.getRange(row-1, 1, 1, lastLedgerCol);
          currentLot = currentRow.getCell(1, LOT_I+1);
          }
        }
      }
    }
    // move up a row
    currentRow = ledgerSheet.getRange(row-1, 1, 1, lastLedgerCol);
    currentLot = currentRow.getCell(1, LOT_I+1);
  }
  Logger.log("resetting to ascending date");
  ledgerSheet.sort(DATE_I+1);
}
  // check for empty lot field
  // for (var row = 2; row <= lastLedgerRow; row++){ 
  //   var lotCell = ledgerSheet.getRange(row, LOT_I+1);
  //   if (lotCell.isBlank()){
  //     Logger.log("is blank");
  //     // check for negative qty
  //     var currentQty = ledgerSheet.getRange(row, QTY_I+1).getValue();
  //     if (currentQty < 0){
  //       Logger.log("is negative");
  //     // check PO/Job/Order for "*(C/L)**ledgerSheet
  //     var jobNumber = ledgerSheet.getRange(row, POJOB_I+1).getValue();
  //     if (jobNumber.toString().length === 6){
  //       Logger.log("length is 6");
  //       if (jobNumber[1] === 'C' || jobNumber[1] === 'L'){
  //         Logger.log("second entry is L or C");
  //         // check box
  //         ledgerSheet.getRange(row, CHECK_I+1).check();
  //         boxesChecked++;

  //       }
  //     }
  //   }
  // }
  // Logger.log("skip");
  // continue;
  // }
  



