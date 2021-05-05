/*
@author Ethan Cloin
@version 2021-05-04 -- NEED TO UPDATE USAGE ETC

goal of this function is to identify and mark what materials are allocated by SKU
This should only need to run if the ledgering system isn't in proper order, and
at least once before I update the ledgering function to consider the checkbox
*/

function detectAllocatedMaterials() {  
  var boxesChecked = 0;
  var boxesSkipped = 0;
  ledgerSheet.sort(LOT_I+1);

  //var activeLedgerRange = ledgerSheet.getActiveRange();
  var currentRow = ledgerSheet.getRange(ledgerSheet.getLastRow(), 1, 1, lastLedgerCol);
  var currentLot = currentRow.getCell(1, LOT_I+1);
  //Logger.log(currentRow.getRow());

  while (currentLot.isBlank()){
    var row = currentRow.getRow();
    //Logger.log(row);

    // skip empty description
    if (currentRow.getCell(1, DESC_I+1).isBlank()){
      // move up a row
    currentRow = ledgerSheet.getRange(row-1, 1, 1, lastLedgerCol);
    currentLot = currentRow.getCell(1, LOT_I+1);
    }
    // check for negative qty
    var currentQty = ledgerSheet.getRange(row, QTY_I+1).getValue();
    if (currentQty < 0){
      // check PO/Job/Order for "*(C/L)**ledgerSheet
      var jobNumber = ledgerSheet.getRange(row, POJOB_I+1).getValue();
      if (jobNumber.toString().length === 6){
        if (jobNumber[1] === 'C' || jobNumber[1] === 'L'){
          // checked box
          if (ledgerSheet.getRange(row, CHECK_I+1).isChecked()) {
            // move up a row
            boxesSkipped++;
          }
          // unchecked box
          else{
          ledgerSheet.getRange(row, CHECK_I+1).check();
          boxesChecked++;
          }
        }
      }
    }
    // move up a row
    currentRow = ledgerSheet.getRange(row-1, 1, 1, lastLedgerCol);
    currentLot = currentRow.getCell(1, LOT_I+1);
  }
  Logger.log("checked " + boxesChecked + " new boxes");
  Logger.log("confirmed " + boxesSkipped + " already checked"); 
  Logger.log("resetting to ascending date");
  ledgerSheet.sort(DATE_I+1);
}


