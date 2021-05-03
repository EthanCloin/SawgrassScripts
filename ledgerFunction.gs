/*
@author Ethan Cloin
@version 2021-03-12

Contents:
  ledgerItemsWithLot()
  ledgerItemsWithSKU()
  onOpen()
  resetInputSheet()

Usage:
  This set of functions is designed specifically for use with SNL Master Inventory as it currently
  exists within GSuite. The functions efficiently perform inventory actions that previously required extensive user 
  interaction with the spreadsheet. 

  The function "ledgerItemsWithSKU" performs allocation on the main Ledger sheet as defined by the contents of
  inputTable sheet. Fill the inputTable with a list of materials that need to be allocated along with quantity,
  jobNumber, and comments. Leave the lot number column blank. Select "Allocate Materials by SKU" from the Ledgering
  dropdown menu and wait for the script to execute. Manual adjustment may be required based on the error messages 
  returned by the function. 
  The intended purpose of this function is to adjust the Ledger such that materials allocated for a job appear in a seperate 
  section in the Stock Check pivot table from those materials already consumed. This will allow purchasing team to be aware
  of what materials are in stock, and what needs to be sourced to fulfill a job. 

  The function "ledgerItemsWithLot" performs ledgering on the main Ledger sheet as defined by the contents of 
  inputTable sheet. Fill the inputTable with a list of materials that need to be ledgered along with quantity
  and job number, select "Ledger Materials by Lot" from the Ledgering dropdown menu, and wait for the script to 
  execute. Manual adjustment may be required based on the error messages returned by the function.
  The intended purpose of this function is to adjust the Ledger such that materials consumed by production are allocated
  and removed to reflect current stock. 
  To ensure that StockCheck is accurate, users will have to manually remove any entries made with prior use of the "ledgerItemsWithSKU" 
  function. This way, the reconciled usage by lot replaces the allocation instead of double-ledgering materials

  The function "onOpen" runs upon opening Master Inventory and creates the "Ledgering" dropdown menu. This menu is how users
  interface with the functions defined within this file. 

  The function "resetInputSheet" clears the contents of inputSheet
*/

/*
@author Ethan Cloin
@version 2021-02-15

This script scans the contents of inventory sheet "Ledger" until it locates a targeted row. 
The target row is one which matches a lot number AND description in the input sheet "inputTable". 

Upon locating, the function will duplicate and edit the duplicate such that the amount in the 
Quantity field of the inputTable is reflected. Whether the input is positive or negative, the script
will always provide a negative quantity value.
The date of the new ledger entry reflects the current date. The comment and job cells reflect cell values
from the input.

In the case that a lot/description combination is not found, the script will highlight the offending input row red
and notify the user with a pop up message. For all combinations successfully located, the respective input 
value will be highlighted green. 

There is also a function "resetInput" to reset the inputTable to empty, unhighlighted cells, and a function "onOpen"
providing a dropdown with options to execute either function. There are some simple checks to ensure that input 
is valid.

@input: input is taken from spreadsheet with columns identifying Lot, Description, Quantity, JobNumber and Comment
@output: output is reflected in ledger spreadsheet as new entries and can be easily accessed by 
locating the JobNumber in the Allocation Pivot in Master Inventory.

*/
function ledgerItemsWithLot(){

  // confirm data is input correctly
  if (!validInputSheet()){
    return;
  }

  var errorRowNumbers = [];

  for (var row = 2; row <= lastAssistantRow; row++){
    // read data from inputSheet
    var currentInputRow = assistantSheet.getRange(row, 1, 1, lastAssistantCol);
    var currentInputLot = currentInputRow.getCell(1, INPUT_LOT).getValue();
    var currentInputDesc = currentInputRow.getCell(1, INPUT_DESC).getValue();
    var currentInputQty = -1 * Math.abs(currentInputRow.getCell(1, INPUT_QTY).getValue());
    var currentInputJob = currentInputRow.getCell(1, INPUT_JOB).getValue();
    var currentInputComment = currentInputRow.getCell(1, INPUT_COMMENT).getValue();

    // create a TextFinder for the current Lot
    Logger.log(currentInputLot);
    var finder = ledgerSheet.getRange(2, LOT_I+1, lastLedgerRow, 1).createTextFinder(currentInputLot).matchEntireCell(true);
    var matchingLot = finder.findNext();
    Logger.log(matchingLot);
    // check for no lot found
    if (matchingLot === null){
          errorRowNumbers.push(row);
          continue;
    }

    while (true) {
      // get row containing matchingLot
      var rowNumber = matchingLot.getRow();
      var currentLedgerRow = ledgerSheet.getRange(rowNumber, 1, 1, lastLedgerCol);

      // check for matching Description
      var currentLedgerDesc = currentLedgerRow.getCell(1, DESC_I+1).getValue();
      if (currentLedgerDesc === currentInputDesc){
        // create new ledger entry
        ledgerSheet.insertRowAfter(rowNumber);
        var newLedgerEntry = ledgerSheet.getRange(rowNumber+1, 1, 1, lastLedgerCol);
        currentLedgerRow.copyTo(newLedgerEntry);
        // update fields of new entry per input sheet
        newLedgerEntry.getCell(1, DATE_I+1).setValue(new Date());
        newLedgerEntry.getCell(1, QTY_I+1).setValue(currentInputQty);
        newLedgerEntry.getCell(1, POJOB_I+1).setValue(currentInputJob);
        newLedgerEntry.getCell(1, COMMENT_I+1).setValue(currentInputComment);
        newLedgerEntry.getCell(1, CHECK_I+1).uncheck();
        break;
      }
      else {
        matchingLot = finder.findNext();
        // no valid match found
        if (matchingLot === null){
          errorRowNumbers.push(row);
          break;
        }
      }
    }
  }
  highlightMissedRows(errorRowNumbers);
}

  /*  ABOVE THIS LINE IS REFACTORED CODE 
      BELOW THIS LINE IS OLD AND SLOW */

/*
@author Ethan Cloin
@version 2021-03-12

This script scans the contents of inventory sheet "Ledger" until it locates a targeted row. 
The target row is one which matches a description in the input sheet "inputTable". 

Upon locating, the function will duplicate and edit the duplicate such that the amount in the 
Quantity field of the inputTable is reflected. Whether the input is positive or negative, the script
will always provide a negative quantity value.
The date of the new ledger entry reflects the current date. The comment and job cells reflect values
from the input.

In the case that a description is not found, the script will highlight the offending input row red
and notify the user with a pop up message. For all items successfully located and ledgered, the respective 
input value will be highlighted green. 

@input: input is taken from spreadsheet with columns identifying Lot, Description, Quantity, JobNumber and Comment
@output: output is reflected in ledger spreadsheet as new entries and can be easily accessed by 
locating the JobNumber in the Allocation Pivot in Master Inventory.
*/
function ledgerItemsWithSKU(){

  // confirm data is input correctly
  if (!validInputSheet()){
    return;
  }

  var errorRowNumbers = [];

  for (var row = 2; row <= lastAssistantRow; row++){
    // read data from inputSheet
    var currentInputRow = assistantSheet.getRange(row, 1, 1, lastAssistantCol);
    var currentInputDesc = currentInputRow.getCell(1, INPUT_DESC).getValue();
    var currentInputQty = -1 * Math.abs(currentInputRow.getCell(1, INPUT_QTY).getValue());
    var currentInputJob = currentInputRow.getCell(1, INPUT_JOB).getValue();
    var currentInputComment = currentInputRow.getCell(1, INPUT_COMMENT).getValue();

    // create a TextFinder for the current Description
    Logger.log(currentInputDesc);
    var finder = ledgerSheet.getRange(2, DESC_I+1, lastLedgerRow, 1).createTextFinder(currentInputDesc).matchEntireCell(true);
    var matchingDesc = finder.findNext();
    Logger.log(matchingDesc);
    // check for no lot found
    if (matchingDesc === null){
          errorRowNumbers.push(row);
          continue;
    }

    // get row containing matchingDesc
    var rowNumber = matchingDesc.getRow();
    var currentLedgerRow = ledgerSheet.getRange(rowNumber, 1, 1, lastLedgerCol);
    // create new ledger entry
    ledgerSheet.insertRowAfter(rowNumber);
    var newLedgerEntry = ledgerSheet.getRange(rowNumber+1, 1, 1, lastLedgerCol);
    currentLedgerRow.copyTo(newLedgerEntry);
    // update fields of new entry per input sheet
    newLedgerEntry.getCell(1, DATE_I+1).setValue(new Date());
    newLedgerEntry.getCell(1, QTY_I+1).setValue(currentInputQty);
    newLedgerEntry.getCell(1, POJOB_I+1).setValue(currentInputJob);
    newLedgerEntry.getCell(1, COMMENT_I+1).setValue(currentInputComment);
    newLedgerEntry.getCell(1, CHECK_I+1).check();
  }
  highlightMissedRows(errorRowNumbers);
}

// //check for actual input
// if (lastAssistantRow === 1) {
//   ui.alert("Provide values that you want to ledger!")
//   return;
// }

// //check for already highlighted rows
// if (assistantSheet.getActiveRange().getBackground() != 'white' && assistantSheet.getActiveRange().getBackground() != '#ffffff'){
//   //prompt user for approval
//   var response = ui.alert('The input sheet is not reset! Are you sure you want to continue?', ui.ButtonSet.YES_NO);
//   if (response == ui.Button.YES ){
//     //continue
//   }else if (response == ui.Button.NO){
//     return;
//   }
// }

//   var targetRange = []; // holds range on inputSheet, each index is a Range representing a single row
//   var lastTargetIndex; // denotes end of targetRange
  
//   //fill targetRange with values from input sheet
//   for (var i = 2; i <= lastAssistantRow; i++){
//     index = i - 2; // have to offset by 2 thanks to sheet syntax conflict with JS array syntax
    
//     targetRange[index] = assistantSheet.getRange(i, 1, 1, lastAssistantCol);
//     lastTargetIndex = index; // mysterious error when using targetRange.length, this index solves
//   }
//   var replacementRow; // holds the data for new ledger entry 
//   var curLot; 
//   var curDesc;
  
//   //loop through ledger
//   for (var i = 2; i <= lastLedgerRow; i++){

//     //end if targetRange is depleted
//     if (lastTargetIndex < 0){
//       break;
//     }
//     //check description of current row
//     curDesc = ledgerSheet.getRange(i, DESC_I+1).getValue();
    
//     //check input array for that description
//     for (var j = 0; j <= lastTargetIndex; j++){
    
//     //variables for inputTable values
//      var ledgerDesc = targetRange[j].getValues()[0][1];
//      var ledgerQty = Math.abs(targetRange[j].getValues()[0][2]); //absolute value to ensure positive
//      var ledgerJob = targetRange[j].getValues()[0][3];
//      var ledgerComment = targetRange[j].getValues()[0][4];
      
//      //if found description
//      if (ledgerDesc === curDesc){
//       //proceed with ledgering
//         //insert blank row in Ledger
//         ledgerSheet.insertRowAfter(i);
      
//         //update replacementRow to duplicate of current row
//         replacementRow = ledgerSheet.getRange(i, 1, 1, 15).getValues();
        
//         //update replacementRow to desired ledger quantity
//         replacementRow[0][QTY_I] = ledgerQty * -1; //switch sign to negative

//         //update replacementRow to desired comment
//         replacementRow[0][COMMENT_I] = ledgerComment;

//         //update replacementRow to desired JobNumber
//         replacementRow[0][POJOB_I] = ledgerJob;

//         //update replacementRow to current date
//         replacementRow[0][DATE_I] = new Date();

//         //update replacementRow to blank Lot Number
//         replacementRow[0][LOT_I] = "";

//         //insert updated row aka complete ledgering
//         ledgerSheet.getRange(i+1, 1, replacementRow.length, 15).setValues(replacementRow);

//         //remove row from targetRange
//         targetRange.splice(j, 1);

//         //decrement loop boundaries to prevent searching out of bounds
//         lastTargetIndex--;//using .length breaks it...idk why
//       }// end if matching desc block
//     }// end loop through input block
//   }// end loop through Ledger sheet block
  
// // loop through targetRange and grab remaining descriptions to store in array
// var remainingDesc = [];
// var isIncomplete = false;
// for (var i = 0; i < targetRange.length; i++){
//   remainingDesc[i] = targetRange[i].getValues()[0][1]; //may cause error
// }

// // loop through inputSheet and check for descriptions in remainingDesc
// for (var i = 2; i <= lastAssistantRow; i++){
//     var current = assistantSheet.getRange(i, 2, 1, 1); 
//     if (remainingDesc.includes(current.getValue())){
//       assistantSheet.getRange(i, 1, 1, lastAssistantCol).setBackground('red');
//       isIncomplete = true;
//     }
//     else{
//       assistantSheet.getRange(i, 1, 1, lastAssistantCol).setBackground('green');
//     }
// }

// // announce whether entries failed to ledger
// if (isIncomplete){
//   var message = "UNLEDGERED ITEMS!\nCheck inputSheet for details";
//     ui.alert(message);
// }else{
//   var message = "SUCCESSFULLY LEDGERED BY SKU!\nReset inputSheet for next user!";
//     ui.alert(message);
// }

//}


/*
creates menu
*/
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Inventory Functions")
    .addSubMenu(ui.createMenu("Ledgering")
    .addItem("Allocate Materials by SKU", "ledgerItemsWithSKU")
    .addItem("Ledger Materials by Lot", "ledgerItemsWithLot")
    .addItem("Reset Input Sheet", "resetInputSheet"))
    .addSeparator()
    .addItem("Relocator", "determineRelocationType")
    .addToUi(); 
}
/*
resets input sheet to blank
*/
function resetInputSheet(){
  var target = assistantSheet.getRange(2, 1, assistantSheet.getLastRow(), assistantSheet.getLastColumn());
  target.setBackground('white');
  target.clearContent();
}

/*
returns false if there is no data or a highlight in active range
*/
function validInputSheet(){
  //check for existing input
  if (lastAssistantRow === 1) {
    
    ui.alert("Provide values that you want to ledger!")
    return false;
  }

  //check for highlighted rows
  if (assistantSheet.getActiveRange().getBackground() != 'white' && assistantSheet.getActiveRange().getBackground() != '#ffffff'){
    //prompt user for approval
    var response = ui.alert('The input sheet is not reset! Are you sure you want to continue?', ui.ButtonSet.YES_NO);

    if (response == ui.Button.YES ){// ignore highlights
      return true;
    }
    else if (response == ui.Button.NO){// cancel script
      return false;
    }
  }
  return true;
}

/*
highlight the error rows
*/
function highlightMissedRows(errorRowNumbers){
  var missingAny = false;

  for (var row = 2; row <= lastAssistantRow; row++){
    if (errorRowNumbers.includes(row)){
      assistantSheet.getRange(row, 1, 1, lastAssistantCol).setBackground("crimson");
      missingAny = true;
      continue;
    }
    assistantSheet.getRange(row, 1, 1, lastAssistantCol).setBackground("mediumspringgreen");
  }

  if (missingAny){
    var message = "UNLEDGERED ITEMS!\nCheck ledgeringAssistant for details.";
    ui.alert(message);
  }
  else{
    var message = "SUCCESSFULLY LEDGERED!\nReset ledgeringAssistant for next use.";
    ui.alert(message);
  }
}