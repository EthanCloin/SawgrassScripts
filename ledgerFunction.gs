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

//Constants to represent what columns are which index values
const DATE_I = 0, SKU_I = 1, DESC_I = 2, COA_I = 3, EXP_I = 4, QTY_I = 5, PACK_I = 6, SECTION_I = 7, LOT_I = 8, 
    VENDOR_I = 9, PRICE_I = 10, POJOB_I = 11, COMMENT_I = 12, COMPONENT_I = 13, CHECK_I = 14;

var ss = SpreadsheetApp.getActiveSpreadsheet();
  
var sheet = ss.getSheetByName("Ledger"); // ledger target
var input = ss.getSheetByName("inputTable"); // ledger input
var ui = SpreadsheetApp.getUi();

var lastRow = sheet.getLastRow(); // final ledger entry
var lastCol = sheet.getLastColumn(); // final ledger column
var lastInputRow = input.getLastRow(); // final input entry
var lastInputCol = input.getLastColumn(); // final input column

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

//check for actual input
if (lastInputRow === 1) {
  
  ui.alert("Provide values that you want to ledger!")
  return;
}

//check for already highlighted rows
if (input.getActiveRange().getBackground() != 'white' && input.getActiveRange().getBackground() != '#ffffff'){
  //prompt user for approval
  var response = ui.alert('The input sheet is not reset! Are you sure you want to continue?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES ){
    //continue
  }else if (response == ui.Button.NO){
    return;
  }
}

  var targetRange = []; // holds range on inputSheet, each index is a Range representing a single row
  var lastTargetIndex; // denotes end of targetRange
  
  //fill targetRange with values from input sheet
  for (var i = 2; i <= lastInputRow; i++){
    index = i - 2; // have to offset by 2 thanks to sheet syntax conflict with JS array syntax
    
    targetRange[index] = input.getRange(i, 1, 1, lastInputCol);
    lastTargetIndex = index; // unknown error when using targetRange.length, this index solves
  }
  var replacementRow; // holds the data for new ledger entry 
  var curLot; 
  var curDesc;
  
  for (var i = 2; i <= lastRow; i++){

    //end if targetRange is depleted
    if (lastTargetIndex < 0){
      break;
    }
    //check lot of current row
    curLot = sheet.getRange(i, LOT_I+1).getValue();

    //check input array for that lot
    for (var j = 0; j <= lastTargetIndex; j++){
      //variables for inputTable values
     var ledgerLot = targetRange[j].getValues()[0][0];
     var ledgerDesc = targetRange[j].getValues()[0][1];
     var ledgerQty = Math.abs(targetRange[j].getValues()[0][2]);
     var ledgerJob = targetRange[j].getValues()[0][3];
     var ledgerComment = targetRange[j].getValues()[0][4];
      
     //if found lot
     if (ledgerLot == curLot){
       
       //fetch and compare descriptions
       curDesc = sheet.getRange(i, DESC_I+1).getValue();
      
      //if matching lot && matching description
       if (ledgerDesc == curDesc){
  
        
        //proceed with ledgering

        //insert blank row in Ledger
        sheet.insertRowAfter(i);
      
        //update replacementRow to duplicate of current row
        replacementRow = sheet.getRange(i, 1, 1, 15).getValues();
        
        //update replacementRow to desired ledger quantity
        replacementRow[0][QTY_I] = ledgerQty * -1; //switch sign to negative

        //update replacementRow to desired comment
        replacementRow[0][COMMENT_I] = ledgerComment;

        //update replacementRow to desired JobNumber
        replacementRow[0][POJOB_I] = ledgerJob;

        //update replacement Row to current date
        replacementRow[0][DATE_I] = new Date();

        //insert updated row aka complete ledgering
        sheet.getRange(i+1, 1, replacementRow.length, 15).setValues(replacementRow);

        //remove row from targetRange
        targetRange.splice(j, 1);

        //decrement loop boundaries to prevent searching out of bounds
        j--;
        lastTargetIndex--;//using .length breaks it...idk why

       }
     }// end if matching lot
    }// end loop through input
  }// end loop through Ledger sheet  
  
// loop through targetRange and grab remaining lot numbers to store in array
var remainingLots = [];
var isIncomplete = false;
for (var i = 0; i < targetRange.length; i++){
  remainingLots[i] = targetRange[i].getValue();
}

// loop through inputSheet and check for lots in remainingLots
for (var i = 2; i <= lastInputRow; i++){
    var current = input.getRange(i, 1, 1, 1);

    if (remainingLots.includes(current.getValue())){
      input.getRange(i, 1, 1, lastInputCol).setBackground('red');
      isIncomplete = true;
    }
    else{
      input.getRange(i, 1, 1, lastInputCol).setBackground('green');
    }
}

// announce whether entries failed to ledger
if (isIncomplete){
  var message = "UNLEDGERED ITEMS!\nCheck inputSheet for details";
    ui.alert(message);
}else{
  var message = "SUCCESSFULLY LEDGERED BY LOT!\nReset inputSheet for next user!";
    ui.alert(message);
}

}
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

//check for actual input
if (lastInputRow === 1) {
  
  ui.alert("Provide values that you want to ledger!")
  return;
}

//check for already highlighted rows
if (input.getActiveRange().getBackground() != 'white' && input.getActiveRange().getBackground() != '#ffffff'){
  //prompt user for approval
  var response = ui.alert('The input sheet is not reset! Are you sure you want to continue?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES ){
    //continue
  }else if (response == ui.Button.NO){
    return;
  }
}

  var targetRange = []; // holds range on inputSheet, each index is a Range representing a single row
  var lastTargetIndex; // denotes end of targetRange
  
  //fill targetRange with values from input sheet
  for (var i = 2; i <= lastInputRow; i++){
    index = i - 2; // have to offset by 2 thanks to sheet syntax conflict with JS array syntax
    
    targetRange[index] = input.getRange(i, 1, 1, lastInputCol);
    lastTargetIndex = index; // mysterious error when using targetRange.length, this index solves
  }
  var replacementRow; // holds the data for new ledger entry 
  var curLot; 
  var curDesc;
  
  //loop through ledger
  for (var i = 2; i <= lastRow; i++){

    //end if targetRange is depleted
    if (lastTargetIndex < 0){
      break;
    }
    //check description of current row
    curDesc = sheet.getRange(i, DESC_I+1).getValue();
    
    //check input array for that description
    for (var j = 0; j <= lastTargetIndex; j++){
    
    //variables for inputTable values
     var ledgerDesc = targetRange[j].getValues()[0][1];
     var ledgerQty = Math.abs(targetRange[j].getValues()[0][2]); //absolute value to ensure positive
     var ledgerJob = targetRange[j].getValues()[0][3];
     var ledgerComment = targetRange[j].getValues()[0][4];
      
     //if found description
     if (ledgerDesc === curDesc){
      //proceed with ledgering
        //insert blank row in Ledger
        sheet.insertRowAfter(i);
      
        //update replacementRow to duplicate of current row
        replacementRow = sheet.getRange(i, 1, 1, 15).getValues();
        
        //update replacementRow to desired ledger quantity
        replacementRow[0][QTY_I] = ledgerQty * -1; //switch sign to negative

        //update replacementRow to desired comment
        replacementRow[0][COMMENT_I] = ledgerComment;

        //update replacementRow to desired JobNumber
        replacementRow[0][POJOB_I] = ledgerJob;

        //update replacementRow to current date
        replacementRow[0][DATE_I] = new Date();

        //update replacementRow to blank Lot Number
        replacementRow[0][LOT_I] = "";

        //insert updated row aka complete ledgering
        sheet.getRange(i+1, 1, replacementRow.length, 15).setValues(replacementRow);

        //remove row from targetRange
        targetRange.splice(j, 1);

        //decrement loop boundaries to prevent searching out of bounds
        lastTargetIndex--;//using .length breaks it...idk why
      }// end if matching desc block
    }// end loop through input block
  }// end loop through Ledger sheet block
  
// loop through targetRange and grab remaining descriptions to store in array
var remainingDesc = [];
var isIncomplete = false;
for (var i = 0; i < targetRange.length; i++){
  remainingDesc[i] = targetRange[i].getValues()[0][1]; //may cause error
}

// loop through inputSheet and check for descriptions in remainingDesc
for (var i = 2; i <= lastInputRow; i++){
    var current = input.getRange(i, 2, 1, 1); 
    if (remainingDesc.includes(current.getValue())){
      input.getRange(i, 1, 1, lastInputCol).setBackground('red');
      isIncomplete = true;
    }
    else{
      input.getRange(i, 1, 1, lastInputCol).setBackground('green');
    }
}

// announce whether entries failed to ledger
if (isIncomplete){
  var message = "UNLEDGERED ITEMS!\nCheck inputSheet for details";
    ui.alert(message);
}else{
  var message = "SUCCESSFULLY LEDGERED BY SKU!\nReset inputSheet for next user!";
    ui.alert(message);
}

}


/*
creates menu
*/
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Custom Functions")
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
  var target = input.getRange(2, 1, input.getLastRow(), input.getLastColumn());
  target.setBackground('white');
  target.clearContent();
}
