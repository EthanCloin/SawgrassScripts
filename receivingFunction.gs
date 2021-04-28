/*
@version 4-23-2021
@author Ethan Cloin

@input event object upon editing cells
@output formatted info transfer from sourceSheet(PO LOG) to destSheet (Ledger/Inventory Input)

This function automates the task of translating data provided by Inventory Team in the PO LOG
to our main Ledger by gathering and sorting the description, lot, expiration, etc into appropriate
format. Executes upon checking an unchecked Received box.
*/

//indices for PO LOG columns
//add 1 if using in Range, leave if using in array
const LOG_PO_I = 6, LOG_VENDOR_I = 7, LOG_DESC_I = 9, LOG_QTY_I = 10, LOG_PACK_I = 11, LOG_PRICE_I = 12,
LOG_RECEIVED_I = 16, LOG_LOT_I = 17, LOG_EXP_I = 18, LOG_COMMENT_I = 19, LOG_TO_INPUT_I = 20;
//indices for Ledger/Input colums
//add 1 if using in Range, leave if using in array
const IN_DATE_I = 0, IN_SKU_I = 1, IN_DESC_I = 2, IN_COA_I = 3, IN_EXP_I = 4, IN_QTY_I = 5, IN_PACK_I = 6, IN_SECTION_I = 7, 
IN_LOT_I = 8, IN_VENDOR_I = 9, IN_PRICE_I = 10, IN_POJOB_I = 11, IN_COMMENT_I = 12, IN_COMPONENT_I = 13, IN_CHECK_I = 14;

//global sheet definitions
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sourceSheet = ss.getSheetByName("PO LOG");
var sourceLastCol = sourceSheet.getLastColumn;
var destSheet = ss.getSheetByName("testSheet");
var destLastRow = destSheet.getLastRow();
var destLastCol = destSheet.getLastColumn();

/*
Primary function that triggers upon every edit to PO LOG sheet
Checks whether the edited range is a checkbox in received column,
highlights that cell yellow and copies the information into the destSheet
with appropriate formatting.  
*/
function onEdit(e){
  var spreadSheet = e.source;
    var range = e.range;
    if (!validEditedCell(range)){
      return;
    }

    var row = range.getRow();
    var col = range.getColumn();
    var sheetName = spreadSheet.getActiveSheet().getName();
    var input = e.value;
    
    if (sheetName == "PO LOG") {
      range.setBackground("yellow");  
      var editedRow = sourceSheet.getRange(row, 1, 1, LOG_TO_INPUT_I);
      var rowValues = editedRow.getValues();
      
      //function to sort values into Ledger format
      var sortedDataValues = sortLogRow(rowValues);

      //paste into Input sheet
      var targetRange = destSheet.getRange(destLastRow+1, 1, 1, IN_CHECK_I+1);
      targetRange.setValues(sortedDataValues);
    }
}

/*
Use this function to test the sorting function
*/
function test(){
  var testData = sourceSheet.getRange(344, 1, 1, LOG_TO_INPUT_I).getValues();
  Logger.log("INPUT: " + testData[0]);
  var result = sortLogRow(testData);
  Logger.log("RESULT: " + result[0]);
}

/*
Returns true if the edited range is in the Received Checkbox column
*/
function validEditedCell(range) {
    column = range.getColumn();
    if (column === LOG_RECEIVED_I + 1 && range.isChecked()){
      return true;
    }
    else return false;
}

/*
receives a 2d array of values from the row
*/
function sortLogRow(rowDataValues) {
  var sortedDataValues = rowDataValues;
  sortedDataValues[0][IN_DATE_I] = new Date();
  sortedDataValues[0][IN_SKU_I] = "";
  sortedDataValues[0][IN_DESC_I] = rowDataValues[0][LOG_DESC_I]; 
  sortedDataValues[0][IN_COA_I] = "";
  sortedDataValues[0][IN_EXP_I] = rowDataValues[0][LOG_EXP_I]; 
  sortedDataValues[0][IN_QTY_I] = rowDataValues[0][LOG_QTY_I]; 
  sortedDataValues[0][IN_PACK_I] = rowDataValues[0][LOG_PACK_I]; 
  sortedDataValues[0][IN_SECTION_I] = "";
  sortedDataValues[0][IN_LOT_I] = rowDataValues[0][LOG_LOT_I];
  sortedDataValues[0][IN_VENDOR_I] = rowDataValues[0][LOG_VENDOR_I]; 
  sortedDataValues[0][IN_PRICE_I] = rowDataValues[0][LOG_PRICE_I]; 
  sortedDataValues[0][IN_POJOB_I] = rowDataValues[0][LOG_PO_I]; 
  sortedDataValues[0][IN_COMMENT_I] = rowDataValues[0][LOG_COMMENT_I]; 
  sortedDataValues[0][IN_COMPONENT_I] = "";
  sortedDataValues[0].splice(15, 5);
  
  Logger.log(sortedDataValues[0]);
  return sortedDataValues;
}
