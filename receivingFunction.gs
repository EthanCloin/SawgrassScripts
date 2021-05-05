/*
@version 2021-05-04 --NEED TO UPDATE USAGE ETC
@author Ethan Cloin

@input event object upon editing cells
@output formatted info transfer from sourceSheet(PO LOG) to destSheet (Ledger/Inventory Input)

This function automates the task of translating data provided by Inventory Team in the PO LOG
to our main Ledger by gathering and sorting the description, lot, expiration, etc into appropriate
format. Executes upon checking an unchecked Received box.
*/


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
      var editedRow = POLogSheet.getRange(row, 1, 1, LOG_TO_INPUT_I);
      var rowValues = editedRow.getValues();
      
      //function to sort values into Ledger format
      var sortedDataValues = sortLogRow(rowValues);

      //paste into Input sheet
      var targetRange = inventoryInputSheet.getRange(inventoryInputLastRow+1, 1, 1, CHECK_I+1);
      targetRange.setValues(sortedDataValues);
    }
}

/*
Use this function to test the sorting function
*/
function test(){
  var testData = POLogSheet.getRange(344, 1, 1, LOG_TO_INPUT_I).getValues();
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
  sortedDataValues[0][DATE_I] = new Date();
  sortedDataValues[0][SKU_I] = "";
  sortedDataValues[0][DESC_I] = rowDataValues[0][LOG_DESC_I]; 
  sortedDataValues[0][COA_I] = "";
  sortedDataValues[0][EXP_I] = rowDataValues[0][LOG_EXP_I]; 
  sortedDataValues[0][QTY_I] = rowDataValues[0][LOG_QTY_I]; 
  sortedDataValues[0][PACK_I] = rowDataValues[0][LOG_PACK_I]; 
  sortedDataValues[0][SECTION_I] = "";
  sortedDataValues[0][LOT_I] = rowDataValues[0][LOG_LOT_I];
  sortedDataValues[0][VENDOR_I] = rowDataValues[0][LOG_VENDOR_I]; 
  sortedDataValues[0][PRICE_I] = rowDataValues[0][LOG_PRICE_I]; 
  sortedDataValues[0][POJOB_I] = rowDataValues[0][LOG_PO_I]; 
  sortedDataValues[0][COMMENT_I] = rowDataValues[0][LOG_COMMENT_I]; 
  sortedDataValues[0][COMPONENT_I] = "";
  sortedDataValues[0].splice(15, 5);
  
  Logger.log(sortedDataValues[0]);
  return sortedDataValues;
}
