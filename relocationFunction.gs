/*
Goal of the function:

Allow the user to update the location of materials in a few ways:
  1. Move ALL of one ingredient from Section A --> Section B
  2. Move ALL of one pallet from Section A --> Section B
  3. Move SOME of one ingredient from Section A --> Section B

@input  1. Data specifying an ingredient (sku, lot, desired section)
        2. Data specifying a pallet (section)
        3. Data specifying an ingredient (sku, lot, qty, desired section)

@ouput  1. N Updated ledger entries to the desired section
        2. N Updated ledger entries to the desired section
        3. Updated ledger entry with decreased quantity in original section
        AND new ledger entry with the specified relocated quantity and section
*/
// Column Constants (index, so add 1 for row use)
const SRC_SECTION = 0, DEST_SECTION = 1, TARGET_DESC = 2, TARGET_LOT = 3, TARGET_QTY = 4, 
      RELOC_TYPE = 5, OUTPUT = 6;

// Global References
var ss = SpreadsheetApp.getActiveSpreadsheet();
var ui = SpreadsheetApp.getUi();
var inputSheet = ss.getSheetByName("Relocator");
var ledgerSheet = ss.getSheetByName("Ledger");
var lastLedgerRow = ledgerSheet.getLastRow();
var lastLedgerCol = ledgerSheet.getLastColumn();
var lastInputRow = inputSheet.getLastRow();
var lastInputCol = inputSheet.getLastColumn();

/*

*/
function relocateFullQty(sourceRowNum) {
  if (!validInputDataForFull(sourceRowNum)){
    inputSheet.getRange(sourceRowNum, 1, 1, lastInputCol).setBackground("gray");
    return;
  }

  // scan Relocator and store description, lot, source, and destination section
  var sourceRow = inputSheet.getRange(sourceRowNum, SRC_SECTION+1, 1, OUTPUT+1);
  var outputCell = sourceRow.getCell(1, OUTPUT+1);
  var targetDescription = sourceRow.getCell(1, TARGET_DESC + 1).getValue();
  var targetLot = sourceRow.getCell(1, TARGET_LOT + 1).getValue();
  var sourceSection = sourceRow.getCell(1, SRC_SECTION + 1).getValue();
  var destinationSection = sourceRow.getCell(1, DEST_SECTION + 1).getValue();
  var editedMaterials = [];
  
  
  // search Ledger for matching lot
  var finder =  ledgerSheet.getRange(2, LOT_I+1, lastLedgerRow, 1)
                .createTextFinder(targetLot)
                .matchEntireCell(true)
                
  var matchingLotsRange = finder.findAll();
  var foundCount = matchingLotsRange.length;

  for (var i = 0; i < foundCount; i++){
    // get the whole row
    var currentCell = matchingLotsRange[i];
    var matchingRow = ledgerSheet.getRange(currentCell.getRow(), 1, 1, lastLedgerCol);

    // check row for positive qty
    var currentQty = matchingRow.getCell(1, QTY_I+1).getValue();
    if (currentQty <= 0){continue;}

    // check row for matching desc
    var currentDesc = matchingRow.getCell(1, DESC_I+1).getValue();
    if (currentDesc != targetDescription){continue;}

    // check row for matching source
    var sectionCell = matchingRow.getCell(1, SECTION_I+1);
    var currentSection = sectionCell.getValue();
    if (sourceSection != "" && currentSection != sourceSection){continue;}

    // update Section value
    sectionCell.setValue(destinationSection);

    var result = "Moved from " + currentSection.toString() + " to " + destinationSection.toString() + " on row " + matchingRow.getRow() + "\n";
    editedMaterials.push(result);
  }
  outputCell.setValue(editedMaterials.toString());
}

/*
need to run tests
*/
function relocatePartialQty(sourceRowNum){
  if (!validInputDataForPartial(sourceRowNum)){
      inputSheet.getRange(sourceRowNum, 1, 1, lastInputCol).setBackground("gray");
      return;
    }  
  // scan Relocator and store description, lot, source, destination, and qty
  var sourceRow = inputSheet.getRange(sourceRowNum, SRC_SECTION+1, 1, OUTPUT+1);
  var outputCell = sourceRow.getCell(1, OUTPUT+1);
  var targetDescription = sourceRow.getCell(1, TARGET_DESC + 1).getValue();
  var targetLot = sourceRow.getCell(1, TARGET_LOT + 1).getValue();
  var sourceSection = sourceRow.getCell(1, SRC_SECTION + 1).getValue();
  var destinationSection = sourceRow.getCell(1, DEST_SECTION + 1).getValue();
  var targetQty = Math.abs(sourceRow.getCell(1, TARGET_QTY+1).getValue());
  var editedMaterials = [];
  var possibleErrors = [];

  // search Ledger for matching line items
  var finder =  ledgerSheet.getRange(2, LOT_I+1, lastLedgerRow, 1)
                  .createTextFinder(targetLot)
                  .matchEntireCell(true)
  var matchingLotsRange = finder.findAll();
  var foundCount = matchingLotsRange.length;

  for (var i = 0; i < foundCount; i++){
    // get the whole row
    var currentCell = matchingLotsRange[i];
    var matchingRow = ledgerSheet.getRange(currentCell.getRow(), 1, 1, lastLedgerCol);

    // check description
    var currentDesc = matchingRow.getCell(1, DESC_I+1).getValue();
    if (currentDesc != targetDescription){
      possibleErrors.push("wrong desc on row " + matchingRow.getRow());
      continue;
      }

    // check section
    var sectionCell = matchingRow.getCell(1, SECTION_I+1);
    var currentSection = sectionCell.getValue();
    if (sourceSection != "" && currentSection != sourceSection){
      possibleErrors.push("wrong section on row " + matchingRow.getRow());
      continue;
      }
    var prevSection = currentSection;

    // check that qty > qty to remove
    var currentQty = matchingRow.getCell(1, QTY_I+1).getValue();
    if (currentQty <= targetQty){
      possibleErrors.push("not enough available qty on row " + matchingRow.getRow());
      continue;
      }


  // duplicate matching line
    ledgerSheet.insertRowAfter(matchingRow.getRow());
    var newRow = ledgerSheet.getRange(matchingRow.getRow() + 1, 1, 1, lastLedgerCol); 
    matchingRow.copyTo(newRow);

  // update duplicate to the qty and section from Relocator
    newRow.getCell(1, SECTION_I + 1).setValue(destinationSection);
    newRow.getCell(1, QTY_I+1).setValue(targetQty);

  // update matching line by subtracting qty from Relocator
    matchingRow.getCell(1, QTY_I+1).setValue(currentQty - targetQty);
    var currentUnit = matchingRow.getCell(1, PACK_I+1).getValue();

  // mark as complete and show edited row number for easy confirmation
    var result = "Moved " + targetQty + currentUnit + " of " + targetDescription +  
    " on row " + matchingRow.getRow() + " from " + currentSection + " to " + destinationSection +"\n"
    editedMaterials.push(result);
 }
 // check whether a result was found
 if (editedMaterials.length > 0){
    outputCell.setValue(editedMaterials);
    outputCell.setBackground("mediumspringgreen");

 }
 else{
   outputCell.setValue(possibleErrors);
   outputCell.setBackground("crimson");
 }
}

/*
trying to brainstorm a better way than searching whole sheet because thattt is a long process
*/
function relocatePallet(sourceRowNum){
  // scan Relocator and store source and destination section
  var sourceRow = inputSheet.getRange(sourceRowNum, SRC_SECTION+1, 1, OUTPUT+1);
  var outputCell = sourceRow.getCell(1, OUTPUT+1);
  var sourceSection = sourceRow.getCell(1, SRC_SECTION + 1);
  var destinationSection = sourceRow.getCell(1, DEST_SECTION + 1);
  var editedMaterials = [];
  // data validation
  if (!validInputDataForPallet(sourceRowNum)){
    inputSheet.getRange(sourceRowNum, 1, 1, lastInputCol).setBackground("gray");
    return;
  }
  
  // find all matching cells
  var finder = ledgerSheet.getRange(2, SECTION_I+1, lastLedgerRow).createTextFinder(sourceSection.getValue()).matchEntireCell(true);
  var foundCells = finder.findAll();
  Logger.log(foundCells);
  var foundCount = foundCells.length; 
  if (foundCount === 0){
    editedMaterials.push("NONE FOUND");
    Logger.log(editedMaterials);
    outputCell.setValue(editedMaterials);
    return;
  }
  
  for (var i = 0; i < foundCount; i++){
    var currentCell = foundCells[i];
    if (currentCell.getColumn() === SECTION_I+1){
      // update value
      currentCell.setValue(destinationSection.getValue());

      // get description and rownum
      var currentRowNum = currentCell.getRow();
      var currentDesc = ledgerSheet.getRange(currentRowNum, DESC_I+1).getValue();
      var result = currentDesc + " on Row " + currentRowNum;
      editedMaterials.push(result);
    }
  }
  outputCell.setValue("Moved " + editedMaterials.length + " items: " + editedMaterials.toString());
}

/*
Check the Relocation Type column for all rows and call appropriate function
Full 
Partial
Pallet
*/
function determineRelocationType(){
  for (var i = 2; i <= lastInputRow; i++) {
    var typeCell = inputSheet.getRange(i, RELOC_TYPE + 1);
    relocationType = typeCell.getValue();
    Logger.log("Relocation type: " + relocationType);

    switch (relocationType){
      case "Full":
        relocateFullQty(i);
        break;
      case "Partial":
        relocatePartialQty(i);
        break;
      case "Pallet":
        relocatePallet(i);
        break;
      default:
        ui.alert("Missing Relocation Type! Fix row " + i);
    }

  }
}

/*
Ensures that the required data is present in a given row
*/
function validInputDataForFull(sourceRowNum) {
  var sourceRow = inputSheet.getRange(sourceRowNum, 1, 1, lastInputCol);
  var outputCell = sourceRow.getCell(1, OUTPUT+1);
  var values = sourceRow.getValues();
  var result = [];

  if (values[0][DEST_SECTION] === "")  {
    result.push("Missing Destination! ");
  }
  if (values[0][TARGET_DESC] === "") {
    result.push("Missing Description! ");
  }
  if (values[0][TARGET_LOT] === "") {
    result.push("Missing Lot! ");
  }
  if (values[0][SRC_SECTION] === ""){
    // prompt whether user is okay with changing values regardless of source section
    var response = ui.alert('The source section is blank! Are you sure you want to move all materials, regardless of current location?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES ){
      //continue
    }else if (response == ui.Button.NO){
      return false;
    }
  }
  
  if (result.length != 0){
    outputCell.setValue(result.toString());
    return false;
  }
  else
  return true;
}

/*
Ensures that the required data is present in a given row
*/
function validInputDataForPartial(sourceRowNum) {
  var sourceRow = inputSheet.getRange(sourceRowNum, 1, 1, lastInputCol);
  var outputCell = sourceRow.getCell(1, OUTPUT+1);
  var values = sourceRow.getValues();
  var result = [];

  if (values[0][SRC_SECTION] === "")  {
      result.push("Missing Source! ");
    }
  if (values[0][DEST_SECTION] === "")  {
    result.push("Missing Destination! ");
  }
  if (values[0][TARGET_DESC] === "") {
    result.push("Missing Description! ");
  }
  if (values[0][TARGET_LOT] === "") {
    result.push("Missing Lot! ");
  }
  if (values[0][TARGET_QTY] === "")  {
    result.push("Missing Quantity! ");
  }
  
  if (result.length != 0){
    outputCell.setValue(result.toString());
    return false;
  }
  else
  return true;
}

/*
Ensures that the required data is present in a given row
*/
function validInputDataForPallet(sourceRowNum) {
  var sourceRow = inputSheet.getRange(sourceRowNum, 1, 1, lastInputCol);
  var outputCell = sourceRow.getCell(1, OUTPUT+1);
  var values = sourceRow.getValues();
  var result = [];

  if (values[0][SRC_SECTION] === "")  {
    // prompt whether user is okay with changing values regardless of source section
    var response = ui.alert('The source section is blank! Are you sure you want to move this material without considering current location?', ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES ){
      //continue
    }else if (response == ui.Button.NO){
      return false;
    }
    }
  if (values[0][DEST_SECTION] === "")  {
    result.push("Missing Destination! ");
  }
 
  if (result.length != 0){
    outputCell.setValue(result.toString());
    return false;
  }
  else
  return true;
}

/*
resets input sheet to blank
*/
function resetInputSheet(){
  var target = inputSheet.getRange(2, 1, inputSheet.getLastRow(), inputSheet.getLastColumn());
  target.clearContent();
  target.setBackground('white');
}

