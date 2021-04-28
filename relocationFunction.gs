/*
Goal of the function:

Allow the user to update the location of materials in a few ways:
  1. Move ALL of one ingredient from Section A --> Section B
  2. Move ALL of one pallet from Section A --> Section B
  3. Move SOME of one ingredient from Section A --> Section B

@input  1. Data specifying an ingredient (sku, lot, desired section)
        2. Data specifying a pallet (section)
        3. Data specifying an ingredient (sku, lot, qty, desired section)

@ouput  1. Updated ledger entry to the desired section
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
var lastInputRow = inputSheet.getLastRow();

function test(){
  determineRelocationType();
}

function relocateFullQty(sourceRowNum) {

  inputSheet.getRange(sourceRowNum, OUTPUT+1).setValue("DETECTED FULL");
  // scan Relocator and store description, lot, source, and destination section
  
  // search Ledger for matching line item

  // update Section value

  // mark as complete and show edited row number for easy confirmation
  
}

function relocatePartialQty(sourceRowNum){
  inputSheet.getRange(sourceRowNum, OUTPUT+1).setValue("DETECTED PARTIAL");
  // scan Relocator and store description, lot, source, destination, and qty

  // search Ledger for matching line item

  // duplicate matching line

  // update duplicate to the qty and section from Relocator

  // update matching line by subtracting qty from Relocator

  // mark as complete and show edited row number for easy confirmation


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
  
  // find all matching cells
  var finder = ledgerSheet.createTextFinder(sourceSection.getValue()).matchEntireCell(true);
  var foundCells = finder.findAll();
  Logger.log(foundCells);
  var foundCount = foundCells.length; 
  if (foundCount === 0){
    editedMaterials.push("NONE FOUND");
    Logger.log(editedMaterials);
    outputCell.setValue(editedMaterials);
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
  // // search Ledger for all line items with source section
  // for (var row = lastLedgerRow; row > 0; row--){
  //   // check section value
  //   currentSectionCell = ledgerSheet.getRange(row, SECTION_I + 1);
  //   currentSection = currentSectionCell.getValue();

  //   if (currentSection === sourceSection){
  //     // update section to destination
  //     currentSectionCell.setValue(destinationSection);

  //     // add row number and desc to 
  //     var rowNum = currentSectionCell.getRow();
  //     var description = ledgerSheet.getRange(row, DESC_I+1).getValue();
  //     var edited = [rowNum, description];
  //     editedMaterials.push(edited);
  //   }

  // }

  // update to destination section

  // return description and rows updated for easy confirmation
  
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



