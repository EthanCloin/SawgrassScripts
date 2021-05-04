// Constants to represent what columns are which index values (offset by +1 to use as column number)
const DATE_I = 0, SKU_I = 1, DESC_I = 2, COA_I = 3, EXP_I = 4, QTY_I = 5, PACK_I = 6, SECTION_I = 7, LOT_I = 8, 
      VENDOR_I = 9, PRICE_I = 10, POJOB_I = 11, COMMENT_I = 12, COMPONENT_I = 13, CHECK_I = 14;

// Constants to represent columns in LedgeringAssistant (no need for offset)
const INPUT_LOT = 1, INPUT_DESC = 2, INPUT_QTY = 3, INPUT_JOB = 4, INPUT_COMMENT = 5;

// Constants to represent columns in Relocator sheet (offset by +1 to use as column number)
const SRC_SECTION = 0, DEST_SECTION = 1, TARGET_DESC = 2, TARGET_LOT = 3, TARGET_QTY = 4, 
      RELOC_TYPE = 5, OUTPUT = 6;

// Constants to represent columns in PO LOG sheet (offset by +1 to use as column number)
const LOG_PO_I = 6, LOG_VENDOR_I = 7, LOG_DESC_I = 9, LOG_QTY_I = 10, LOG_PACK_I = 11, LOG_PRICE_I = 12,
      LOG_RECEIVED_I = 16, LOG_LOT_I = 17, LOG_EXP_I = 18, LOG_COMMENT_I = 19, LOG_TO_INPUT_I = 20;
// spreadsheet
var ss = SpreadsheetApp.getActiveSpreadsheet();
var ui = SpreadsheetApp.getUi();

// sheet
var ledgerSheet = ss.getSheetByName("Ledger");
var assistantSheet = ss.getSheetByName("LedgeringAssistant");
var POLogSheet = ss.getSheetByName("PO LOG");
var inventoryInputSheet = ss.getSheetByName("Inventory Input");
var relocatorSheet = ss.getSheetByName("Relocator");


// row / col
var lastLedgerRow = ledgerSheet.getLastRow(); // final ledger entry
var lastLedgerCol = ledgerSheet.getLastColumn(); // final ledger column
var lastAssistantRow = assistantSheet.getLastRow(); // final input entry
var lastAssistantCol = assistantSheet.getLastColumn(); // final input column
var POLogLastCol = POLogSheet.getLastColumn();
var inventoryInputLastRow = inventoryInputSheet.getLastRow();
var inventoryInputLastCol = inventoryInputSheet.getLastColumn();
var lastRelocatorRow = relocatorSheet.getLastRow();
var lastRelocatorCol = relocatorSheet.getLastColumn();

/*
creates menu
*/
function onOpen() {
  ui.createMenu("InventoryAssistant")
    .addSubMenu(ui.createMenu("LedgeringAssistant")
    .addItem("Allocate Materials by SKU", "ledgerItemsWithSKU")
    .addItem("Reconcile Materials by Lot", "ledgerItemsWithLot")
    .addItem("Undo Allocation", "removeAllocation"))
    .addSeparator()
    .addItem("Relocator", "determineRelocationType")
    .addToUi(); 
}