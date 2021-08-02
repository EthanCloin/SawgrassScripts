/*
Function to move data from requisition log to PO Log
*/
function requisitionToPOLog() {
// indices for array version of Requisition Log Columns
  const reqMap = {
    sku:0, snlDescription:1, component:3, dueDate:5, vendor:6, // desc is merged 1 and 2
    quantity:7, pack:8, price:9, job:10, purchase:11, comment:12, status:13
  }

// indices for array version of PO Log Columns
  const poMap = {
    date:0, requestor:1, poApproved:2, poConfirmed:3, poPaid:4, job:5, poNumber:6, 
    vendor:7, sku:8, snlDescription:9, component:10, vendorDescription:11, quantity:12,
    pack:13, price:14, dueDate:15, tracking:16, ledgerAdded:17, received:18, lot:19,
    expiration:20, note:21
  }  

// grab the requisition entries that are marked for purchasing
  let requisitionSheet = ss.getSheetByName(REQ_SHEET_NAME)
  let lastRow = getFirstEmptyRowInB(REQ_SHEET_NAME) - 1
  let reqEntries = requisitionSheet.getRange(2, 1, lastRow, requisitionSheet.getLastColumn())
                    .getValues()
  
  // filter to only entries marked as purchased
  let purchaseTagged = reqEntries.filter(function(entry){
    if (entry[reqMap.purchase] === true){
      return entry
    }
  })

  // build entries for PO Log based on Req Log data
  let newLogEntries = purchaseTagged.map(function(entry){
    let newEntry = new Array(poMap.note+1)
    newEntry[poMap.date]           = new Date()
    newEntry[poMap.sku]            = POLogSheet.getRange("I2").getFormulaR1C1()
    newEntry[poMap.snlDescription] = entry[reqMap.snlDescription]
    newEntry[poMap.component]      = entry[reqMap.component]
    newEntry[poMap.dueDate]        = entry[reqMap.dueDate]
    newEntry[poMap.vendor]         = entry[reqMap.vendor]
    newEntry[poMap.quantity]       = entry[reqMap.quantity]
    newEntry[poMap.pack]           = entry[reqMap.pack]
    newEntry[poMap.price]          = entry[reqMap.price]
    newEntry[poMap.job]            = entry[reqMap.job]
    newEntry[poMap.received]       = false
    newEntry[poMap.ledgerAdded]    = false
    return newEntry
    })

  // add entries
  let firstAvailableRow = getFirstEmptyRowInB(POLogSheet.getName())
  POLogSheet.insertRowsAfter(firstAvailableRow-1, newLogEntries.length)
  POLogSheet.getRange(firstAvailableRow, 1, newLogEntries.length, POLogSheet.getLastColumn())
            .setValues(newLogEntries)
  
  // Update the status column for those copied to PO Log
  let outputRange = requisitionSheet.getRange(1, reqMap.purchase+1, lastRow, 3)
  let outputValues = outputRange.getValues()
  let outputUpdate = outputValues.map(function(entry){
    if (entry[0] === true){
      entry[0] = false // purchase
      entry[2] = "ADDED TO PO LOG" // status
    }
    return entry
  })
  outputRange.setValues(outputUpdate)
  
  ui.alert( "Copied entries marked for purchase to PO LOG! Be sure to add your initials." )


}

// 
function getFirstEmptyRowInB(sheetName) {
  var sheet = ss.getSheetByName(sheetName)
  var column = sheet.getRange('B:B')
  var values = column.getValues(); // get all data in one call
  // Logger.log(values)
  var ct = 0
  while ( values[ct] && values[ct][0] != "" ) {
    ct++
  }
  // Logger.log(ct+1)
  return (ct+1)
}
