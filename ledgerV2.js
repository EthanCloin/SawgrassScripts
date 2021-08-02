// Object with column indices in Ledger.getValues
const ledgerMap =  { 
                      date:0,sku:1,description:2,coa:3,expiration:4,
                      quantity:5,packUnit:6, section:7,lot:8,vendor:9,
                      price:10,job:11,comment:12,componentType:13, allocated:14,
                    }
// Object with column indices in assistant.getValues
const assistantMap = {
                        lot:0, description:1, quantity:2, job:3, comment:4, status:5
                      }
/*
  Works as intended, 10x faster than old ledgerFunction
*/
function reconcileByLot() {
  // if designated box is checked, change all quantities to negative
  var namedRanges = assistantSheet.getNamedRanges()
  var negativeToggle = namedRanges.filter(function(range){
    if (range.getName() === "NegativeToggle"){
      return range
    }})
  var negativeStatus = negativeToggle[0].getRange().getValue() 


  // collect info from LedgeringAssistant
  let assistantRange = assistantSheet.getRange(2, 1, getFirstEmptyRowInB(assistantSheet.getSheetName())- 2, assistantMap.status+1)
  let assistantData = assistantRange.getValues()
  Logger.log("input: " + assistantData)

  // cast to negative if needed
  assistantData = assistantData.map(function(entry){
    if (negativeStatus){
      entry[assistantMap.quantity] = -1 * Math.abs(Number(entry[assistantMap.quantity]))
    }
    return entry
  })
  // collect info from Ledger
  var ledgerData = ledgerSheet.getDataRange().getValues()
  
  // for each entry in assistant, find a matching lot + description in ledger
  var newEntries = assistantData.map(function(entry){
    // find index of matching lot using loop
    let errors = []
    while (true){
      // find row in ledger with matching lot

      var matchLotIndex = ledgerData.findIndex(el => el[ledgerMap.lot] == entry[assistantMap.lot])
      
      // none found, update status info with errors and break from loop
      if (matchLotIndex === -1 || entry[assistantMap.lot] == ""){
        Logger.log("none found for " + entry[assistantMap.description] + " of lot " + entry[assistantMap.lot])
        entry[assistantMap.status] = "NOT FOUND! Consider: " + errors.toString() 
        break
      }
      // check whether also matching description
      if (ledgerData[matchLotIndex][ledgerMap.description] === entry[assistantMap.description]){
        // form a new entry and return it to the new array
        var newEntry = ledgerData[matchLotIndex]
        Logger.log("copy: " + newEntry.toString())
        newEntry[ledgerMap.date] = new Date()
        newEntry[ledgerMap.quantity] = entry[assistantMap.quantity]
        newEntry[ledgerMap.job] = entry[assistantMap.job]
        newEntry[ledgerMap.comment] = entry[assistantMap.comment]
        newEntry[ledgerMap.sku] = ledgerSheet.getRange("B2").getFormulaR1C1() // could make this into a constant for speed
        Logger.log("newEntry: " + newEntry)
        entry[assistantMap.status] = "LEDGERED"
        return newEntry

      }else{
        // CURRENTLY DOESN"T WORK AS INTENDED
        // JUST MEANS WE HAVE DUPLICATED ERROR MESSAGES
        // remove that entry, so it is not reconsidered in next interation and push info to error array
        if (errors.includes(ledgerData[matchLotIndex][ledgerMap.description].toString())){
              // remove from ledger data
              ledgerData[matchLotIndex][ledgerMap.lot] = ""
              Logger.log("already found")
              
        }else {
        // add to errors if doesn't exist
        errors.push([
            ledgerData[matchLotIndex][ledgerMap.description], 
            " - ", ledgerData[matchLotIndex][ledgerMap.lot], "\n"])
        ledgerData[matchLotIndex][ledgerMap.lot] = ""
        Logger.log("partial found for " + entry[assistantMap.description])
          }
        }
    }
  }) 

  newEntries = newEntries.filter(function(entry) {
    if (entry != undefined) {return entry}
  })
  // update Ledger
  let firstAvailableRow = getFirstEmptyRowInC(ledgerSheet.getName())  
  ledgerSheet.insertRowsAfter(firstAvailableRow-1, newEntries.length)
  ledgerSheet.getRange(firstAvailableRow, 1, newEntries.length, ledgerSheet.getLastColumn())
              .setValues(newEntries)
  // update Assistant 
  assistantRange.setValues(assistantData)
  Logger.log("updated assistant")
}

// CURRENT BUG DOESNT WORK WITH MULTIPLES OF SAME DESCRIPTION
function allocateBySKU() {
  // if designated box is checked, change all quantities to negative
  var namedRanges = assistantSheet.getNamedRanges()
  var negativeToggle = namedRanges.filter(function(range){
    if (range.getName() === "NegativeToggle"){
      return range
    }})
  var negativeStatus = negativeToggle[0].getRange().getValue() 


  // collect info from LedgeringAssistant
  let assistantRange = assistantSheet.getRange(2, 1, getFirstEmptyRowInB(assistantSheet.getSheetName())- 2, assistantMap.status+1)
  let assistantData = assistantRange.getValues()

  // cast to negative if needed
  assistantData = assistantData.map(function(entry){
    if (negativeStatus){
      entry[assistantMap.quantity] = -1 * Math.abs(Number(entry[assistantMap.quantity]))
    }
    return entry
  })
  Logger.log("input: " + assistantData)
  // collect info from Ledger
  var ledgerData = ledgerSheet.getDataRange().getValues()
  
  // for each entry in assistant, find a matching lot + description in ledger
  let newEntries = assistantData.map(function(entry){
    // find index of matching lot using loop
    while (true){
      // find row in ledger with matching description
      var matchDescIndex = ledgerData.findIndex(el => el[ledgerMap.description] == entry[assistantMap.description])
      
      // none found, ask for permission to create a new entry
      if (matchDescIndex === -1){
        Logger.log("none found for " + entry[assistantMap.description])
        let response = ui.alert("There is no item in Ledger with this description. Would you like to allocate anyway?"
                                , ui.ButtonSet.YES_NO)
        if (response === ui.Button.YES){
          // form a new entry and return it to the new array
          var newEntry = ledgerData[0]
          newEntry.forEach(function(cell){
            cell = ""
          })
          
          Logger.log("copy: " + newEntry.toString())
          newEntry[ledgerMap.date]          = new Date()
          newEntry[ledgerMap.description]   = entry[assistantMap.description]
          newEntry[ledgerMap.quantity]      = entry[assistantMap.quantity]
          newEntry[ledgerMap.job]           = entry[assistantMap.job]
          newEntry[ledgerMap.comment]       = entry[assistantMap.comment]
          newEntry[ledgerMap.sku]           = ledgerSheet.getRange("B2").getFormulaR1C1() // could make this into a constant for speed
          newEntry[ledgerMap.allocated]     = true
          entry[assistantMap.status]        = "ALLOCATED"
          return newEntry
          }
          else {
            entry[assistantMap.status] = "NOT FOUND!"
          }
      }
      else {
      // form a new entry and return it to the new array
      let newEntry = ledgerData[matchDescIndex]
      newEntry[ledgerMap.date]      = new Date()
      newEntry[ledgerMap.lot]       = ""
      newEntry[ledgerMap.price]     = ""
      Logger.log("Entry: " + entry.toString())
      newEntry[ledgerMap.quantity]  = entry[assistantMap.quantity]
      Logger.log("Entry2: " + entry.toString())
      newEntry[ledgerMap.job]       = entry[assistantMap.job]
      newEntry[ledgerMap.comment]   = entry[assistantMap.comment]
      newEntry[ledgerMap.sku]       = ledgerSheet.getRange("B2").getFormulaR1C1() // could make this into a constant for speed
      newEntry[ledgerMap.allocated] = true
      entry[assistantMap.status]    = "ALLOCATED"
      Logger.log("newEntry: " + newEntry)
      return newEntry
      }
    }
  }) 
  newEntries = newEntries.filter(function(entry) {
    if (entry != undefined) {return entry}
  })
  
  // update Ledger
  let firstAvailableRow = getFirstEmptyRowInC(ledgerSheet.getName())  
  ledgerSheet.insertRowsAfter(firstAvailableRow-1, newEntries.length)
  ledgerSheet.getRange(firstAvailableRow, 1, newEntries.length, ledgerSheet.getLastColumn())
              .setValues(newEntries)
  // update Assistant 
  assistantRange.setValues(assistantData)
  Logger.log("updated assistant") 
}

function getFirstEmptyRowInC(sheetName) {
  var sheet = ss.getSheetByName(sheetName)
  var column = sheet.getRange('C:C')
  var values = column.getValues(); // get all data in one call
  // Logger.log(values)
  var ct = 0
  while ( values[ct] && values[ct][0] != "" ) {
    ct++
  }
  // Logger.log(ct+1)
  return (ct+1)
}

function getFirstEmptyRowInB(sheetName) {
  var sheet = ss.getSheetByName(sheetName)
  var column = sheet.getRange('B:B');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}
