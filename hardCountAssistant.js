/*
 *
 * Goal of this function is to locate all Lot Numbers associated with a given description
 * 
 * CURRENT ISSUE: not successfully grabbing associated lots and descriptions from the primary function
 *                output seems to work when run from testing function ???
*/

function readHardCount(){
  const hcMap       = {
                      component:0, description:1, quantity:2, pack:3, section:4,
                      lot:5, comment:6, status:7, tips:8
                    }

  const ledgerMap   = { 
                      date:0,sku:1,description:2,coa:3,expiration:4,
                      quantity:5,packUnit:6, section:7,lot:8,vendor:9,
                      price:10,job:11,comment:12,component:13, allocated:14,
                    }

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const hardCountSheet = ss.getSheetByName("Script Test")
  let countDataRange = hardCountSheet.getRange(2, 1, getFirstEmptyRowByA(hardCountSheet.getName()), hcMap.tips+1)
  let countData = countDataRange.getValues()
  const LedgerData = ss.getSheetByName("Copy of Ledger").getDataRange().getValues()
  
 countData.forEach(function(entry) {
    let workingData = LedgerData.map((entry) => entry)

    /*
    LOT CHECK
    */
    while (true){
      let matchingLotIndex = workingData.findIndex(el => el[ledgerMap.lot] == entry[hcMap.lot])
      // lot doesn't exist
      if (matchingLotIndex === -1 || entry[hcMap.lot] == "na" || entry[hcMap.lot] == "") {
        entry[hcMap.status] = "Lot/Description not found!"
        break
      }

      // check lot row for matching description
      if (workingData[matchingLotIndex][ledgerMap.description] == entry[hcMap.description]){
        entry[hcMap.status] = "Lot/Description confirmed!"
        break
      }
      // blank out this lot to continue search in next loop iteration
      workingData.splice(matchingLotIndex, 1)
    }
    // reset the working data since some may have been removed
    workingData = LedgerData.map((entry) => entry)

    /*
    DESCRIPTION CHECK
    */
    while (true){
      // check whether already confirmed
      if (entry[hcMap.status] == "Lot/Description confirmed!"){break}
      let matchingDescIndex = workingData.findIndex(el => el[ledgerMap.description] == entry[hcMap.description])
      // lot doesn't exist
      if (matchingDescIndex === -1) {
        entry[hcMap.status] = "Lot/Description not found!"
        break
      }

      // check description row for matching lot
      if (workingData[matchingDescIndex][ledgerMap.lot] == entry[hcMap.lot]){
        entry[hcMap.status] = "Lot/Description confirmed!"
        break
      }
      // blank out this desc to continue search in next loop iteration
      workingData.splice(matchingDescIndex, 1)        
    }

    // give tips if needed
    if (entry[hcMap.status] == "Lot/Description not found!"){
      let linkedLots = findLinkedLots(ledgerMap, LedgerData, entry[hcMap.description])
      Logger.log(linkedLots)
      let linkedDesc = findLinkedDescriptions(ledgerMap, LedgerData, entry[hcMap.lot])
      Logger.log(linkedDesc)
      entry[hcMap.tips] = "Linked Lots: " + linkedLots.toString() + "\nLinked Descriptions: " + linkedDesc.toString()
    }
  })
  countDataRange.setValues(countData)

}

/*
 * Returns an array with all associated Lots for a given description
*/
function findLinkedLots(ledgerMap, LedgerData, thisDescription) {
  Logger.log("received: " + thisDescription.toString())
  if (thisDescription == "") {
    return [""]
  }
  let linkedLots = LedgerData.map(function(entry){
    //Logger.log(entry[ledgerMap.lot])
    if (entry[ledgerMap.description] == thisDescription){
      Logger.log(entry[ledgerMap.lot])
      return entry[ledgerMap.lot]
    }
    else return ""
  })
  Logger.log(linkedLots)
  return Array.from(new Set(linkedLots))
}

/*
 * Returns an array with all associated Descriptions for a given lot
*/
function findLinkedDescriptions(ledgerMap, LedgerData, thisLot) {
  Logger.log("received: " + thisLot.toString())
  if (thisLot == "" || thisLot == "na"){
    return [""]
  }
  
  let linkedDesc = LedgerData.map(function(entry){
    if (entry[ledgerMap.lot] == thisLot){
      return entry[ledgerMap.description]
    }
    else return ""
    
  })
  Logger.log(linkedDesc)
  return Array.from(new Set(linkedDesc))
}

function testing() {
  const hcMap       = {
                      component:0, description:1, quantity:2, pack:3, section:4,
                      lot:5, comment:6, status:7, tips:8
                    }

  const ledgerMap   = { 
                      date:0,sku:1,description:2,coa:3,expiration:4,
                      quantity:5,packUnit:6, section:7,lot:8,vendor:9,
                      price:10,job:11,comment:12,component:13, allocated:14,
                    }
  const LedgerData = ss.getSheetByName("Copy of Ledger").getDataRange().getValues()

  let linkedLots = findLinkedLots(ledgerMap, LedgerData, "Capsule - Gel - Clear - Size 00E")
                    
  // const linkedDesc = findLinkedDescriptions(hcMap, ledgerMap, LedgerData, "APC8511")

  Logger.log(linkedLots.toString())
  // Logger.log(linkedDesc.toString())
}

function getFirstEmptyRowByA(sheetName) {
  var sheet = ss.getSheetByName(sheetName)
  var column = sheet.getRange('A:A');
  var values = column.getValues(); // get all data in one call
  var ct = 0;
  while ( values[ct] && values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);
}

