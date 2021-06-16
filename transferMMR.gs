/*
Goal of this function is to read the form responses sheet
and move the file to the appropriate folder in MMR Shared Drive

needs to trigger on form completion
*/


const MMRDriveID = "";
const FormSheetName = "Transfer Form Responses";
const MMRSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
const WIP_FOLDER = "1-WIP", QC_REVIEW_FOLDER = "2-QC_Approval", A_TEST_FOLDER = "3-A_Test",
      PRE_PROD_FOLDER = "4-PreProduction", PRODUCTION_FOLDER = "5-Production", 
      RECONCILIATION_FOLDER = "6-Reconciliation";
var formResponse = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FormSheetName);

/*
Trigger upon form submission. Moves the MMR to the chosen location and notifies necessary
parties. 
*/
function main(){
  // do nothing if the form response sheet is empty
  if (formResponseEmpty()){return;}
  
  // get the parent folder of current file
  var thisFile = DriveApp.getFileById(MMRSpreadsheet.getId());
  var parentFolderIterator = thisFile.getParents();
  var parentFolder = parentFolderIterator.next();

  // get the latest response from sheet
  var latestResponse = formResponse.getRange(formResponse.getLastRow(), 2)
                      .getValue();
  var responseComment = formResponse.getRange(formResponse.getLastRow(), 3)
                      .getValue();
  
  // if folder is in chosen folder, do nothing
  if (parentFolder.getName() === latestResponse){return;}

  // otherwise relocate and notify team
  moveToTargetFolder(thisFile, latestResponse);
  notifyTeam(thisFile.getName(), latestResponse, responseComment);
}

// checks whether the form response sheet has any data
function formResponseEmpty(){
  if (formResponse.getLastRow() > 1) {
    return false;
  }
  return true;
}

// relocates file into the target folder (parameter)
function moveToTargetFolder(thisFile, latestResponse){
  var mmrDrive = DriveApp.getFolderById(MMRDriveID);

  // iterate thru folders in mmrDrive
  var folders = mmrDrive.getFolders();
   while(folders.hasNext()){
     var folder = folders.next();
     if (folder.getName() === latestResponse){
       thisFile.moveTo(folder);
       break;
     }
   }
}

// sends an email describing what action has been taken
function notifyTeam(filename, newLocation, responseComment){
  var recipients;
  var subject;
  var body;
  var options = {
        noReply: true
      };
  var fileLink = MMRSpreadsheet.getUrl();
  var isComment = false;

  // set boolean true if a comment was added
  if (responseComment != ""){isComment = true;}

  // initial location is WIP, no notification
  if (newLocation === WIP_FOLDER){return;}
  // QC REVIEW - notify QC Team (Amy)
  else if (newLocation === QC_REVIEW_FOLDER){
    recipients = ""
    subject = "";

    if (isComment){
      body = "";
    }else {
      body = "";
    }
  }
  // A TESTING - Notify whoever needs to pull and make lab form
  // (Mike, Diana, Alicia, Ethan)
  else if (newLocation === A_TEST_FOLDER){
    recipients = "";
    subject = "";
    
    if (isComment){
      body = "";
    }else {
      body = "";
    }

  // PRE PRODUCTION - Notify whoever needs to prepare materials or machines
  else if (newLocation === PRE_PROD_FOLDER){
    recipients = "";
    subject = "";
    
    if (isComment){
      body = "";
    }else {
      body = "";
    }
  }
  // PRODUCTION - Notify production and special services
  else if (newLocation === PRODUCTION_FOLDER){
    recipients = "";
    subject = "";
    
    if (isComment){
      body = ";
    }else {
      body = "";
    }
  }
  // RECONCILIATION - Notify inventory, production, and documentation people
  else if (newLocation === RECONCILIATION_FOLDER){
    recipients = "";
    subject = ";
    body = "";
  }
  // do nothing if any unexpected input
  else{return;}
  
  // recipients = "ecloin@sawgrassnutralabs.com";
  MailApp.sendEmail(recipients, subject, body, options);
}

// check if a form submit trigger exists and create it if not
function triggerCheck(){
  if (ScriptApp.getProjectTriggers().length === 0){
    ScriptApp.newTrigger("main")
              .forSpreadsheet(MMRSpreadsheet)
              .onFormSubmit()
              .create();
  MMRSpreadsheet.toast("Transfer form trigger is initialized and ready for use!");
  }
  else {return;}
}

// check if user has authorized the script to function
function authorizationCheck(){
  if (ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationUrl() != null){
    SpreadsheetApp.getUi().alert("You need to authorize the script! Paste this link into a new tab!\n\n" + ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL).getAuthorizationUrl());
  }
}

// trigger to build
function onOpen(e){
  triggerCheck();
  authorizationCheck();
}
