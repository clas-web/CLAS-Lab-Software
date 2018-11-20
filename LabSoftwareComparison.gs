//Lab software Excel sheets folder: https://drive.google.com/open?id=1QHnqC8YNWfXIUA2yOPNKOhSLszh61ZBU
//Google Sheet imports folder: https://drive.google.com/drive/u/1/folders/1IwJosiO60iiGvzXhkXH1l55QycLGIEmg

//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************************************************
//Add functions to spreadsheet menu
/*function onOpen() {
var spreadsheet = SpreadsheetApp.openById('1J4boWrZ0ACwpM7Yq-1vp6zVLp1vDAUH8-mJuvx05nvg');
var menuItems = [
{name: 'Update Master Sheet', functionName: 'updateMasterSheet'},
{name: 'Import All New Lab Software (slow)', functionName: 'importNewSoftware'}, 
];
spreadsheet.addMenu('Run Functions', menuItems);
}
//************************************************************************************************************************************************************************
*/    
//************************************************************************************************************************************
//print updated software to master sheet: https://docs.google.com/spreadsheets/d/1J4boWrZ0ACwpM7Yq-1vp6zVLp1vDAUH8-mJuvx05nvg/edit#gid=0
//@NotOnlyCurrentDoc
function updateMasterSheet(){
  var master = SpreadsheetApp.openById('1J4boWrZ0ACwpM7Yq-1vp6zVLp1vDAUH8-mJuvx05nvg');
  var newSoftware = "New software";
  var oldSoftware = "No longer listed software";
  //Get the current time and date
  var currentDate = new Date();
  
  //Format sheet
  master.getActiveSheet().clear();
  master.setFrozenRows(1);  
  master.getRange("A1").setValue("This sheet was last updated "+currentDate);
  master.getRange("A2").activate();
  
  //Get labs
  var spreadsheets = DriveApp.getFolderById('1IwJosiO60iiGvzXhkXH1l55QycLGIEmg').getFiles();
  var spreadsheet;
  var spreadsheetURL;
  var sheetName = "";
  var newSoftwareLR;
  var oldSoftwareLR;
  var ssRow = SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getLastRow() + 1;
  var ssCol = SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getLastColumn();
  var dateNote;
  while (spreadsheets.hasNext()){        
    spreadsheet = spreadsheets.next();           
    spreadsheetURL = spreadsheet.getUrl();
    dateNote = SpreadsheetApp.openById(spreadsheet.getId()).getSheetByName("Old").getRange("A1").getNote();
    Logger.log("dateNote "+dateNote);
    sheetName = spreadsheet.getName().replace(' Monthly Software Comparison','');      
    newSoftwareLR = SpreadsheetApp.openById(spreadsheet.getId()).getSheetByName("New software").getLastRow();
    Logger.log("newSoftwareLR" + newSoftwareLR);
    Logger.log("oldSoftwareLR" + oldSoftwareLR);
    newSoftware = SpreadsheetApp.openById(spreadsheet.getId()).getSheetByName("New software").getRange(2, 1, newSoftwareLR, 3).getValues();
    oldSoftwareLR = SpreadsheetApp.openById(spreadsheet.getId()).getSheetByName("No longer listed software").getLastRow();
    oldSoftware = SpreadsheetApp.openById(spreadsheet.getId()).getSheetByName("No longer listed software").getRange(2, 1, oldSoftwareLR, 3).getValues();
    ssRow = SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getLastRow() + 1;
    if(newSoftwareLR > 1){
      SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getRange(ssRow, 1, 1, 1).setValue('=HYPERLINK("'+spreadsheetURL+'","'+sheetName+' New Software")').setBackground("LightGreen");
      SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getRange(ssRow, 2, 1, 1).setValue(spreadsheet.getLastUpdated());
      ssRow = SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getLastRow() + 1;          
      SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getRange(ssRow, 1, newSoftware.length, 3).setValues(newSoftware);
      ssRow = SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getLastRow() + 1;
    }
    if(oldSoftwareLR > 1){
      SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getRange(ssRow, 1, 1, 1).setValue('=HYPERLINK("'+spreadsheetURL+'","'+sheetName+' No longer listed software")').setBackground("LightGray");
      SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getRange(ssRow, 2, 1, 1).setValue(dateNote);
      ssRow = SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getLastRow() + 1;
      SpreadsheetApp.openById(master.getId()).getSheetByName("Master").getRange(ssRow, 1, oldSoftware.length, 3).setValues(oldSoftware);    
    }
  }
  
  //Save and email
  print(master);
  
}
//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************
//save and email the report in Google Drive
function print(master){
  
  var sourceSheet = master;
  var outputSheet = sourceSheet.getSheetByName("Master");
  var parentFolder; //Folder to save PDF in
  var currentDate = new Date();
  
  //Checks if folder exists, if it doesn't, create it
  try {
    //Folder exists
    parentFolder = DriveApp.getFoldersByName('Lab Software Change Reports').next();   
    Logger.log('folder exists');
  }
  catch(e) {
    //Folder doesn't exist, create folder
    parentFolder = DriveApp.createFolder('Lab Software Change Reports');
    Logger.log('folder does not exist, creating folder');
  }
  
  //nane PDF
  var PDFName = "Lab Software Change Report-" + currentDate;
  SpreadsheetApp.getActiveSpreadsheet().toast('Creating PDF '+PDFName);
  
  // export url
  var PDFurl = 'https://docs.google.com/spreadsheets/d/'+sourceSheet.getId()+'/export?exportFormat=pdf&format=pdf' // export as pdf
  + '&size=letter'                           // paper size legal / letter / A4
  + '&portrait=true'                     // orientation, false for landscape
  + '&fitw=true'                        // fit to page width, false for actual size
  + '&sheetnames=true&printtitle=true' // hide optional headers and footers
  + '&pagenum=CENTER&gridlines=false' // hide page numbers and gridlines
  + '&fzr=false'                         // do not repeat row headers (frozen rows) on each page
  + '&top_margin=.75&bottom_margin=.75&left_margin=.25&right_margin=.25' //Narrow margins
  + '&gid='+outputSheet.getSheetId();    // the sheet's Id
  
  //authorize script
  var token = ScriptApp.getOAuthToken();
  
  // request export url
  var response = UrlFetchApp.fetch(PDFurl, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
  
  //name PDF blob
  var PDFBlob = response.getBlob().setName(PDFName+'.pdf');
  
  // delete pdf if it already exists
  var files = parentFolder.getFilesByName(PDFName);
  while (files.hasNext())
  {
    files.next().setTrashed(true);
  }
  
  // create pdf file from blob
  var createPDFFile = parentFolder.createFile(PDFBlob);  
  var folderURL = parentFolder.getUrl();
  emailPDF(PDFBlob, folderURL);
}

//************************************************************************************************************************************************************************
//Email report to user
function emailPDF(PDFBlob, folderURL){
  
  // Send the PDF of the spreadsheet to this email address
  var email = Session.getActiveUser().getEmail(); 
  var currentDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM YYYY");
  
  
  // Subject of email message
  var subject = "Lab Software Change Report-" + currentDate;
  
  // Email Body
  var body = "This has also been saved to your Google Drive at "+folderURL;
  
  // If allowed to send emails, send the email with the PDF attachment
  if (MailApp.getRemainingDailyQuota() > 0) 
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[PDFBlob]     
    });  
}

//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************
//just to edit all spreadsheets simultaneously, delete/create functions as necessary
function deleteAndCreateSheets(){
  
  var spreadsheets = DriveApp.getFolderById('1IwJosiO60iiGvzXhkXH1l55QycLGIEmg').getFiles();
  var spreadsheet;
  var sheetName = "";
  var formula1 = '=IFERROR({{INDIRECT("'+sheetName+'!$A$1:$C$1")};{FILTER({INDIRECT("Old!$A:$C")},ISERROR(MATCH(INDIRECT("Old!$A:$A"),INDIRECT("'+sheetName+'!$A:$A"),0)),len(INDIRECT("Old!$A:$A")))}},{INDIRECT("Old!$A$1:$C$1")})';
  var formula2 = '=IFERROR({{INDIRECT("Old!$A$1:$C$1")};{FILTER({INDIRECT("'+sheetName+'!$A:$C")},ISERROR(MATCH(INDIRECT("'+sheetName+'!$A:$A"),INDIRECT("Old!$A:$A"),0)),len(INDIRECT("'+sheetName+'!$A:$A")))}},{INDIRECT("'+sheetName+'!$A$1:$C$1")})';
  
  while (spreadsheets.hasNext()){
    
    spreadsheet = spreadsheets.next();
    try{ 
      sheetName = spreadsheet.getName().replace(' Monthly Software Comparison','');
      formula1 = '=IFERROR({{INDIRECT("'+sheetName+'!$A$1:$C$1")};{FILTER({INDIRECT("Old!$A:$C")},ISERROR(MATCH(INDIRECT("Old!$A:$A"),INDIRECT("'+sheetName+'!$A:$A"),0)),len(INDIRECT("Old!$A:$A")))}},{INDIRECT("Old!$A$1:$C$1")})';
      formula2 = '=IFERROR({{INDIRECT("Old!$A$1:$C$1")};{FILTER({INDIRECT("'+sheetName+'!$A:$C")},ISERROR(MATCH(INDIRECT("'+sheetName+'!$A:$A"),INDIRECT("Old!$A:$A"),0)),len(INDIRECT("'+sheetName+'!$A:$A")))}},{INDIRECT("'+sheetName+'!$A$1:$C$1")})';
      SpreadsheetApp.openById(spreadsheet.getId()).getSheetByName("No longer listed software").getRange("A1").setFormula(formula1);      
      SpreadsheetApp.openById(spreadsheet.getId()).getSheetByName("New software").getRange("A1").setFormula(formula2);      
    } catch (e){
      Logger.log(e.toString());
    }
  }
  
}

//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************

//@NotOnlyCurrentDoc
function importNewSoftware(){
  
  //Get a list of labs
  var labs = listLabs();
  
  //Get the current time and date
  var currentDate = new Date();
  
  //iterate for each lab
  for (var n = 0; n < labs.length; n++){
    
    //Get the original excel file with lab software for current lab
    var labExcelFile = DriveApp.getFolderById('1QHnqC8YNWfXIUA2yOPNKOhSLszh61ZBU').getFilesByName(labs[n]+'.xlsx').next();
    Logger.log("labExcelFile is "+labExcelFile.getName());
    
    //Get the Google Sheet with lab data for current lab
    if (DriveApp.getFolderById('1IwJosiO60iiGvzXhkXH1l55QycLGIEmg').getFilesByName(labs[n]+' Monthly Software Comparison').hasNext()){      
      var excelToSheet = DriveApp.getFolderById('1IwJosiO60iiGvzXhkXH1l55QycLGIEmg').getFilesByName(labs[n]+' Monthly Software Comparison').next(); 
    } else {
      var excelToSheet = SpreadsheetApp.create(labs[n]+' Monthly Software Comparison');
      //Add created Google spreadsheet for lab to folder
      DriveApp.getFolderById('1IwJosiO60iiGvzXhkXH1l55QycLGIEmg').addFile(DriveApp.getFileById(excelToSheet.getId()));        
    }    
    
    //Get primary sheet tab of current lab Google Spreadsheet
    var sheet = SpreadsheetApp.openById(excelToSheet.getId()).getSheetByName(labs[n]) || SpreadsheetApp.openById(excelToSheet.getId()).insertSheet(labs[n]);
    
    //convert to Google Sheet, back up sheet to archive folder
    var convertedFile = {
      title: labExcelFile.getName()+"_"+currentDate,
      parents: [{ id: "1q7vu0oubneesj7P49Fb3X3U1mh1-25YS" }]
  };
  convertedFile = Drive.Files.insert(convertedFile,labExcelFile, {
    convert:true
  });
  
  //import Google Sheet
  var SSSheets = SpreadsheetApp.openById(convertedFile.id);
  // Get full range of data
  var sheetRange = SSSheets.getDataRange();
  // get the data values in range
  var sheetData = sheetRange.getValues();
  
  //Back up sheet before clearing it
  backUpOldSheet(labs[n],excelToSheet.getId());
  
  //clear the sheet before import new range
  sheet.clear();
  Logger.log("The sheet has been cleared, importing new sheet...");
  Logger.log("sheetRange.getValues().toString() is " + sheetData.toString());
  Logger.log("sheetRange.getValues().length is " + sheetData.length);
  
  sheet.getRange(1, 1, SSSheets.getLastRow(), SSSheets.getLastColumn()).setValues(sheetData);     
  
  //Update A1 with note featuring the date of the import
  sheet.getRange("A1").setNote("Imported " + currentDate);
  
}

}
//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************
//Back up sheet before clearing it so both sheets can be contrasted and compared
function backUpOldSheet(lab,excelToSheet) {
  //declare both sheet variables
  var backupSheet = SpreadsheetApp.openById(excelToSheet).getSheetByName("Old") || SpreadsheetApp.openById(excelToSheet).insertSheet("Old");
  //SpreadsheetApp.openById(excelToSheet).setActiveSheet(SpreadsheetApp.openById(excelToSheet).getSheetByName("Sheet1"));
  //SpreadsheetApp.openById(excelToSheet).renameActiveSheet(lab);
  var newSheet = SpreadsheetApp.openById(excelToSheet).getSheetByName(lab) || SpreadsheetApp.openById(excelToSheet).renameActiveSheet(lab) || SpreadsheetApp.openById(excelToSheet).insertSheet(lab);  
  
  //remove contents of backup sheet
  backupSheet.clear();
  
  //grab data from the newest sheet before it is cleared so it becomes our backup
  Logger.log("newSheet.getDataRange().getValues().toString() is " + newSheet.getDataRange().getValues().toString());
  Logger.log("newSheet.getDataRange().getValues().length is " + newSheet.getDataRange().getValues().length);
  if (newSheet.getDataRange().getValues().toString() != ""){
    var newData = newSheet.getDataRange().getValues();
    var notes = newSheet.getDataRange().getNotes();
    
    
    //import data
    backupSheet.getRange(1, 1, newSheet.getLastRow(), newSheet.getLastColumn()).setValues(newData);
    backupSheet.getRange(1, 1, newSheet.getLastRow(), newSheet.getLastColumn()).setNotes(notes);
  } else {
    Logger.log("Sheet was blank, no backup performed");
  }
  
  //Remove blank Sheet1
  try{
    SpreadsheetApp.openById(excelToSheet).deleteSheet(SpreadsheetApp.openById(excelToSheet).getSheetByName("Sheet1"));
  } catch(e){
    Logger.log(e.toString());
  }
}
//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************
//************************************************************************************************************************************
//List all the labs in the lab software folder
//@NotOnlyCurrentDoc
function listLabs(){
  var labsArray = DriveApp.getFolderById("1QHnqC8YNWfXIUA2yOPNKOhSLszh61ZBU").getFiles();
  var names = [];
  var moddedNames = [];
  var iterator = 0;
  while (labsArray.hasNext()){
    //Return labs (with the Excel file format removed from the end of the name)
    lab = labsArray.next();
    names.push(lab.getName());
    moddedNames.push((lab.getName()).replace('.xlsx',''));
    Logger.log(lab.getName());
  }
  return moddedNames;
}