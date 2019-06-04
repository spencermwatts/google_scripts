// ----- The following functions set up the plugin to be used in the sheet. 

// On install, open.
function onInstall(e) {
  onOpen(); 
}

// On open, show these menu items
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Rented.com Appraisal Tools')
  .addItem('Pull Appraisal Data','fetchAppraisalData')
  .addItem('Generate document', 'showSidebar')
  .addItem('Appraisal Troubleshooting Tips','displayDialog')
  .addToUi();
}

// How to show a sidebar
function showSidebar() {
  var template = HtmlService
  .createTemplateFromFile('sidemenu')
  
  var htmlOutput = template.evaluate()
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setTitle('Document Merge Tool')
  .setWidth(400);
  
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showSidebar(htmlOutput);
}

// Not sure if this is still relevant... 
function getListOptions() {
  // In production code, get an array of options by
  // reading a spreadsheet.
  // var options = ['Audi', 'Saab']; // SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var options = new Array()
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) {
    options.push(  sheets[i].getName()  )
  }
  
  return( options );
}

// Show the help dialog
function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
  .setWidth(400)
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showSidebar(html);
}

// Set up the help dialog.
function displayDialog() {
  var htmlApp = HtmlService
  .createHtmlOutputFromFile('errorSheet') 
  .setTitle("We can't gather comp data for this address :(")
  .setWidth(600)
  .setHeight(700);
  
  SpreadsheetApp.getActiveSpreadsheet().show(htmlApp);
}


function fetchAppraisalData(){ 

  var sheet = getInputSheet();
  
  var address = sheet.getRange('Address').getValue();
  var now = Utilities.formatDate(new Date(), "EST", "yyyy-MM-dd HH:mm");
  renameDoc(address + " - " + Session.getActiveUser().getEmail() + " " + "[" + now + "]");
  moveFile();
//  SpreadsheetApp.getActiveSpreadsheet().toast("up to date");
  
  
  try { 
    fetchAirDNAData();
    writeInternalCompsToSheet();
  } catch(error) {
    Logger.log(error);
    SpreadsheetApp.getActiveSpreadsheet().toast(error);
    displayDialog();
  }
}

function getInputSheet(){  
  var sheet = SpreadsheetApp.getActive().getSheetByName('Appraisal');
  return sheet;
}
