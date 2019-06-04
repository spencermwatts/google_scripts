// ----- The following are some helper methods 

function renameDoc(name) {
//  Logger.log("Renaming Doc to");
//  Logger.log(name);
  var doc = SpreadsheetApp.getActiveSpreadsheet();
//  var name = doc.getName();
  doc.rename(name);
}

function moveFile(){
  var file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  DriveApp.getFolderById("XXXXX").addFile(file);
}

function getCurrentMonth(){
  var date = new Date();
  return date.getMonth() + 1;
}
