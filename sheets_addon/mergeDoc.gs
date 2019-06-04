// template id, template name, input sheet id, input sheet workbook name
//  - preso?
function createMergedDoc(offer_presentation, pff, owner_agreement, sig_contract, sig_presentation_owner, sig_presentation_manager) {
  
  
  Logger.log(sig_presentation_manager);
  Logger.log(sig_presentation_owner);
  // If neither the standard document boxes are checked and the custom create buton was clicked, create a copy of the template and merge
  
  // If nothing selected, tell user to select something!
  if (offer_presentation == false && pff == false && owner_agreement == false && sig_contract == false && sig_presentation_owner == false && sig_presentation_manager == false) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Select PFF, Offer Presentation, or make a custom merge!');
  } else {
    // If any of the standard document boxes are checked.
    //  if (offer_presentation != false || pff != false || owner_agreement != false) {
    var dataSpreadsheetId = SpreadsheetApp.getActive().getSheetByName('Generate');
    var range = 'Generate!1:4';  
    var values = SpreadsheetApp.getActive().getRange(range).getValues();
    var address = values[3][0];
    var keys = values[0];
    var value = values[3];
  }
  
  if (pff) {    
    Logger.log("Generating PFF");
    var pff_template = DriveApp.getFileById('XXX'); // This is the PFF template ID
    var pff_driveResponse = pff_template.makeCopy(address + ' - Property Funding Form');
    Logger.log("Here we go!");
    generateGoogleDoc(keys, value, pff_driveResponse);    
  }
  
  if (offer_presentation) {
    Logger.log("Generating offer presentation");
    var offer_presentation_template = DriveApp.getFileById('XXX'); // This is the Offer Presentation template ID
    var offer_presentation_driveResponse = offer_presentation_template.makeCopy(address + ' - PM Offer Presentation');
    Logger.log(offer_presentation_driveResponse);
    generateSlides(keys, value, offer_presentation_driveResponse);
  }
  
  if (owner_agreement) {
    Logger.log("Generating offer presentation");
    var offer_presentation_template = DriveApp.getFileById('XXX'); // This is the Owner Offer Presentation template ID
    var offer_presentation_driveResponse = offer_presentation_template.makeCopy(address + ' - Owner Offer Presentation');
    Logger.log(offer_presentation_driveResponse);
    generateSlides(keys, value, offer_presentation_driveResponse);
  }
  
  if (sig_contract) {    
    Logger.log("Generating Secure Income Contract");
    var sig_contract_template = DriveApp.getFileById('XXX'); // This is the SIG Contract template ID
    var sig_contract_driveResponse = sig_contract_template.makeCopy(address + ' - Secure Income Guarantee Contract');
    Logger.log("Here we go!");
    generateGoogleDoc(keys, value, sig_contract_driveResponse);    
  }
  
  if (sig_presentation_owner) {
    Logger.log("Generating Secure Income presentation");
    var sig_presentation_template = DriveApp.getFileById('XXX-6qKgyQw'); // This is the SIG Owner Presentation template ID
    var sig_presentation_driveResponse = sig_presentation_template.makeCopy(address + ' - Secure Income Guarantee Owner Presentation');
    Logger.log(offer_presentation_driveResponse);
    generateSlides(keys, value, sig_presentation_driveResponse);
    
  }
  if (sig_presentation_manager) {
    Logger.log("Generating Secure Income presentation");
    var sig_presentation_template = DriveApp.getFileById('XXX'); // This is the SIG PM Presentation template ID
    var sig_presentation_driveResponse = sig_presentation_template.makeCopy(address + ' - Secure Income Guarantee Manager Presentation');
    Logger.log(offer_presentation_driveResponse);
    generateSlides(keys, value, sig_presentation_driveResponse);
    
  }
  
  
}

function generateSlides(keys, value, driveResponse) {
  Logger.log("Generating your google slides");
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Creating your Slides document!');
  var requests = [];     
  
  for (var i = 0; i < keys.length; ++i) {  
    var _Text = String(keys[i]) || "NO DATA";
    var _replaceText = String(value[i]) || "NO DATA";
    
    requests.push({
      replaceAllText: {
        containsText: {
          text: _Text,
          matchCase: true
        },
        replaceText: _replaceText
      }
    });
  }
  
  Slides.Presentations.batchUpdate({
    requests: requests
  }, driveResponse.getId());
  SpreadsheetApp.getActiveSpreadsheet().toast('Done. Check the "Recents" folder in your Google Drive for your document.', 'Complete!', -1);  
}

function generateGoogleDoc(keys, value, driveResponse) {
  
  Logger.log("Generating your google doc");
  
  //  Logger.log("keys are %s", keys);
  
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Creating your Google document!');  
  var body = DocumentApp.openById(driveResponse.getId()).getBody();
  for (var i = 0; i < keys.length; ++i) {  
    
    var _Text = String(keys[i]);
    var _replaceText = String(value[i]) || ' ';
    //    Logger.log("Text is %s", _Text);
    //    Logger.log("Value is %s", _replaceText);
    replacement = body.replaceText(_Text, _replaceText);
  }
  SpreadsheetApp.getActiveSpreadsheet().toast('Done. Check the "Recents" folder in your Google Drive for your document.');  
  
  emitJSON();
}
