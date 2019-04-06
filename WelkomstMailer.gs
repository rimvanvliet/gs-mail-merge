var ledenadminMail = "Ledenadministratie ETT <ledenadministratie@ett-twello.nl>";
var keyMandateTemplate = '1UmNKez4gsa5T5i8IPLYcsNcAXRi4nzbose_uHN9xPfM';


function sendWelcomeEmails() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var wmSheet = ss.getSheetByName('welkomst mail');
  var emailHTML = DocumentApp.openById(wmSheet.getRange("berichtWelkomstMail").getValue()).getBody().getText();
  var emailTXT = "Dit is een HTML mail maar uw email programma begrijpt dat blijkbaar niet. Neem contact op met de afzender ...";
  var emailSubject = wmSheet.getRange("onderwerpWelkomstMail").getValue();
  var emailFrom = wmSheet.getRange("vanWelkomstMail").getValue();
  var emailTest = wmSheet.getRange("testWelkomstMail").getValue();
  var emailAttachments = wmSheet.getRange("bijlagenWelkomstMail").getValue().split(",");

  var attachments = new Array();
  for (var i = 0; i < emailAttachments.length; ++i) {
    attachments.push(DriveApp.getFileById(emailAttachments[i].trim()).getAs(MimeType.PDF));
  }
  
  if ((wmSheet.getRange("vanWelkomstMail").getValue() == "") || (wmSheet.getRange("onderwerpWelkomstMail").getValue() == "") || (wmSheet.getRange("berichtWelkomstMail").getValue() == "")) {
    Browser.msgBox("Verplichte velden", "Je moet alle verplichte velden invullen!", Browser.Buttons.OK);
    return;
  }

  wmSheet.getRange(2, 1, wmSheet.getMaxRows(), 1).clearContent();
  
  var ledenArray = haalLedenLijstOp('ledenlijst actueel');
  var leden = converteerArrayToObjects(ledenArray, 'userName'); // leden["<lid.gebruikers.naam>"] = <lid>
  var welkomstLeden = haalWelkomstLijstOp(ss); // array met de gebruikersnamen van de nieuwe leden
  
  if (emailTest != "") {
    numberOfMails = 1;
  } else {
    numberOfMails = welkomstLeden.length;
  }
  
  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < numberOfMails; ++i) {
    // Get a row object
    var rowData = leden[welkomstLeden[i]["gebruikersnaam"]];
    
    if (rowData == undefined) {
      wmSheet.getRange(i+2,1).setValue(">>>> NOK: " + welkomstLeden[i]["gebruikersnaam"].toUpperCase());
    } else {

      // Generate a personalized email.
      // Given a template string, replace markers (for instance ${"First Name"}) with
      // the corresponding value in a row object (for instance rowData.firstName).
      
      var file = personaliseAttachment(keyMandateTemplate, rowData);
      
      var advancedArgs = {name:emailFrom, htmlBody:fillInTemplateFromObject(emailHTML, rowData), attachments: [file.getAs(MimeType.PDF)].concat(attachments)}; 
      
      if (emailTest != "") {
        MailApp.sendEmail(rowData.displayName + " <" + emailTest + ">", emailSubject, emailTXT, advancedArgs);
        wmSheet.getRange(i+2,1).setValue("ok: " + rowData.displayName + " (test)");
        SpreadsheetApp.flush();
      }
      else {
        advancedArgs.cc = ledenadminMail;
        Logger.log(rowData);
        MailApp.sendEmail(rowData.displayName + " <" + rowData.email + ">", emailSubject, emailTXT, advancedArgs);
        wmSheet.getRange(i+2,1).setValue("ok: " + rowData.displayName);
        SpreadsheetApp.flush();
        
      }
      file.setTrashed(true);
    }
  } 
}

function haalWelkomstLijstOp(ss) {
  
  var dataSheet = ss.getSheetByName('welkomst mail');
  var dataRange = dataSheet.getRange(2, 2, dataSheet.getMaxRows(), 1);

  return getRowsData(dataSheet, dataRange);
}
