function sendIncassoEmails(incSheet,debtors) {
  
  /*var ss = SpreadsheetApp.getActiveSpreadsheet();*/
  
  var incassoDatum = Utilities.formatDate(incSheet.getRange("datumIncasso").getValue(), "CET", "yyyy-MM-dd");
  var ddId = incSheet.getRange("kenmerkIncasso").getValue();
  var emailHtmlINC = DocumentApp.openById(incSheet.getRange("berichtINCIncasso").getValue()).getBody().getText();
  var emailTXT = "Dit is een HTML mail maar uw email programma begrijpt dat blijkbaar niet. Neem contact op met de afzender ...";
  var emailSubject = incSheet.getRange("onderwerpIncasso").getValue();
  var emailFrom = incSheet.getRange("vanIncasso").getValue();
  var emailTest = incSheet.getRange("testIncasso").getValue();
  
  var IncassoTestSent = false;
  var FactuurTestSent = false;
  
  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var i = 0; i < debtors.length; ++i) {
    // Generate a personalized email.
    // Given a template string, replace markers (for instance ${{First Name}}) with
    // the corresponding value in a row object (for instance rowData.firstName).
    
    //switch (debtors[i]["betaling"]) {
    //  case "INCASSO":
    //  case "SEC-INC":
        var emailHtml = emailHtmlINC;
    //    break;
    //  default:
    //    Logger.log("Onbekend betalingssoort:"+debtors[i]["betaling"]);
    //}
    
    var advancedArgs = {name:emailFrom, htmlBody:fillInTemplateFromObject(emailHtml, debtors[i])}; 
    
    if (emailTest != "" ) {
      if(debtors[i]["betaling"] == 'INCASSO' && ! IncassoTestSent) {
        Logger.log(debtors[i]['displayName'])
        MailApp.sendEmail(debtors[i].displayName + " <" + emailTest + ">", emailSubject, emailTXT, advancedArgs);
        incSheet.getRange(i+2,1).setValue("ok: " + debtors[i].displayName + " (test)");
        SpreadsheetApp.flush();
        var IncassoTestSent = true;
      } 
    }
    else {
      // advancedArgs.cc = ledenadminMail;
      MailApp.sendEmail(debtors[i].displayName + " <" + debtors[i].email + ">", emailSubject, emailTXT, advancedArgs);
      incSheet.getRange(i+2,1).setValue("ok: " + debtors[i].displayName);
      SpreadsheetApp.flush();
      
    }
  } 
}