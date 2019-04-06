function testBijlagen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gmmSheet = ss.getSheetByName('welkomst mail');
  var emailBijlagen = gmmSheet.getRange("bijlagenWelkomstMail").getValue().split(",");
  Logger.log(emailBijlagen);

}

function showAlert() {
  var result = Browser.msgBox(
    'Please confirm',
    'Are you sure you want to continue?',
    Browser.Buttons.YES_NO);

  // Process the user's response.
  if (result == 'yes') {
    // User clicked "Yes".
    Browser.msgBox('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    Browser.msgBox('Permission denied.');
  }
}

function printpdf(){
  
  keyOngetekendeMandatenMap = '0B_JKiFIPnTaGT21NUlRIbXdheTQ';

  var cloneId = DriveApp.getFileById('1CkNAM5O3XKZdFSqRaiUNq5Go0Z9nKnbHcfKvCTmT3P0').makeCopy('clone').getId();
  var clone = DocumentApp.openById(cloneId);

  var body = clone.getBody();
  
  body.replaceText('{kenmerk}', 'ETT00456v1');
  body.replaceText('{displaynaam}', 'Ruud van Vliet');
  body.replaceText('{adres}', 'Kneuterstraat 7');
  body.replaceText('{postcode}', '7384 CM');
  body.replaceText('{plaats}', 'Wilp');
  body.replaceText('{iban}', 'NL99RABO0123456789');
  
  clone.saveAndClose();

  var clonePDF = DriveApp.createFile(clone.getAs('application/pdf'));
  clonePDF.setName("machtiging-ruuds.apekool.pdf")
  clonePDF.removeFromFolder(DriveApp.getRootFolder());

  DriveApp.getFileById(cloneId).setTrashed(true);
}

function testPainGen() {
  
   debtors = [{
    "InstdAmt": "1.00",
    "MndtId": "ETT00010v1",
    "DtOfSgntr": "2013-12-11",
    "AmdmntInd": "false",
    "DbtrNm": "R. van Vliet",
    "DbtrAdrLine1": "Kneuterstraat 7",
    "DbtrAdrLine2": "7384 CM  Wilp",
    "DbtrIBAN": "NL56RABO0146102096",
    "DbtrUstrd": "Dit is testbericht 1"
   }, {
    "InstdAmt": "2.00",
    "MndtId": "ETT00011v1",
    "DtOfSgntr": "2013-12-12",
    "AmdmntInd": "false",
    "DbtrNm": "R. van Vliet",
    "DbtrAdrLine1": "Kneuterstraat 7",
    "DbtrAdrLine2": "7384 CM  Wilp",
    "DbtrIBAN": "NL93RABO0145751260",
    "DbtrUstrd": "Dit is testbericht 2"
   },{
    "InstdAmt": "3.00",
    "MndtId": "ETT00011v1",
    "DtOfSgntr": "2013-12-12",
    "AmdmntInd": "false",
    "DbtrNm": "QreaCom",
    "DbtrAdrLine1": "Kneuterstraat 7",
    "DbtrAdrLine2": "7384 CM  Wilp",
    "DbtrIBAN": "NL93RABO0145751260",
    "DbtrUstrd": "Dit is testbericht 3"
   }];
  
  createPainMessage('2014-01-06', debtors);
  
}

function reserve() {
  // haal alle leden op in leden
  var ss = SpreadsheetApp.openById(keyLedenlijst);
  var leden = haalLedenLijstOp('ledenlijst actueel');
  
  for (var i = 0; i < leden.length; ++i) {
    var lidSoort = leden[i]["soort"];
    if (leden[i]["gebdatum"] > jongerDan18Jaar) {
      leden[i]["bijzlid"] = "JL";
    }
    if (leden[i]["bijzlid"]) {
      lidSoort += "-" + leden[i]["bijzlid"];
    }
    Logger.log(berekenContributie(lidSoort) + "\t" + lidSoort + "\t" + leden[i]["userName"]);
  }
}

function testMatch() {
  s = '<div style="background-color:#f1f1f1; overflow:hidden;"><table width=530 align=center bgcolor="#ffffff" border=0>'+
    '<div style="font-family:Calibri,Verdana,serif;font-size: 12pt;text-align: left; color: #666666;">'+
'<tr><td colspan="3" bgcolor="#ffffff"><img src="http://www.ett-twello.nl/wp-content/uploads/2014/02/welkom-bij-de-ett.jpg" border="0" alt="Aankoniging Eerste Twellose Toerclub"></td></tr>'+
'<tr><td colspan="3" bgcolor="#ffffff"><img src="http://www.ett-twello.nl/wp-content/public/underban.jpg" border="0" alt="." width="530" height="22"></td> </tr>'+
'<tr><td width="15" valign="top" bgcolor="#ffffff"><img src="http://www.ett-twello.nl/wp-content/public/spacer.gif" border="0" alt="1" width="15" height="1"></td>'+
'<td width="500" valign="top" bgcolor="#ffffff">'+
'<p>Beste ${"voornaam"},</p>'

vars = s.match(/\$\{\"[^\"]+\"\}/g);
  
  Logger.log(vars);
}

function testXml() {
// Log an XML document in human-readable form.
 var xml = '<root><a><b>Text!</b><b>More text!</b></a></root>';
 var document = XmlService.parse(xml).getRootElement();
 var output = XmlService.getPrettyFormat()
     .format(document);
 Logger.log(output);
 }

function testDiakrieten() {
  Logger.log(replaceDiacritics('André'));
}

function testVoorletters() {
  Logger.log(formatVoorletters('jwr'));
}

function testUsername() {
  Logger.log(formatUserName('André Mulder-Theunissen'));
}

function testPostcode() {
  Logger.log(formatPostcode('1234   fh'));
}

function checkLedenLijst() {
  var leden = haalLedenLijstOp('ledenlijst actueel');
  
  Logger.log(leden.length)
  
  for (var i = 0; i < 20; ++i) { 
    Logger.log(leden[i])
  }
}

function getQuota() {
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
}

function tesModulo() {
  Logger.log(modulo("111111111111111111111111111111111111111111111111110","7"))
}

function testBankcode() {
  Logger.log(convertLatin("INGB"));
}

function testFiles() {
  getFileNamesAndIds(keyMandaatFolder,15);
}

function testNormaliseWord() {
  Logger.log(normalizeWord("displayName"));
}

function testSepaMandaat() {
  var member = {
    "kenmerk":"ETT0123v01",
    "tenNameVan":"Ruud",
    "adres":"straat 2",
    "postcode":"7384CM",
    "plaats":"Wilp",
    "iban":"NL01RABO0123456789",
    "userName":"piet.puk"
  }
  personaliseAttachment('1ohAndHsiEL9tYrGVOGzP0ZzWJlMedtphGI5iNALZ-kM', member);
}
function testMimeType() {
  Logger.log(DriveApp.getFileById('1Baeb62CaahpLTUVqp5llHny3k9c_MWxzFC6gQ-Dcf0c').getMimeType());
}

function testIban() {
  Logger.log(validateIban('NL47INGB0007539857'))
}

function checkUserNamesInIncassoLijst() {
  var incSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('incasso');
  var incassos =  haalIncassoLijstOp(incSheet);
  
  var ledenArray = haalLedenLijstOp("ledenlijst actueel");
  var leden = converteerArrayToObjects(ledenArray, 'userName');
  
  for (var i = 0; i < incassos.length; ++i) { 
    Logger.log(leden[incassos[i]['gebruikersnaam']]['voornaam']);
    Logger.log(incassos[i]['gebruikersnaam']);
  }
}

function testValidateEmail() {
  var validated = validateEmail(undefined);
  Logger.log(validated);
}

function testGetMembers() {
  var member = getMembers('adres')[0];
  Logger.log(member);
}

function testCreateDraft() {
  createDraftFromDocument('1iZ30ayOMNIqqhh5jtY3lOLXSh7CYFfcLcXCU4GMZFR0', 'ETT Nieuwsbrief')
}