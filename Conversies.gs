var offset = 8;

function List2Site() {
  var leden = haalLedenLijstOp('ledenlijst actueel');
  
  
  var site = "";
  
  for (var i = 0; i < leden.length; ++i) { 
    if (leden[i]["id"] != undefined) {
      var achternaam = leden[i]["achternaam"];
      if (leden[i]["tussenv"] != undefined) {
        achternaam = leden[i]["tussenv"].concat(" ",achternaam);
      }
      Logger.log(achternaam);
      site = site.concat(leden[i]["id"],',',leden[i]["userName"],',,',leden[i]["voornaam"],',',achternaam,',',leden[i]["displayName"],',',leden[i]["email"],'\n');
    }
  }
 
  siteDoc = DocumentApp.create("list2site.csv");
  siteDoc.getBody().setText(site);
  siteDoc.saveAndClose();
}

function Form2List() {
  var nieuweLeden = haalLedenLijstOp('form import');
  var importSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('form import');  
  for (var i = 0; i <= nieuweLeden.length; i++) {

    var displayName = (importSheet.getRange(i+1,5).getValue() + ' ' + importSheet.getRange(i+1,3).getValue() + ' ' + importSheet.getRange(i+1,4).getValue()).replace(/ +(?= )/g,'');

    importSheet.getRange(offset+i+1,1).setValue(importSheet.getRange(i+1,1).getValue());
    importSheet.getRange(offset+i+1,2).setValue(formatVoorletters(importSheet.getRange(i+1,2).getValue()));
    importSheet.getRange(offset+i+1,3).setValue(importSheet.getRange(i+1,5).getValue());
    importSheet.getRange(offset+i+1,4).setValue(importSheet.getRange(i+1,3).getValue());
    importSheet.getRange(offset+i+1,5).setValue(importSheet.getRange(i+1,4).getValue());
    importSheet.getRange(offset+i+1,6).setValue(importSheet.getRange(i+1,9).getValue());
    importSheet.getRange(offset+i+1,7).setValue(formatPostcode(importSheet.getRange(i+1,10).getValue()));
    importSheet.getRange(offset+i+1,8).setValue(importSheet.getRange(i+1,11).getValue());
    importSheet.getRange(offset+i+1,9).setValue(importSheet.getRange(i+1,7).getValue());
    importSheet.getRange(offset+i+1,11).setValue(importSheet.getRange(i+1,6).getValue());
    importSheet.getRange(offset+i+1,12).setValue(importSheet.getRange(i+1,8).getValue());

    importSheet.getRange(offset+i+1,14).setValue('H');
    if (importSheet.getRange(i+1,12).getValue()) { importSheet.getRange(offset+i+1,15).setValue('GL');}
    if (importSheet.getRange(i+1,13).getValue()) { importSheet.getRange(offset+i+1,15).setValue('M');}
    importSheet.getRange(offset+i+1,18).setValue((new Date()).toLocaleDateString());

    importSheet.getRange(offset+i+1,21).setValue('INCASSO');
    importSheet.getRange(offset+i+1,22).setValue(validateIban(importSheet.getRange(i+1,16).getValue()));
    importSheet.getRange(offset+i+1,23).setValue(importSheet.getRange(i+1,17).getValue());
    //importSheet.getRange(offset+i+1,23).setValue("=IF(J"+String(offset+i+1)+"<>\"\";YEAR(now()-J"+String(offset+i+1)+")-1900;\"-\")");

    importSheet.getRange(offset+i+1,27).setValue(importSheet.getRange(i+1,20).getValue());
    importSheet.getRange(offset+i+1,28).setValue(formatUserName(displayName));
    importSheet.getRange(offset+i+1,29).setValue(displayName);
    importSheet.getRange(offset+i+1,30).setValue('ETT' + Utilities.formatString("%05d", importSheet.getRange(i+1,20).getValue()) + 'v1');
  }
}

function newForm2List() {
  var nieuweLeden = haalLedenLijstOp('form import');
  var importSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('form import');
  var headerRow = 8;
  for (var i = 0; i <= nieuweLeden.length; i++) {
    var nieuwLid = {};
    
    nieuwLid.displayname = (importSheet.getRange(i+1,5).getValue() + ' ' + importSheet.getRange(i+1,3).getValue() + ' ' + importSheet.getRange(i+1,4).getValue()).replace(/ +(?= )/g,'');
    nieuwLid.titel = importSheet.getRange(i+1,1).getValue();
    nieuwLid.voorletter = formatVoorletters(importSheet.getRange(i+1,2).getValue());
    
    Logger.log(nieuwLid['Voorletter']);
    
    
    for (var j=1; j<=importSheet.getLastColumn(); j++) {
      var headerName = importSheet.getRange(headerRow,j).getValue().toLowerCase().replace(/ /g,'');
      if (nieuwLid[headerName]) {
        importSheet.getRange(offset+i+1,j).setValue(nieuwLid[headerName])
      }
    }    
  }
}

function Username2Email() {
  var leden = haalLedenLijstOp('ledenlijst actueel');
  var ledenOpUsername = converteerArrayToObjects(leden, "userName");
  var importSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('form import');  
  Logger.log(ledenOpUsername);

  for (var i = 1; i <= importSheet.getMaxRows(); i++) {
    if (ledenOpUsername[importSheet.getRange(i,1).getValue()]) {
       importSheet.getRange(i,2).setValue(ledenOpUsername[importSheet.getRange(i,1).getValue()].email);
    }
  }
}

function formatVoorletters(voorletters){
  var formattedVL = "";
  for (var i = 0; i < voorletters.length; ++i) {
    var letter = voorletters[i];
    if (!isAlnum(letter)) {
      continue;
    }
    formattedVL += letter.toUpperCase() + ".";
  }
  return formattedVL;
}

function formatUserName(username){
  return replaceDiacritics(username.replace(/[\s\-]/g, '.').toLowerCase());
}

function formatPostcode(postcode) {
  var res = postcode.replace(/\s+/g, '');
  return res.slice(0,4)+' '+res.slice(4,6).toUpperCase();
}

function validateIban(iban){
  ibanNum = convertLatin(iban.toUpperCase().substring(4,8))+iban.substring(8,19)+"232100";
  if (Number(modulo(ibanNum,"97")) + Number(iban.substring(2,4)) != 98) {
    return "*** FOUTE IBAN ***"
  }
  return (iban.toUpperCase());
}

function convertLatin(bankcode) {
  var result = "";
  for(var i=0; i<bankcode.length; i++) {
    result += (bankcode[i].charCodeAt(0)-55).toString();
  }
  return result;
}

function birthDayCurrentYear(date) {
  date.setFullYear(2016);
  return date;
}

function testB() {
  var d = new Date(1957, 7, 30);
  Logger.log(d);
  Logger.log(d.setFullYear(2016));
  Logger.log(d);

  
}