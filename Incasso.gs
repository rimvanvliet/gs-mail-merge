function incasso() {
  
  var incSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('incasso');
  
  var datumIncasso = Utilities.formatDate(incSheet.getRange("datumIncasso").getValue(), "CET", "dd-MM-yyyy");
  var msgId = incSheet.getRange("kenmerkIncasso").getValue();
  var factuurNummer = incSheet.getRange("factuurNummer").getValue();
  
  var ledenArray = haalLedenLijstOp("ledenlijst actueel");
  var leden = converteerArrayToObjects(ledenArray, 'userName');
  
  if (incSheet.getRange("soortIncasso").getValue() == 'jaar') {
    var incassos = creeerJaarContributieLijst(msgId, ledenArray);
  } else {
    var incassos = haalIncassoLijstOp(incSheet);
  }
  var debtors = [];
  
  for (var i = 0; i < incassos.length; ++i) { 
    if (incassos[i]['bedrag'] != undefined) {
      var debtor = {};
      if (! leden[incassos[i]['gebruikersnaam']]) {
        Logger.log(incassos[i]['gebruikersnaam'] + " niet gevonden.");
      } else {
        debtor["betaling"] =  leden[incassos[i]['gebruikersnaam']]['betaling'];
        if (debtor["betaling"] == "FACTUUR") {
          incassos[i]['omschrijving'] += " + admin. kosten (2.50)";
          debtor["factuurnummer"] = 'D'+Utilities.formatString("%03d", factuurNummer);
          factuurNummer++;
        }
        
        // kopieer alle properties van het lid naar de debtor
        for (var property in leden[incassos[i]['gebruikersnaam']]) {
          if (leden[incassos[i]['gebruikersnaam']].hasOwnProperty(property)) {
            debtor[property] =  leden[incassos[i]['gebruikersnaam']][property];
          }
        }
        
        debtor["incassodatum"] = datumIncasso;
        debtor["msgid"] = msgId;
        
        debtor["mndtid"] = leden[incassos[i]['gebruikersnaam']]['kenmerk'];
        
        if (leden[incassos[i]['gebruikersnaam']]['datumMachtiging']) {
          debtor["dtofsgntr"] = Utilities.formatDate(leden[incassos[i]['gebruikersnaam']]['datumMachtiging'], "CET", "yyyy-MM-dd");
        }
        debtor["amdmntind"] = false;
        debtor["dbtrnm"] = leden[incassos[i]['gebruikersnaam']]['tenNameVan'];
        debtor["dbtriban"] = leden[incassos[i]['gebruikersnaam']]['iban'];
        debtor["dbtrustrd"] = incassos[i]['omschrijving'];
        debtor["instdamt"] = Utilities.formatString('%.2f', incassos[i]['bedrag']);
        
        debtors.push(debtor);
      }
    }
  }
  
  // vooeg dubbelen per mandaat id samen
  debtors.sort(compare);
  
  for (var i = debtors.length-2; i >= 0; i--) { 
    if(debtors[i].mndtid.substr(0,10) == debtors[i+1].mndtid.substr(0,10)) {
      debtors[i].instdamt = Utilities.formatString('%.2f',parseFloat(debtors[i].instdamt) + parseFloat(debtors[i+1].instdamt));
      debtors[i].dbtrustrd += ", " + debtors[i+1].dbtrustrd;
      debtors.splice(i+1,1);
    }
  }
  
  
  createPainMessage(incSheet, debtors);
  
  sendIncassoEmails(incSheet, debtors);
}

function haalIncassoLijstOp(dataSheet) {
  
  var dataRange = dataSheet.getRange(2, 1, dataSheet.getMaxRows(), dataSheet.getMaxColumns()-2);
  return getRowsData(dataSheet, dataRange);
}

function creeerJaarContributieLijst(msgId, leden) {
  var objects = [];
  for (var i = 0; i < leden.length; ++i) {
   //Logger.log(leden[i]);
    if (leden[i]["opgezegd"] == undefined) {
      var object = {};
      var hasData = false;
      var lidSoort = leden[i]["soort"];
      if (leden[i]["gebdatum"] > jongerDan18Jaar) {
        leden[i]["bijzlid"] = "JL";
      }
      if (leden[i]["bijzlid"]) {
        lidSoort += "-" + leden[i]["bijzlid"];
      }
      var contributie = berekenContributie(lidSoort);
      if (contributie["bedrag"] > 0) {
        object["gebruikersnaam"] = leden[i]["userName"];
        object["bedrag"] = contributie["bedrag"];
        object["omschrijving"] = replaceDiacritics(leden[i]["voornaam"]) + " " + contributie["soort"] + " (" +contributie["bedrag"] + ")"; 
      }
      objects.push(object);
    }
  }
  return objects;
}

function berekenContributie(lidSoort) {
  switch (lidSoort) {
    case "H":
      return {soort:"Hoofdlid","bedrag":60.0};
    case "GL":
      return {soort:"Gezinslid","bedrag":60.0};
    case "H-JL":
    case "GL-JL":
      return {soort:"Jeugdlid","bedrag":30.0};
    case "H-LvV":
    case "GL-LvV":
      return {soort:"Lid van verdienste","bedrag":30.0};
    case "M":
      return {soort:"M-lid","bedrag":30.0};
    case "NFL":
      return {soort:"Niet-fietsend lid","bedrag":20.0};
    case "NFL-LvV":
      return {soort:"Niet-fietsend lid van versdienste","bedrag":0.0};
      
    default:
      return ("Ongeldige lidSoort: "+lidSoort);
  }
}
