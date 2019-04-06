
function checkMandaatBestandenTegenLedenlijst() {
  var controleSheet = SpreadsheetApp.getActive().getSheetByName('mandaat controles');
  controleSheet.activate();
  controleSheet.getDataRange().clearContent();
  controleSheet.getRange(1,1).setValue("Overzicht");
  controleSheet.getRange(2,1).setValue("moment geduld ...");
  controleSheet.getRange(1,2).setValue("Incasso maar geen machtiging");
  controleSheet.getRange(1,3).setValue("Machtiging maar geen lid/incasso");
  controleSheet.getRange(1,4).setValue("Machtiging maar geen datum");
  
  SpreadsheetApp.flush();
  
  var leden = haalLedenLijstOp('ledenlijst actueel');
  var userNames = [];
  for(var i = 0; i < leden.length; ++i) {
    Logger.log(leden[i]);
    if (leden[i]['userName'] != undefined) {
      userNames.push(normalizeWord(leden[i]['userName']));
    }
  };
  
  var fileNames = getFileNames();

  var countBestanden = 0;
  var countLedenMetIncasso = 0;
  var countLedenMetSecIncasso = 0;
  var countLedenZonderIncasso = 0;
  var countMachtigingMaarGeenLid = 0;
  var countGeenMachtiging = 0;
  var countGeenDatum = 0;
  
 
  // doorloop de bestanden en zoek het lid erbij
  for (var i = 0, len = fileNames.length; i < len; ++i)  {
    // aantal gevonden machtigingsbestanden
    ++countBestanden;
    // staat het lid ook in de ledenlijst? let op: de fileName is 'normalised'!!
    if (userNames.indexOf(fileNames[i]) == -1) {
      ++countMachtigingMaarGeenLid;
      controleSheet.getRange(countMachtigingMaarGeenLid+1,3).setValue(fileNames[i]);
    }
    SpreadsheetApp.flush();
  }
  
  // doorloop de leden en zoek het bestand erbij
  for (var i = 0; i < leden.length; ++i) { 
    if (leden[i]["betaling"] == "INCASSO") {
      // Lid betaalt per incasso
      ++countLedenMetIncasso;
      // is de machtiging erbij?
      if(fileNames.indexOf(normalizeWord(leden[i]['userName'])) == -1) {
        ++countGeenMachtiging; 
        controleSheet.getRange(countGeenMachtiging+1,2).setValue(leden[i]["displayName"]+" - "+leden[i]["kenmerk"]);
      }
    } else if (leden[i]["betaling"] == "SEC-INC") {
      // Lid betaalt per incasso van een gezinslid
      ++countLedenMetSecIncasso;
    } else {
      // Lid betaald niet per incasso
      ++countLedenZonderIncasso;
      if (leden[i]["userName"] && fileNames.indexOf(normalizeWord(leden[i]["userName"])) >= 0) {
        // en er is toch nog een machtiging 
        ++countMachtigingMaarGeenLid; 
        controleSheet.getRange(countMachtigingMaarGeenLid+1,3).setValue(leden[i]["userName"]);
      }
    }
    
    if (!(typeof(leden[i]["datumMachtiging"]) == "object" && leden[i]["datumMachtiging"].getYear())  && leden[i]["userName"] && fileNames.indexOf(normalizeWord(leden[i]["userName"])) >= 0) {
      // er geen datum maar wel een bestand gevonden
      ++countGeenDatum; 
      controleSheet.getRange(countGeenDatum+1,4).setValue(leden[i]["displayName"]);
    }
    SpreadsheetApp.flush();
  }
  
  controleSheet.getRange(2,1).setValue(leden.length + " leden in de ledenlijst.");
  controleSheet.getRange(3,1).setValue(countLedenMetIncasso + " leden met INCASSO, nog "  + countGeenMachtiging + " zonder machtiging.");
  controleSheet.getRange(4,1).setValue(countLedenMetSecIncasso + " leden op de incasso van een gezinslid.");
  controleSheet.getRange(5,1).setValue(countLedenZonderIncasso + " leden zonder incasso (factuur of geen contributie verschuldigd).");
  controleSheet.getRange(6,1).setValue(countBestanden + " getekende machtigingen gevonden, waarvan "  + countMachtigingMaarGeenLid + " zonder lid.");
}

function archiveFilesWithoutMember() {
  var controleSheet = SpreadsheetApp.getActive().getSheetByName('mandaat controles');
  members2BArchived = [].concat.apply([], controleSheet.getRange(2,3,controleSheet.getLastRow()).getValues());
  var files = mandaatFolder.getFiles();
  
  while (files.hasNext()) {
    var currentFile = files.next();
    var fileName = currentFile.getName().substr(11,(currentFile.getName().length-15));
    if (members2BArchived.indexOf(normalizeWord(fileName)) >= 0) {
      archiefMandaatFolder.addFile(currentFile);
      mandaatFolder.removeFile(currentFile);
    }
  }
}

// Maakt een array aan met het username deel uit de naam van het mandaatbestand
function getFileNames() { 
  var mandaatFolder = DriveApp.getFolderById('0B-c8daT0x_MvcFpCYmI3V3UwWjQ');
  var files = mandaatFolder.getFiles();
  var fileNames = [];
  
  while (files.hasNext()) {
    var currentFile = files.next();
    var fileName = currentFile.getName().substr(11,(currentFile.getName().length-15));
    fileNames.push(normalizeWord(fileName));
  }
  return fileNames;
}