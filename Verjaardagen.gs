/*
 * Genereert de lijst met verjaardagen in de sheet 'verjaardagslijst'
 * Verwacht het kwartaal en het jaar (huidig of volgend) in de sidebar
 *
 * De Sidebar UI zit in VerjaardagenUI.html
 *
 * Auteur: Ruud van Vliet - ruud.van.vliet@ett-twello.nl
 * Datum:  6 januari 2018
 */

function verjaardagen() {
  // Gegevens ophalen uit de sheet, de verjaardagenlijst leegmaken en de sidebar openen
  var verjaardagslijstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('verjaardagslijst');
  maakLijstLeeg(verjaardagslijstSheet);

  var html = HtmlService.createHtmlOutputFromFile('VerjaardagenUI').setTitle('Verjaardagenlijst');
  SpreadsheetApp.getUi().showSidebar(html);
}

function toonVerjaardagenLijst(kwartaal, jaar) {
  var verjaardagslijstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('verjaardagslijst');
  maakLijstLeeg();

  // Constanten vastleggen
  var nlMaandNamen = ['januari', 'februari', 'maart', 'april', 'mei', 'juni', 'juli', 'augustus', 'september', 'oktober', 'november', 'december'];
  var offset = 2; // Uitvoer begint op regel 2

  // Ledenlijst omzetten naar de verjaardagslijst
  var ledenArray = haalLedenLijstOp("ledenlijst actueel");
  var verjaardagslijst = [];
  for (var i = 0, len = ledenArray.length; i < len; i++) {
    if (parseInt(ledenArray[i].gebdatum.getMonth() / 3) + 1 == kwartaal) {
      var verjaardag = {};
      verjaardag.displayName = ledenArray[i].displayName;
      verjaardag.gebDatumDag = ledenArray[i].gebdatum.getDate();
      verjaardag.gebDatumMaand = ledenArray[i].gebdatum.getMonth();
      verjaardag.gebDatumJaar = ledenArray[i].gebdatum.getFullYear();
      verjaardagslijst.push(verjaardag);
    }
  };
  verjaardagslijst.sort(function(a, b) {
    return ((a.gebDatumMaand * 100 + a.gebDatumDag) - (b.gebDatumMaand * 100 + b.gebDatumDag))
  });

  // Verjaardagslijst printen
  for (var i = 0, len = verjaardagslijst.length; i < len; i++) {
    verjaardagslijstSheet.getRange(offset + i, 1).setValue(verjaardagslijst[i].displayName);
    verjaardagslijstSheet.getRange(offset + i, 2).setValue(verjaardagslijst[i].gebDatumDag + " " + nlMaandNamen[verjaardagslijst[i].gebDatumMaand]);
    //verjaardagslijstSheet.getRange(offset + i, 3).setValue(jaar - verjaardagslijst[i].gebDatumJaar);
  }
}

function maakLijstLeeg(verjaardagslijstSheet) {
  if (! verjaardagslijstSheet) {
    verjaardagslijstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('verjaardagslijst');
  }
  verjaardagslijstSheet.activate();
  verjaardagslijstSheet.getDataRange().clearContent();
  verjaardagslijstSheet.getRange(1, 1).setValue("Naam");
  verjaardagslijstSheet.getRange(1, 2).setValue("Datum");
  //verjaardagslijstSheet.getRange(1, 3).setValue("Leeftijd");
}
