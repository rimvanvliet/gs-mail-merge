function checkSlackTegenLedenlijst() {
  var controleSheet = SpreadsheetApp.getActive().getSheetByName('Slack gebruik');
  maakControleSheetLeeg(controleSheet);
  
  var slackQuery = "https://slack.com/api/users.list"
  var slackHeaders = {
       Authorization:"Bearer " + slackKey
  }
  var slackOptions = {
        "method" : "get",
        "headers": slackHeaders,
        "muteHttpExceptions": false
      };

  var response = JSON.parse(UrlFetchApp.fetch(slackQuery, slackOptions)).members;
  var slackLedenLijst = [];
  for (var i = 0; i < response.length; i++) {
        slackLedenLijst.push(response[i].real_name);
    }

  response = haalLedenLijstOp('ledenlijst actueel');
  var ettLedenLijst = [];
  for (var i = 0; i < response.length; i++) {
        ettLedenLijst.push(response[i].displayName);
    }
  ettLedenLijst.sort();
  
  var inSlack = ettLedenLijst.filter(function(n) {
    return slackLedenLijst.indexOf(n) !== -1;
  });

  var nietSlack = ettLedenLijst.filter(function(n) {
    return slackLedenLijst.indexOf(n) === -1;
  });

  var misMatch = slackLedenLijst.filter(function(n) {
    return ettLedenLijst.concat(['Slackbot', undefined]).indexOf(n) === -1;
  });

  writeColumn(controleSheet, 2, misMatch);
  writeColumn(controleSheet, 3, nietSlack);
  writeColumn(controleSheet, 4, inSlack);

  summarize(controleSheet, inSlack.length, nietSlack.length, misMatch.length);
}


function maakControleSheetLeeg(controleSheet) {
  controleSheet.activate();
  controleSheet.getDataRange().clearContent();
  controleSheet.getRange(1,1).setValue("Overzicht");
  controleSheet.getRange(2,1).setValue("moment geduld ...");
  controleSheet.getRange(1,2).setValue("Naamverschillen");
  controleSheet.getRange(1,3).setValue("Niet in Slack");
  controleSheet.getRange(1,4).setValue("Wel in Slack");
}

function logInSheet(controleSheet, row, column, text) {
  controleSheet.getRange(row, column).setValue(text);
}

function writeColumn(controleSheet, column, list) {
    for (var i = 0; i < list.length; i++) {
        logInSheet(controleSheet, i+2, column, list[i]);
    }
}

function summarize(controleSheet, aantalslack, aantalNietSlack, naamsVerschil) {
  logInSheet(controleSheet, 2, 1, "Naamsverschillen: " + naamsVerschil);
  logInSheet(controleSheet, 3, 1, "Wel in Slack: " + aantalslack);
  logInSheet(controleSheet, 4, 1, "Niet in Slack: " + aantalNietSlack);
}
    
