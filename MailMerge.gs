/* MailMerge
 *
 * mailTemplate       object with name, cc, bcc, subject, body, htmlBody, attachments
 *                    attachments = array of keys of Drive documents
 * to                 object with the addressesses including attributes to substitute in the body, htmlBodt and attachments
 * To be able to use MailMerge, a member (the receiver of the mail, so the 'to') must have 3 properties:
 *                   - email
 *                   - displayName (column name: Display Name)
 *                   - userName (column name: User Name)
 */

mailMerge = function() {
  // de mailinglijst leegmaken en de sidebar openen
  var mailMergeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('mail merge');
  clearMailMergeSheet(mailMergeSheet);

  var ui = HtmlService
    .createTemplateFromFile('MailMergeUI')
    .evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Mail merge');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function getDraftMessages() {
  try {
    draftMessages = GmailApp.getDrafts();
    drafts = [];
    for (var i = 0; i < draftMessages.length; i++) {
      drafts.push({
        'id': draftMessages[i].getId(),
        'subject': draftMessages[i].getMessage().getSubject()
      })
    }
  } catch (err) {
    throw (err);
  }
  return drafts;
}

var newsletter = {
  'cc': '',
  'subject': 'ETT Nieuwsbrief',
  'htmlBodyKey': '1iZ30ayOMNIqqhh5jtY3lOLXSh7CYFfcLcXCU4GMZFR0',
  'attachmentFiles': [],
  'attachmentDocumentTemplates': []
}
var welcome = {
  'cc': 'Ledenadministratie ETT <ledenadministratie@ett-twello.nl>',
  'subject': 'Welkom bij de ETT',
  'htmlBodyKey': '1ccwgRaqOcFmO5nH0W0vCEPXKaMqbuF19wop8TDD-1Tc',
  'attachmentFiles': ['0B-c8daT0x_MvZmI4VVNwZ1VIcjA', '0B-c8daT0x_MvUUJIMFlsM0VZT1U', '0B--hLDgyENTEVGpOMjZ3cURkbFU'],
  'attachmentDocumentTemplates': ['1UmNKez4gsa5T5i8IPLYcsNcAXRi4nzbose_uHN9xPfM']
}
var privacy = {
  'cc': 'Secretaris ETT <secretaris@ett-twello.nl>',
  'subject': 'Geheimhoudingsverklaring ETT',
  'htmlBodyKey': '1TVzvyMG0wVFPF8OzcntwAH0hrxqOMLlzIkQpNyFJIdQ',
  'attachmentFiles': [],
  'attachmentDocumentTemplates': ['1UsSHE2-og6WzCqLUP4aK-K_swf6wE5DoXjApRcvgleg']
}
var permission = {
  'cc': 'Secretaris ETT <secretaris@ett-twello.nl>',
  'subject': 'Toestemmingsverklaring ETT',
  'htmlBodyKey': '1dAyZopNWVyLxL-y9wJT6ISFaLVgY4O_wobxq1r9BR7o',
  'attachmentFiles': [],
  'attachmentDocumentTemplates': ['1OEU8NnmbhH_UIyAVI7TM2hkPqOcEzPuVEVw5WzD3PPU']
}

function createDraftFromDocument(documentTypeName) {
  try {
    var documentType = '';
    switch (documentTypeName) {
      case 'newsletter': documentType = newsletter; break;
      case 'welcome': documentType = welcome; break;
      case 'privacy': documentType = privacy; break;
      case 'permission': documentType = permission; break;
      default: throw('Programmeerfout: Onbekend document soort opgegeven');
    }
    var htmlBody = DocumentApp.openById(documentType.htmlBodyKey).getBody().getText();
    var attachments = [];
    for (var i=0; i< documentType.attachmentFiles.length; i++) {
       attachments.push(DriveApp.getFileById(documentType.attachmentFiles[i]).getAs(MimeType.PDF));
    }
    var draft = GmailApp.createDraft(null, documentType.subject, '', {'cc':documentType.cc, 'htmlBody': htmlBody, 'attachments': attachments});
  } catch (err) {
    throw (err);
  }
}

function getEmail() {
  return Session.getActiveUser().getEmail();
}

function sendFromDraft(draftId, to, testEmail) {
  var mailSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('mail merge');
  clearMailMergeSheet(mailSheet);

  var gmailMessage = GmailApp.getDraft(draftId).getMessage();
  var messageTemplate = {
    cc: gmailMessage.getCc(),
    bcc: gmailMessage.getBcc(),
    subject: gmailMessage.getSubject(),
  }
  var extendedTos = getExtendedTos(to);

  if (testEmail != "") {
    var extendedTo = extendedTos[Math.floor(Math.random() * extendedTos.length)];
    extendedTo.email = testEmail;
    messageTemplate.cc = '';
    messageTemplate.body = gmailMessage.getPlainBody();
    messageTemplate.htmlBody = gmailMessage.getBody();
    messageTemplate.attachments = gmailMessage.getAttachments();

    sendPersonalisedMail(messageTemplate, extendedTo);
    mailSheet.getRange(2,1).setValue(extendedTo.displayName);
    mailSheet.getRange(2,2).setValue("test ok");
  } else {
    for (var i = 0; i < extendedTos.length; i++) {
      messageTemplate.body = gmailMessage.getPlainBody();
      messageTemplate.htmlBody = gmailMessage.getBody();
      messageTemplate.attachments = gmailMessage.getAttachments();
      sendPersonalisedMail(messageTemplate, extendedTos[i]);
      mailSheet.getRange(i+2,1).setValue(extendedTos[i].displayName);
      mailSheet.getRange(i+2,2).setValue("ok");
      SpreadsheetApp.flush();
    }
  }
}

function getExtendedTos(to) {
  var members = haalLedenLijstOp('ledenlijst actueel');
  return members.filter(function(member){
    return to.indexOf(member.userName) >= 0;
  });
}

/* sendPersonalisedMail
 *
 * Sends email (singular) with personalised body & attachments
 */
  sendPersonalisedMail = function(message, extendedTo) {
  message.to = extendedTo.displayName + '<' + extendedTo.email + '>';
  message.body = fillInTemplateFromObject(message.body, extendedTo);
  message.htmlBody = fillInTemplateFromObject(message.htmlBody, extendedTo);

  var welcomKey = 'Welkom bij de ETT'
  if (message.subject.substring(0, welcomKey.length) === welcomKey) {
     Logger.log('in sendPersonalisedMail');
    for (var i = 0; i < welcome.attachmentDocumentTemplates.length; i++) {
      message.attachments.push(personaliseAttachment(welcome.attachmentDocumentTemplates[i], extendedTo));
    }
  };

  var privacyKey = 'Geheimhoudingsverklaring ETT'
   if (message.subject.substring(0, privacyKey.length) === privacyKey) {
     Logger.log('in sendPersonalised privacy Mail');
    for (var i = 0; i < privacy.attachmentDocumentTemplates.length; i++) {
      message.attachments.push(personaliseAttachment(privacy.attachmentDocumentTemplates[i], extendedTo));
    }
  };

  var permissionKey = 'Toestemmingsverklaring ETT'
   if (message.subject.substring(0, privacyKey.length) === permissionKey) {
     Logger.log('in sendPersonalised permission Mail');
    for (var i = 0; i < permission.attachmentDocumentTemplates.length; i++) {
      message.attachments.push(personaliseAttachment(permission.attachmentDocumentTemplates[i], extendedTo));
    }
  };



  MailApp.sendEmail(message);
}

/* personaliseAttachment
 *
 * Personalises Google docs attachments
 * extendedTo:    member object, receiver of the email with all his properties
 * attachments:   array with file keys in Drive;
 *                if mimeType = GOOGLE_DOCS then all {{template_variables}} are substituted
 *                                          else the document is attached unchanged
 */

personaliseAttachment = function(attachmentKey, extendedTo) {

  if (DriveApp.getFileById(attachmentKey).getMimeType() != MimeType.GOOGLE_DOCS) {
    return DriveApp.getFileById(attachmentKey);
  }
  var cloneId = DriveApp.getFileById(attachmentKey).makeCopy('MergeMailerCloneAttachment').getId();
  var clone = DocumentApp.openById(cloneId);
  var body = clone.getBody();
  var fileName = (DriveApp.getFileById(attachmentKey).getName()) + '-' + extendedTo.userName + '.pdf';

  for (var property in extendedTo) {
    if (extendedTo.hasOwnProperty(property)) {
      var replace2 = extendedTo[property] instanceof Date ? formatDate(extendedTo[property]) : extendedTo[property]
      body.replaceText("{{" + property + "}}", replace2);
    }
  }
  clone.saveAndClose();

  var clonePDF = DriveApp.createFile(clone.getAs('application/pdf'));
  clonePDF.setName(fileName);

  DriveApp.getFileById(cloneId).setTrashed(true);

  return clonePDF;
}

function getMembers(order, filter) {
  var members = haalLedenLijstOp('ledenlijst actueel');
  var filtered = [];
  if (order) {
    if (['datumMachtiging', 'gebdatum', 'lidVanaf'].indexOf(order) >= 0) {
      members = members.sort(function(a, b) {
        return new Date(b[order]) - new Date(a[order])
      });
    } else {
      members = members.sort(function(a, b) {
        return String.localeCompare(new String(a[order]), new String(b[order]))
      });
    }
  }
  switch (filter) {
    case 'tfDigitaal':
      filtered = members.filter(function(member) {return member.wijkTF == 'D'})
      break;
    case 'jeugd':
      Logger.log(members[0])
      filtered = members.filter(function(member) {return member.bijzlid == 'JL'})
      break;
    default:
      filtered = members;
  }
  return filtered.map(function(member) {
    return {
      'userName': member.userName,
      'displayName': member.displayName
    }
  });
}

function clearMailMergeSheet(mailMergeSheet) {
  if (!mailMergeSheet) {
    mailMergeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('verjaardagslijst');
  }
  mailMergeSheet.activate();
  mailMergeSheet.getDataRange().clearContent();
  mailMergeSheet.getRange(1, 1).setValue("Naam");
  mailMergeSheet.getRange(1, 2).setValue("Mail verstuurd");
  SpreadsheetApp.flush();
}

// Helper function that puts external JS / CSS into the HTML file.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function formatDate(date) {
  var month =['januari','februari','maart','april','mei','juni','juli','augustus','september','oktober','november','december'];
  return '' + date.getDate() + ' ' + month[date.getMonth()] + ' ' + date.getFullYear();
}
