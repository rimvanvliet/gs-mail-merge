var member1 = {
  email:"rimvanvliet@gmail.com",
  tenNameVan: "Member één",
  userName: "memberEen",
  kenmerk:""
};
var member2 = {
  email:"ruud@vliet.io",
  displayName: "Member twee",
  userName: "memberTwee"
};
testMembers = [member1, member2];

var mandaatKey = '1UmNKez4gsa5T5i8IPLYcsNcAXRi4nzbose_uHN9xPfM';


function testPersonaliseAttachment() {
  personaliseAttachment(mandaatKey, member1,"test personaliseer");
}

function testMatch() {
  
    var cloneId = DriveApp.getFileById(mandaatKey).makeCopy('MergeMailerCloneAttachment').getId();
    var clone = DocumentApp.openById(cloneId);
    var body = clone.getBody();

  var templateVars = body.findText(/\{\{[^\}]+\}\}/g);
  Logger.log(templateVars);
}