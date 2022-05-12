function emailCannon() {

var emailAddresses = [];
var subjects = [];
var bodies = [];
var attachements = [];
var checkEmpty = [];

for (let i = 3; i < SpreadsheetApp.getActive().getSheetByName('Email Cannon').getLastRow()+1; i++){
  if (SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('A'+i).getValue() == ''){
    emailAddresses.push(['NONE']);
  } else {
    emailAddresses.push([SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('A'+i).getValue()]);
  };

  if (SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('B'+i).getValue() == ''){
    subjects.push(['NONE']);
  } else {
  subjects.push([SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('B'+i).getValue()]);
  };

  if (SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('C'+i).getValue() == ''){
    bodies.push(['NONE']);
  } else {
  bodies.push([SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('C'+i).getValue()]);
  };


  if (SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('K'+i).getValue() == ''){
    checkEmpty.push(0);
  } else if (SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('L'+i).getValue() == ''){
    checkEmpty.push(1);
  } else if (SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('M'+i).getValue() == ''){
    checkEmpty.push(2);
  } else if (SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('N'+i).getValue() == ''){
    checkEmpty.push(3);
  } else if (SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('O'+i).getValue() == ''){
    checkEmpty.push(4);
  } else if (SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('P'+i).getValue() == ''){
    checkEmpty.push(5);
  } else if (SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('Q'+i).getValue() == ''){
    checkEmpty.push(6); 
  } else {
    checkEmpty.push(7);
  };

};
//AFTER THE FIRST FOR LOOP



for (let e = 0; e < SpreadsheetApp.getActive().getSheetByName('Email Cannon').getLastRow(); e++){


  if (checkEmpty[e] == 7){
    attachements.push([
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('K'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('L'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('M'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('N'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('O'+(e+3)).getValue(),    
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('P'+(e+3)).getValue(),    
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('Q'+(e+3)).getValue() ]);
  } else if (checkEmpty[e] == 6){
    attachements.push([
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('K'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('L'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('M'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('N'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('O'+(e+3)).getValue(),    
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('P'+(e+3)).getValue() ]);
  } else if (checkEmpty[e] == 5){
    attachements.push([
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('K'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('L'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('M'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('N'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('O'+(e+3)).getValue() ]);
  } else if (checkEmpty[e] == 4){
    attachements.push([
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('K'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('L'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('M'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('N'+(e+3)).getValue() ]);
  } else if (checkEmpty[e] == 3){
    attachements.push([
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('K'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('L'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('M'+(e+3)).getValue() ]);
  } else if (checkEmpty[e] == 2){
    attachements.push([
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('K'+(e+3)).getValue(),
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('L'+(e+3)).getValue() ]);
  } else if (checkEmpty[e] == 1){
    attachements.push([
    SpreadsheetApp.getActive().getSheetByName('Email Cannon').getRange('K'+(e+3)).getValue() ]);
  } else {
    attachements.push(['NONE']);
  };

  if (checkEmpty[e] == 7 && emailAddresses[e][0] != 'NONE'){
          GmailApp.sendEmail(emailAddresses[e][0].toString(), subjects[e][0].toString(), bodies[e][0].toString(),{ attachments: [
          DriveApp.getFileById(attachements[e][0].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][1].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][2].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][3].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][4].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][5].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][6].toString()).getAs(MimeType.PDF) ]});
  } else if (checkEmpty[e] == 6 && emailAddresses[e][0] != 'NONE'){
          GmailApp.sendEmail(emailAddresses[e][0].toString(), subjects[e][0].toString(), bodies[e][0].toString(),{ attachments: [
          DriveApp.getFileById(attachements[e][0].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][1].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][2].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][3].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][4].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][5].toString()).getAs(MimeType.PDF) ]});
  } else if (checkEmpty[e] == 5 && emailAddresses[e][0] != 'NONE'){
          GmailApp.sendEmail(emailAddresses[e][0].toString(), subjects[e][0].toString(), bodies[e][0].toString(),{ attachments: [
          DriveApp.getFileById(attachements[e][0].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][1].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][2].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][3].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][4].toString()).getAs(MimeType.PDF) ]});
  } else if (checkEmpty[e] == 4 && emailAddresses[e][0] != 'NONE'){
          GmailApp.sendEmail(emailAddresses[e][0].toString(), subjects[e][0].toString(), bodies[e][0].toString(),{ attachments: [
          DriveApp.getFileById(attachements[e][0].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][1].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][2].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][3].toString()).getAs(MimeType.PDF) ]});
  } else if (checkEmpty[e] == 3 && emailAddresses[e] != 'NONE'){
          GmailApp.sendEmail(emailAddresses[e][0].toString(), subjects[e][0].toString(), bodies[e][0].toString(),{ attachments: [
          DriveApp.getFileById(attachements[e][0].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][1].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][2].toString()).getAs(MimeType.PDF) ]});
  } else if (checkEmpty[e] == 2 && emailAddresses[e] != 'NONE'){
          GmailApp.sendEmail(emailAddresses[e][0].toString(), subjects[e][0].toString(), bodies[e][0].toString(),{ attachments: [
          DriveApp.getFileById(attachements[e][0].toString()).getAs(MimeType.PDF),
          DriveApp.getFileById(attachements[e][1].toString()).getAs(MimeType.PDF) ]});
  } else if (checkEmpty[e] == 1 && emailAddresses[e] != 'NONE'){
          GmailApp.sendEmail(emailAddresses[e][0].toString(), subjects[e][0].toString(), bodies[e][0].toString(),{ attachments: [
          DriveApp.getFileById(attachements[e][0].toString()).getAs(MimeType.PDF) ]});
  } else if (checkEmpty[e] == 0 && emailAddresses[e] != 'NONE'){
          GmailApp.sendEmail(emailAddresses[e][0].toString(), subjects[e][0].toString(), bodies[e][0].toString());
  }

};
//AFTER THE SECOND FOR LOOP



}






