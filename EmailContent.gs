/**
 * Tests the schema.
 */
function testSchemas() {
  var htmlBody = HtmlService.createHtmlOutputFromFile('mail_template').getContent();

  MailApp.sendEmail({
    to: 't107590017@ntut.org.tw',
    subject: '【Unity 社課】 行前通知信 ' + new Date(),
    htmlBody: htmlBody,
  });
}

function myFunction(){
  var url = 'https://docs.google.com/spreadsheets/d/1RdGtWJL_RGtKhJUA9XZe7KtyT6boV2nD0c8jiJZNL5M/edit#gid=1146367932';
  var name = '表單回應 1';
  var SpreadSheet = SpreadsheetApp.openByUrl(url)
  var SheetName = SpreadSheet.getSheetByName(name)

  for(i = 2;i <= SheetName.getLastRow();i++){
    var email = SheetName.getSheetValues(i,2,1,1).toString();
    testSchemas(email)
  }
}

function sendEmail(email) {
  var htmlBody = HtmlService.createHtmlOutputFromFile('mail_template').getContent();
  MailApp.sendEmail({
    to: email,
    subject: '【Unity 社課】 行前通知信 ' + new Date(),
    htmlBody: htmlBody,
  });
}


