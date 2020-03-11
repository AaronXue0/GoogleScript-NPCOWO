function onSubmit(e) {
  var range = e.range;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(range.getLastRow() <= 47){
    sendingEmail(range,sheet); 
  }
  else {
    var af = FormApp.openByUrl('https://docs.google.com/forms/d/1F3ofA1cl5qRKGydLh8w0Zlq9E0dnMcH5icJPAOV-Mxw/edit');
    af.setCustomClosedFormMessage("很抱歉，為保證上課品質，報名人數已滿\n請關注社團資訊，或等待之後的相關課程！");
    af.setAcceptingResponses(false);
  }
}

function sendingEmail(range,sheet){
  Logger.log(sheet.getRange(range.getLastRow(), 2).getValue());
  var emailAddress = sheet.getRange(range.getLastRow(), 2).getValue();
  
  var subject = '【 報名成功通知 】Unity 主題課程課程 by NPC 北科程式設計研究社';
  

              
  MailApp.sendEmail(emailAddress, subject, message);

}

function waitingList(range,sheet){
  var emailAddress = sheet.getRange(range.getLastRow(), 2).getValue();
  
  var subject = '【備取】OOP 物件導向 進階課程 by NPC 北科程式設計研究社 ';
  
  var message = 'Hi, ' + '\n\n' +
                '感謝您報名 NPC 北科程式設計研究社 - OOP 物件導向課程\n' + 
                '為了保證上課品質，我們人數已達到上限，如果有人放棄資格，我們會儘速通知您！\n\n' +

                '若有任何疑問，歡迎隨時連絡我們。\n' + 
                '由衷感謝您:)\n\n' + 
                'Best regards,\n' +
                'NPC 北科程式設計研究社';
              
  MailApp.sendEmail(emailAddress, subject, message);
}


function myFunction(){
  var url = 'https://docs.google.com/spreadsheets/d/1RdGtWJL_RGtKhJUA9XZe7KtyT6boV2nD0c8jiJZNL5M/edit#gid=1146367932';
  var name = '表單回應 1';
  var SpreadSheet = SpreadsheetApp.openByUrl(url)
  var SheetName = SpreadSheet.getSheetByName(name)
  var message = '您好, ' + '\n\n' +
                '提醒您，Unity 主題課程 即將到來！\n' + 
                '第一堂課程請您務必出席並繳交保證金。\n' +
                '若臨時無法參與課程請務必提前告知，好將機會讓給他人。\n\n' + 
                '【課程資訊】\n' +
                '一. 基礎操作與物件控制\n'+
                '1. 基本操作及環境介紹\n'+
                '2. 重力及碰撞Collider 及 rigidbody\n'+
                '3. script(PlayerControl) 及 Start 及 Update 及 transform.position\n\n'+
                '二. Scene 場景控制\n'+
                '1. 物件移動及跳躍\n'+
                '2. 場景轉換及結束(勝利或死亡)\n\n'+
                '三. Animation 與 state control\n'+
                '1. 陷阱及機關\n'+
                '2. 設計及擺放\n'+
                '3. 發布Build\n\n' +
                '【時    間】：11/28（四）19:00~21:00 ( 18:30 入場 )\n' + 
                '【地    點】：臺北科技大學 宏裕科技研究大樓 教室1223\n\n' +
                '【備註】\n' + 
                '1. 現場提供電腦；於教室內時，請勿飲食\n' +
                '2. 若需提前離開，請務必在當天告知工作人員\n\n' +
                '非常期待與您在課程上相見！\n' + 
                'Best regards\n' +
                'NPC 北科程式設計研究社';
                
  var subject = '【Unity 主題課程】行前通知信 by NPC ';
  for(i = 2;i <= SheetName.getLastRow();i++){
    var emailAddress = SheetName.getSheetValues(i,2,1,1);
    console.log(emailAddress);
    MailApp.sendEmail(emailAddress,subject,message);
  }
}
