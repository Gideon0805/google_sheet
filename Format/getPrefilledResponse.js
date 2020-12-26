function getPrefilledResponse()
{
  // 取得工作表之對應表單網址 https://developers.google.com/apps-script/reference/spreadsheet/sheet#getFormUrl()
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formURL = ss.getFormUrl();
  Logger.log(formURL);  //執行後... 查看->記錄 
  
  // 依表單網址開啟form(物件) https://developers.google.com/apps-script/reference/forms/form-app#openbyurlurl
  var form = FormApp.openByUrl(formURL);
  
  // 取得表單的所有題項(陣列) https://developers.google.com/apps-script/reference/forms/form#getitems
  var items = form.getItems();
 
  // 產生表單的預先填入連結 https://www.codetw.com/xhkyxhk.html  
  var formResponse = form.createResponse();  
  var prefilledResponse = formResponse.withItemResponse(items[0].asTextItem().createResponse("0000000000")).toPrefilledUrl();
  Logger.log(prefilledResponse);  //執行後... 查看->記錄
  
  // 調整表單的預先填入連結 http://bionicteaching.com/silent-submission-of-google-forms/
  // 將特殊符號(,/?:@&=+$#)轉編碼 https://www.w3schools.com/jsref/jsref_encodeURIComponent.asp
  prefilledResponse = encodeURIComponent(prefilledResponse).replace("viewform", "formResponse");
  Logger.log(prefilledResponse);  //執行後... 查看->記錄
  
  // 將prefilledResponse傳給callAppsLaunch()做後續處理
  callAppsLaunch(prefilledResponse);

}


function callAppsLaunch(prefilledResponse)
{
  // QR Droid Private(Android版) https://play.google.com/store/apps/details?id=la.droid.qr.priva
  // 啟動該App的語法 https://stackoverflow.com/questions/15840659/
  var qrdroidAppLaunch = "http://qrdroid.com/scan?q=" + prefilledResponse.replace("0000000000", "{CODE}");
  Logger.log(qrdroidAppLaunch);  //執行後... 查看->記錄

  // Zxing(Android版) https://play.google.com/store/apps/details?id=com.google.zxing.client.android
  // 啟動該App的語法 https://github.com/zxing/zxing/wiki/Scanning-From-Web-Pages
  // var zxingAppLaunch = "http://zxing.appspot.com/scan?SCAN_FORMATS=QR&ret=" + prefilledResponse.replace("0000000000", "{CODE}");
  // Logger.log(zxingAppLaunch);  //執行後... 查看->記錄
  
  // QRafter(iOS版) https://itunes.apple.com/us/app/id416098700
  // 啟動該App的語法 https://qrafter.com/qrafter-x-callback-url/
  var qrafterAppLaunch = "https://qrafter.com/x-callback-url/scan?x-success=" + prefilledResponse.replace("0000000000", "{CODE}");
  Logger.log(qrafterAppLaunch);  //執行後... 查看->記錄

  // pic2shop(Android版) https://play.google.com/store/apps/details?id=com.visionsmarts.pic2shop
  // pic2shop(iOS版) https://itunes.apple.com/mo/app/id308740640
  // 啟動該App的語法 http://www.pic2shop.com/demo/scan_source.html
  // var pic2shopAppLaunch = "pic2shop://scan?formats=QR&callback=" + prefilledResponse.replace("0000000000", "CODE");
  // Logger.log(pic2shopAppLaunch);  //執行後... 查看->記錄
  
  
  // 將這幾個App的資料存入message4buddy陣列，再傳給sendmail2buddy()寄出報到處工作程序說明信件
  var message4buddy = [];
  message4buddy.push({"App" : "QR Droid Private(Android版)",
                      "Url" : "https://play.google.com/store/apps/details?id=la.droid.qr.priva",
                      "Call": qrdroidAppLaunch
                     });
  message4buddy.push({"App" : "QRafter(iOS版)",
                      "Url" : "https://itunes.apple.com/us/app/id416098700",
                      "Call": qrafterAppLaunch
                     });
  
  sendmail2buddy(message4buddy);
    
}


function sendmail2buddy(message4buddy)
{
  // 用GAS發出郵件的每日額度(一般Gmail是100人次、G Suite教育版是1500人次) https://developers.google.com/apps-script/guides/services/quotas
  // 用GAS發出郵件的每日餘額 https://developers.google.com/apps-script/reference/mail/mail-app#getremainingdailyquota
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();

  // 跳出對話框 https://developers.google.com/apps-script/reference/base/ui#prompt(String,String,ButtonSet) 
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('準備發給報到處工作人員的程序說明信件', '請輸入收件人 ( 多人以逗號","分隔  至多寄給'+emailQuotaRemaining+'人 )', ui.ButtonSet.OK_CANCEL);
  
  // 測試時要到Spreadsheet從跳出的對話框輸入收件人信箱
  if (response.getSelectedButton() == ui.Button.OK) {
      var emailTo = response.getResponseText();
      Logger.log(emailTo);  //執行後... 查看->記錄
    
      // 用Gmail寄信 https://developers.google.com/apps-script/reference/mail/mail-app#sendEmail(Object)
      MailApp.sendEmail({
        to: emailTo,
        subject: "報到",
        htmlBody: "報到處的工作夥伴們好!<br><br>" +
                  "請事先安裝下列其中一個掃描二維條碼的App，而點擊下一行的網址就能自動呼叫該App執行掃描QRcode報到程序。<br>" +
                  "<ul>" +
                     "<li><a href='" + message4buddy[0].Url + "'>" + message4buddy[0].App + "</a><br>報到掃描點擊這個: " + message4buddy[0].Call + "</li>" +
                     "<li><a href='" + message4buddy[1].Url + "'>" + message4buddy[1].App + "</a><br>報到掃描點擊這個: " + message4buddy[1].Call + "</li>" +
                  "</ul>" +
                  "<p style='color:red;'>***此郵件為系統自動傳送，請勿直接回覆! ***</p>"
      });
  
      // 右下方快顯通知(toast notification) https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#toastmsg-title
      SpreadsheetApp.getActiveSpreadsheet().toast("寄給了"+emailTo , "已完成寄出說明信件");
  }  
  
}




function showModalDialog4CC() {
  var html = '<div id="all">'
           +   '<table style="width:100%;background-color:#eee;">'
           +     '<tr>'
           +       '<td>'
           +         '<img src="http://creativecommons.tw/sites/creativecommons.tw/files/cc-by-88x31.png">'
           +         '<p>本授權條款允許使用者重製、散布、傳輸以及修改著作(包括商業性利用)，惟使用時必須按照著作人或授權人所指定的方式，表彰其姓名。</p>'
           +       '</td>'
           +     '</tr>'
           +   '</table>'
           +   '<p>此Google試算表(及連動的Google表單)之Google Apps Script開發者如下:</p>'
           +   '<ul>'
           +      '<li><a href="http://nchu-eucl.blogspot.com/" target="_blank">NCHU EUCL</a></li>'
           +      '<li><a href="http://cc.nchu.edu.tw/" target="_blank">NCHU CINC</a></li>'
           +      '<li><a href="https://www.pintech.com.tw" target="_blank">Pintech.com.tw</a></li>'
           +   '</ul>'
           +   '<p style="color:red;font-size:small">(也就是說在不移除此聲明的情況下，允許重製、散布、傳輸、修改、以及商業性利用。)</p>'
           + '</div>';
  // 在ModalDialog置入HTML語法 https://developers.google.com/apps-script/reference/base/ui#showmodaldialoguserinterface-title
  htmlOutput = HtmlService.createHtmlOutput(html).setWidth(520).setHeight(280);
  var ui = SpreadsheetApp.getUi();
  ui.showModalDialog(htmlOutput,"創用CC授權(姓名標示)");
}





