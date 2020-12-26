
function generateQRcode() {
  // 存取工作表(ByName) https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheetbynamename
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("名單");
  var lastrow = sheet.getLastRow();
  
  // 讓"名單"工作表被看到 https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#setactivesheetsheet
  ss.setActiveSheet(sheet);

  // 取得B欄資料 https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getrangea1notation
  var B = sheet.getRange("B2:B").getValues();
  // 取得A欄資料
  var A = sheet.getRange("A2:A");
  var QRcode = A.getValues();

  // 準備產生的QRcode除了uid並附加20碼隨機字元 https://stackoverflow.com/questions/1349404/
  var charString = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'; 
  for (var i = 0; i < lastrow-1 ; i++) {
    var randString = '';
    for (var j = 0; j < 20; j++) randString += charString.charAt(Math.floor(Math.random() * charString.length));
    // 以前是用Google Chart API的QRcode產生方式(但該API隨時都可能停用) https://developers.google.com/chart/infographics/docs/overview
    // QRcode[i][0] = '=IMAGE("http://chart.apis.google.com/chart?cht=qr&chs=500x500&chl=' + B[i][0] + '$' + randString + '")';
    // iPower 舊版
    QRcode[i][0] = '=IMAGE("https://chart.googleapis.com/chart?chs=200x200&cht=qr&chl=' + B[i][0] + '")';
    // 現在使用QR Gode Generator API的QRcode產生方式 https://www.qrcoder.co.uk/api/v1/docs/
    // QRcode[i][0] = '=IMAGE("https://www.qrcoder.co.uk/api/v1/?size=500&text=' + B[i][0] + '$' + randString + '")';
  }
  // 將QRcode存至A欄
  A.setValues(QRcode); 
  
  // 右下方快顯通知(toast notification) https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#toastmsg-title
  ss.toast("共計"+i+"個...", "已產生每個人的報到QRcode");
  
}