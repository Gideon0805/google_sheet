/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 開啟Spreadsheet就觸發執行onOpen()函式 https://developers.google.com/apps-script/guides/triggers/#onopene

function onOpen() {
  // 選單增加自訂功能 https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#addMenu(String,Object)
  var menuEntries = [ {name: "發給報到處工作人員的程序說明信件", functionName: "getPrefilledResponse"},
                      {name: "整批產生名單內每個人的報到QRcode", functionName: "generateQRcode"},
                      null,
                      {name: "創用CC授權(姓名標示)", functionName: "showModalDialog4CC"} ];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("[自訂功能]", menuEntries); 
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
