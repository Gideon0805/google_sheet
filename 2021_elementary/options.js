/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// 開啟Spreadsheet就觸發執行onOpen()函式 https://developers.google.com/apps-script/guides/triggers/#onopene

function onOpen() {
  // 選單增加自訂功能 https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#addMenu(String,Object)
  var menuEntries = [ {name: "整理名單", functionName: "getNameList"},
  					  {name: "產生繳費名單", functionName: "getRemittance"},
  					  {name: "人口普查", functionName: "getSummary"},
                      null,];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("[自訂功能]", menuEntries); 
}

/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
