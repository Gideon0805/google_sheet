function resetFunction() {
  // 存取工作表(ByName) https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheetbynamename
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //獲取“報到”的最新一筆資料
  var tempSheet = ss.getSheetByName("大會"); 
  tempSheet.clear();
  tempSheet = ss.getSheetByName("白傑睿"); 
  tempSheet.clear();
  tempSheet = ss.getSheetByName("紅麥克"); 
  tempSheet.clear();

  for (var key in sheetByName) {
    let index = sheetByName[key];
    tempSheet = ss.getSheetByName(index[0]);
    let content = [['名稱', key], ['積分', 250], ['行動數', 0]];
    tempSheet.getRange(1, index[1], 3, 2).setValues(content);
  }

  for (var key in outputUrl) {
    let tempUrl = outputUrl[key];
    let tag = '積分 動作';
    let tempSS = SpreadsheetApp.openByUrl(tempUrl);
    let tempSheet = tempSS.getSheetByName(tag);
    tempSheet.clear();
    let content = [['名稱', key], ['積分', 250], ['行動數', 0]];
    tempSheet.getRange(1, 1, 3, 2).setValues(content);
  }

  var result = ss.getSheetByName("大會");
  var temp = ss.getSheetByName("白傑睿");
  var tempContent = temp.getRange(1, 1, 2, 11).getValues();
  result.getRange(1, 1, 2, 11).setValues(tempContent);
  temp = ss.getSheetByName("紅麥克");
  tempContent = temp.getRange(1, 1, 2, 11).getValues();
  result.getRange(3, 1, 2, 11).setValues(tempContent);

}
