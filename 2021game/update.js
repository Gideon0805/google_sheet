function updateFunction() {
  // 存取工作表(ByName) https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheetbynamename
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //獲取“報到”的最新一筆資料
  var result = ss.getSheetByName("大會");
  var temp = ss.getSheetByName("白傑睿");
  var tempContent = temp.getRange(1, 1, 2, 11).getValues();
  result.getRange(1, 1, 2, 11).setValues(tempContent);
  temp = ss.getSheetByName("紅麥克");
  tempContent = temp.getRange(1, 1, 2, 11).getValues();
  result.getRange(3, 1, 2, 11).setValues(tempContent);


  var nameList = ['白傑睿一', '白傑睿二', '白傑睿三', '白傑睿四',
                  '紅麥克一', '紅麥克二', '紅麥克三', '紅麥克四'];

  for (var i in nameList) {
    let key = nameList[i];
    let tempUrl = outputUrl[key];
    let tempIndex = sheetByName[key];
    var teamSheet = ss.getSheetByName(tempIndex[0]);
    let actionNums = teamSheet.getRange(3, tempIndex[1]+1, 1, 1).getValues()[0][0];
    let tag = '積分 動作';
    let tempSS = SpreadsheetApp.openByUrl(tempUrl);
    let tempSheet = tempSS.getSheetByName(tag);
    let tempContent = ss.getSheetByName(tempIndex[0]).getRange(1, tempIndex[1], 3+actionNums, 2).getValues();
    tempSheet.clear();
    tempSheet.getRange(1, 1, 3+actionNums, 2).setValues(tempContent);

  } 

}
