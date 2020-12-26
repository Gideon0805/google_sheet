function mainFunction() {
  // DataProcess();//存取當下新增的人名與時間
  var url = "https://docs.google.com/spreadsheets/d/1R08B6p163dOwKM-t4278U4xCMttHvfgHIP_gce3fC4I/edit#gid=1138466960";
  var tag = "名單"
  GetData(url, tag)
}

function DataProcess() {
  // 存取工作表(ByName) https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheetbynamename
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //獲取“報到”的最新一筆資料
  var sheet = ss.getSheetByName("表單回應"); 
  var lastrow = sheet.getLastRow();
  var lastcol = sheet.getLastColumn();
  var rawData = sheet.getRange(lastrow, 1,1, lastcol).getValues()[0];//(start row, start col, row num, col num)//二維矩陣 取[0] 轉一維
  //2019/6/10 下午 8:54:29
  //var date = getDate("2019/6/10 下午 8:54:29");
  
  // 欲填入保險表格
  var insurance = ss.getSheetByName("保險");
  var in_row = insurance.getLastRow();
  // var lastcol = insurance.getLastColumn();
  // 保險欄位
  var data_nums = 6;
  var data_start = 1;
  for (var i = data_start; i <= data_nums; i++) {
    var temp = rawData[i];
    insurance.getRange(in_row+1, i).setValue(temp);
    // console.log(temp);
  }  
}

function GetData(url, tag) {
	var sheetTag = tag;
	var sheetUrl = url;
	var SpreadSheet = SpreadsheetApp.openByUrl(sheetUrl);
	var Sheet = SpreadSheet.getSheetByName(sheetTag);
	var lastrow = sheet.getLastRow();
	var lastcol = sheet.getLastColumn();
	var rawData = sheet.getRange(lastrow, 1,1, lastcol).getValues()[0];
	console.log(rawData)

}