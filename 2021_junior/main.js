function mainFunction() {
  DataProcess();
  // var DateName=DataProcess();//存取當下新增的人名與時間
  // DataCheck(DateName);//檢查人名與時間，送入資料格式：（月份/日期,人名）
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
  // var summary = 0;
  // var summary = new Array(DataString); // 回傳
  // return summary; 
  
}


// //在"簽到表"填入適當的內容
// function DataCheck(input) {//input data type is array [date,name] -> [月份/日期, 人名]
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = ss.getSheetByName("簽到表");
//   var lastrow = sheet.getLastRow();
//   var lastcol = sheet.getLastColumn();
//   // 取得 姓名那行，作為檢查姓名是否符合的檢查表
//   var checkName = sheet.getRange(1, 1, lastrow, 1).getValues();//(start row, start col, row num, col num)
//   checkName = twoD_to_oneD (checkName) //資料整理，把資料整理成可以做 indexOf的形式
//   //取得日期列（自己訂的日期），照報到的日期是否與自己預設的日期吻合
//   //[0]用來獲取一維陣列，getValues()為二維陣列
//   var checkDate = sheet.getRange(1, 1,1, lastcol).getValues()[0];//(start row, start col, row num, col num)
//   var rowIndex = checkName.indexOf(input[1]);//相對應的姓名當作列指標
//   var colIndex = checkDate.indexOf(input[0]);//相對應的時間當作行指標
//   if(rowIndex != -1 && colIndex != -1 ){//如果日期與時間對得上，就在 該人該日期打勾
//     sheet.getRange(rowIndex+1, colIndex+1).setValue("V");//表單row, col 起始為1，但index起始為0
//   }
// }

// function twoD_to_oneD (Array){//transform 2D array to 1D；2D is [ [2], [1], [3] ], 3x1
//   var OneD =[];
//   for (var i=0; i<Array.length; i++){
//     OneD.push(Array[i][0]);
//   }
//   return OneD; 
// }
  