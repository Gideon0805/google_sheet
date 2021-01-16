//在陣列運算時，起始值為 0
//在填入或取出表格資料時，起始值為1

// 產生繳費名單
function getRemittance() {
  checkRemittance();
}

// debug:
// console.log(varname)

function checkRemittance() {
  // 存取工作表(ByName) https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheetbynamename
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //獲取“報到”的最新一筆資料
  const sheet = ss.getSheetByName("名單"); 
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const startRow = 3;
  const titleName = '110年1/25-26 國中全員逃走FUN寒假_保險';

  const rawData = sheet.getRange(startRow, 2, lastRow-1, lastCol).getValues();//(start row, start col, row num, col num)//二維矩陣 取[0] 轉一維
  //2019/6/10 下午 8:54:29
  //var date = getDate("2019/6/10 下午 8:54:29");
  
  // 整理繳費名單
  // var remittResult = rawData.filter(isRemittance);
  var tempRow = sheet.getRange(startRow-1, 1, 1, lastCol).getValues()[0];
  const remitIndes = tempRow.indexOf('已繳費');
  var remittResult = rawData.filter(item => {
    if (item[remitIndes] != "" & item[remitIndes] != " " & item[remitIndes] != "  ") 
      return true;
  });
  // 欲填入保險表格
  var insurance = ss.getSheetByName("保險");

  // 獲取已繳費的index
  // col:18 is col:R (繳費欄位)
  // var Pay = sheet.getRange(2, 18, lastRow, 1).getValues(); //(start row, start col, row num, col num)
  // var Index = getRemitIndex(Pay);

  // 保險欄位
  const dataNums = 6;
  const dataStart = 1;
  // 清空保險分頁資料
  insurance.clear();
  // 欄colTitle位名稱
  var colTitle = sheet.getRange(startRow-1, 2, 1, lastCol).getValues();
  insurance.getRange(1, 1, 1, lastCol).setValues(colTitle);
  // 填入繳費資料
  insurance.getRange(2, 1, remittResult.length, lastCol).setValues(remittResult);
  // 刪除多餘欄位 G~Q 
  insurance.deleteColumns(7, 11); //deleteColumns(columnPosition, howMany) or no howMany
  insurance.insertColumns(1);
  insurance.getRange(1, 1).setValue("編號");
  // insurance.getRange(1, dataNums+2).setValue("已繳費");
  // 填入資料的起始列
  // var inRow = insurance.getLastRow() + 1;

  for (var i=0; i < remittResult.length; i++) {
    // 編號欄，填入編號
    insurance.getRange(2+i, 1).setValue(1+i);
  }

  console.log(colTitle);
  const validRange = colTitle[0].indexOf('緊急聯絡人電話') + 2;
  // 插入標頭
  insurance.insertRows(1);
  insurance.getRange(1, 1).setValue(titleName);
  insurance.getRange(1, 1).setHorizontalAlignment("center");
  // Border
  insurance.getRange(1, 1, remittResult.length+2, validRange).setBorder(true, true, true, true, true, true);//top, left, bottom, right, vertical, horizontal
  // 填入顏色
  insurance.getRange(1, 1, 1, validRange).setBackground('#94EDFF');// 第一列
  // 合併第一列
  insurance.getRange(1, 1, 1, validRange).merge();

  // for (var i = 0; i < Index.length; i++) {
  //   // Index[i]
  //   // 照 Index[i] 取得有繳費的欄位
  //   var tempRow = sheet.getRange(Index[i]+1, dataStart, 1, dataNums).getValues();
  //   // console.log(tempRow);
  //   //編號欄，填入編號
  //   insurance.getRange(inRow+i, 1).setValue(i+1);    
  //   //保險欄
  //   insurance.getRange(inRow+i, 2, 1, dataNums).setValues(tempRow);
  //   //繳費欄
  //   var tempPay = sheet.getRange(Index[i]+1, 18, 1, 1).getValues();
  //   // console.log(tempPay);
  //   insurance.getRange(inRow+i, dataNums+2).setValue(tempPay);
  // }
  
}

function isRemittance(sheetValue) {
  // 起始欄為0，第 18 欄為繳費欄(17), 第 19 欄為繳費方式欄(18)
  // 不為空就判定為已繳費
  // console.log(sheetValue);
  if (sheetValue[17] != "" & sheetValue[17] != " " & sheetValue[17] != "  ") 
    return true;
  // if (sheetValue[18] != "" & sheetValue[18] != " " & sheetValue[18] != "  ") 
  //   return true;
}

function getRemitIndex(PayCol) {
  var indexList = [];
  for (var i = 0; i < PayCol.length; i++) {
    if (PayCol[i] != "" & PayCol[i] != " " & PayCol[i] != "  ") {
      indexList.push(i+1);  
    }
  }
  return indexList;
}
