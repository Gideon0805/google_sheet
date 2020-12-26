//在陣列運算時，起始值為 0
//在填入或取出表格資料時，起始值為1

function getSummary() {
  checkSummary();
}


function checkSummary() {
  // 存取工作表(ByName) https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheetbynamename
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 欲統計名單表格
  const nameSheet = ss.getSheetByName("名單");
  const lastRow = nameSheet.getLastRow();
  const lastCol = nameSheet.getLastColumn();

  const rawData = nameSheet.getRange(2, 1, lastRow-1, lastCol-1).getValues();
  // titleRow 來找陣列的index;
  const titleRow = nameSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const gradeIndex = titleRow.indexOf('年級');
  const genderIndex = titleRow.indexOf('性別');
  const remitIndex = titleRow.indexOf('已繳費');
  const gradeCol = nameSheet.getRange(2, gradeIndex+1, lastRow-1, 1).getValues();
  const remitCol = nameSheet.getRange(2, remitIndex+1, lastRow-1, 1).getValues();
  const set = new Set();
  var tempContent;
  var content = [];
  tempContent = ['國小ABC總動員人數統計（含桂冠）','','',''];
  content.push(tempContent);
  // 取出已繳費民數
  const remitResult = remitCol.filter(item => item[0] != "" 
                                              & item[0] != " " 
                                              & item[0] != "  ");

  tempContent = ['已繳費', remitResult.length, 
                 '未繳費', rawData.length - remitResult.length];
  content.push(tempContent);

  tempContent = ['年級', '男', '女', '共計'];
  content.push(tempContent);  
  // 取出要統計的年級
  const gradeResult = gradeCol.filter(item => !set.has(item[0]) ? set.add(item[0]) : false);
  var maleSum = 0, femaleSum = 0;
  var maleTemp, femaleTemp;
  var maleNum = 0, femaleNum = 0;
  for (var i=0; i < gradeResult.length; i++) {
    maleTemp = rawData.filter(item => item[gradeIndex] == gradeResult[i] 
                                      & item[genderIndex] == '男');
    femaleTemp = rawData.filter(item => item[gradeIndex] == gradeResult[i] 
                                      & item[genderIndex] == '女');
    /*
    3年級 男自動+1 女+1 / 4年級 男+1 / 5年級 男+8 女+3 / 6年級 男+6 女+4
    */
    switch (gradeResult[i][0]) {
      case 3:
        maleNum = maleTemp.length + 1;
        femaleNum = femaleTemp.length + 1;
        break;
      case 4:
        maleNum = maleTemp.length + 1;
        femaleNum = femaleTemp.length;
        break;
      case 5:
        maleNum = maleTemp.length + 8;
        femaleNum = femaleTemp.length + 3;
        break;
      case 6:
        maleNum = maleTemp.length + 6;
        femaleNum = femaleTemp.length + 4;
        break;
      default:
        maleNum = maleTemp.length;
        femaleNum = femaleTemp.length
        break;
    }
    tempContent = [gradeResult[i], maleNum, femaleNum, maleNum+femaleNum];
    content.push(tempContent);
    maleSum = maleSum + maleNum;
    femaleSum = femaleSum + femaleNum;
  }
  tempContent = ['共計', maleSum, femaleSum, maleSum+femaleSum];
  content.push(tempContent);

  // 欲統計表單
  var summarySheet = ss.getSheetByName("統計");
  // 清空分頁, 列
  summarySheet.clear();
  // summarySheet.deleteRows(2, content.length);
  // Get range
  var sheetRange = summarySheet.getRange(1, 1, content.length, content[0].length);
  // 填入 content
  sheetRange.setValues(content);
  // Border
  sheetRange.setBorder(true, true, true, true, true, true);//top, left, bottom, right, vertical, horizontal
  // 置中
  sheetRange.setHorizontalAlignment("center");
  // 填入顏色
  summarySheet.getRange(1, 1, 1, content[0].length).setBackground('#FFD000');// 第一列
  summarySheet.getRange(content.length, content[0].length).setBackground('#FFFB00');// 共計格
  // 合併第一列
  summarySheet.getRange(1, 1, 1, content[0].length).merge();


}
