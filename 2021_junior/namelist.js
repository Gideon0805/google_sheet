//在陣列運算時，起始值為 0
//在填入或取出表格資料時，起始值為1

// 重新產生名單
function getNameList() {
  checkNameList();
}

// // 提交時更新
// function updateNameList() {
//   checkLast();
// }

// debug:
// console.log(varname)

function checkNameList() {
  // 存取工作表(ByName) https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheetbynamename
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  //獲取“報到”的最新一筆資料
  const sheet = ss.getSheetByName("表單回應"); 
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const titleName = '2021全員逃走FUN寒假_報名資料';
  var tempRow;
  // var rawData = sheet.getRange(lastRow, 1,1, lastCol).getValues()[0];//(start row, start col, row num, col num)//二維矩陣 取[0] 轉一維  

  // ==== 刪掉有重複的名字
  const rawData = sheet.getRange(2, 2, lastRow-1, lastCol-1).getValues()
  const nonRepeat = delRepeat(rawData);
  // var repIndex = getRepeat(namCol);

  // 欲填入名單表格
  var nameSheet = ss.getSheetByName("名單");
  // 清空分頁
  nameSheet.clear();
  // 填入Title
  var colTitle = sheet.getRange(1, 2, 1, lastCol-1).getValues();
  nameSheet.getRange(1, 1, 1, lastCol-1).setValues(colTitle);
  // 整理男女
  tempRow = nameSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var genderIndex = tempRow.indexOf('性別');
  // console.log(tempRow[genderIndex]);
  var male = nonRepeat.filter(function(item){
    return item[genderIndex] == '男'
  });
  var female = nonRepeat.filter(function(item){
    return item[genderIndex] == '女'
  });
  
  // === 填入無重複且整理完的男女名單 ===
  // nameSheet.getRange(2, 1, nonRepeat.length, lastCol-1).setValues(nonRepeat);
  nameSheet.getRange(2, 1, male.length, lastCol-1).setValues(male);
  nameSheet.getRange(male.length+2, 1, female.length, lastCol-1).setValues(female);



  // ==== 將班級轉換成年級，插入名單
  // 班級欄位
  // nameSheet: col:8
  // 找出年級的欄位 index
  tempRow = nameSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var classIndex = tempRow.indexOf('年級') + 1; // array index begin is 0;
  var classCol = nameSheet.getRange(2, classIndex, nonRepeat.length, 1).getValues();//(start row, start col, row num, col num)
  var gradeCol = classCol;
  // 班級轉年級
  for (var i = 0; i < classCol.length; i++) {
      if (classCol[i][0]>100) {gradeCol[i][0] = Math.round(classCol[i][0]/100)} 
      else {gradeCol[i][0] = classCol[i][0]}
  }
  // 插入年級col在第8欄前
  nameSheet.insertColumns(classIndex);
  nameSheet.getRange(2, classIndex, nonRepeat.length, 1).setValues(gradeCol);
  nameSheet.getRange(1, classIndex).setValue("年級");
  nameSheet.getRange(1, classIndex+1).setValue("班級");

  // 整理年級
  var lastNameRow = nameSheet.getLastRow();
  var lastNameCol = nameSheet.getLastColumn();

  var temp = nameSheet.getRange(2, 1, lastNameRow-1, lastNameCol).getValues();
  temp.sort(sortFunction); 
  // === 填入整理過年級的名單 ===
  nameSheet.getRange(2, 1, lastNameRow-1, lastNameCol).setValues(temp);

  // 插入排序編號
  // gradeCol = nameSheet.getRange(2, 8, lastNameRow-1, 1).getValues();
  var insertCol = 1;
  nameSheet.insertColumns(insertCol);
  nameSheet.getRange(1, insertCol).setValue("序號");
  // numbering_ref(nameSheet, gradeCol, 1);
  numbering(nameSheet, insertCol);

  const endCol = nameSheet.getLastColumn();
  // 插入標頭
  nameSheet.insertRows(1);
  nameSheet.getRange(1, 1).setValue(titleName);
  nameSheet.getRange(1, 1).setHorizontalAlignment("center");
  // Border
  nameSheet.getRange(1, 1, 1, endCol).setBorder(true, true, true, true, null, null);//top, left, bottom, right, vertical, horizontal
  // 填入顏色
  nameSheet.getRange(1, 1, 1, endCol).setBackground('#94EDFF');// 第一列
  // 合併第一列
  nameSheet.getRange(1, 1, 1, endCol).merge();


}
/*
function checkLast() {
  // 存取工作表(ByName) https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheetbynamename
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //獲取“報到”的最新一筆資料
  var sheet = ss.getSheetByName("表單回應"); 
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var latestData = sheet.getRange(lastRow, 2, 1, lastCol).getValues()[0];//(start row, start col, row num, col num)//二維矩陣 取[0] 轉一維
  
  console.log(latestData);
  // 欲填入名單表格
  var nameSheet = ss.getSheetByName("名單");
  var rowNums = nameSheet.getLastRow();
  var colNums = nameSheet.getLastColumn();
  var nameList = nameSheet.getRange(2, 1, rowNums-1, lastCol).getValues();
  const titleName = '2021全員逃走FUN寒假_報名資料';

  // ==== 判斷是否存在
  var insert = true;
  // 用姓名及生日判斷
  nameList.forEach(function(item){
    if (item.includes(latestData[0]) || item.includes(latestData[3]))
      insert = false;
  });
  console.log(insert)

  // ==== 將班級轉換成年級，插入名單
  // 找出年級的欄位 index
  var tempRow = nameSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var index = tempRow.indexOf('年級');
  // 班級轉年級
  var tempGrade = latestData[index]>100 ? Math.round(latestData[index]/100) : latestData[index];
  // 插入年級到 index 8
  latestData.splice(index, 0, tempGrade);

  // 判斷數量是否一致
  if (latestData.length > colNums) {
    // delete 多餘值
    var tooMany = latestData.length - colNums;
    latestData.splice(colNums, tooMany);
  }
  if (latestData.length < colNums) {
    // 補值
    var howMany = colNums - latestData.length;
    var addIndex = latestData.length
    for (var i = 0; i < howMany; i++) {
      latestData.splice(addIndex+i, 1, '');
    }
  }
  // console.log(latestData);
  // console.log(colNums);
  // console.log(latestData.length);
  // 插入新一列，並填值
  if (insert) nameSheet.getRange(rowNums+1, 1, 1, colNums).setValues([latestData]);

  // 整理年級
  var lastNameRow = nameSheet.getLastRow();
  var lastNameCol = nameSheet.getLastColumn();
  var temp = nameSheet.getRange(2, 1, lastNameRow-1, lastNameCol).getValues();
  temp.sort(sortFunction); 
  nameSheet.getRange(2, 1, lastNameRow-1, lastNameCol).setValues(temp);

  // 插入排序編號
  // gradeCol = nameSheet.getRange(2, 8, lastNameRow-1, 1).getValues();
  var insertCol = 1;
  // 已有序號，不用插入
  // nameSheet.insertColumns(insertCol);
  // nameSheet.getRange(1, insertCol).setValue("序號");
  // numbering_ref(nameSheet, gradeCol, 1);
  numbering(nameSheet, insertCol)
  
}
*/

function delRepeat(sheetValues) {
  var set = new Set();
  // 根據姓名欄 刪除重複
  var result = sheetValues.filter(item => !set.has(item[0]) ? set.add(item[0]) : false);
  // console.log(result);
  // console.log(sheetValues);
  return result;
}

function getRepeat(sheetValues) {
    var set = new Set();
    var result = sheetValues.filter(item => {
      if (set.has(item[4])) {
        return true;
      } else {
        set.add(item[4]);
        console.log(item[4]);
      }
    });

    // var repeat = sheetValues.filter(item => set.has(JSON.stringify(item)) ? 
    //                                   true : (set.add(JSON.stringify(item)), false));
    // console.log(result);
    return result;
}

// sort 2D array
function sortFunction(a, b) {
  // comapre by first col
  if (a[7] === b[7]) {
      return 0;
  }
  else {
      return (a[7] < b[7]) ? -1 : 1;
  }
}

// 普通編號
function numbering(inSheet, colIndex) {
  var rowIndex = 2;
  var j = inSheet.getLastRow();
  for (var i=1; i < j; i++) {
    inSheet.getRange( i+1, colIndex).setValue(i);
  }
}

// 在inSheet 中，根據 refCol 在 index欄 排序號
// 例如 refCol 是年級，就會排序出每個年級個別的編號
function numbering_ref(inSheet, refCol, index) {
  var numCount = 1
  var rowIndex = 2
  console.log(refCol);
  for (var i=0; i < refCol.length-1; i++) {
    if (refCol[i][0] == refCol[i+1][0]) {
      inSheet.getRange(rowIndex, index).setValue(numCount);
      numCount ++;
    }
    else {
      inSheet.getRange(rowIndex, index).setValue(numCount);
      numCount = 1;
    }
    rowIndex ++;
  }
  if (refCol[refCol.length-1][0] == refCol[refCol.length-2][0]) inSheet.getRange(rowIndex, index).setValue(numCount);
  else inSheet.getRange(rowIndex, index).setValue(1);
}

function isExist(sheetValues, value) {
  return false;
}

function twoD_to_oneD (Array){//transform 2D array to 1D；2D is [ [2], [1], [3] ], 3x1
  var OneD =[];
  for (var i=0; i<Array.length; i++){
    OneD.push(Array[i][0]);
  }
  return OneD; 
}
