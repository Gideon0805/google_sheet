
// 發動方效果成功
const activeEffectT = new Map([
                    ['流氓卡', 800], ['肌肉人卡', 800],
                    ['辛苦站立的侍衛卡', 400], ['辛勤耕田的農夫卡', 400],
                    ['豺狼的圍攻卡', 200], ['積分獵手卡', 400],
                    ['面罩糾察卡', 400], ['勝利之劍', 50]
    ]);
// 被動方效果成功
const passiveEffectT = new Map([
                    ['流氓卡', -200], ['肌肉人卡', -200],
                    ['辛苦站立的侍衛卡', -300], ['辛勤耕田的農夫卡', -300],
                    ['豺狼的圍攻卡', -200], ['積分獵手卡', -300],
                    ['面罩糾察卡', -300], ['逆刃刀', -50]
    ]);
// 發動方效果失敗
const activeEffectF = new Map([
                    ['流氓卡', 500], ['肌肉人卡', 500],
                    ['辛苦站立的侍衛卡', 200], ['辛勤耕田的農夫卡', 200],
                    ['積分獵手卡', 200], ['面罩糾察卡', 200]
    ]);
// 被動方效果失敗
const passiveEffectF = new Map([
                    ['流氓卡', 0], ['肌肉人卡', 0],
                    ['辛苦站立的侍衛卡', 0], ['辛勤耕田的農夫卡', 0],
                    ['積分獵手卡', 0], ['面罩糾察卡', 0]
    ]);
const comboEffect = new Map([
                    ['倍增卡', 2], ['羅馬大盾', 0.95],
                    ['振金之盾', 0.5]
    ]);
const sheetByName = {
                  爵士: ['大會', 1],
                  交易市集: ['大會', 4],
                  白傑睿第一分隊: ['白傑睿', 1],
                  白傑睿第二分隊: ['白傑睿', 4],
                  白傑睿第三分隊: ['白傑睿', 7],
                  白傑睿第四分隊: ['白傑睿', 10],
                  紅麥克第一分隊: ['紅麥克', 1],
                  紅麥克第二分隊: ['紅麥克', 4],
                  紅麥克第三分隊: ['紅麥克', 7],
                  紅麥克第四分隊: ['紅麥克', 10]
    }; // 後面的數字為在分頁中的起始欄位
// const sheetByName = new Map([
//                 ['爵士', '大會'],
//                 ['白傑睿第一分隊', '白傑睿'],
//                 ['白傑睿第二分隊', '白傑睿'],
//                 ['白傑睿第三分隊', '白傑睿'],
//                 ['白傑睿第四分隊', '白傑睿'],
//                 ['紅麥克第一分隊', '紅麥克'],
//                 ['紅麥克第二分隊', '紅麥克'],
//                 ['紅麥克第三分隊', '紅麥克'],
//                 ['紅麥克第四分隊', '紅麥克']
//     ]);

const outputUrl = {
                白傑睿第一分隊: 'https://docs.google.com/spreadsheets/d/1LO1aZn3ne-x89BLezNOugfnl2vNuTmpg3_nuHSWjb0E/edit#gid=0',
                白傑睿第二分隊: 'https://docs.google.com/spreadsheets/d/1v1rkI1yipurRm3ibLiWgPG-e32Hw_dFM4ueAdUgMEZ8/edit#gid=0',
                白傑睿第三分隊: 'https://docs.google.com/spreadsheets/d/1c5x1ne7TOuGhQa9E2voWbAR80edtPdIZzrdXV20Pcn4/edit#gid=0',
                白傑睿第四分隊: 'https://docs.google.com/spreadsheets/d/1vfNMHeqdQxhVN1Su_d6YhAX26iCshGCMkRwEuwlMWOs/edit#gid=0',
                紅麥克第一分隊: 'https://docs.google.com/spreadsheets/d/1hM-DdBPyMAxTfHHLslt28B-2sKE_Smk1mSUiVLemOKs/edit#gid=0',
                紅麥克第二分隊: 'https://docs.google.com/spreadsheets/d/1PUasKvXIZKKPtGwx2vRwNDHZRKzR5wk-Z61lkEegwHg/edit#gid=0',
                紅麥克第三分隊: 'https://docs.google.com/spreadsheets/d/1zE_y6txsYE8hU6fxfrnHp8IQANWAgrM6rO2B5cJ1BRQ/edit#gid=0',
                紅麥克第四分隊: 'https://docs.google.com/spreadsheets/d/13KVHMb6Us2H84LQV8b1zv_XxAUJm8FDbVnymDZB-rqw/edit#gid=0'
    };


function mainFunction() {
  // 存取工作表(ByName) https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getsheetbynamename
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //獲取“報到”的最新一筆資料
  var sheet = ss.getSheetByName("表單回應 1"); 
  var lastrow = sheet.getLastRow();
  var lastcol = sheet.getLastColumn();
  var rawData = sheet.getRange(lastrow, 1,1, lastcol).getValues()[0];//(start row, start col, row num, col num)//二維矩陣 取[0] 轉一維
  // console.log(rawData)
  if (rawData[1]=='交易') {
    var updateContent = tradeFun(ss, rawData);
  }
}

function DataProcess() {
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

function tradeFun(ss, tradeRec){
    var seller = tradeRec[2];
    var sIndex = sheetByName[seller];
    var buyer = tradeRec[7];
    var bIndex = sheetByName[buyer];
    var salePoint = tradeRec[6];
    var buyPoint = tradeRec[8]
    var items = ''; 
    for (var i =  3; i <= 5; i++) {
        if (tradeRec[i]) {
            items = items + tradeRec[i] + ',';
        }
    }
    var sSheet = ss.getSheetByName(sIndex[0]);
    var bSheet = ss.getSheetByName(bIndex[0]);
    // 獲取買賣方表格資料 s: Seller, b: Buyer
    // content = [[ '名稱', '大會' ], [ '積分', 0 ], [ '行動數', 0 ]]
    var sContent = sSheet.getRange(1, sIndex[1], 3, 2).getValues();
    var bContent = bSheet.getRange(1, bIndex[1], 3, 2).getValues();
    // 買賣積分與行動數更新
    var sTotal = buyPoint + sContent[1][1] - salePoint;
    sContent[1][1] = sTotal;
    var bTotal = salePoint + bContent[1][1] - buyPoint;
    bContent[1][1] = bTotal;
    var sAction = sContent[2][1] + 1; // 行動數加一
    var bAction = bContent[2][1] + 1; // 行動數加一
    sContent[2][1] = sAction;
    bContent[2][1] = bAction;
    // 回填回表單
    sSheet.getRange(1, sIndex[1], 3, 2).setValues(sContent);
    bSheet.getRange(1, bIndex[1], 3, 2).setValues(bContent);
    // 填入行動
    var timeString = formatDate(tradeRec[0]);
    var sActString = '賣' + items + '給' + buyer;
    sSheet.getRange(3+sAction, sIndex[1], 1, 2).setValues([[timeString, sActString]]);
    var bActString = '向' + seller + '買入' + items; 
    bSheet.getRange(3+bAction, bIndex[1], 1, 2).setValues([[timeString, bActString]]);

    // 更新輸出表單
    // 賣方
    if (sIndex[0] !== '大會'){
        let tempUrl = outputUrl[seller];
        let tag = '積分 動作';
        let tempSS = SpreadsheetApp.openByUrl(tempUrl);
        let tempSheet = tempSS.getSheetByName(tag);
        let tempContent = sSheet.getRange(1, sIndex[1], 3+sAction, 2).getValues();
        tempSheet.clear();
        tempSheet.getRange(1, 1, 3+sAction, 2).setValues(tempContent);
    }
    // 買方
    if (bIndex[0] !== '大會'){
        let tempUrl = outputUrl[buyer];
        let tag = '積分 動作';
        let tempSS = SpreadsheetApp.openByUrl(tempUrl);
        let tempSheet = tempSS.getSheetByName(tag);
        let tempContent = bSheet.getRange(1, bIndex[1], 3+sAction, 2).getValues();        
        tempSheet.clear();
        tempSheet.getRange(1, 1, 3+bAction, 2).setValues(tempContent);
    }


}

function formatDate(now) { 
// 　　var year=now.getYear(); 
　　var month=now.getMonth()+1; 
　　var date=now.getDate(); 
　　var hour=now.getUTCHours()+8; 
　　var minute=now.getMinutes(); 
　　var second=now.getSeconds(); 
　　return month+"/"+date+" "+hour+":"+minute+":"+second; 
} 

/*
'流氓卡'
'肌肉人卡'
'至尊神卡'

'辛苦站立的侍衛卡'
'辛勤耕田的農夫卡'
'豺狼的圍攻卡'
'積分獵手卡'
'面罩糾察卡'

'逆刃刀'
'勝利之劍'
'羅馬大盾'
'振金之盾'



流氓卡
肌肉人卡
強盜卡
狂暴卡
至尊神卡
金牌特務卡
守護者卡

下課鐘聲卡
傳奇三連抽
束脩卡
時間暫停卡
增值卡
辛苦站立的侍衛卡
辛勤耕田的農夫卡
絕對先手卡
豺狼的圍攻卡
隱形卡
田裡的稻草人卡
浴火重生卡
積分獵手卡
面罩糾察卡

逆刃刀
勝利之劍
戰術求生斧
羅馬大盾
振金之盾*/