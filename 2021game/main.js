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
const attackEffect = new Map([
                   ['倍增卡', 2]
    ]);
const defenseEffect = new Map([
                    ['羅馬大盾', 0.95], ['振金之盾', 0.5]
    ]);
const sheetByName = {
                  爵士: ['大會', 1],
                  交易市集: ['大會', 4],
                  白傑睿一: ['白傑睿', 1],
                  白傑睿二: ['白傑睿', 4],
                  白傑睿三: ['白傑睿', 7],
                  白傑睿四: ['白傑睿', 10],
                  紅麥克一: ['紅麥克', 1],
                  紅麥克二: ['紅麥克', 4],
                  紅麥克三: ['紅麥克', 7],
                  紅麥克四: ['紅麥克', 10]
    }; // 後面的數字為在分頁中的起始欄位
// const sheetByName = new Map([
//                 ['爵士', '大會'],
//                 ['白傑睿一', '白傑睿'],
//                 ['白傑睿二', '白傑睿'],
//                 ['白傑睿三', '白傑睿'],
//                 ['白傑睿四', '白傑睿'],
//                 ['紅麥克一', '紅麥克'],
//                 ['紅麥克二', '紅麥克'],
//                 ['紅麥克三', '紅麥克'],
//                 ['紅麥克四', '紅麥克']
//     ]);

const outputUrl = {
                白傑睿一: 'https://docs.google.com/spreadsheets/d/1LO1aZn3ne-x89BLezNOugfnl2vNuTmpg3_nuHSWjb0E/edit#gid=0',
                白傑睿二: 'https://docs.google.com/spreadsheets/d/1v1rkI1yipurRm3ibLiWgPG-e32Hw_dFM4ueAdUgMEZ8/edit#gid=0',
                白傑睿三: 'https://docs.google.com/spreadsheets/d/1c5x1ne7TOuGhQa9E2voWbAR80edtPdIZzrdXV20Pcn4/edit#gid=0',
                白傑睿四: 'https://docs.google.com/spreadsheets/d/1vfNMHeqdQxhVN1Su_d6YhAX26iCshGCMkRwEuwlMWOs/edit#gid=0',
                紅麥克一: 'https://docs.google.com/spreadsheets/d/1hM-DdBPyMAxTfHHLslt28B-2sKE_Smk1mSUiVLemOKs/edit#gid=0',
                紅麥克二: 'https://docs.google.com/spreadsheets/d/1PUasKvXIZKKPtGwx2vRwNDHZRKzR5wk-Z61lkEegwHg/edit#gid=0',
                紅麥克三: 'https://docs.google.com/spreadsheets/d/1zE_y6txsYE8hU6fxfrnHp8IQANWAgrM6rO2B5cJ1BRQ/edit#gid=0',
                紅麥克四: 'https://docs.google.com/spreadsheets/d/13KVHMb6Us2H84LQV8b1zv_XxAUJm8FDbVnymDZB-rqw/edit#gid=0'
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
        tradeFunction(ss, rawData);
    }
    else{
        actFunction(ss, rawData);
    }

    var result = ss.getSheetByName("大會");
    var temp = ss.getSheetByName("白傑睿");
    var tempContent = temp.getRange(1, 1, 2, 11).getValues();
    result.getRange(1, 1, 2, 11).setValues(tempContent);
    temp = ss.getSheetByName("紅麥克");
    tempContent = temp.getRange(1, 1, 2, 11).getValues();
    result.getRange(3, 1, 2, 11).setValues(tempContent);
}

function tradeFunction(ss, tradeRec){
    var seller = tradeRec[2];
    var sIndex = sheetByName[seller];
    var buyer = tradeRec[7];
    var bIndex = sheetByName[buyer];
    if (seller == buyer){
        return false;
    }
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
    if (bIndex[0] == '大會'){
        bContent[1][1] = 10000000;
        bContent[2][1] = 0;
    }
    if (sIndex[0] == '大會'){
        sContent[1][1] = 10000000;
        sContent[2][1] = 0;
    }
    // 買賣積分更新
    var sTotal = buyPoint + sContent[1][1] - salePoint;
    sContent[1][1] = sTotal;
    var bTotal = salePoint + bContent[1][1] - buyPoint;
    bContent[1][1] = bTotal;
    // 判斷交易失敗
    if (bTotal < 0 || sTotal<0){return false;}

    // 行動數更新
    var sAction = sContent[2][1] + 1; // 行動數加一
    sContent[2][1] = sAction;
    var bAction = bContent[2][1] + 1; // 行動數加一
    bContent[2][1] = bAction;

    if (sIndex[0] !== '大會'){
        // 回填回表單
        sSheet.getRange(1, sIndex[1], 3, 2).setValues(sContent);
        // 填入行動
        let timeString = formatDate(tradeRec[0]);
        let sActString = '賣' + items + '給' + buyer;
        sSheet.getRange(3+sAction, sIndex[1], 1, 2).setValues([[timeString, sActString]]);
    }

    if (bIndex[0] !== '大會'){
        // 回填回表單
        bSheet.getRange(1, bIndex[1], 3, 2).setValues(bContent);
        // 填入行動
        let timeString = formatDate(tradeRec[0]);
        let bActString = '向' + seller + '買入' + items; 
        if (items==''){
            if (salePoint >= buyPoint) { bActString = '獲得' + salePoint + '分';}
            else {bActString = '扣' + buyPoint + '分';}
            
        }
        bSheet.getRange(3+bAction, bIndex[1], 1, 2).setValues([[timeString, bActString]]);
    }

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
        let tempContent = bSheet.getRange(1, bIndex[1], 3+bAction, 2).getValues();        
        tempSheet.clear();
        tempSheet.getRange(1, 1, 3+bAction, 2).setValues(tempContent);
    }
}

function actFunction(ss, actRec){
    var attacker = actRec[9];
    var aIndex = sheetByName[attacker];
    var defender = actRec[13];
    var dIndex = sheetByName[defender];
    if (attacker == defender){
        return false;
    }
    var flag = false;
    if (actRec[17] == '是'){
        flag = true;
    }
    var aCards = ''; 
    for (var i = 10; i <= 12; i++) {
        if (actRec[i]) {
            aCards = aCards + actRec[i] + ',';
        }
    }
    var dCards = ''; 
    for (var i = 14; i <= 16; i++) {
        if (actRec[i]) {
            dCards = dCards + actRec[i] + ',';
        }
    }
    var getPoints = 0;
    var hitPoints = 0;
    [getPoints, hitPoints] = actPoint(aCards, dCards, flag);

// spreadSheet, teamIndex, points, actString, timeString
    // 填入行動
    var timeString = formatDate(actRec[0]);
    var actString = attacker + ' 向 ' + defender + ' \n發動 ' + aCards;
    if (flag) {
        actString = '成功：' + actString;
    }
    else{
        actString = '失敗：' + actString;        
    }

    var actionNums = updateACT(ss, aIndex, getPoints, actString, timeString);
    actionNums = updateACT(ss, dIndex, hitPoints, actString, timeString);

    // 更新輸出表單
    // 攻方
    let tempUrl = outputUrl[attacker];
    let tag = '積分 動作';
    let tempSS = SpreadsheetApp.openByUrl(tempUrl);
    let tempSheet = tempSS.getSheetByName(tag);
    let tempContent = ss.getSheetByName(aIndex[0]).getRange(1, aIndex[1], 3+actionNums, 2).getValues();
    tempSheet.clear();
    tempSheet.getRange(1, 1, 3+actionNums, 2).setValues(tempContent);

    // 守方
    tempUrl = outputUrl[defender];
    tag = '積分 動作';
    tempSS = SpreadsheetApp.openByUrl(tempUrl);
    tempSheet = tempSS.getSheetByName(tag);
    tempContent = ss.getSheetByName(dIndex[0]).getRange(1, dIndex[1], 3+actionNums, 2).getValues();
    tempSheet.clear();
    tempSheet.getRange(1, 1, 3+actionNums, 2).setValues(tempContent);

}

// activeEffectT, activeEffectF
// passiveEffectT, passiveEffectF
function actPoint(attacks, defenses, flag){
    var atkList = attacks.split(',');
    var defList = defenses.split(',');
    var earn = 0;
    var hit = 0;
    var defFactor = 1;
    var atkFactor = 1;
    var card ='';
    for (var i in defList) {
        card = defList[i].trim();
        if (defenseEffect.has(card)) {
            defFactor = defFactor * defenseEffect.get(card);
        }
    }
    for (var i in atkList) {
        card = atkList[i].trim();
        if (attackEffect.has(card)) {
            atkFactor = atkFactor * attackEffect.get(card);
        }
    }
    if (flag) {
        for (var i in atkList) {
            card = atkList[i].trim();
            if (activeEffectT.has(card)) {
                earn = earn + activeEffectT.get(card);
            }
            if (passiveEffectT.has(card)) {
                hit = hit + passiveEffectT.get(card);
            }
        }
    }
    else{
        for (var i in atkList) {
            card = atkList[i].trim();
            if (activeEffectF.has(card)) {
                earn = earn + activeEffectF.get(card);
                hit = hit + passiveEffectF.get(card);
            }
        }        
    }
    hit = hit * defFactor * atkFactor;
    return [earn, hit];
}

function updateACT(spreadSheet, teamIndex, points, actString, timeString){
    var updateSheet = spreadSheet.getSheetByName(teamIndex[0]);
    // 獲取表格資料
    // content = [[ '名稱', '大會' ], [ '積分', 0 ], [ '行動數', 0 ]]
    var content = updateSheet.getRange(1, teamIndex[1], 3, 2).getValues();
    // 買賣積分與行動數更新
    var total = content[1][1] + points;
    var actions = content[2][1] + 1; // 行動數加一
    content[1][1] = total;
    content[2][1] = actions;
    // 回填回表單
    updateSheet.getRange(1, teamIndex[1], 3, 2).setValues(content);
    // 填入行動
    updateSheet.getRange(3+actions, teamIndex[1], 1, 2).setValues([[timeString, actString]]);
    return actions;
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
