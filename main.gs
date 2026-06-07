function myFunction() {
  
}
/**
 * =========================================================================
 * デイサービス専用 AI自動配車＆マスタ自動化フルコンプシステム
 * =========================================================================
 */

// ==========================================================
// 💡 【共通ヘルパー関数】スペースを消して完全一致でNG乗客をチェック
// ==========================================================
function checkNgPassengerConflict(uId, uInfo, currentRunUserIds, userMap) {
  if (currentRunUserIds.length === 0) return false;
  
  // 比較用に、文字の全角・半角スペースを全て消し去って小文字にする補助関数
  var cleanStr = function(str) {
    if (!str) return "";
    return String(str).replace(/[\s　]/g, "").toLowerCase();
  };

  for (var k = 0; k < currentRunUserIds.length; k++) {
    var exId = currentRunUserIds[k];
    var exInfo = userMap[exId];
    if (!exInfo) continue;
    
    // 1. 新しく乗ろうとしている人のNG乗客チェック
    if (uInfo.ngPassenger && uInfo.ngPassenger !== "") {
      var ngTarget = cleanStr(uInfo.ngPassenger);
      if (cleanStr(exInfo.name) === ngTarget || cleanStr(exId) === ngTarget) return true;
    }
    // 2. 既に車にいる人のNG乗客チェック（双方向ガード）
    if (exInfo.ngPassenger && exInfo.ngPassenger !== "") {
      var ngTarget2 = cleanStr(exInfo.ngPassenger);
      if (cleanStr(uInfo.name) === ngTarget2 || cleanStr(uId) === ngTarget2) return true;
    }
  }
  return false;
}

// 比較用スペース除去の共通関数（メイン処理用）
function cleanStringForMatch(str) {
  if (!str) return "";
  return String(str).replace(/[\s　]/g, "").toLowerCase();
}

// ==========================================================
// 機能1：【マスタ同期】基本マスタから契約・条件シートへ新メンバーを自動追加
// ==========================================================
function syncUserMasters() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var basicSheet = ss.getSheetByName("m_user_basics");
  var contractSheet = ss.getSheetByName("t_user_contracts");
  var conditionSheet = ss.getSheetByName("t_user_pickup_conditions");
  
  if (!basicSheet || !contractSheet || !conditionSheet) {
    Browser.msgBox("エラー", "マスタシートが見つかりません。シート名を確認してください。", Browser.Buttons.OK);
    return;
  }
  
  var basicData = basicSheet.getDataRange().getValues();
  var contractData = contractSheet.getDataRange().getValues();
  var conditionData = conditionSheet.getDataRange().getValues();
  
  var contractIds = new Set();
  for (var i = 2; i < contractData.length; i++) {
    if (contractData[i][0]) contractIds.add(String(contractData[i][0]).trim());
  }
  
  var conditionIds = new Set();
  for (var i = 2; i < conditionData.length; i++) {
    if (conditionData[i][0]) conditionIds.add(String(conditionData[i][0]).trim());
  }
  
  var newContracts = [];
  var newConditions = [];
  
  for (var i = 2; i < basicData.length; i++) {
    var uId = String(basicData[i][0]).trim();
    var uName = basicData[i][1];
    if (!uId || uId === "null" || uId === "") continue;
    
    if (!contractIds.has(uId)) {
      newContracts.push([uId, uName, false, false, false, false, false, false, false]);
    }
    if (!conditionIds.has(uId)) {
      newConditions.push([uId, uName, "一般席", "", ""]);
    }
  }
  
  var msg = "【マスタ同期完了】\n";
  var updated = false;
  if (newContracts.length > 0) {
    contractSheet.getRange(contractSheet.getLastRow() + 1, 1, newContracts.length, 9).setValues(newContracts);
    msg += " ・ 契約マスタに " + newContracts.length + " 名追加しました。\n";
    updated = true;
  }
  if (newConditions.length > 0) {
    conditionSheet.getRange(conditionSheet.getLastRow() + 1, 1, newConditions.length, 5).setValues(newConditions);
    msg += " ・ 送迎条件マスタに " + newConditions.length + " 名追加しました。\n";
    updated = true;
  }
  if (!updated) msg += "新しく追加する利用者はいませんでした。（すべて同期済みです）";
  Browser.msgBox("完了", msg, Browser.Buttons.OK);
}

// ==========================================================
// 機能2：【名簿作成】曜日契約から本日の出席者を上書き展開する
// ==========================================================
function generateDailyAttendance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName("t_daily_attendance");
  var contractSheet = ss.getSheetByName("t_user_contracts");
  
  var targetDate = attendanceSheet.getRange("H1").getValue();
  if (!targetDate || !(targetDate instanceof Date)) {
    Browser.msgBox("エラー", "H1セルに正しい日付を入力してください。", Browser.Buttons.OK);
    return;
  }
  
  var dayOfWeekIdx = targetDate.getDay(); 
  var days = ["日", "月", "火", "水", "木", "金", "土"];
  var targetDayName = days[dayOfWeekIdx];
  
  var lastRow = attendanceSheet.getLastRow();
  if (lastRow >= 3) {
    attendanceSheet.getRange(3, 1, lastRow - 2, 6).clearContent();
  }
  
  var contractData = contractSheet.getDataRange().getValues();
  var weekCols = {"月":2, "火":3, "水":4, "木":5, "金":6, "土":7, "日":8};
  var targetColIdx = weekCols[targetDayName];
  
  var insertValues = [];
  for (var i = 2; i < contractData.length; i++) { 
    var uId = contractData[i][0];   
    var uName = contractData[i][1]; 
    var isTargetDay = contractData[i][targetColIdx]; 
    if (uId && isTargetDay === true) {
      insertValues.push([targetDate, uId, uName, "", "", ""]);
    }
  }
  if (insertValues.length > 0) {
    attendanceSheet.getRange(3, 1, insertValues.length, 6).setValues(insertValues);
    Browser.msgBox("完了", "本日（" + targetDayName + "曜日）の出席者 " + insertValues.length + " 名を展開しました！", Browser.Buttons.OK);
  }
}
// ==========================================================
// 機能3：【往路（朝のお迎え）】AI自動配車（J列マスタ完全連動・最終完成版）
// ==========================================================
function runAiRoutingV7() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName("t_daily_attendance");
  var basicSheet = ss.getSheetByName("m_user_basics");
  var vehicleSheet = ss.getSheetByName("m_vehicles");
  var conditionSheet = ss.getSheetByName("t_user_pickup_conditions"); 
  var outputSheet = ss.getSheetByName("t_routing_outputs");
  
  var MAX_WAIT_MINUTES = 15; // 早く着きすぎた時の最大現地待機時間（15分）
  
  var targetDate = attendanceSheet.getRange("H1").getValue();
  var officeLat = attendanceSheet.getRange("H2").getValue(); 
  var officeLng = attendanceSheet.getRange("H3").getValue(); 
  
  if (!targetDate || !(targetDate instanceof Date) || officeLat === "" || officeLng === "") {
    Browser.msgBox("エラー", "日付または施設の緯度経度を確認してください。", Browser.Buttons.OK);
    return;
  }
  
  var OFFICE_LAT = Number(officeLat); var OFFICE_LNG = Number(officeLng);
  var attendanceData = attendanceSheet.getDataRange().getValues();
  var todayUserIds = [];
  for (var i = 2; i < attendanceData.length; i++) {
    var rowDate = attendanceData[i][0]; var userId = String(attendanceData[i][1]).trim();
    var status = attendanceData[i][3] ? String(attendanceData[i][3]).trim() : ""; 
    if (rowDate instanceof Date && rowDate.getTime() === targetDate.getTime() && userId && status !== "欠席") {
      todayUserIds.push(userId);
    }
  }
  
  if (todayUserIds.length === 0) { Browser.msgBox("お知らせ", "本日の利用者がいません。", Browser.Buttons.OK); return; }
  
  var basicData = basicSheet.getDataRange().getValues();
  var userMap = {};
  for (var i = 2; i < basicData.length; i++) {
    var uId = String(basicData[i][0]).trim();
    if (uId && uId !== "") {
      userMap[uId] = { name: basicData[i][1], lat: Number(basicData[i][4]), lng: Number(basicData[i][5]), timeStart: null, timeEnd: null, wheelchairType: "一般席", ngPassenger: "", ngDriver: "" };
    }
  }
  
  var conditionRows = conditionSheet.getDataRange().getDisplayValues();
  var headers = conditionRows[0]; var headerRowIdx = 0;
  for (var r = 0; r < 2; r++) {
    if (conditionRows[r].join("").indexOf("ID") !== -1 || conditionRows[r].join("").indexOf("id") !== -1) { headers = conditionRows[r]; headerRowIdx = r; break; }
  }
  
  var colId = 0, colWhType = 2, colStart = 3, colEnd = 4, colNgP = 5, colNgD = 6;
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).trim().toLowerCase();
    if (h.indexOf("id") !== -1 && h.indexOf("vehicle") === -1 && h.indexOf("driver") === -1 && h.indexOf("passenger") === -1) colId = c;
    if (h.indexOf("車椅子") !== -1 || h.indexOf("wheelchair") !== -1 || h.indexOf("区分") !== -1) colWhType = c;
    
    if (h.indexOf("dropoff") === -1 && h.indexOf("復路") === -1 && h.indexOf("送り") === -1) {
      if (h.indexOf("開始") !== -1 || h.indexOf("start") !== -1 || h.indexOf("迎え希望") !== -1) colStart = c;
      if (h.indexOf("終了") !== -1 || h.indexOf("end") !== -1 || h.indexOf("期限") !== -1 || h.indexOf("まで") !== -1) colEnd = c;
    }
    if (h.indexOf("ng乗客") !== -1 || h.indexOf("passenger") !== -1) colNgP = c;
    if (h.indexOf("ngドライバー") !== -1 || h.indexOf("driver") !== -1) colNgD = c;
  }
  
  for (var i = headerRowIdx + 1; i < conditionRows.length; i++) {
    var cId = String(conditionRows[i][colId]).trim();
    if (cId && userMap[cId]) {
      userMap[cId].wheelchairType = conditionRows[i][colWhType] ? String(conditionRows[i][colWhType]).trim() : "一般席";
      userMap[cId].timeStart = conditionRows[i][colStart] ? String(conditionRows[i][colStart]).trim() : "";
      userMap[cId].timeEnd = conditionRows[i][colEnd] ? String(conditionRows[i][colEnd]).trim() : "";
      userMap[cId].ngPassenger = conditionRows[i][colNgP] ? String(conditionRows[i][colNgP]).trim() : "";
      userMap[cId].ngDriver = conditionRows[i][colNgD] ? String(conditionRows[i][colNgD]).trim() : "";
    }
  }
  
  function getLimitTime(baseDate, timeStr) {
    if (!timeStr || timeStr === "" || timeStr === "undefined" || timeStr.indexOf("時間指定") !== -1) return null;
    var d = new Date(baseDate.getTime());
    var cleanStr = timeStr.replace(/[^0-9:]/g, ""); if (cleanStr.indexOf(":") === -1) return null;
    var parts = cleanStr.split(':'); var hours = parseInt(parts[0], 10); var minutes = parseInt(parts[1], 10);
    d.setHours(hours, minutes, 0, 0); return d;
  }
  
  // ─── 車両マスタから「往路」の行だけを狙い撃ちで選別 ───
  var vehicleRows = vehicleSheet.getDataRange().getDisplayValues();
  var vehicles = [];
  var vehicleRunCount = {}; // 車ごとの便数カウント用
  
  for (var i = 2; i < vehicleRows.length; i++) {
    var vId = String(vehicleRows[i][0]).trim(); 
    var vType = String(vehicleRows[i][9]).trim(); // J列：往復区分
    
    // 【重要】J列が「往路」の行だけをパズル対象にする
    if (vId && vType === "往路" && (Number(vehicleRows[i][3]) > 0 || Number(vehicleRows[i][4]) > 0)) {
      if (!vehicleRunCount[vId]) { vehicleRunCount[vId] = 1; } else { vehicleRunCount[vId]++; }
      
      vehicles.push({ 
        id: vId, 
        name: String(vehicleRows[i][1]).trim(), 
        capacityRegular: Number(vehicleRows[i][3]), 
        capacityWheelchair: Number(vehicleRows[i][4]), 
        departureTimeStr: vehicleRows[i][7], 
        returnLimitStr: vehicleRows[i][8],
        tripName: vehicleRunCount[vId] + "便目" // 「1便目」「2便目」を自動命名
      });
    }
  }
  
  function getDistance(lat1, lng1, lat2, lng2) {
    return Math.sqrt(Math.pow((lng1 - lng2) * 91, 2) + Math.pow((lat1 - lat2) * 111, 2));
  }
  
  var unassignedUsers = todayUserIds.filter(function(id) { return userMap[id]; });
  unassignedUsers.sort(function(a, b) {
    var timeA = userMap[a].timeEnd ? getLimitTime(targetDate, userMap[a].timeEnd) : null;
    var timeB = userMap[b].timeEnd ? getLimitTime(targetDate, userMap[b].timeEnd) : null;
    var tA = timeA ? timeA.getTime() : new Date(targetDate.getTime()).setHours(23,59,0,0);
    var tB = timeB ? timeB.getTime() : new Date(targetDate.getTime()).setHours(23,59,0,0);
    return tA - tB;
  });
  
  var outputValues = [];
  
  for (var v = 0; v < vehicles.length; v++) {
    var car = vehicles[v]; var currentLat = OFFICE_LAT; var currentLng = OFFICE_LNG;
    var regularCount = 0; var wheelchairCount = 0; var stopOrder = 1;
    var currentTime = getLimitTime(targetDate, car.departureTimeStr) || new Date(targetDate.getTime()).setHours(8, 15, 0, 0);
    var myLineLimit = car.returnLimitStr !== "" ? getLimitTime(targetDate, car.returnLimitStr) : (v < vehicles.length - 1 && vehicles[v+1].id === car.id ? getLimitTime(targetDate, vehicles[v+1].departureTimeStr) : null);
    
    var currentRunUserIds = []; 
    
    while (unassignedUsers.length > 0) {
      var nearestUserIdx = -1; var minDistance = 99999; var nextArrivalTime = null;
      
      // Mode A: 時間制限あり
      for (var u = 0; u < unassignedUsers.length; u++) {
        var uId = unassignedUsers[u]; var uInfo = userMap[uId];
        if (!uInfo.timeEnd || !getLimitTime(targetDate, uInfo.timeEnd)) continue;
        
        var hasSeatCapacity = (uInfo.wheelchairType === "車椅子") ? (wheelchairCount < car.capacityWheelchair) : (regularCount < car.capacityRegular);
        if (!hasSeatCapacity) continue;
        
        var dist = getDistance(currentLat, currentLng, uInfo.lat, uInfo.lng);
        var tempArrivalTime = new Date(currentTime.getTime()); tempArrivalTime.setMinutes(tempArrivalTime.getMinutes() + Math.ceil(dist * 2) + 3);
        
        var limitStart = getLimitTime(targetDate, uInfo.timeStart); 
        if (limitStart) {
          var waitMinutes = (limitStart.getTime() - tempArrivalTime.getTime()) / (1000 * 60);
          if (waitMinutes > MAX_WAIT_MINUTES) continue; 
          if (tempArrivalTime < limitStart) tempArrivalTime = limitStart; 
        }
        
        var isTimeOk = (tempArrivalTime <= getLimitTime(targetDate, uInfo.timeEnd));
        var isReturnOk = true;
        if (myLineLimit) {
          var tempReturnTime = new Date(tempArrivalTime.getTime()); tempReturnTime.setMinutes(tempReturnTime.getMinutes() + Math.ceil(getDistance(uInfo.lat, uInfo.lng, OFFICE_LAT, OFFICE_LNG) * 2) + 3);
          if (tempReturnTime > myLineLimit) isReturnOk = false;
        }
        
        if (isTimeOk && isReturnOk && dist < minDistance) { 
          minDistance = dist; nearestUserIdx = u; nextArrivalTime = tempArrivalTime; 
        }
      }
      
      // Mode B: フリー（最終救済）
      if (nearestUserIdx === -1) {
        for (var u = 0; u < unassignedUsers.length; u++) {
          var uId = unassignedUsers[u]; var uInfo = userMap[uId];
          var hasSeatCapacity = (uInfo.wheelchairType === "車椅子") ? (wheelchairCount < car.capacityWheelchair) : (regularCount < car.capacityRegular);
          if (!hasSeatCapacity) continue;
          
          var dist = getDistance(currentLat, currentLng, uInfo.lat, uInfo.lng);
          var tempArrivalTime = new Date(currentTime.getTime()); tempArrivalTime.setMinutes(tempArrivalTime.getMinutes() + Math.ceil(dist * 2) + 3);
          
          var limitStart = getLimitTime(targetDate, uInfo.timeStart); 
          if (limitStart && tempArrivalTime < limitStart) tempArrivalTime = limitStart;
          
          var isReturnOk = true;
          if (myLineLimit) {
            var tempReturnTime = new Date(tempArrivalTime.getTime()); tempReturnTime.setMinutes(tempReturnTime.getMinutes() + Math.ceil(getDistance(uInfo.lat, uInfo.lng, OFFICE_LAT, OFFICE_LNG) * 2) + 3);
            if (tempReturnTime > myLineLimit) isReturnOk = false;
          }
          if (isReturnOk && dist < minDistance) { minDistance = dist; nearestUserIdx = u; nextArrivalTime = tempArrivalTime; }
        }
      }
      
      if (nearestUserIdx !== -1) {
        var assignedId = unassignedUsers[nearestUserIdx]; var assignedInfo = userMap[assignedId];
        if (assignedInfo.wheelchairType === "車椅子") { wheelchairCount++; } else { regularCount++; }
        currentTime = nextArrivalTime;
        currentRunUserIds.push(assignedId); 
        
        var displayName = assignedInfo.name;
        if (assignedInfo.wheelchairType === "車椅子") displayName += " ♿";
        else if (assignedInfo.wheelchairType === "助手席") displayName += " 💺";
        
        // ★【列ズレ修正】D列に「1便目」等の便名、E列に「回る順番」を寸分違わず美しく格納！
        outputValues.push([targetDate, car.id, car.name, car.tripName, stopOrder, "往路", "", assignedId, displayName, Utilities.formatDate(currentTime, "JST", "HH:mm")]);
        currentLat = assignedInfo.lat; currentLng = assignedInfo.lng; stopOrder++;
        unassignedUsers.splice(nearestUserIdx, 1);
      } else { break; }
    }
    if (stopOrder > 1) {
      currentTime.setMinutes(currentTime.getMinutes() + Math.ceil(getDistance(currentLat, currentLng, OFFICE_LAT, OFFICE_LNG) * 2) + 3);
      outputValues.push([targetDate, car.id, car.name, car.tripName, stopOrder, "往路", "", "🏢 OFFICE", "ーーー 施設に帰着 ーーー", Utilities.formatDate(currentTime, "JST", "HH:mm")]);
    }
  }
  
  if (outputValues.length > 0) {
    var lastRow = outputSheet.getLastRow();
    var keepValues = [];
    if (lastRow >= 3) {
      var outData = outputSheet.getRange(3, 1, lastRow - 2, 10).getValues();
      for (var i = 0; i < outData.length; i++) {
        if (outData[i][5] !== "往路") keepValues.push(outData[i]); // 復路を救出して退避
      }
      outputSheet.getRange(3, 1, lastRow - 2, 10).clearContent(); 
    }
    if (outputValues.length > 0) {
      outputSheet.getRange(3, 1, outputValues.length, 10).setValues(outputValues);
    }
    if (keepValues.length > 0) {
      outputSheet.getRange(3 + outputValues.length, 1, keepValues.length, 10).setValues(keepValues);
    }
    Browser.msgBox("完了", "【V7決定版：往路配車完了】J列マスタと連動し、ルートを完全出力しました！", Browser.Buttons.OK);
  }
}
// ==========================================================
// 機能4：【復路（夕方の送り）】AI自動配車（V9：昼便ガード＆暴走セーフティ撤去・最終完成版）
// ==========================================================
function runAiDropoffRoutingV7() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName("t_daily_attendance");
  var basicSheet = ss.getSheetByName("m_user_basics");
  var vehicleSheet = ss.getSheetByName("m_vehicles");
  var conditionSheet = ss.getSheetByName("t_user_pickup_conditions"); 
  var outputSheet = ss.getSheetByName("t_routing_outputs");
  
  var targetDate = attendanceSheet.getRange("H1").getValue();
  var officeLat = attendanceSheet.getRange("H2").getValue(); 
  var officeLng = attendanceSheet.getRange("H3").getValue(); 
  
  if (!targetDate || !(targetDate instanceof Date) || officeLat === "" || officeLng === "") {
    Browser.msgBox("エラー", "日付または施設の緯度経度を確認してください。", Browser.Buttons.OK);
    return;
  }
  
  var OFFICE_LAT = Number(officeLat); var OFFICE_LNG = Number(officeLng);
  var attendanceData = attendanceSheet.getDataRange().getValues();
  var todayUserIds = [];
  for (var i = 2; i < attendanceData.length; i++) {
    var rowDate = attendanceData[i][0]; var userId = String(attendanceData[i][1]).trim();
    var status = attendanceData[i][3] ? String(attendanceData[i][3]).trim() : ""; 
    if (rowDate instanceof Date && rowDate.getTime() === targetDate.getTime() && userId && status !== "欠席") {
      todayUserIds.push(userId);
    }
  }
  
  if (todayUserIds.length === 0) { Browser.msgBox("お知らせ", "本日の利用者がいません。", Browser.Buttons.OK); return; }
  
  var basicData = basicSheet.getDataRange().getValues();
  var userMap = {};
  for (var i = 2; i < basicData.length; i++) {
    var uId = String(basicData[i][0]).trim();
    if (uId && uId !== "") {
      userMap[uId] = { name: basicData[i][1], lat: Number(basicData[i][4]), lng: Number(basicData[i][5]), wheelchairType: "一般席", dropoffTime: "" };
    }
  }
  
  var conditionRows = conditionSheet.getDataRange().getDisplayValues();
  var headers = conditionRows[0]; var headerRowIdx = 0;
  for (var r = 0; r < 2; r++) {
    if (conditionRows[r].join("").indexOf("ID") !== -1 || conditionRows[r].join("").indexOf("id") !== -1) { headers = conditionRows[r]; headerRowIdx = r; break; }
  }
  
  var colId = 0, colWhType = 2, colDropTime = -1;
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).trim().toLowerCase();
    if (h.indexOf("id") !== -1 && h.indexOf("vehicle") === -1 && h.indexOf("driver") === -1 && h.indexOf("passenger") === -1) colId = c;
    if (h.indexOf("車椅子") !== -1 || h.indexOf("wheelchair") !== -1 || h.indexOf("区分") !== -1) colWhType = c;
    
    if (h.indexOf("dropoff") !== -1 || h.indexOf("復路") !== -1 || h.indexOf("送り") !== -1) {
      if ((h.indexOf("希望") !== -1 || h.indexOf("時間") !== -1 || h.indexOf("time") !== -1 || h.indexOf("start") !== -1) && colDropTime === -1) { 
        colDropTime = c; 
      }
    }
  }
  if (colDropTime === -1) colDropTime = 5; 
  
  for (var i = headerRowIdx + 1; i < conditionRows.length; i++) {
    var cId = String(conditionRows[i][colId]).trim();
    if (cId && userMap[cId]) {
      userMap[cId].wheelchairType = conditionRows[i][colWhType] ? String(conditionRows[i][colWhType]).trim() : "一般席";
      userMap[cId].dropoffTime = conditionRows[i][colDropTime] ? String(conditionRows[i][colDropTime]).trim() : "";
    }
  }
  
  function getLimitTime(baseDate, timeStr) {
    if (!timeStr || timeStr === "" || timeStr === "undefined" || timeStr.indexOf("時間指定") !== -1) return null;
    var d = new Date(baseDate.getTime());
    var cleanStr = timeStr.replace(/[^0-9:]/g, ""); if (cleanStr.indexOf(":") === -1) return null;
    var parts = cleanStr.split(':'); var hours = parseInt(parts[0], 10); var minutes = parseInt(parts[1], 10);
    d.setHours(hours, minutes, 0, 0); return d;
  }
  
  var vehicleRows = vehicleSheet.getDataRange().getDisplayValues();
  var vehicleHeaders = vehicleRows[0];
  var colVDeparture = 7; 
  for (var c = 0; c < vehicleHeaders.length; c++) {
    var vh = String(vehicleHeaders[c]).toLowerCase();
    if ((vh.indexOf("復路") !== -1 || vh.indexOf("夕方") !== -1 || vh.indexOf("送り") !== -1 || vh.indexOf("dropoff") !== -1) && 
        (vh.indexOf("出発") !== -1 || vh.indexOf("start") !== -1 || vh.indexOf("time") !== -1)) {
      colVDeparture = c;
      break;
    }
  }
  
  var tempVehicles = [];
  for (var i = 2; i < vehicleRows.length; i++) {
    var vId = String(vehicleRows[i][0]).trim(); 
    var vType = String(vehicleRows[i][9]).trim(); 
    if (vId && vType === "復路" && (Number(vehicleRows[i][3]) > 0 || Number(vehicleRows[i][4]) > 0)) {
      tempVehicles.push({ id: vId, name: String(vehicleRows[i][1]).trim(), capacityRegular: Number(vehicleRows[i][3]), capacityWheelchair: Number(vehicleRows[i][4]), departureTimeStr: vehicleRows[i][colVDeparture] });
    }
  }
  
  // 出発時間が「早い順」にタイムラインをソート
  tempVehicles.sort(function(a, b) {
    var timeA = getLimitTime(targetDate, a.departureTimeStr); var timeB = getLimitTime(targetDate, b.departureTimeStr);
    var tA = timeA ? timeA.getTime() : 0; var tB = timeB ? timeB.getTime() : 0;
    return tA - tB;
  });
  
  var vehicles = []; var vehicleRunCount = {};
  for (var v = 0; v < tempVehicles.length; v++) {
    var car = tempVehicles[v];
    if (!vehicleRunCount[car.id]) { vehicleRunCount[car.id] = 1; } else { vehicleRunCount[car.id]++; }
    car.tripName = vehicleRunCount[car.id] + "便目";
    vehicles.push(car);
  }
  
  function getDistance(lat1, lng1, lat2, lng2) {
    return Math.sqrt(Math.pow((lng1 - lng2) * 91, 2) + Math.pow((lat1 - lat2) * 111, 2));
  }
  
  var unassignedUsers = todayUserIds.filter(function(id) { return userMap[id]; });
  var outputValues = [];
  
  for (var v = 0; v < vehicles.length; v++) {
    var car = vehicles[v]; var currentLat = OFFICE_LAT; var currentLng = OFFICE_LNG;
    var regularCount = 0; var wheelchairCount = 0; var stopOrder = 1;
    var currentTime = getLimitTime(targetDate, car.departureTimeStr) || new Date(targetDate.getTime()).setHours(15, 30, 0, 0);
    var currentRunUserIds = []; 
    
    while (unassignedUsers.length > 0) {
      var nearestUserIdx = -1; var bestScore = -1; var minDistance = 99999; var nextArrivalTime = null;
      
      // Mode A: 希望時間5分前後にジャストフィットする人を探索
      for (var u = 0; u < unassignedUsers.length; u++) {
        var uId = unassignedUsers[u]; var uInfo = userMap[uId];
        
        var hasSeatCapacity = (uInfo.wheelchairType === "車椅子") ? (wheelchairCount < car.capacityWheelchair) : (regularCount < car.capacityRegular);
        if (!hasSeatCapacity) continue;
        
        // 🚨【現場ルール】15:00前の「昼便」の場合、希望なし(空欄)の人は夕方まで残るため、昼便の選択肢から完全除外！
        var carDepartureTime = getLimitTime(targetDate, car.departureTimeStr);
        var isLunchCar = carDepartureTime && carDepartureTime.getHours() < 15;
        if (isLunchCar && (!uInfo.dropoffTime || String(uInfo.dropoffTime).trim() === "")) continue;
        
        var dist = getDistance(currentLat, currentLng, uInfo.lat, uInfo.lng);
        var tempArrivalTime = new Date(currentTime.getTime()); tempArrivalTime.setMinutes(tempArrivalTime.getMinutes() + Math.ceil(dist * 2) + 3);
        
        var currentScore = 0;
        
        if (uInfo.dropoffTime && getLimitTime(targetDate, uInfo.dropoffTime)) {
          var targetTime = getLimitTime(targetDate, uInfo.dropoffTime);
          var limitStart = new Date(targetTime.getTime()); limitStart.setMinutes(limitStart.getMinutes() - 5); 
          var limitEnd = new Date(targetTime.getTime()); limitEnd.setMinutes(limitEnd.getMinutes() + 5);     
          
          if (tempArrivalTime >= limitStart && tempArrivalTime <= limitEnd) {
            currentScore = 3000 - dist; 
          } else if (tempArrivalTime < limitStart) {
            currentScore = 0; 
          } else {
            currentScore = -100; 
          }
        } else {
          currentScore = 1000 - dist;
        }
        
        if (currentScore > 0 && currentScore > bestScore) {
          bestScore = currentScore; nearestUserIdx = u; nextArrivalTime = tempArrivalTime;
        }
      }
      
      // Mode B: フリー枠（最終救済）
      if (nearestUserIdx === -1) {
        for (var u = 0; u < unassignedUsers.length; u++) {
          var uId = unassignedUsers[u]; var uInfo = userMap[uId];
          
          var hasSeatCapacity = (uInfo.wheelchairType === "車椅子") ? (wheelchairCount < car.capacityWheelchair) : (regularCount < car.capacityRegular);
          if (!hasSeatCapacity) continue;
          
          // 🚨【現場ルール】昼便には、希望なしの人は乗せない
          var carDepartureTime = getLimitTime(targetDate, car.departureTimeStr);
          var isLunchCar = carDepartureTime && carDepartureTime.getHours() < 15;
          if (isLunchCar && (!uInfo.dropoffTime || String(uInfo.dropoffTime).trim() === "")) continue;
          
          var dist = getDistance(currentLat, currentLng, uInfo.lat, uInfo.lng);
          var tempArrivalTime = new Date(currentTime.getTime()); tempArrivalTime.setMinutes(tempArrivalTime.getMinutes() + Math.ceil(dist * 2) + 3);
          
          if (uInfo.dropoffTime && getLimitTime(targetDate, uInfo.dropoffTime)) {
            var targetTime = getLimitTime(targetDate, uInfo.dropoffTime);
            
            if (targetTime < carDepartureTime) continue;
            
            var limitStartBuffer = new Date(targetTime.getTime()); limitStartBuffer.setMinutes(limitStartBuffer.getMinutes() - 30);
            if (tempArrivalTime < limitStartBuffer) continue;
          }
          
          if (dist < minDistance) { minDistance = dist; nearestUserIdx = u; nextArrivalTime = tempArrivalTime; }
        }
      }
      
      // ⚠️【大改造】暴走誘拐セーフティブロックを完全撤去！
      // 誰も該当者がいなければ、無理に誰かを乗せずに即座にbreakして施設へ帰着させます。
      
      if (nearestUserIdx !== -1) {
        var assignedId = unassignedUsers[nearestUserIdx]; var assignedInfo = userMap[assignedId];
        if (assignedInfo.wheelchairType === "車椅子") { wheelchairCount++; } else { regularCount++; }
        currentTime = nextArrivalTime;
        currentRunUserIds.push(assignedId); 
        
        var displayName = assignedInfo.name;
        if (assignedInfo.wheelchairType === "車椅子") displayName += " ♿";
        else if (assignedInfo.wheelchairType === "助手席") displayName += " 💺";
        
        outputValues.push([targetDate, car.id, car.name, car.tripName, stopOrder, "復路", "", assignedId, displayName, Utilities.formatDate(currentTime, "JST", "HH:mm")]);
        currentLat = assignedInfo.lat; currentLng = assignedInfo.lng; stopOrder++;
        unassignedUsers.splice(nearestUserIdx, 1);
      } else { break; }
    }
    if (stopOrder > 1) {
      currentTime.setMinutes(currentTime.getMinutes() + Math.ceil(getDistance(currentLat, currentLng, OFFICE_LAT, OFFICE_LNG) * 2) + 3);
      outputValues.push([targetDate, car.id, car.name, car.tripName, stopOrder, "復路", "", "🏢 OFFICE", "ーーー 施設に帰着 ーーー", Utilities.formatDate(currentTime, "JST", "HH:mm")]);
    }
  }
  
  if (outputValues.length > 0) {
    var lastRow = outputSheet.getLastRow();
    var keepValues = [];
    if (lastRow >= 3) {
      var outData = outputSheet.getRange(3, 1, lastRow - 2, 10).getValues();
      for (var i = 0; i < outData.length; i++) {
        if (outData[i][5] !== "復路") keepValues.push(outData[i]); 
      }
      outputSheet.getRange(3, 1, lastRow - 2, 10).clearContent(); 
    }
    if (keepValues.length > 0) {
      outputSheet.getRange(3, 1, keepValues.length, 10).setValues(keepValues);
    }
    outputSheet.getRange(3 + keepValues.length, 1, outputValues.length, 10).setValues(outputValues);
    
    Browser.msgBox("完了", "【V9最終完成版：復路時系列配車完了】昼の臨時便と、夕方の定期送り便が完全に仕分けられました！", Browser.Buttons.OK);
  }
}
// ==========================================
// 機能5：【マスタ自動化】住所 ➔ 緯度経度・マップ絵文字リンク一発変換
// ==========================================
function convertAddressToLatLng() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var basicSheet = ss.getSheetByName("m_user_basics"); 
  if (!basicSheet) return;
  
  var basicData = basicSheet.getDataRange().getValues();
  var successCount = 0;
  
  for (var i = 2; i < basicData.length; i++) {
    var address = basicData[i][3] ? String(basicData[i][3]).trim() : ""; 
    var currentLat = basicData[i][4];
    var currentLng = basicData[i][5];
    
    // 【修正】緯度か経度の「どちらか片方でも」空文字（空欄）なら、確実に再取得を走らせる
    if (address && address !== "" && (currentLat === "" || currentLng === "")) {
      try {
        var response = Maps.newGeocoder().geocode(address);
        if (response.status === "OK" && response.results && response.results.length > 0) {
          var loc = response.results[0].geometry.location;
          
          // 1. 緯度と経度を書き込み（5列目＝E列、6列目＝F列）
          basicSheet.getRange(i + 1, 5).setValue(loc.lat); 
          basicSheet.getRange(i + 1, 6).setValue(loc.lng);
          
          // 2. 【ここを修正！】長ったらしいURLを「🗺️」の絵文字リンクに変身させる
          var mapUrl = "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(address);
          // スプレッドシートの「=HYPERLINK("URL", "表示文字")」という数式を組み立てる
          var hyperlinkFormula = '=HYPERLINK("' + mapUrl + '", "🗺️")';
          
          // 7列目（＝G列）に数式として書き込みます
          basicSheet.getRange(i + 1, 7).setFormula(hyperlinkFormula); 
          
          successCount++; 
          Utilities.sleep(100);
        }
      } catch (e) { 
        console.error("エラー: " + e.message); 
      }
    }
  }
  if (successCount > 0) Browser.msgBox("完了", successCount + " 名の緯度経度とマップリンク(🗺️)を自動取得しました！", Browser.Buttons.OK);
}
/**
 * 【実績自動蓄積システム】今日の運行表を t_past_records の一番下に自動でコピペ追記する
 */
function saveToPastRecords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = ss.getSheetByName("t_routing_outputs");
  var pastSheet = ss.getSheetByName("t_past_records");
  
  if (!outputSheet || !pastSheet) {
    Browser.msgBox("エラー", "t_routing_outputs または t_past_records タブが見つかりません。", Browser.Buttons.OK);
    return;
  }
  
  var outputData = outputSheet.getDataRange().getValues();
  
  // 3行目以降にデータがあるかチェック（ヘッダーだけの時はスキップ）
  if (outputData.length < 3) {
    Browser.msgBox("お知らせ", "保存する配車データがありません。", Browser.Buttons.OK);
    return;
  }
  
  // 3行目以降のデータを抽出
  var valuesToSave = [];
  for (var i = 2; i < outputData.length; i++) {
    valuesToSave.push(outputData[i]);
  }
  
  // t_past_records の一番下の行を取得して、そのすぐ下に追記
  var lastRow = pastSheet.getLastRow();
  var startRow = (lastRow < 3) ? 3 : lastRow + 1;
  
  pastSheet.getRange(startRow, 1, valuesToSave.length, 10).setValues(valuesToSave);
  
  Browser.msgBox("完了", "本日の配車結果（" + valuesToSave.length + "行）を、実績データ（t_past_records）の末尾に完全保存しました！", Browser.Buttons.OK);
}
/**
 * LINEにそのまま貼り付けられる送迎運行表テキストを生成する関数（マスタ連動・マップリンク付き）
 */
function generateLineCopyText() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName('t_daily_attendance'); // ボタンと出力先
  var outputSheet = ss.getSheetByName('t_routing_outputs');     // 配車データ元
  var userSheet = ss.getSheetByName('m_user_basics');           // 利用者マスタ
  var vehicleSheet = ss.getSheetByName('m_vehicles');           // 車両マスタ
  
  if (!attendanceSheet || !outputSheet || !userSheet || !vehicleSheet) {
    Browser.msgBox("エラー", "必要なシート（t_daily_attendance, t_routing_outputs, m_user_basics, m_vehicles）のいずれかが見つかりません。", Browser.Buttons.OK);
    return;
  }
  
  // ─── 1. 各種マスタデータを読み込んで「ID ➔ 日本語」の変換辞書を作る ───
  
  // ① 車両マスタの読み込み (A列: ID, B列: 車両名)
  var vehicleMap = {};
  var vehicleData = vehicleSheet.getDataRange().getValues();
  for (var i = 1; i < vehicleData.length; i++) {
    var vId = String(vehicleData[i][0]).trim();
    var vName = String(vehicleData[i][1]).trim();
    if (vId) vehicleMap[vId] = vName;
  }
  
  // ② 利用者マスタの読み込み (A列: ID, B列: 名前, D列: 住所)
  var userMap = {};
  var userData = userSheet.getDataRange().getValues();
  for (var i = 1; i < userData.length; i++) {
    var uId = String(userData[i][0]).trim();
    var uName = String(userData[i][1]).trim();
    var uAddress = userData[i][3] ? String(userData[i][3]).trim() : ""; // D列(インデックス3)が住所
    
    if (uId) {
      userMap[uId] = {
        name: uName,
        address: uAddress
      };
    }
  }
  
  // ─── 2. 配車結果データの読み込みと、ズレない列特定 ───
  var data = outputSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Browser.msgBox("通知", "配車データがありません。先に配車計算を実行してください。", Browser.Buttons.OK);
    return;
  }
  
  var header = data[0];
  var colVehicle = -1;
  var colTime = -1;
  var colUserId = -1;
  var colDirection = -1;
  
  // ヘッダーの文字から列の位置を正確に自動スキャン
  for (var c = 0; c < header.length; c++) {
    var hName = String(header[c]).toLowerCase().trim();
    if (hName.indexOf('vehicle_id') !== -1) colVehicle = c;
    if (hName.indexOf('time') !== -1 || hName.indexOf('arrival') !== -1 || hName === '時間') colTime = c;
    if (hName.indexOf('user_id') !== -1 || hName.indexOf('passenger_id') !== -1 || hName === '利用者id') colUserId = c;
    if (hName.indexOf('direction') !== -1 || hName === '方向') colDirection = c;
  }
  
  // 【重要】万が一自動検出に失敗した場合の、シート構造から逆算した手動フォールバック設定
  if (colVehicle === -1) colVehicle = 1;   // B列 (vehicle_id)
  if (colTime === -1) colTime = 4;          // E列（通常、ここに時間が入ります）
  if (colUserId === -1) colUserId = 7;      // H列 (user_id) ※前回ここにM002が出ていたため確実
  if (colDirection === -1) colDirection = 5; // F列 (direction) ※前回ここに「往路」が出ていたため確実
  
  // ─── 3. LINE用テキストの組み立て（IDを日本語に変換しながら結合） ───
  var dateStr = Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
  var lineText = "🚐 【本日の送迎運行表】 " + dateStr + "\n";
  lineText += "━━━━━━━━━━━━━━\n";
  
  var currentVehicleId = "";
  var currentDirection = "";
  
  for (var i = 1; i < data.length; i++) {
    var vId = data[i][colVehicle] ? String(data[i][colVehicle]).trim() : "";
    var time = data[i][colTime];
    var uId = data[i][colUserId] ? String(data[i][colUserId]).trim() : "";
    var direction = data[i][colDirection] ? String(data[i][colDirection]).trim() : "";
    
    // ヘッダー行や空行、おかしなシステム文字の行はスキップ
    if (vId === "vehicle_id" || !uId || uId === "user_id" || uId === "direction") continue;
    
    // ① 車両IDから「本物の車両名」をマスタ引き
    var vehicleName = vehicleMap[vId] || vId;
    
    // ② 利用者IDから「本物の名前」と「住所」をマスタ引き
    var userName = uId;
    var mapUrl = "";
    if (userMap[uId]) {
      userName = userMap[uId].name;
      if (userMap[uId].address) {
        // LINEでタップして直接一発起動できるGoogleマップのURLを生成
        mapUrl = "https://www.google.com/maps/search/?api=1&query=" + encodeURIComponent(userMap[uId].address);
      }
    }
    
    // 時間の表示形式を「HH:mm」に整える
    var timeStr = "--:--";
    if (time instanceof Date) {
      timeStr = Utilities.formatDate(time, "JST", "HH:mm");
    } else if (time && String(time).trim() !== "") {
      timeStr = String(time).trim();
      // 万が一、列のズレで「往路」などの文字を時間を拾ってしまっていた場合のセーフティ
      if (timeStr === "往路" || timeStr === "復路") timeStr = "時間指定なし";
    }
    
    // 車両が切り替わったら、新しい車両名で見出しを入れる
    if (vId !== currentVehicleId) {
      currentVehicleId = vId;
      lineText += "\n🚗 【" + vehicleName + "】\n";
      currentDirection = ""; // 便をリセット
    }
    
    // 往路/復路（便）が切り替わったら見出しを入れる
    if (direction !== currentDirection) {
      currentDirection = direction;
      lineText += "  ── " + currentDirection + " ──\n";
    }
    
    // 利用者の送迎予定を1行ずつ追加
    lineText += "  🟢 " + timeStr + "  " + userName + " 様\n";
    // Googleマップのリンクがあれば、次の行にインデント付きで差し込む
    if (mapUrl) {
      lineText += "     🗺️ マップ: " + mapUrl + "\n";
    }
  }
  
  lineText += "\n━━━━━━━━━━━━━━\n🏁 今日も安全運転でお願いします！";
  
  // 4. 【出力】t_daily_attendance シートの「J1セル」にドカンと書き込み
  attendanceSheet.getRange("J1").setValue(lineText);
  
  // 5. 画面中央に完了ポップアップを表示
  var ui = SpreadsheetApp.getUi();
  ui.alert("生成完了", "t_daily_attendanceのJ1セルに、マスタ連動＆マップリンク付きのテキストを保存しました！", ui.ButtonSet.OK);
}
