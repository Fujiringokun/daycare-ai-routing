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
// 機能3：【往路（朝のお迎え）】AI自動配車
// ==========================================================
function runAiRouting() {
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
      userMap[uId] = { name: basicData[i][1], lat: Number(basicData[i][4]), lng: Number(basicData[i][5]), timeStart: null, timeEnd: null, wheelchairType: "一般席", ngPassenger: "", ngDriver: "", dropoffTimeEnd: null };
    }
  }
  
  var conditionRows = conditionSheet.getDataRange().getDisplayValues();
  var headers = conditionRows[0]; var headerRowIdx = 0;
  for (var r = 0; r < 2; r++) {
    if (conditionRows[r].join("").indexOf("ID") !== -1 || conditionRows[r].join("").indexOf("id") !== -1) { headers = conditionRows[r]; headerRowIdx = r; break; }
  }
  var colId = 0, colWhType = 2, colStart = 3, colEnd = 4, colNgP = 5, colNgD = 6, colDropEnd = 7;
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).trim().toLowerCase();
    if (h.indexOf("id") !== -1) colId = c;
    if (h.indexOf("車椅子") !== -1 || h.indexOf("wheelchair") !== -1) colWhType = c;
    if (h.indexOf("往路開始") !== -1 || h.indexOf("start") !== -1) colStart = c;
    if (h.indexOf("往路終了") !== -1 || h.indexOf("end") !== -1) colEnd = c;
    if (h.indexOf("ng乗客") !== -1 || h.indexOf("passenger") !== -1) colNgP = c;
    if (h.indexOf("ngドライバー") !== -1 || h.indexOf("driver") !== -1) colNgD = c;
    if (h.indexOf("復路終了") !== -1 || h.indexOf("dropoff") !== -1 || h.indexOf("届けて") !== -1) colDropEnd = c;
  }
  
  for (var i = headerRowIdx + 1; i < conditionRows.length; i++) {
    var cId = String(conditionRows[i][colId]).trim();
    if (cId && userMap[cId]) {
      userMap[cId].wheelchairType = conditionRows[i][colWhType] ? String(conditionRows[i][colWhType]).trim() : "一般席";
      userMap[cId].timeStart = conditionRows[i][colStart] ? conditionRows[i][colStart].trim() : null;
      userMap[cId].timeEnd = conditionRows[i][colEnd] ? conditionRows[i][colEnd].trim() : null;
      userMap[cId].ngPassenger = conditionRows[i][colNgP] ? conditionRows[i][colNgP].trim() : "";
      userMap[cId].ngDriver = conditionRows[i][colNgD] ? conditionRows[i][colNgD].trim() : "";
      userMap[cId].dropoffTimeEnd = conditionRows[i][colDropEnd] ? conditionRows[i][colDropEnd].trim() : null;
    }
  }
  
  function getLimitTime(baseDate, timeStr) {
    if (!timeStr) return null;
    var d = new Date(baseDate.getTime());
    var isPm = (timeStr.indexOf("午後") !== -1 || timeStr.toLowerCase().indexOf("pm") !== -1);
    var cleanStr = timeStr.replace(/[^0-9:]/g, ""); if (cleanStr.indexOf(":") === -1) return null;
    var parts = cleanStr.split(':'); var hours = parseInt(parts[0], 10); var minutes = parseInt(parts[1], 10);
    if (isPm && hours < 12) hours += 12;
    d.setHours(hours, minutes, 0, 0); return d;
  }
  
  var vehicleRows = vehicleSheet.getDataRange().getDisplayValues();
  var vehicles = [];
  for (var i = 2; i < vehicleRows.length; i++) {
    var vId = String(vehicleRows[i][0]).trim(); 
    if (vId && (Number(vehicleRows[i][3]) > 0 || Number(vehicleRows[i][4]) > 0)) {
      vehicles.push({ id: vId, name: String(vehicleRows[i][1]).trim(), capacityRegular: Number(vehicleRows[i][3]), capacityWheelchair: Number(vehicleRows[i][4]), departureTimeStr: vehicleRows[i][7], returnLimitStr: vehicleRows[i][8] });
    }
  }
  
  function getDistance(lat1, lng1, lat2, lng2) {
    return Math.sqrt(Math.pow((lng1 - lng2) * 91, 2) + Math.pow((lat1 - lat2) * 111, 2));
  }
  
  var unassignedUsers = todayUserIds.filter(function(id) { return userMap[id]; });
  unassignedUsers.sort(function(a, b) {
    var timeA = userMap[a].timeEnd ? getLimitTime(targetDate, userMap[a].timeEnd).getTime() : new Date(targetDate.getTime()).setHours(23,59,0,0);
    var timeB = userMap[b].timeEnd ? getLimitTime(targetDate, userMap[b].timeEnd).getTime() : new Date(targetDate.getTime()).setHours(23,59,0,0);
    return timeA - timeB;
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
        if (!uInfo.timeEnd) break;
        
        var hasSeatCapacity = (uInfo.wheelchairType === "車椅子") ? (wheelchairCount < car.capacityWheelchair) : (regularCount < car.capacityRegular);
        if (!hasSeatCapacity) continue;
        
        // ⭐【修正】NGドライバーの完全一致チェック（スペース無視）
        if (uInfo.ngDriver && uInfo.ngDriver !== "") {
          var cleanedNgD = cleanStringForMatch(uInfo.ngDriver);
          if (cleanStringForMatch(car.name) === cleanedNgD || cleanStringForMatch(car.id) === cleanedNgD) continue;
        }
        if (checkNgPassengerConflict(uId, uInfo, currentRunUserIds, userMap)) continue;
        
        var dist = getDistance(currentLat, currentLng, uInfo.lat, uInfo.lng);
        var tempArrivalTime = new Date(currentTime.getTime()); tempArrivalTime.setMinutes(tempArrivalTime.getMinutes() + Math.ceil(dist * 2) + 3);
        var limitStart = getLimitTime(targetDate, uInfo.timeStart); if (limitStart && tempArrivalTime < limitStart) tempArrivalTime = limitStart;
        
        var isTimeOk = (tempArrivalTime <= getLimitTime(targetDate, uInfo.timeEnd));
        var isReturnOk = true;
        if (myLineLimit) {
          var tempReturnTime = new Date(tempArrivalTime.getTime()); tempReturnTime.setMinutes(tempReturnTime.getMinutes() + Math.ceil(getDistance(uInfo.lat, uInfo.lng, OFFICE_LAT, OFFICE_LNG) * 2) + 3);
          if (tempReturnTime > myLineLimit) isReturnOk = false;
        }
        if (isTimeOk && isReturnOk) { nearestUserIdx = u; nextArrivalTime = tempArrivalTime; break; }
      }
      
      // Mode B: フリー
      if (nearestUserIdx === -1) {
        for (var u = 0; u < unassignedUsers.length; u++) {
          var uId = unassignedUsers[u]; var uInfo = userMap[uId];
          if (uInfo.timeEnd) continue;
          
          var hasSeatCapacity = (uInfo.wheelchairType === "車椅子") ? (wheelchairCount < car.capacityWheelchair) : (regularCount < car.capacityRegular);
          if (!hasSeatCapacity) continue;
          
          // ⭐【修正】NGドライバーの完全一致チェック
          if (uInfo.ngDriver && uInfo.ngDriver !== "") {
            var cleanedNgD = cleanStringForMatch(uInfo.ngDriver);
            if (cleanStringForMatch(car.name) === cleanedNgD || cleanStringForMatch(car.id) === cleanedNgD) continue;
          }
          if (checkNgPassengerConflict(uId, uInfo, currentRunUserIds, userMap)) continue;
          
          var dist = getDistance(currentLat, currentLng, uInfo.lat, uInfo.lng);
          var tempArrivalTime = new Date(currentTime.getTime()); tempArrivalTime.setMinutes(tempArrivalTime.getMinutes() + Math.ceil(dist * 2) + 3);
          var limitStart = getLimitTime(targetDate, uInfo.timeStart); if (limitStart && tempArrivalTime < limitStart) tempArrivalTime = limitStart;
          
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
        
        outputValues.push([targetDate, car.id, car.name, "", "", "往路", stopOrder, assignedId, displayName, Utilities.formatDate(currentTime, "JST", "HH:mm")]);
        currentLat = assignedInfo.lat; currentLng = assignedInfo.lng; stopOrder++;
        unassignedUsers.splice(nearestUserIdx, 1);
      } else { break; }
    }
    if (stopOrder > 1) {
      currentTime.setMinutes(currentTime.getMinutes() + Math.ceil(getDistance(currentLat, currentLng, OFFICE_LAT, OFFICE_LNG) * 2) + 3);
      outputValues.push([targetDate, car.id, car.name, "", "", "往路", stopOrder, "🏢 OFFICE", "ーーー 施設に帰着 ーーー", Utilities.formatDate(currentTime, "JST", "HH:mm")]);
    }
  }
  
  if (outputValues.length > 0) {
    var lastRow = outputSheet.getLastRow();
    if (lastRow >= 3) outputSheet.getRange(3, 1, lastRow - 2, outputSheet.getLastColumn() || 10).clearContent();
    outputSheet.getRange(3, 1, outputValues.length, 10).setValues(outputValues);
    var msg = "【AI配車完了】 最新の往路ルートに上書きしました！";
    if (unassignedUsers.length > 0) {
      var unassignedNames = unassignedUsers.map(function(id) { return " ・ " + (userMap[id] ? userMap[id].name : id); }).join("\n");
      msg += "\n\n⚠️ 【ルートが組めなかった人: " + unassignedUsers.length + "名】\n" + unassignedNames;
    }
    Browser.msgBox("完了", msg, Browser.Buttons.OK);
  }
}

// ==========================================
// 機能4：【復路（夕方のお送り）】AI自動配車
// ==========================================
function runAiRoutingReturn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var attendanceSheet = ss.getSheetByName("t_daily_attendance");
  var basicSheet = ss.getSheetByName("m_user_basics");
  var vehicleSheet = ss.getSheetByName("m_vehicles");
  var conditionSheet = ss.getSheetByName("t_user_pickup_conditions"); 
  var outputSheet = ss.getSheetByName("t_routing_outputs");
  
  var targetDate = attendanceSheet.getRange("H1").getValue();
  var officeLat = attendanceSheet.getRange("H2").getValue(); var officeLng = attendanceSheet.getRange("H3").getValue(); 
  if (!targetDate || !(targetDate instanceof Date)) return;
  
  var OFFICE_LAT = Number(officeLat); var OFFICE_LNG = Number(officeLng);
  var attendanceData = attendanceSheet.getDataRange().getValues();
  var todayUserIds = [];
  for (var i = 2; i < attendanceData.length; i++) {
    var rowDate = attendanceData[i][0]; var userId = String(attendanceData[i][1]).trim();
    if (rowDate instanceof Date && rowDate.getTime() === targetDate.getTime() && userId && attendanceData[i][3] !== "欠席") {
      todayUserIds.push(userId);
    }
  }
  
  if (todayUserIds.length === 0) return;
  
  var basicData = basicSheet.getDataRange().getValues();
  var userMap = {};
  for (var i = 2; i < basicData.length; i++) {
    var uId = String(basicData[i][0]).trim();
    if (uId && uId !== "") userMap[uId] = { name: basicData[i][1], lat: Number(basicData[i][4]), lng: Number(basicData[i][5]), wheelchairType: "一般席", ngPassenger: "", ngDriver: "", dropoffTimeEnd: null };
  }
  
  var conditionRows = conditionSheet.getDataRange().getDisplayValues();
  var headers = conditionRows[0]; var headerRowIdx = 0;
  for (var r = 0; r < 2; r++) {
    if (conditionRows[r].join("").indexOf("ID") !== -1 || conditionRows[r].join("").indexOf("id") !== -1) { headers = conditionRows[r]; headerRowIdx = r; break; }
  }
  var colId = 0, colWhType = 2, colNgP = 5, colNgD = 6, colDropEnd = 7;
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c]).trim().toLowerCase();
    if (h.indexOf("id") !== -1) colId = c;
    if (h.indexOf("車椅子") !== -1 || h.indexOf("wheelchair") !== -1) colWhType = c;
    if (h.indexOf("ng乗客") !== -1 || h.indexOf("passenger") !== -1) colNgP = c;
    if (h.indexOf("ngドライバー") !== -1 || h.indexOf("driver") !== -1) colNgD = c;
    if (h.indexOf("復路終了") !== -1 || h.indexOf("dropoff") !== -1 || h.indexOf("届けて") !== -1) colDropEnd = c;
  }
  
  for (var i = headerRowIdx + 1; i < conditionRows.length; i++) {
    var cId = String(conditionRows[i][colId]).trim();
    if (cId && userMap[cId]) {
      userMap[cId].wheelchairType = conditionRows[i][colWhType] ? String(conditionRows[i][colWhType]).trim() : "一般席";
      userMap[cId].ngPassenger = conditionRows[i][colNgP] ? conditionRows[i][colNgP].trim() : "";
      userMap[cId].ngDriver = conditionRows[i][colNgD] ? conditionRows[i][colNgD].trim() : "";
      userMap[cId].dropoffTimeEnd = conditionRows[i][colDropEnd] ? conditionRows[i][colDropEnd].trim() : null;
    }
  }
  
  function getLimitTime(baseDate, timeStr) {
    if (!timeStr || timeStr === "") return null;
    var d = new Date(baseDate.getTime());
    var cleanStr = timeStr.replace(/[^0-9:]/g, ""); if (cleanStr.indexOf(":") === -1) return null;
    var parts = cleanStr.split(':'); d.setHours(parseInt(parts[0], 10), parseInt(parts[1], 10), 0, 0); return d;
  }
  
  var vehicleRows = vehicleSheet.getDataRange().getDisplayValues();
  var vehicles = [];
  for (var i = 2; i < vehicleRows.length; i++) {
    var vId = String(vehicleRows[i][0]).trim(); 
    if (vId && (Number(vehicleRows[i][3]) > 0 || Number(vehicleRows[i][4]) > 0)) {
      vehicles.push({ id: vId, name: String(vehicleRows[i][1]).trim(), capacityRegular: Number(vehicleRows[i][3]), capacityWheelchair: Number(vehicleRows[i][4]), returnDepartureTimeStr: vehicleRows[i][9] });
    }
  }
  
  function getDistance(lat1, lng1, lat2, lng2) {
    return Math.sqrt(Math.pow((lng1 - lng2) * 91, 2) + Math.pow((lat1 - lat2) * 111, 2));
  }
  
  var unassignedUsers = todayUserIds.filter(function(id) { return userMap[id]; });
  
  unassignedUsers.sort(function(a, b) {
    var timeA = userMap[a].dropoffTimeEnd ? getLimitTime(targetDate, userMap[a].dropoffTimeEnd).getTime() : new Date(targetDate.getTime()).setHours(23,59,0,0);
    var timeB = userMap[b].dropoffTimeEnd ? getLimitTime(targetDate, userMap[b].dropoffTimeEnd).getTime() : new Date(targetDate.getTime()).setHours(23,59,0,0);
    return timeA - timeB;
  });
  
  var outputValues = [];
  
  for (var v = 0; v < vehicles.length; v++) {
    var car = vehicles[v]; var currentLat = OFFICE_LAT; var currentLng = OFFICE_LNG;
    var regularCount = 0; var wheelchairCount = 0; var stopOrder = 1;
    var currentTime = getLimitTime(targetDate, car.returnDepartureTimeStr);
    var myLineLimit = (v < vehicles.length - 1 && vehicles[v+1].id === car.id) ? getLimitTime(targetDate, vehicles[v+1].returnDepartureTimeStr) : null;
    
    var currentRunUserIds = []; 
    
    while (unassignedUsers.length > 0) {
      var nearestUserIdx = -1; var minDistance = 99999; var nextArrivalTime = null;
      
      // 復路・Mode A
      for (var u = 0; u < unassignedUsers.length; u++) {
        var uId = unassignedUsers[u]; var uInfo = userMap[uId];
        if (!uInfo.dropoffTimeEnd) break; 
        
        var hasSeatCapacity = (uInfo.wheelchairType === "車椅子") ? (wheelchairCount < car.capacityWheelchair) : (regularCount < car.capacityRegular);
        if (!hasSeatCapacity) continue;
        
        // ⭐【修正】NGドライバーの完全一致チェック
        if (uInfo.ngDriver && uInfo.ngDriver !== "") {
          var cleanedNgD = cleanStringForMatch(uInfo.ngDriver);
          if (cleanStringForMatch(car.name) === cleanedNgD || cleanStringForMatch(car.id) === cleanedNgD) continue;
        }
        if (checkNgPassengerConflict(uId, uInfo, currentRunUserIds, userMap)) continue;
        
        var dist = getDistance(currentLat, currentLng, uInfo.lat, uInfo.lng);
        var tempArrivalTime = new Date(currentTime.getTime()); tempArrivalTime.setMinutes(tempArrivalTime.getMinutes() + Math.ceil(dist * 2) + 3);
        
        var isTimeOk = (tempArrivalTime <= getLimitTime(targetDate, uInfo.dropoffTimeEnd));
        var isReturnOk = true;
        if (myLineLimit) {
          var tempReturnTime = new Date(tempArrivalTime.getTime()); tempReturnTime.setMinutes(tempReturnTime.getMinutes() + Math.ceil(getDistance(uInfo.lat, uInfo.lng, OFFICE_LAT, OFFICE_LNG) * 2) + 3);
          if (tempReturnTime > myLineLimit) isReturnOk = false;
        }
        if (isTimeOk && isReturnOk) { nearestUserIdx = u; nextArrivalTime = tempArrivalTime; break; }
      }
      
      // 復路・Mode B
      if (nearestUserIdx === -1) {
        for (var u = 0; u < unassignedUsers.length; u++) {
          var uId = unassignedUsers[u]; var uInfo = userMap[uId];
          if (uInfo.dropoffTimeEnd) continue; 
          
          var hasSeatCapacity = (uInfo.wheelchairType === "車椅子") ? (wheelchairCount < car.capacityWheelchair) : (regularCount < car.capacityRegular);
          if (!hasSeatCapacity) continue;
          
          // ⭐【修正】NGドライバーの完全一致チェック
          if (uInfo.ngDriver && uInfo.ngDriver !== "") {
            var cleanedNgD = cleanStringForMatch(uInfo.ngDriver);
            if (cleanStringForMatch(car.name) === cleanedNgD || cleanStringForMatch(car.id) === cleanedNgD) continue;
          }
          if (checkNgPassengerConflict(uId, uInfo, currentRunUserIds, userMap)) continue;
          
          var dist = getDistance(currentLat, currentLng, uInfo.lat, uInfo.lng);
          var tempArrivalTime = new Date(currentTime.getTime()); tempArrivalTime.setMinutes(tempArrivalTime.getMinutes() + Math.ceil(dist * 2) + 3);
          
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
        
        outputValues.push([targetDate, car.id, car.name, "", "", "復路", stopOrder, assignedId, displayName, Utilities.formatDate(currentTime, "JST", "HH:mm")]);
        currentLat = assignedInfo.lat; currentLng = assignedInfo.lng; stopOrder++;
        unassignedUsers.splice(nearestUserIdx, 1);
      } else { break; }
    }
    if (stopOrder > 1) {
      currentTime.setMinutes(currentTime.getMinutes() + Math.ceil(getDistance(currentLat, currentLng, OFFICE_LAT, OFFICE_LNG) * 2) + 3);
      outputValues.push([targetDate, car.id, car.name, "", "", "復路", stopOrder, "🏢 OFFICE", "ーーー 施設に帰着 ーーー", Utilities.formatDate(currentTime, "JST", "HH:mm")]);
    }
  }
  
  if (outputValues.length > 0) {
    var startRow = (outputSheet.getLastRow() < 3) ? 3 : outputSheet.getLastRow() + 1;
    outputSheet.getRange(startRow, 1, outputValues.length, 10).setValues(outputValues);
    var msg = "【AI復路配車完了】 夕方の送りルートを往路の下に追記しました！";
    if (unassignedUsers.length > 0) {
      var unassignedNames = unassignedUsers.map(function(id) { return " ・ " + (userMap[id] ? userMap[id].name : id); }).join("\n");
      msg += "\n\n⚠️ 【送りきれなかった人: " + unassignedUsers.length + "名】\n" + unassignedNames;
    }
    Browser.msgBox("完了", msg, Browser.Buttons.OK);
  }
}

// ==========================================
// 機能5：【マスタ自動化】住所 ➔ 緯度経度一発変換
// ==========================================
function convertAddressToLatLng() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var basicSheet = ss.getSheetByName("m_user_basics"); if (!basicSheet) return;
  var basicData = basicSheet.getDataRange().getValues();
  var successCount = 0;
  for (var i = 2; i < basicData.length; i++) {
    var address = basicData[i][3] ? String(basicData[i][3]).trim() : ""; 
    var currentLat = basicData[i][4];
    var currentLng = basicData[i][5];
    
    // ⭐【修正】緯度か経度の「どちらか片方でも」空文字（空欄）なら、確実に再取得を走らせる
    if (address && address !== "" && (currentLat === "" || currentLng === "")) {
      try {
        var response = Maps.newGeocoder().geocode(address);
        if (response.status === "OK" && response.results && response.results.length > 0) {
          var loc = response.results[0].geometry.location;
          basicSheet.getRange(i + 1, 5).setValue(loc.lat); 
          basicSheet.getRange(i + 1, 6).setValue(loc.lng);
          successCount++; Utilities.sleep(100);
        }
      } catch (e) { console.error("エラー: " + e.message); }
    }
  }
  if (successCount > 0) Browser.msgBox("完了", successCount + " 名の緯度経度を自動取得しました！", Browser.Buttons.OK);
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
