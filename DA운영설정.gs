// DB_RAW 시트의 원본 전환 데이터를 집계해서 DA운영설정 시트 로그(L~Q열)에
// "조회일" 기준 스냅샷 한 묶음을 새로 적재한다. (당일현황 버튼)
function saveMonitoringSnapshot() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingSheet = ss.getSheetByName("DA운영설정");
  var rawSheet = ss.getSheetByName("DB_RAW");

  var settings = settingSheet.getDataRange().getValues();
  var rawData = rawSheet.getDataRange().getValues();

  var now = new Date();
  var currentTime = Utilities.formatDate(now, "Asia/Seoul", "HH:mm:ss");

  var results = [];

  var headerRow = settingSheet.getRange(1, 1, 1, settingSheet.getLastColumn()).getValues()[0];
  var targetDateCol = headerRow.indexOf("조회일");

  if (targetDateCol === -1) {
    throw new Error("DA운영설정 시트에서 '조회일' 헤더를 찾을 수 없습니다.");
  }

  var targetDate = settingSheet.getRange(2, targetDateCol + 1).getValue();

  if (!(targetDate instanceof Date)) {
    targetDate = new Date();
  }

  var targetDateStr = Utilities.formatDate(targetDate, "Asia/Seoul", "yyyy-MM-dd");
  var todayStr = Utilities.formatDate(now, "Asia/Seoul", "yyyy-MM-dd");

  // rawData를 Map으로 미리 인덱싱 (날짜_매체_보종 → count)
  var rawMap = {};

  for (var j = 1; j < rawData.length; j++) {
    var rawDate = rawData[j][3];
    if (!rawDate) continue;

    var rawDateStr = Utilities.formatDate(new Date(rawDate), "Asia/Seoul", "yyyy-MM-dd");
    if (rawDateStr !== targetDateStr) continue;

    var rawMedia = String(rawData[j][11]).trim();
    var rawProduct = String(rawData[j][12]).trim();
    var key = rawMedia + "_" + rawProduct;

    rawMap[key] = (rawMap[key] || 0) + 1;
  }

  for (var i = 1; i < settings.length; i++) {

    var media = settings[i][0];
    var product = settings[i][1];
    var active = settings[i][2];
    var inputCost = Number(settings[i][4]) || 0;
    var vat = settings[i][5];
    var markupRate = Number(settings[i][6]) || 0;

    if (active !== true) continue;

    var key = String(media).trim() + "_" + String(product).trim();
    var dbCount = rawMap[key] || 0;

    var finalCost = inputCost;
    if (String(vat).trim() !== "포함") finalCost = finalCost * 1.1;
    finalCost = finalCost * (1 + markupRate);
    finalCost = Math.round(finalCost);

    var cpa = dbCount > 0 ? Math.round(finalCost / dbCount) : 0;
    var saveDate = new Date(targetDateStr + " " + currentTime);

    results.push([saveDate, media, product, finalCost, dbCount, cpa]);
  }

  if (results.length === 0) {
    Logger.log("저장할 데이터 없음");
    return;
  }

  var logColumn = settingSheet.getRange("L:L").getValues().flat();
  var startRow = 2;

  for (var r = logColumn.length - 1; r >= 0; r--) {
    if (logColumn[r] !== "") {
      startRow = r + 2;
      break;
    }
  }

  settingSheet.getRange(startRow, 12, results.length, 6).setValues(results);
  Logger.log(results.length + "건 저장 완료");

}

// DB_RAW 시트의 원본 전환 데이터를 집계해서 DA운영설정 시트 로그(L~Q열)에
// "전일" 기준 최종마감(00:00) 스냅샷을 새로 적재한다. (전일마감 버튼)
// 같은 날짜의 기존 00:00 최종마감 행은 먼저 삭제한 뒤 다시 넣으므로, 같은 전일 날짜로
// 여러 번 눌러도(정정 포함) 중복되지 않는다.
function saveMonitoringSnapshot_final() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingSheet = ss.getSheetByName("DA운영설정");
  var rawSheet = ss.getSheetByName("DB_RAW");

  var settings = settingSheet.getDataRange().getValues();
  var rawData = rawSheet.getDataRange().getValues();

  var now = new Date();
  var results = [];

  var headerRow = settingSheet.getRange(1, 1, 1, settingSheet.getLastColumn()).getValues()[0];
  var targetDateCol = headerRow.indexOf("전일");

  if (targetDateCol === -1) {
    throw new Error("DA운영설정 시트에서 '전일' 헤더를 찾을 수 없습니다.");
  }

  var targetDate = settingSheet.getRange(2, targetDateCol + 1).getValue();

  if (!(targetDate instanceof Date)) {
    throw new Error("전일 날짜가 올바르지 않습니다.");
  }

  var targetDateStr = Utilities.formatDate(targetDate, "Asia/Seoul", "yyyy-MM-dd");

  // 기존 최종마감 데이터 삭제 (00:00 시간대만)
  var logData = settingSheet.getRange(2, 12, settingSheet.getLastRow() - 1, 6).getValues();

  var rowsToDelete = [];

  for (var r = logData.length - 1; r >= 0; r--) {
    var logDate = logData[r][0];
    if (!logDate) continue;

    var logDateStr = Utilities.formatDate(new Date(logDate), "Asia/Seoul", "yyyy-MM-dd");
    var logTime = Utilities.formatDate(new Date(logDate), "Asia/Seoul", "HH:mm");

    if (logDateStr === targetDateStr && logTime === "00:00") {
      rowsToDelete.push(r + 2);
    }
  }

  rowsToDelete.sort(function(a, b) { return b - a; });
  rowsToDelete.forEach(function(rowNum) {
    settingSheet.deleteRow(rowNum);
  });

  // rawData를 Map으로 미리 인덱싱
  var rawMap = {};

  for (var j = 1; j < rawData.length; j++) {
    var rawDate = rawData[j][3];
    if (!rawDate) continue;

    var rawDateStr = Utilities.formatDate(new Date(rawDate), "Asia/Seoul", "yyyy-MM-dd");
    if (rawDateStr !== targetDateStr) continue;

    var rawMedia = String(rawData[j][11]).trim();
    var rawProduct = String(rawData[j][12]).trim();
    var key = rawMedia + "_" + rawProduct;

    rawMap[key] = (rawMap[key] || 0) + 1;
  }

  for (var i = 1; i < settings.length; i++) {

    var media = settings[i][0];
    var product = settings[i][1];
    var active = settings[i][2];
    var inputCost = Number(settings[i][4]) || 0;
    var vat = settings[i][5];
    var markupRate = Number(settings[i][6]) || 0;

    if (active !== true) continue;

    var key = String(media).trim() + "_" + String(product).trim();
    var dbCount = rawMap[key] || 0;

    var finalCost = inputCost;
    if (String(vat).trim() !== "포함") finalCost = finalCost * 1.1;
    finalCost = finalCost * (1 + markupRate);
    finalCost = Math.round(finalCost);

    var cpa = dbCount > 0 ? Math.round(finalCost / dbCount) : 0;
    var saveDate = new Date(targetDateStr + " 00:00:00");

    results.push([saveDate, media, product, finalCost, dbCount, cpa]);
  }

  if (results.length === 0) {
    Logger.log("저장할 데이터 없음");
    return;
  }

  var logColumn = settingSheet.getRange("L:L").getValues().flat();
  var startRow = 2;

  for (var r = logColumn.length - 1; r >= 0; r--) {
    if (logColumn[r] !== "") {
      startRow = r + 2;
      break;
    }
  }

  settingSheet.getRange(startRow, 12, results.length, 6).setValues(results);
  Logger.log(results.length + "건 최종마감 저장 완료");

}

// 당일현황 버튼: DB_RAW 기준으로 "조회일" 스냅샷을 로그에 적재하고, 그 날짜 블록만 다시 그린다.
// 추이 대시보드 팝업은 여기서 자동으로 뜨지 않는다 (DA운영현황 시트 상단에 별도로 놓은
// 버튼(그림)에서 showRecentTrendDashboard를 직접 호출해서 연다).
function updateDAReport() {
  saveMonitoringSnapshot();
  renderDashboardForHeaderDate_("조회일");
}

// 전일마감 버튼: DB_RAW 기준으로 "전일" 최종마감 스냅샷을 로그에 적재(기존 00:00 행 교체)하고,
// 그 날짜 블록만 다시 그린다. 월요일에 금/토/일을 몰아 처리할 때는 전일 셀 값을 바꿔가며
// 이 함수를 순서대로 여러 번 실행한다. 추이 대시보드 팝업은 여기서 자동으로 뜨지 않는다.
function updateDAReport_final() {
  saveMonitoringSnapshot_final();
  renderDashboardForHeaderDate_("전일");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DA 대시보드')
    .addItem('당일 현황 업데이트', 'updateDAReport')
    .addItem('전일 마감 확인', 'updateDAReport_final')
    .addItem('최근 32일 재검증 (과거 로그 수정 반영)', 'rebuildRecentDashboard_v2')
    .addItem('최근 ' + TREND_DASHBOARD_DAYS + '일 추이 대시보드 보기', 'showRecentTrendDashboard')
    .addToUi();
}
