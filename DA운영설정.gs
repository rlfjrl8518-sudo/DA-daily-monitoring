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

// 당일현황 버튼: DB_RAW 기준으로 "조회일" 스냅샷을 로그에 적재하고, 로그 전체를 기준으로
// 대시보드를 처음부터 다시 그린다. 추이 대시보드 팝업은 여기서 자동으로 뜨지 않는다
// (DA운영현황 시트 상단에 별도로 놓은 버튼(그림)에서 showRecentTrendDashboard를 직접 호출해서 연다).
function updateDAReport() {
  saveMonitoringSnapshot();
  renderFullDashboard();
}

// 전일마감 버튼: DB_RAW 기준으로 "전일" 최종마감 스냅샷을 로그에 적재(기존 00:00 행 교체)하고,
// 로그 전체를 기준으로 대시보드를 처음부터 다시 그린다. 월요일에 금/토/일을 몰아 처리할 때는
// 전일 셀 값을 바꿔가며 이 함수를 순서대로 여러 번 실행한다. 추이 대시보드 팝업은 여기서
// 자동으로 뜨지 않는다.
function updateDAReport_final() {
  saveMonitoringSnapshot_final();
  renderFullDashboard();
}

// 조정사항 인덱스(매체별 보종별 조정 기록)를 DA운영설정 시트 안에 별도 컬럼으로 마련한다.
// 기존에 쓰이고 있는 마지막 컬럼에서 한 칸 띄워 새로 배치하므로 "조회일"/"전일" 같은 기존
// 헤더 위치와 겹치지 않는다. 헤더가 이미 있으면(재실행해도) 아무 것도 하지 않는다.
function setupAdjustmentLogColumns() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);

  var headerRow = settingSheet.getRange(1, 1, 1, settingSheet.getLastColumn()).getValues()[0];

  if (headerRow.indexOf("조정일자") !== -1) {
    Logger.log("조정사항 인덱스가 이미 설정되어 있습니다.");
    return;
  }

  var startCol = settingSheet.getLastColumn() + 2; // 버퍼 1칸
  var headers = ["조정일자", "조정시간", "매체", "보종", "카테고리", "세부내용"];

  settingSheet.getRange(1, startCol, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold");

  // 카테고리/시간 목록은 하드코딩된 유효성 검사가 아니라 시트 위 참조 범위로 두어서,
  // 나중에 사용자가 이 범위 안의 값만 고쳐도(늘리거나 줄여도) 드롭다운에 바로 반영되게 한다.
  var categories = ["예산 상향", "예산 하향", "입찰가 상향", "입찰가 하향", "타겟팅 추가", "타겟팅 제외", "소재 추가", "소재 제외", "기타"];
  var times = [];
  for (var h = 8; h <= 22; h++) {
    times.push((h < 10 ? "0" + h : h) + ":00");
  }

  var catListCol = startCol + headers.length + 1; // 버퍼 1칸 두고 우측에 배치
  var timeListCol = catListCol + 2; // 카테고리 목록에서 버퍼 1칸 두고 그 옆에 배치

  settingSheet.getRange(1, catListCol).setValue("카테고리 목록(수정 가능)").setFontWeight("bold");
  settingSheet.getRange(2, catListCol, categories.length, 1)
    .setValues(categories.map(function(c) { return [c]; }));

  settingSheet.getRange(1, timeListCol).setValue("시간 목록(수정 가능)").setFontWeight("bold");
  settingSheet.getRange(2, timeListCol, times.length, 1)
    .setValues(times.map(function(t) { return [t]; }))
    .setNumberFormat("@"); // 시간 값이 아니라 "14:00" 같은 문자열 그대로 저장/표시

  var timeCol = startCol + 1; // headers 배열에서 "조정시간"의 위치(0-based 1)
  var timeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settingSheet.getRange(2, timeListCol, times.length, 1), true)
    .setAllowInvalid(true)
    .build();
  settingSheet.getRange(2, timeCol, 2000, 1).setDataValidation(timeRule).setNumberFormat("@");

  var categoryCol = startCol + 4; // headers 배열에서 "카테고리"의 위치(0-based 4)
  var catRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settingSheet.getRange(2, catListCol, categories.length, 1), true)
    .setAllowInvalid(true)
    .build();
  settingSheet.getRange(2, categoryCol, 2000, 1).setDataValidation(catRule);

  // 세부내용은 "10%"처럼 숫자+기호로 보이는 값도 자주 들어가는데, 서식을 안 정해두면 시트가
  // 이를 숫자(0.1, 퍼센트 서식)로 자동 변환해버려 나중에 읽을 때 "10%"가 아니라 "0.1"로
  // 보이는 문제가 있었다. 텍스트 서식으로 고정해 애초에 자동 변환이 안 일어나게 한다.
  var detailCol = startCol + 5; // headers 배열에서 "세부내용"의 위치(0-based 5)
  settingSheet.getRange(2, detailCol, 2000, 1).setNumberFormat("@");

  Logger.log("조정사항 인덱스 컬럼을 " + startCol + "번째 컬럼부터 설정했습니다.");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DA 대시보드')
    .addItem('당일 현황 업데이트', 'updateDAReport')
    .addItem('전일 마감 확인', 'updateDAReport_final')
    .addItem('전체 기간 재검증 (과거 로그 수정 반영)', 'renderFullDashboard')
    .addItem('최근 ' + TREND_DASHBOARD_DAYS + '일 추이 대시보드 보기', 'showRecentTrendDashboard')
    .addSeparator()
    .addItem('조정사항 인덱스 설정 (최초 1회)', 'setupAdjustmentLogColumns')
    .addToUi();
}
