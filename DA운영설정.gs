// 카탈로그로 운영되는 매체(보종별 DB는 들어오지만 매체 전체로만 비용이 청구되어, 설정 시트에는
// 편의상 매체 전체 비용이 보종 행마다 그대로 중복 입력돼 있음). buildSnapshotResults_에서
// 이 매체들은 보종별 행의 비용을 그대로 더하지 않고, 매체 전체 비용을 CATALOG_TOTAL_PRODUCT
// 이름으로 한 번만 별도 적재해 DA운영현황의 매체 소계에서만 반영되도록 한다.
var CATALOG_MEDIA = ["크리테오 다이나믹"];
var CATALOG_TOTAL_PRODUCT = "__전체__";

// settings(DA운영설정 원본 값)와 rawMap("날짜_매체_보종" → DB건수)을 받아 saveDate 시각으로
// 찍을 로그 행(날짜, 매체, 보종, 비용, DB, 단가)들을 계산한다.
// saveMonitoringSnapshot()과 saveMonitoringSnapshot_final()이 공유해서 쓴다.
//
// CATALOG_MEDIA에 속한 매체는 보종별 비용을 나눌 수 없으므로, 보종별 행은 비용/단가를 비워
// DB 확인용으로만 남기고, 매체 전체 비용은 매체당 한 번만 CATALOG_TOTAL_PRODUCT 보종으로
// 별도 적재한다 (DA운영현황.gs의 renderDateBlock_이 이 값을 매체 소계에만 반영한다).
function buildSnapshotResults_(settings, rawMap, saveDate) {

  var results = [];
  var catalogTotals = {}; // media -> { cost: 매체 전체 비용, db: 보종별 DB 합계 }

  for (var i = 1; i < settings.length; i++) {

    var media = String(settings[i][0]).trim();
    var product = String(settings[i][1]).trim();
    var active = settings[i][2];
    var inputCost = Number(settings[i][4]) || 0;
    var vat = settings[i][5];
    var markupRate = Number(settings[i][6]) || 0;

    if (active !== true) continue;

    var key = media + "_" + product;
    var dbCount = rawMap[key] || 0;

    var finalCost = inputCost;
    if (String(vat).trim() !== "포함") finalCost = finalCost * 1.1;
    finalCost = finalCost * (1 + markupRate);
    finalCost = Math.round(finalCost);

    if (CATALOG_MEDIA.indexOf(media) !== -1) {
      results.push([saveDate, media, product, "", dbCount, ""]);

      // 보종 행마다 매체 전체 비용이 그대로(동일하게) 입력돼 있다는 전제이므로, 어느 행에
      // 입력하든 상관없도록 그 중 최댓값을 매체 전체 비용으로 채택한다.
      if (!catalogTotals[media]) catalogTotals[media] = { cost: 0, db: 0 };
      catalogTotals[media].cost = Math.max(catalogTotals[media].cost, finalCost);
      catalogTotals[media].db += dbCount;
      continue;
    }

    var cpa = dbCount > 0 ? Math.round(finalCost / dbCount) : 0;
    results.push([saveDate, media, product, finalCost, dbCount, cpa]);
  }

  Object.keys(catalogTotals).forEach(function(media) {
    var totals = catalogTotals[media];
    var cpa = totals.db > 0 ? Math.round(totals.cost / totals.db) : 0;
    results.push([saveDate, media, CATALOG_TOTAL_PRODUCT, totals.cost, totals.db, cpa]);
  });

  return results;
}

// 단계별 소요 시간을 실행 로그(Apps Script 편집기 "실행 기록"/Logger)에 남기는 헬퍼.
// 어느 구간이 느린지 추측 대신 실측으로 확인하기 위한 용도로, 로직에는 영향을 주지 않는다.
function _perfLog(label, startMs) {
  Logger.log("[perf] " + label + ": " + (Date.now() - startMs) + "ms");
}

// L열(로그 날짜)에 데이터가 있는 마지막 행 다음 번호를 찾는다. 예전에는 "L:L"(시트에 잡혀있는
// 전체 행 범위 — 실제 로그 행 수보다 훨씬 큰 경우가 흔함)을 통째로 읽어서 느렸는데, getLastRow()로
// 실제 마지막 행까지만 읽도록 범위를 좁혔다. saveMonitoringSnapshot()/saveMonitoringSnapshot_final()이
// 공유해서 쓴다.
function findNextLogRow_(settingSheet) {
  var lastRow = settingSheet.getLastRow();
  if (lastRow < 2) return 2;

  var logColumn = settingSheet.getRange(2, 12, lastRow - 1, 1).getValues().flat();

  for (var r = logColumn.length - 1; r >= 0; r--) {
    if (logColumn[r] !== "") return r + 3; // logColumn[r]는 시트의 (r+2)행이므로, 다음 빈 행은 그 다음
  }

  return 2;
}

// DB_RAW 시트의 원본 전환 데이터를 집계해서 DA운영설정 시트 로그(L~Q열)에
// "조회일" 기준 스냅샷 한 묶음을 새로 적재한다. (당일현황 버튼)
function saveMonitoringSnapshot() {

  var tStart = Date.now();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingSheet = ss.getSheetByName("DA운영설정");
  var rawSheet = ss.getSheetByName("DB_RAW");

  var settings = settingSheet.getDataRange().getValues();
  var rawData = rawSheet.getDataRange().getValues();
  _perfLog("설정/DB_RAW 읽기 (DB_RAW " + rawData.length + "행)", tStart);

  var now = new Date();
  var currentTime = Utilities.formatDate(now, "Asia/Seoul", "HH:mm:ss");

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

  // rawData를 Map으로 미리 인덱싱 (날짜_매체_보종 → count)
  var tRawMap = Date.now();
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
  _perfLog("DB_RAW 집계(rawMap 구성)", tRawMap);

  var saveDate = new Date(targetDateStr + " " + currentTime);
  var results = buildSnapshotResults_(settings, rawMap, saveDate);

  if (results.length === 0) {
    Logger.log("저장할 데이터 없음");
    return;
  }

  var tFindRow = Date.now();
  var startRow = findNextLogRow_(settingSheet);
  _perfLog("다음 빈 행 탐색", tFindRow);

  var tWrite = Date.now();
  settingSheet.getRange(startRow, 12, results.length, 6).setValues(results);
  // L열(로그 날짜)에 새로 확장된 행은 서식이 비어 있거나 "숫자"로 남아있어 Date 값이
  // 46212 같은 시리얼 넘버로 표시되는 문제가 있어, 쓸 때마다 날짜 서식을 명시적으로 고정한다.
  settingSheet.getRange(startRow, 12, results.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  _perfLog("로그 " + results.length + "건 쓰기", tWrite);

  Logger.log(results.length + "건 저장 완료 (전체 " + (Date.now() - tStart) + "ms)");

}

// DB_RAW 시트의 원본 전환 데이터를 집계해서 DA운영설정 시트 로그(L~Q열)에
// "전일" 기준 최종마감(00:00) 스냅샷을 새로 적재한다. (전일마감 버튼)
// 같은 날짜의 기존 00:00 최종마감 행은 먼저 삭제한 뒤 다시 넣으므로, 같은 전일 날짜로
// 여러 번 눌러도(정정 포함) 중복되지 않는다.
function saveMonitoringSnapshot_final() {

  var tStart = Date.now();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingSheet = ss.getSheetByName("DA운영설정");
  var rawSheet = ss.getSheetByName("DB_RAW");

  var settings = settingSheet.getDataRange().getValues();
  var rawData = rawSheet.getDataRange().getValues();
  _perfLog("설정/DB_RAW 읽기 (DB_RAW " + rawData.length + "행)", tStart);

  var now = new Date();

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
  var tDelete = Date.now();
  var lastRow = settingSheet.getLastRow();
  var logData = lastRow >= 2 ? settingSheet.getRange(2, 12, lastRow - 1, 6).getValues() : [];

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

  // rowsToDelete는 이미 내림차순으로 모이지만, 연속된 행 구간을 하나로 묶어 deleteRows()를
  // 구간당 한 번씩만 호출한다. 예전에는 지울 행마다 deleteRow()를 개별 호출했는데, 행 삭제는
  // 그 아래 모든 행을 밀어올리는 무거운 연산이라 여러 번 반복하면 특히 느렸다(보통 한 번의
  // 최종마감 저장이 여러 매체/보종에 걸쳐 연속된 행으로 찍히므로 대부분 한 구간으로 묶인다).
  var di = 0;
  while (di < rowsToDelete.length) {
    var runEnd = rowsToDelete[di];
    var dj = di;
    while (dj + 1 < rowsToDelete.length && rowsToDelete[dj + 1] === rowsToDelete[dj] - 1) {
      dj++;
    }
    var runStart = rowsToDelete[dj];
    settingSheet.deleteRows(runStart, runEnd - runStart + 1);
    di = dj + 1;
  }
  _perfLog("기존 최종마감(00:00) " + rowsToDelete.length + "건 삭제", tDelete);

  // rawData를 Map으로 미리 인덱싱
  var tRawMap = Date.now();
  var rawMap = {};

  for (var k = 1; k < rawData.length; k++) {
    var rawDate = rawData[k][3];
    if (!rawDate) continue;

    var rawDateStr = Utilities.formatDate(new Date(rawDate), "Asia/Seoul", "yyyy-MM-dd");
    if (rawDateStr !== targetDateStr) continue;

    var rawMedia = String(rawData[k][11]).trim();
    var rawProduct = String(rawData[k][12]).trim();
    var key = rawMedia + "_" + rawProduct;

    rawMap[key] = (rawMap[key] || 0) + 1;
  }
  _perfLog("DB_RAW 집계(rawMap 구성)", tRawMap);

  var saveDate = new Date(targetDateStr + " 00:00:00");
  var results = buildSnapshotResults_(settings, rawMap, saveDate);

  if (results.length === 0) {
    Logger.log("저장할 데이터 없음");
    return;
  }

  var tFindRow = Date.now();
  var startRow = findNextLogRow_(settingSheet);
  _perfLog("다음 빈 행 탐색", tFindRow);

  var tWrite = Date.now();
  settingSheet.getRange(startRow, 12, results.length, 6).setValues(results);
  // L열(로그 날짜)에 새로 확장된 행은 서식이 비어 있거나 "숫자"로 남아있어 Date 값이
  // 46212 같은 시리얼 넘버로 표시되는 문제가 있어, 쓸 때마다 날짜 서식을 명시적으로 고정한다.
  settingSheet.getRange(startRow, 12, results.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  _perfLog("로그 " + results.length + "건 쓰기", tWrite);

  Logger.log(results.length + "건 최종마감 저장 완료 (전체 " + (Date.now() - tStart) + "ms)");

}

// 당일현황 버튼: DB_RAW 기준으로 "조회일" 스냅샷을 로그에 적재하고, 대시보드는 최근
// DASHBOARD_RECENT_WINDOW_DAYS일 구간만 다시 그린다(renderRecentDashboard, DA운영현황.gs 참고).
// 과거 로그를 직접 수정한 경우엔 메뉴의 "전체 기간 재검증"을 별도로 실행해야 한다. 추이 대시보드
// 팝업은 여기서 자동으로 뜨지 않는다(DA운영현황 시트 상단에 별도로 놓은 버튼(그림)에서
// showRecentTrendDashboard를 직접 호출해서 연다).
function updateDAReport() {
  saveMonitoringSnapshot();
  renderRecentDashboard();
}

// 전일마감 버튼: DB_RAW 기준으로 "전일" 최종마감 스냅샷을 로그에 적재(기존 00:00 행 교체)하고,
// 대시보드는 최근 DASHBOARD_RECENT_WINDOW_DAYS일 구간만 다시 그린다. 월요일에 금/토/일을 몰아
// 처리할 때는 전일 셀 값을 바꿔가며 이 함수를 순서대로 여러 번 실행하는데, 그 3일 모두 최근
// 구간(기본 7일) 안에 들어오므로 매번 정상적으로 갱신된다. 추이 대시보드 팝업은 여기서 자동으로
// 뜨지 않는다.
function updateDAReport_final() {
  saveMonitoringSnapshot_final();
  renderRecentDashboard();
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

// L열(로그 날짜)에 이미 찍혀 있던 값들이 서식 누락으로 46212 같은 시리얼 넘버로 보이는
// 기존 문제를 한 번에 복구한다. saveMonitoringSnapshot()/saveMonitoringSnapshot_final()이
// 이제 매번 서식을 고정하므로, 이 함수는 과거에 이미 잘못 찍힌 행들을 위한 1회성 복구용이다.
function fixLogDateFormat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);

  var lastRow = settingSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("복구할 로그 행이 없습니다.");
    return;
  }

  settingSheet.getRange(2, 12, lastRow - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  Logger.log("L열 날짜 서식을 " + (lastRow - 1) + "행에 복구했습니다.");
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DA 대시보드')
    .addItem('당일 현황 업데이트', 'updateDAReport')
    .addItem('전일 마감 확인', 'updateDAReport_final')
    .addItem('전체 기간 재검증 (과거 로그 수정 반영)', 'renderFullDashboard')
    .addItem('최근 ' + TREND_DASHBOARD_DEFAULT_DAYS + '일 추이 대시보드 보기', 'showRecentTrendDashboard')
    .addSeparator()
    .addItem('조정사항 인덱스 설정 (최초 1회)', 'setupAdjustmentLogColumns')
    .addItem('로그 날짜 서식 복구 (1회성)', 'fixLogDateFormat')
    .addToUi();
}
