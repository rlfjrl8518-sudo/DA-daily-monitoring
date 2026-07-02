var DASH_SHEET_NAME = "DA운영현황";
var SETTING_SHEET_NAME = "DA운영설정";
var CHART_HELPER_SHEET_NAME = "DA차트데이터";
var MEDIA_FILTER_CELL = "B1";
var PRODUCT_FILTER_CELL = "D1";
var ALL_LABEL = "전체";
var CHART_ROWS_PER_CHART = 14;
var CHART_START_ROW = 3; // 1행 필터 + 여백 다음부터 차트 시작
var CHART_WIDTH = 900;
var CHART_HEIGHT = 260;
var TIMEZONE = "Asia/Seoul";

function buildDashboard_v2() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName(DASH_SHEET_NAME);
  var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);

  // 재실행 시에도 기존에 선택해둔 필터 값 유지
  var prevMediaFilter = dashboardSheet.getRange(MEDIA_FILTER_CELL).getValue();
  var prevProductFilter = dashboardSheet.getRange(PRODUCT_FILTER_CELL).getValue();

  // clear()는 차트 객체를 지우지 않으므로 먼저 제거 (중복 삽입 방지)
  removeAllCharts_(dashboardSheet);

  dashboardSheet.clear();

  var data = loadSettingData_(settingSheet);
  var logs = data.logs;
  var mediaOrder = data.mediaOrder;
  var activeProducts = data.activeProducts;

  // ---- 최상단: 필터 컨트롤 + 최근 7일 비용/DB/단가 차트 ----
  var contentStartRow = renderTop7DayCharts_(
    ss, dashboardSheet, logs, mediaOrder, activeProducts,
    prevMediaFilter, prevProductFilter
  );

  // ---- 기존 일자별 상세 표 ----
  var sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 32); // 33일 기준

  var dateMap = {};

  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    var logDateOnly = new Date(formatDateKey_(logDate));

    if (logDateOnly < sevenDaysAgo) continue;

    var dateKey = formatDateKey_(logDate);

    if (!dateMap[dateKey]) dateMap[dateKey] = [];
    dateMap[dateKey].push(logs[i]);
  }

  var row = contentStartRow;

  Object.keys(dateMap).sort().forEach(function(dateKey) {

    var dayLogs = dateMap[dateKey];

    dashboardSheet.getRange(row, 1)
      .setValue(dateKey)
      .setBackground("#444444")
      .setFontColor("white")
      .setFontWeight("bold");

    row++;

    var times = [];

    dayLogs.forEach(function(log) {
      var t = formatTimeKey_(log[0]);
      if (times.indexOf(t) === -1) times.push(t);
    });

    times.sort();

    // [최종마감] 00:00 맨 앞 고정
    var finalIdx = times.indexOf("00:00");
    if (finalIdx !== -1) {
      times.splice(finalIdx, 1);
      times.unshift("00:00");
    }

    // dayLogs를 Map으로 미리 인덱싱 (시간_매체_보종 → log)
    var logMap = {};

    dayLogs.forEach(function(log) {
      var t = formatTimeKey_(log[0]);
      var key = t + "_" + String(log[1]).trim() + "_" + String(log[2]).trim();
      logMap[key] = log;
    });

    var header1 = ["", ""];
    var header2 = ["매체", "보종"];

    times.forEach(function(t) {
      var label = (t === "00:00") ? "[최종마감]" : t;
      header1.push(label);
      header1.push("");
      header1.push("");
      header2.push("비용");
      header2.push("DB");
      header2.push("단가");
    });

    dashboardSheet.getRange(row, 1, 1, header1.length)
      .setValues([header1])
      .setNumberFormat("@");

    row++;

    dashboardSheet.getRange(row, 1, 1, header2.length)
      .setValues([header2])
      .setBackground("#EFEFEF")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    row++;

    var grandCost = {};
    var grandDb = {};

    times.forEach(function(t) {
      grandCost[t] = 0;
      grandDb[t] = 0;
    });

    mediaOrder.forEach(function(media) {

      // 해당 날짜 로그에 실제 데이터가 있는 것만 표시
      var mediaRows = activeProducts.filter(function(x) {
        if (x.media !== media) return false;
        return times.some(function(t) {
          var key = t + "_" + x.media + "_" + x.product;
          return !!logMap[key];
        });
      });

      if (mediaRows.length === 0) return;

      var mediaCost = {};
      var mediaDb = {};

      times.forEach(function(t) {
        mediaCost[t] = 0;
        mediaDb[t] = 0;
      });

      var mediaStartRow = row;

      mediaRows.forEach(function(item) {

        var rowData = [item.media, item.product];

        times.forEach(function(t) {

          var key = t + "_" + item.media + "_" + item.product;
          var found = logMap[key];

          if (found) {
            var cost = Number(found[3]) || 0;
            var db = Number(found[4]) || 0;
            var cpa = Number(found[5]) || 0;

            rowData.push(cost);
            rowData.push(db);
            rowData.push(cpa);

            mediaCost[t] += cost;
            mediaDb[t] += db;
            grandCost[t] += cost;
            grandDb[t] += db;
          } else {
            rowData.push("");
            rowData.push("");
            rowData.push("");
          }

        });

        dashboardSheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);

        // 숫자 서식 적용 (C열부터)
        dashboardSheet.getRange(row, 3, 1, rowData.length - 2).setNumberFormat("#,##0");

        var dataIndex = 4;
        var sheetCol = 5;

        times.forEach(function(t) {
          var cpaValue = Number(rowData[dataIndex]);

          if (!isNaN(cpaValue) && cpaValue > 0) {
            var cell = dashboardSheet.getRange(row, sheetCol);
            cell.setNote("실제CPA=" + cpaValue + " / 목표CPA=" + item.targetCPA);

            if (cpaValue <= item.targetCPA) {
              cell.setBackground("#B6D7A8");
            } else if (cpaValue <= item.targetCPA * 1.2) {
              cell.setBackground("#FFD966");
            } else {
              cell.setBackground("#EA9999");
            }
          }

          dataIndex += 3;
          sheetCol += 3;
        });

        row++;

      });

      if (mediaRows.length > 1) {
        dashboardSheet.getRange(mediaStartRow, 1, mediaRows.length, 1)
          .mergeVertically()
          .setVerticalAlignment("middle")
          .setHorizontalAlignment("center")
          .setFontWeight("bold");
      }

      var subtotal = [media + " 소계", ""];

      times.forEach(function(t) {
        var cost = mediaCost[t];
        var db = mediaDb[t];
        subtotal.push(cost);
        subtotal.push(db);
        subtotal.push(db > 0 ? Math.round(cost / db) : "");
      });

      dashboardSheet.getRange(row, 1, 1, subtotal.length)
        .setValues([subtotal])
        .setBackground("#EAEAEA")
        .setFontWeight("bold");

      // 숫자 서식 적용 (C열부터)
      dashboardSheet.getRange(row, 3, 1, subtotal.length - 2).setNumberFormat("#,##0");

      row++;

    });

    var totalRow = ["전체합계", ""];

    times.forEach(function(t) {
      var cost = grandCost[t];
      var db = grandDb[t];
      totalRow.push(cost);
      totalRow.push(db);
      totalRow.push(db > 0 ? Math.round(cost / db) : "");
    });

    dashboardSheet.getRange(row, 1, 1, totalRow.length)
      .setValues([totalRow])
      .setBackground("#D9EAD3")
      .setFontWeight("bold");

    // 숫자 서식 적용 (C열부터)
    dashboardSheet.getRange(row, 3, 1, totalRow.length - 2).setNumberFormat("#,##0");

    row += 3;

  });

  // A열 볼드 + 가운데 정렬 (필터/차트 영역 제외, 상세 표 영역만)
  var lastRowNum = dashboardSheet.getLastRow();
  if (lastRowNum >= contentStartRow) {
    dashboardSheet.getRange(contentStartRow, 1, lastRowNum - contentStartRow + 1, 1)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
  }

  // 열 너비
  dashboardSheet.setColumnWidth(1, 120);
  dashboardSheet.setColumnWidth(2, 130);

  var lastCol = dashboardSheet.getLastColumn();
  if (lastCol >= 3) {
    dashboardSheet.setColumnWidths(3, lastCol - 2, 80);
  }

}

// 설정 시트 원본 데이터에서 매체 정렬 순서 / 매체·보종·목표CPA 목록을 뽑아낸다
function getMediaAndProducts_(settings) {

  var mediaOrderRaw = [];
  var activeProducts = [];

  for (var i = 1; i < settings.length; i++) {
    var sortNo = settings[i][8];
    var mediaName = settings[i][9];

    if (sortNo !== "" && mediaName !== "") {
      mediaOrderRaw.push({
        order: Number(sortNo),
        media: String(mediaName).trim()
      });
    }
  }

  mediaOrderRaw.sort(function(a, b) { return a.order - b.order; });
  var mediaOrder = mediaOrderRaw.map(function(x) { return x.media; });

  // 사용여부 관계없이 모든 매체/보종 로드 (targetCPA 참조용)
  for (var i = 1; i < settings.length; i++) {
    if (settings[i][0] === "" || settings[i][1] === "") continue;
    activeProducts.push({
      media: String(settings[i][0]).trim(),
      product: String(settings[i][1]).trim(),
      targetCPA: Number(settings[i][3]) || 0
    });
  }

  return { mediaOrder: mediaOrder, activeProducts: activeProducts };
}

// 설정 시트에서 로그(L~Q열) + 매체 정렬 순서 + 매체·보종 목록을 한 번에 읽어온다.
// (buildDashboard_v2 전체 실행 시와 필터 변경으로 인한 차트만 갱신할 때 공통으로 사용)
function loadSettingData_(settingSheet) {
  var settings = settingSheet.getDataRange().getValues();
  var lastRow = settingSheet.getLastRow();
  var logs = settingSheet.getRange(1, 12, lastRow, 6).getValues();
  var meta = getMediaAndProducts_(settings);

  return { logs: logs, mediaOrder: meta.mediaOrder, activeProducts: meta.activeProducts };
}

// 시트에 삽입되어 있는 차트를 모두 제거한다. clear()가 차트 객체까지 지우지는 않으므로,
// 다시 그리기 전에 호출해 중복 삽입을 막는다.
function removeAllCharts_(sheet) {
  sheet.getCharts().forEach(function(c) { sheet.removeChart(c); });
}

function formatDateKey_(date) {
  return Utilities.formatDate(date, TIMEZONE, "yyyy-MM-dd");
}

function formatTimeKey_(date) {
  return Utilities.formatDate(date, TIMEZONE, "HH:mm");
}

// 대시보드 최상단에 매체/보종 필터 컨트롤과 최근 7일(최종마감 기준) 비용·DB·단가 차트 3개를 그린다.
// 반환값: 이 영역 다음에 기존 상세 표를 그리기 시작할 행 번호
function renderTop7DayCharts_(ss, dashboardSheet, logs, mediaOrder, activeProducts, prevMediaFilter, prevProductFilter) {

  // ---- 필터 옵션 목록 ----
  var mediaOptions = [ALL_LABEL].concat(mediaOrder);

  var productSeen = {};
  var productOptions = [ALL_LABEL];
  activeProducts.forEach(function(p) {
    if (!productSeen[p.product]) {
      productSeen[p.product] = true;
      productOptions.push(p.product);
    }
  });

  var mediaFilter = mediaOptions.indexOf(prevMediaFilter) !== -1 ? prevMediaFilter : ALL_LABEL;
  var productFilter = productOptions.indexOf(prevProductFilter) !== -1 ? prevProductFilter : ALL_LABEL;

  // ---- 필터 컨트롤 UI (1행) ----
  dashboardSheet.getRange("A1").setValue("매체 필터:").setFontWeight("bold");
  dashboardSheet.getRange("C1").setValue("보종 필터:").setFontWeight("bold");

  var mediaCell = dashboardSheet.getRange(MEDIA_FILTER_CELL);
  mediaCell.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(mediaOptions, true).setAllowInvalid(false).build()
  );
  mediaCell.setValue(mediaFilter).setBackground("#FFF2CC");

  var productCell = dashboardSheet.getRange(PRODUCT_FILTER_CELL);
  productCell.setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(productOptions, true).setAllowInvalid(false).build()
  );
  productCell.setValue(productFilter).setBackground("#FFF2CC");

  // ---- 필터가 적용된 매체+보종 조합 목록 (차트 시리즈) ----
  var combos = [];
  var comboSeen = {};

  mediaOrder.forEach(function(media) {
    if (mediaFilter !== ALL_LABEL && media !== mediaFilter) return;

    activeProducts.forEach(function(p) {
      if (p.media !== media) return;
      if (productFilter !== ALL_LABEL && p.product !== productFilter) return;

      var key = p.media + "_" + p.product;
      if (comboSeen[key]) return;
      comboSeen[key] = true;
      combos.push({ media: p.media, product: p.product, label: p.media + " " + p.product, key: key });
    });
  });

  // ---- 조회일(오늘) 기준 최근 7일 날짜 ----
  var today = new Date();
  var last7Dates = [];
  for (var d = 6; d >= 0; d--) {
    var dt = new Date(today);
    dt.setDate(dt.getDate() - d);
    last7Dates.push(formatDateKey_(dt));
  }
  var last7Set = {};
  last7Dates.forEach(function(dateKey) { last7Set[dateKey] = true; });

  // ---- 일자별 최종마감(00:00) 로그만 집계 ----
  var finalMap = {};

  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    var t = formatTimeKey_(logDate);
    if (t !== "00:00") continue;

    var dateKey = formatDateKey_(logDate);
    if (!last7Set[dateKey]) continue;

    var key = dateKey + "_" + String(logs[i][1]).trim() + "_" + String(logs[i][2]).trim();
    finalMap[key] = {
      cost: Number(logs[i][3]) || 0,
      db: Number(logs[i][4]) || 0
    };
  }

  // ---- 차트용 데이터 테이블 3개 (비용 / DB / 단가) ----
  var header = ["날짜"].concat(combos.map(function(c) { return c.label; }));

  var costTable = [header.slice()];
  var dbTable = [header.slice()];
  var cpaTable = [header.slice()];

  last7Dates.forEach(function(dateKey) {
    var costRow = [dateKey];
    var dbRow = [dateKey];
    var cpaRow = [dateKey];

    combos.forEach(function(c) {
      var found = finalMap[dateKey + "_" + c.key];
      costRow.push(found ? found.cost : "");
      dbRow.push(found ? found.db : "");
      cpaRow.push(found && found.db > 0 ? Math.round(found.cost / found.db) : "");
    });

    costTable.push(costRow);
    dbTable.push(dbRow);
    cpaTable.push(cpaRow);
  });

  // ---- 숨김 헬퍼 시트에 원본 데이터 기록 (차트가 참조할 range) ----
  var helperSheet = ss.getSheetByName(CHART_HELPER_SHEET_NAME);
  if (!helperSheet) {
    helperSheet = ss.insertSheet(CHART_HELPER_SHEET_NAME);
  }
  helperSheet.clear();
  helperSheet.hideSheet();

  var nCols = header.length;
  var cursorRow = 1;

  var ranges = [costTable, dbTable, cpaTable].map(function(table) {
    var range = helperSheet.getRange(cursorRow, 1, table.length, nCols);
    range.setValues(table);
    cursorRow += table.length + 1; // 표 사이 빈 행 1개
    return range;
  });

  var costRange = ranges[0];
  var dbRange = ranges[1];
  var cpaRange = ranges[2];

  // ---- 차트 3개 삽입 (비용 / DB / 단가, 세로로 쌓음) ----
  var suffix = filterSuffix_(mediaFilter, productFilter);

  var chartConfigs = [
    { range: costRange, title: '최근 7일 비용 추이 (최종마감 기준)' + suffix, anchorRow: CHART_START_ROW },
    { range: dbRange, title: '최근 7일 DB 추이 (최종마감 기준)' + suffix, anchorRow: CHART_START_ROW + CHART_ROWS_PER_CHART },
    { range: cpaRange, title: '최근 7일 단가(CPA) 추이 (최종마감 기준)' + suffix, anchorRow: CHART_START_ROW + CHART_ROWS_PER_CHART * 2 }
  ];

  chartConfigs.forEach(function(cfg) {
    dashboardSheet.insertChart(buildLineChart_(dashboardSheet, cfg.range, cfg.title, cfg.anchorRow));
  });

  return CHART_START_ROW + CHART_ROWS_PER_CHART * 3 + 2; // 이 다음 행부터 기존 상세 표 렌더링
}

function buildLineChart_(sheet, range, title, anchorRow) {
  return sheet.newChart()
    .asLineChart()
    .addRange(range)
    .setNumHeaders(1)
    .setOption('title', title)
    .setOption('width', CHART_WIDTH)
    .setOption('height', CHART_HEIGHT)
    .setOption('useFirstColumnAsDomain', true)
    .setPosition(anchorRow, 1, 0, 0)
    .build();
}

function filterSuffix_(mediaFilter, productFilter) {
  var parts = [];
  if (mediaFilter !== ALL_LABEL) parts.push(mediaFilter);
  if (productFilter !== ALL_LABEL) parts.push(productFilter);
  return parts.length ? " - " + parts.join(" / ") : "";
}

// 매체/보종 필터 셀(B1, D1)을 사용자가 직접 수정하면 차트 3개만 빠르게 다시 그린다.
// (전체 상세 표는 다시 그리지 않음 — buildDashboard_v2 재실행보다 훨씬 빠름)
function onEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== DASH_SHEET_NAME) return;

  var editedCell = e.range.getA1Notation();
  if (editedCell !== MEDIA_FILTER_CELL && editedCell !== PRODUCT_FILTER_CELL) return;

  refreshTop7DayCharts_();
}

function refreshTop7DayCharts_() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName(DASH_SHEET_NAME);
  var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);

  var mediaFilter = dashboardSheet.getRange(MEDIA_FILTER_CELL).getValue();
  var productFilter = dashboardSheet.getRange(PRODUCT_FILTER_CELL).getValue();

  var data = loadSettingData_(settingSheet);

  removeAllCharts_(dashboardSheet);

  renderTop7DayCharts_(ss, dashboardSheet, data.logs, data.mediaOrder, data.activeProducts, mediaFilter, productFilter);
}
