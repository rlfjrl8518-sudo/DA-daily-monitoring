var DASH_SHEET_NAME = "DA운영현황";
var SETTING_SHEET_NAME = "DA운영설정";
var BLOCK_STATE_PROPERTY_KEY = "DA_BLOCK_STATE";
var TIMEZONE = "Asia/Seoul";
var RECHECK_WINDOW_DAYS = 32; // 재검증 시 다시 그릴 최근 일수 (기존 33일 기준과 동일)

// 상단 "조회일 기준 최근 N일 최종마감 추이" 요약표 설정.
// 이 표는 실제 필요한 행 수를 매번 계산해서 그만큼만 차지하고, 그 아래 날짜별 상세 블록들은
// 늘어나거나 줄어드는 만큼 정확히 밀거나 당겨진다(고정 여백을 미리 크게 잡아두지 않음).
var TOP_SUMMARY_START_ROW = 1;
var TOP_SUMMARY_RECENT_DAYS = 3;
var TOP_SUMMARY_ROWS_PROPERTY_KEY = "DA_TOP_SUMMARY_ROWS"; // 지금 실제로 시트에 반영된 요약표 전체 행 수

// 날짜별 상세 블록이 하나도 없는 최초 상태에서, 새 블록을 어느 행부터 시작해야 하는지
// (요약표가 차지한 행 수 + 빈 줄 2개 다음)를 그때그때 계산한다. 요약표 크기가 매번 바뀌므로
// 고정 상수로 두지 않고 항상 현재 반영된 값을 문서 속성에서 읽어온다.
function getDetailContentBaseRow_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(TOP_SUMMARY_ROWS_PROPERTY_KEY);
  var appliedRows = raw !== null ? Number(raw) : 0;
  return TOP_SUMMARY_START_ROW + appliedRows + 2;
}

// 일회성 정리용: 예전에 있었다가 삭제된 매체/보종 필터 드롭다운의 데이터 확인 규칙이
// 그 후 여러 번의 행 삽입/삭제를 거치며 시트 곳곳으로 밀려 남아있을 수 있다. 그 잔재 때문에
// 스크립트가 일반 텍스트/숫자를 쓰려다가 "데이터 확인 규칙 위반" 오류가 날 수 있으므로,
// DA운영현황 시트 전체에서 데이터 확인 규칙만 한 번에 제거한다. 문제가 재발하면 다시 실행해도 안전하다.
function clearAllDataValidationsOnDashboard_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName(DASH_SHEET_NAME);
  var lastRow = Math.max(dashboardSheet.getMaxRows(), 1);
  var lastCol = Math.max(dashboardSheet.getMaxColumns(), 1);
  dashboardSheet.getRange(1, 1, lastRow, lastCol).clearDataValidations();
  Logger.log("DA운영현황 시트 전체의 데이터 확인 규칙을 제거했습니다.");
}

// DA운영설정 시트의 헤더(예: "조회일", "전일") 밑 2행에 있는 날짜 값을 기준으로,
// 그 날짜 하나의 블록만 다시 그린다. saveMonitoringSnapshot()/saveMonitoringSnapshot_final()이
// 로그를 적재한 뒤 어느 날짜를 다시 그려야 하는지는 항상 이 헤더 셀 값이 결정하므로,
// 시스템 시계로 "오늘"/"어제"를 추측하지 않는다 (운영자가 조회일/전일을 직접 입력·변경함).
// 다른 날짜(예: 월요일에 금/토/일을 순서대로)를 처리해야 할 때는 운영자가 헤더 셀 값을 바꿔가며
// 이 함수를 여러 번 호출하면 되므로, 요일 판단 로직이 따로 필요 없다.
// 주의: upsertDateBlock_는 "가장 최근(맨 아래) 블록"을 갱신하는 상황에서만 안전하다. 조회일/전일을
// 이미 지나간 지 오래된 과거 날짜(맨 아래가 아닌 중간 블록)로 설정해서 정정하는 용도로는 쓰지 말고,
// 그런 경우엔 rebuildRecentDashboard_v2()로 최근 32일 구간을 통째로 다시 그릴 것.
function renderDashboardForHeaderDate_(headerName) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName(DASH_SHEET_NAME);
  var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);

  var settings = settingSheet.getDataRange().getValues();
  var lastSettingRow = settingSheet.getLastRow();
  var logs = settingSheet.getRange(1, 12, lastSettingRow, 6).getValues();

  var meta = getMediaAndProducts_(settings);
  var mediaOrder = meta.mediaOrder;
  var activeProducts = meta.activeProducts;

  var targetDate = getHeaderDate_(settingSheet, headerName);
  var dateKey = Utilities.formatDate(targetDate, TIMEZONE, "yyyy-MM-dd");

  var dayLogs = [];
  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    if (Utilities.formatDate(logDate, TIMEZONE, "yyyy-MM-dd") === dateKey) {
      dayLogs.push(logs[i]);
    }
  }

  upsertDateBlock_(dashboardSheet, dateKey, dayLogs, mediaOrder, activeProducts);

  var summaryRefDate = getHeaderDate_(settingSheet, "조회일");
  renderRecentFinalSummary_(dashboardSheet, logs, mediaOrder, activeProducts, summaryRefDate);

  applyColumnWidths_(dashboardSheet);
}

// DA운영설정 시트의 헤더(예: "조회일", "전일") 밑 2행에 있는 날짜 값을 읽어온다.
// 값이 없거나 날짜가 아니면 오늘 날짜로 대체한다.
function getHeaderDate_(settingSheet, headerName) {
  var headerRow = settingSheet.getRange(1, 1, 1, settingSheet.getLastColumn()).getValues()[0];
  var col = headerRow.indexOf(headerName);

  if (col === -1) {
    throw new Error("DA운영설정 시트에서 '" + headerName + "' 헤더를 찾을 수 없습니다.");
  }

  var value = settingSheet.getRange(2, col + 1).getValue();
  return (value instanceof Date) ? value : new Date();
}

// 최근 32일 구간 전체만 다시 그린다 (과거 로그가 수정됐을 때 대시보드에 반영하기 위함).
// 32일보다 오래된 블록(아카이브)은 그대로 두고 절대 다시 쓰지 않는다.
function rebuildRecentDashboard_v2() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName(DASH_SHEET_NAME);
  var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);

  var settings = settingSheet.getDataRange().getValues();
  var lastSettingRow = settingSheet.getLastRow();
  var logs = settingSheet.getRange(1, 12, lastSettingRow, 6).getValues();

  var meta = getMediaAndProducts_(settings);
  var mediaOrder = meta.mediaOrder;
  var activeProducts = meta.activeProducts;

  var cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - RECHECK_WINDOW_DAYS);

  var dateMap = {};

  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    var logDateOnly = new Date(Utilities.formatDate(logDate, TIMEZONE, "yyyy-MM-dd"));
    if (logDateOnly < cutoff) continue;

    var dateKey = Utilities.formatDate(logDate, TIMEZONE, "yyyy-MM-dd");
    if (!dateMap[dateKey]) dateMap[dateKey] = [];
    dateMap[dateKey].push(logs[i]);
  }

  var windowDateKeys = Object.keys(dateMap).sort(); // 오래된 -> 최신

  // 이전에 활성 구간이었지만 이번 재검증에서 윈도우 밖으로 밀려난(=32일보다 오래된) 항목은
  // 시트 내용은 그대로 두고, 그 경계 행 번호만 기억해서 활성 구간 목록에서 제외한다.
  var state = loadBlockState_();
  var archiveBoundaryEndRow = state.archiveBoundaryEndRow;

  state.activeEntries.forEach(function(e) {
    if (windowDateKeys.indexOf(e.dateKey) === -1 && e.endRow > archiveBoundaryEndRow) {
      archiveBoundaryEndRow = e.endRow;
    }
  });

  // 재검증 구간은 아카이브 경계 바로 다음 행부터 시작해서, 시트 끝까지를 통째로 지우고 다시 그린다.
  var startRow = archiveBoundaryEndRow > 0 ? archiveBoundaryEndRow + 3 : getDetailContentBaseRow_();

  var lastRowNum = dashboardSheet.getLastRow();
  if (lastRowNum >= startRow) {
    var maxCol = dashboardSheet.getMaxColumns();
    dashboardSheet.getRange(startRow, 1, lastRowNum - startRow + 1, maxCol).clear({ validationsAlso: true });
  }

  var row = startRow;
  var newActiveEntries = [];

  windowDateKeys.forEach(function(dateKey) {
    var endRow = renderDateBlock_(dashboardSheet, row, dateKey, dateMap[dateKey], mediaOrder, activeProducts);

    dashboardSheet.getRange(row, 1, endRow - row + 1, 1)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    newActiveEntries.push({ dateKey: dateKey, startRow: row, endRow: endRow });
    row = endRow + 3;
  });

  saveBlockState_({ archiveBoundaryEndRow: archiveBoundaryEndRow, activeEntries: newActiveEntries });

  var summaryRefDate = getHeaderDate_(settingSheet, "조회일");
  renderRecentFinalSummary_(dashboardSheet, logs, mediaOrder, activeProducts, summaryRefDate);

  applyColumnWidths_(dashboardSheet);
}

// 특정 날짜 하나의 블록을 시트에 새로 쓰거나(신규) 같은 자리에 다시 그린다(갱신).
// 활성 구간에 없는 날짜(신규)는 마지막 블록 바로 아래에 추가되고, 있는 날짜(갱신)는
// 기존 위치를 지우고 같은 자리에 다시 그린다. 항상 가장 최근 날짜(맨 아래)에서만
// 호출되므로 뒤에 있는 다른 블록을 밀어내거나 침범할 일이 없다.
function upsertDateBlock_(dashboardSheet, dateKey, dayLogs, mediaOrder, activeProducts) {
  if (!dayLogs || dayLogs.length === 0) return;

  var state = loadBlockState_();
  var activeEntries = state.activeEntries;

  var existingIdx = -1;
  for (var i = 0; i < activeEntries.length; i++) {
    if (activeEntries[i].dateKey === dateKey) { existingIdx = i; break; }
  }

  var startRow;

  if (existingIdx !== -1) {
    var existing = activeEntries[existingIdx];
    startRow = existing.startRow;
    var maxCol = dashboardSheet.getMaxColumns();
    dashboardSheet.getRange(startRow, 1, existing.endRow - startRow + 1, maxCol).clear({ validationsAlso: true });
  } else if (activeEntries.length > 0) {
    startRow = activeEntries[activeEntries.length - 1].endRow + 3;
  } else {
    startRow = state.archiveBoundaryEndRow > 0 ? state.archiveBoundaryEndRow + 3 : getDetailContentBaseRow_();
  }

  var endRow = renderDateBlock_(dashboardSheet, startRow, dateKey, dayLogs, mediaOrder, activeProducts);

  dashboardSheet.getRange(startRow, 1, endRow - startRow + 1, 1)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  if (existingIdx !== -1) {
    activeEntries[existingIdx].endRow = endRow;
  } else {
    activeEntries.push({ dateKey: dateKey, startRow: startRow, endRow: endRow });
  }

  saveBlockState_({ archiveBoundaryEndRow: state.archiveBoundaryEndRow, activeEntries: activeEntries });
}

// 날짜 하나에 대한 상세 표(날짜 헤더 + 시간대별 매체/보종 비용·DB·단가 + 소계 + 전체합계)를
// startRow부터 그리고, 이 블록이 차지한 마지막 행(전체합계 행) 번호를 반환한다.
function renderDateBlock_(dashboardSheet, startRow, dateKey, dayLogs, mediaOrder, activeProducts) {

  var row = startRow;

  dashboardSheet.getRange(row, 1)
    .setValue(dateKey)
    .setBackground("#444444")
    .setFontColor("white")
    .setFontWeight("bold");

  row++;

  var times = [];

  dayLogs.forEach(function(log) {
    var t = Utilities.formatDate(log[0], TIMEZONE, "HH:mm");
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
    var t = Utilities.formatDate(log[0], TIMEZONE, "HH:mm");
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

  return row; // 이 블록의 마지막 행(전체합계 행)
}

// 대시보드 최상단에 "조회일 기준 최근 N일([최종마감] 00:00) 추이" 표를 그린다.
// 매체+보종 조합을 행으로, 날짜별 비용/DB/단가를 열로 나란히 배치한다. 조회일 자체가 아니라
// 조회일-1, 조회일-2, 조회일-3(즉 이미 마감이 끝난 지난 3일)을 기준으로 삼는다.
// 표에 실제로 필요한 행 수를 먼저 계산해서, 이전에 반영돼 있던 행 수와 다르면 그 차이만큼
// insertRowsBefore/deleteRows로 아래 날짜별 상세 블록 전체를 정확히 밀거나 당기고, 블록 상태의
// 시작/끝 행 번호도 같이 보정한 뒤에 그린다. 그래서 항상 "딱 필요한 만큼"만 자리를 차지한다.
function renderRecentFinalSummary_(dashboardSheet, logs, mediaOrder, activeProducts, refDate) {

  var recentDates = [];
  for (var d = TOP_SUMMARY_RECENT_DAYS; d >= 1; d--) {
    var dt = new Date(refDate);
    dt.setDate(dt.getDate() - d);
    recentDates.push(Utilities.formatDate(dt, TIMEZONE, "yyyy-MM-dd"));
  }

  // 로그 중 [최종마감](00:00) 항목만, 위 3개 날짜에 해당하는 것만 인덱싱한다.
  var finalMap = {}; // "날짜_매체_보종" -> {cost, db}

  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    var t = Utilities.formatDate(logDate, TIMEZONE, "HH:mm");
    if (t !== "00:00") continue;

    var dateKey = Utilities.formatDate(logDate, TIMEZONE, "yyyy-MM-dd");
    if (recentDates.indexOf(dateKey) === -1) continue;

    var key = dateKey + "_" + String(logs[i][1]).trim() + "_" + String(logs[i][2]).trim();
    finalMap[key] = {
      cost: Number(logs[i][3]) || 0,
      db: Number(logs[i][4]) || 0
    };
  }

  // 매체별로 표시할 조합 그룹 구성 (행 수 계산과 실제 렌더링에 그대로 재사용)
  var mediaGroups = mediaOrder
    .map(function(media) {
      var items = activeProducts.filter(function(x) {
        if (x.media !== media) return false;
        return recentDates.some(function(d) {
          return !!finalMap[d + "_" + x.media + "_" + x.product];
        });
      });
      return { media: media, items: items };
    })
    .filter(function(g) { return g.items.length > 0; });

  var dataRowCount = mediaGroups.reduce(function(sum, g) { return sum + g.items.length + 1; }, 0); // +1 소계행
  var totalRows = 3 + dataRowCount + 1; // 제목1 + 헤더2 + (조합+소계) + 전체합계1

  resizeTopSummaryArea_(dashboardSheet, totalRows);

  var row = TOP_SUMMARY_START_ROW;

  dashboardSheet.getRange(row, 1)
    .setValue("최근 " + TOP_SUMMARY_RECENT_DAYS + "일 최종마감 추이 (조회일 기준)")
    .setBackground("#444444")
    .setFontColor("white")
    .setFontWeight("bold");

  row++;

  var header1 = ["", ""];
  var header2 = ["매체", "보종"];

  recentDates.forEach(function(dateKey) {
    header1.push(dateKey);
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

  recentDates.forEach(function(d) {
    grandCost[d] = 0;
    grandDb[d] = 0;
  });

  mediaGroups.forEach(function(group) {

    var media = group.media;
    var mediaCost = {};
    var mediaDb = {};

    recentDates.forEach(function(d) {
      mediaCost[d] = 0;
      mediaDb[d] = 0;
    });

    var mediaStartRow = row;

    group.items.forEach(function(item) {

      var rowData = [item.media, item.product];

      recentDates.forEach(function(d) {
        var found = finalMap[d + "_" + item.media + "_" + item.product];

        if (found) {
          rowData.push(found.cost);
          rowData.push(found.db);
          rowData.push(found.db > 0 ? Math.round(found.cost / found.db) : 0);

          mediaCost[d] += found.cost;
          mediaDb[d] += found.db;
          grandCost[d] += found.cost;
          grandDb[d] += found.db;
        } else {
          rowData.push("");
          rowData.push("");
          rowData.push("");
        }
      });

      dashboardSheet.getRange(row, 1, 1, rowData.length).setValues([rowData]);
      dashboardSheet.getRange(row, 3, 1, rowData.length - 2).setNumberFormat("#,##0");

      var dataIndex = 4;
      var sheetCol = 5;

      recentDates.forEach(function() {
        var cpaValue = Number(rowData[dataIndex]);

        if (!isNaN(cpaValue) && cpaValue > 0) {
          var cell = dashboardSheet.getRange(row, sheetCol);

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

    if (group.items.length > 1) {
      dashboardSheet.getRange(mediaStartRow, 1, group.items.length, 1)
        .mergeVertically()
        .setVerticalAlignment("middle")
        .setHorizontalAlignment("center")
        .setFontWeight("bold");
    }

    var subtotal = [media + " 소계", ""];

    recentDates.forEach(function(d) {
      var cost = mediaCost[d];
      var db = mediaDb[d];
      subtotal.push(cost);
      subtotal.push(db);
      subtotal.push(db > 0 ? Math.round(cost / db) : "");
    });

    dashboardSheet.getRange(row, 1, 1, subtotal.length)
      .setValues([subtotal])
      .setBackground("#EAEAEA")
      .setFontWeight("bold");

    dashboardSheet.getRange(row, 3, 1, subtotal.length - 2).setNumberFormat("#,##0");

    row++;

  });

  var totalRow = ["전체합계", ""];

  recentDates.forEach(function(d) {
    var cost = grandCost[d];
    var db = grandDb[d];
    totalRow.push(cost);
    totalRow.push(db);
    totalRow.push(db > 0 ? Math.round(cost / db) : "");
  });

  dashboardSheet.getRange(row, 1, 1, totalRow.length)
    .setValues([totalRow])
    .setBackground("#D9EAD3")
    .setFontWeight("bold");

  dashboardSheet.getRange(row, 3, 1, totalRow.length - 2).setNumberFormat("#,##0");

  dashboardSheet.getRange(TOP_SUMMARY_START_ROW, 1, row - TOP_SUMMARY_START_ROW + 1, 1)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

// 요약표가 desiredRows행을 차지하도록 맞춘다. 이전에 실제로 반영돼 있던 행 수(문서 속성에 기억)와
// 다르면 그 차이만큼 아래 날짜별 상세 블록 전체를 insertRowsBefore/deleteRows로 밀거나 당기고,
// 블록 상태(아카이브 경계/활성 구간의 시작·끝 행)도 같이 보정한다. 그런 다음 desiredRows행을 지워서
// 실제 내용을 새로 쓸 자리를 마련한다.
function resizeTopSummaryArea_(dashboardSheet, desiredRows) {
  var props = PropertiesService.getDocumentProperties();
  var appliedRaw = props.getProperty(TOP_SUMMARY_ROWS_PROPERTY_KEY);
  var applied = appliedRaw !== null ? Number(appliedRaw) : 0;

  var delta = desiredRows - applied;

  if (delta !== 0) {
    var state = loadBlockState_();
    var hasExistingContent = state.archiveBoundaryEndRow > 0 || state.activeEntries.length > 0;

    if (hasExistingContent) {
      var boundaryRow = applied + 1;

      if (delta > 0) {
        dashboardSheet.insertRowsBefore(boundaryRow, delta);
      } else {
        dashboardSheet.deleteRows(boundaryRow + delta, -delta);
      }

      var newArchiveBoundaryEndRow = state.archiveBoundaryEndRow > 0 ? state.archiveBoundaryEndRow + delta : 0;
      var newActiveEntries = state.activeEntries.map(function(e) {
        return { dateKey: e.dateKey, startRow: e.startRow + delta, endRow: e.endRow + delta };
      });

      saveBlockState_({ archiveBoundaryEndRow: newArchiveBoundaryEndRow, activeEntries: newActiveEntries });
    }

    props.setProperty(TOP_SUMMARY_ROWS_PROPERTY_KEY, String(desiredRows));
  }

  if (desiredRows > 0) {
    var maxCol = Math.max(dashboardSheet.getMaxColumns(), 20);
    dashboardSheet.getRange(TOP_SUMMARY_START_ROW, 1, desiredRows, maxCol).clear({ validationsAlso: true });
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

function applyColumnWidths_(dashboardSheet) {
  dashboardSheet.setColumnWidth(1, 120);
  dashboardSheet.setColumnWidth(2, 130);

  var lastCol = dashboardSheet.getLastColumn();
  if (lastCol >= 3) {
    dashboardSheet.setColumnWidths(3, lastCol - 2, 80);
  }
}

// ---- 블록 상태 저장 (어떤 날짜가 대시보드 시트 몇 번째 행에 있는지) ----
// 시트를 새로 만들지 않고, 스프레드시트에 종속된 Document Properties에 JSON으로 저장한다.
// 32일보다 오래된 날짜는 개별 항목으로 쌓아두지 않고 "여기까지가 아카이브"라는 경계
// 행 번호(archiveBoundaryEndRow) 하나로만 기억하므로, 저장 용량이 시간이 지나도 늘어나지 않는다.

function loadBlockState_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(BLOCK_STATE_PROPERTY_KEY);
  if (!raw) return { archiveBoundaryEndRow: 0, activeEntries: [] };

  try {
    var parsed = JSON.parse(raw);
    return {
      archiveBoundaryEndRow: Number(parsed.archiveBoundaryEndRow) || 0,
      activeEntries: parsed.activeEntries || []
    };
  } catch (e) {
    return { archiveBoundaryEndRow: 0, activeEntries: [] };
  }
}

function saveBlockState_(state) {
  PropertiesService.getDocumentProperties().setProperty(BLOCK_STATE_PROPERTY_KEY, JSON.stringify(state));
}
