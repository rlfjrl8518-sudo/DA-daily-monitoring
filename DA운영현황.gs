var DASH_SHEET_NAME = "DA운영현황";
var SETTING_SHEET_NAME = "DA운영설정";
var TIMEZONE = "Asia/Seoul";

// DA운영현황 시트 맨 위 3행은 추이 대시보드 팝업을 여는 버튼(그림)을 위해 항상 비워둔다.
// 스크립트는 이 행들을 절대 지우거나 쓰지 않고, 날짜별 상세 블록은 항상 4행부터 쌓인다.
var DASHBOARD_TOP_RESERVED_ROWS = 3;
var DASHBOARD_DETAIL_BASE_ROW = DASHBOARD_TOP_RESERVED_ROWS + 1;

// DA운영설정 시트 로그(L~Q열) 전체를 날짜별로 묶어서, 대시보드 상세 영역(4행부터 시트 끝까지)을
// 통째로 지우고 처음부터 다시 그린다.
//
// 예전에는 최근 32일만 다시 그리고 그보다 오래된 블록은 "아카이브"로 남겨 손대지 않는 방식(부분
// 갱신 + Document Properties에 블록 위치 상태 저장)을 썼는데, 하루 안에서도 활성 매체/보종 구성이
// 바뀌면 같은 날짜 블록의 행 수가 달라질 수 있어 부분 갱신 시 아래쪽에 있는 다른 블록을 침범하거나
// 위치가 밀리고, 다음 갱신 때 그 어긋난 상태가 그대로 누적돼 중복 표시로 이어지는 문제가 있었다.
// 매번 로그 전체를 기준으로 처음부터 다시 그리면 위치 추적 상태를 따로 저장·보정할 필요가 없어
// 이런 문제가 근본적으로 사라진다. (로그가 아주 많아지면 매번 전체를 다시 그리는 비용이 커질 수
// 있으니, 그런 상황이 되면 다시 구간 분할을 고려한다.)
function renderFullDashboard() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName(DASH_SHEET_NAME);
  var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);

  var settings = settingSheet.getDataRange().getValues();
  var lastSettingRow = settingSheet.getLastRow();
  var logs = settingSheet.getRange(1, 12, lastSettingRow, 6).getValues();

  var meta = getMediaAndProducts_(settings);
  var mediaOrder = meta.mediaOrder;
  var activeProducts = meta.activeProducts;

  var dateMap = {};

  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    var dateKey = Utilities.formatDate(logDate, TIMEZONE, "yyyy-MM-dd");
    if (!dateMap[dateKey]) dateMap[dateKey] = [];
    dateMap[dateKey].push(logs[i]);
  }

  var dateKeys = Object.keys(dateMap).sort(); // 오래된 -> 최신

  var lastRowNum = dashboardSheet.getLastRow();
  if (lastRowNum >= DASHBOARD_DETAIL_BASE_ROW) {
    var maxCol = dashboardSheet.getMaxColumns();
    dashboardSheet.getRange(DASHBOARD_DETAIL_BASE_ROW, 1, lastRowNum - DASHBOARD_DETAIL_BASE_ROW + 1, maxCol).clear();
  }

  var row = DASHBOARD_DETAIL_BASE_ROW;

  dateKeys.forEach(function(dateKey) {
    var endRow = renderDateBlock_(dashboardSheet, row, dateKey, dateMap[dateKey], mediaOrder, activeProducts);

    dashboardSheet.getRange(row, 1, endRow - row + 1, 1)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    row = endRow + 3;
  });

  applyColumnWidths_(dashboardSheet);
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

    var isCatalogMedia = CATALOG_MEDIA.indexOf(media) !== -1;
    var mediaStartRow = row;

    mediaRows.forEach(function(item) {

      var rowData = [item.media, item.product];

      times.forEach(function(t) {

        var key = t + "_" + item.media + "_" + item.product;
        var found = logMap[key];

        if (found) {
          var db = Number(found[4]) || 0;

          // 카탈로그형 매체(CATALOG_MEDIA)는 보종별 비용을 나눌 수 없으므로, 보종별 행은
          // 일단 비용/단가를 비워두고 DB만 채운다. 매체 전체 비용/CPA는 이 매체 블록을 다 그린
          // 뒤 CATALOG_TOTAL_PRODUCT 로그를 이용해 비용/단가 칸을 세로 병합해서 채워 넣는다
          // (아래 "카탈로그형 매체는 보종별 행에..." 블록 참고).
          if (isCatalogMedia) {
            rowData.push("");
            rowData.push(db);
            rowData.push("");

            mediaDb[t] += db;
            grandDb[t] += db;
          } else {
            var cost = Number(found[3]) || 0;
            var cpa = Number(found[5]) || 0;

            rowData.push(cost);
            rowData.push(db);
            rowData.push(cpa);

            mediaCost[t] += cost;
            mediaDb[t] += db;
            grandCost[t] += cost;
            grandDb[t] += db;
          }
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

    // 카탈로그형 매체는 보종별 행에 비용을 넣지 않았으므로, CATALOG_TOTAL_PRODUCT로 별도
    // 적재해둔 매체 전체 비용을 여기서 매체 소계에 더한다 (DB는 위에서 이미 보종별로 합산됨).
    // 아울러 보종별 행의 비용/단가 칸을 매체명 칸처럼 세로 병합해서, 빈칸으로 휑하게 두는 대신
    // 그 병합된 자리에 매체 전체 비용/CPA를 한 번만 표시한다.
    if (isCatalogMedia) {
      var catalogSheetCol = 5; // 첫 시간대의 "단가" 컬럼(= "비용" 컬럼 + 2)부터 시작

      times.forEach(function(t) {
        var totalFound = logMap[t + "_" + media + "_" + CATALOG_TOTAL_PRODUCT];

        if (totalFound) {
          var totalCost = Number(totalFound[3]) || 0;
          var totalCpa = Number(totalFound[5]) || 0;
          var costCol = catalogSheetCol - 2;
          var cpaCol = catalogSheetCol;

          var costRange = dashboardSheet.getRange(mediaStartRow, costCol, mediaRows.length, 1);
          var cpaRange = dashboardSheet.getRange(mediaStartRow, cpaCol, mediaRows.length, 1);

          if (mediaRows.length > 1) {
            costRange.mergeVertically();
            cpaRange.mergeVertically();
          }

          costRange.setValue(totalCost).setVerticalAlignment("middle").setHorizontalAlignment("center").setNumberFormat("#,##0");
          cpaRange.setValue(totalCpa).setVerticalAlignment("middle").setHorizontalAlignment("center").setNumberFormat("#,##0");

          mediaCost[t] += totalCost;
          grandCost[t] += totalCost;
        }

        catalogSheetCol += 3;
      });
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

  // 날짜 표시 행(제목 줄)은 제외하고, 그 아래 실제 표 부분에만 옅은 회색 격자 테두리를 그린다.
  dashboardSheet.getRange(startRow + 1, 1, row - startRow, header1.length)
    .setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);

  return row; // 이 블록의 마지막 행(전체합계 행)
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
