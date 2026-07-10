var DASH_SHEET_NAME = "DA운영현황";
var SETTING_SHEET_NAME = "DA운영설정";
var TIMEZONE = "Asia/Seoul";

// DA운영현황 시트 맨 위 3행은 추이 대시보드 팝업을 여는 버튼(그림)을 위해 항상 비워둔다.
// 스크립트는 이 행들을 절대 지우거나 쓰지 않고, 날짜별 상세 블록은 항상 4행부터 쌓인다.
var DASHBOARD_TOP_RESERVED_ROWS = 3;
var DASHBOARD_DETAIL_BASE_ROW = DASHBOARD_TOP_RESERVED_ROWS + 1;

// renderRecentDashboard()가 매번 다시 그리는 "최근" 구간의 폭(오늘 포함 며칠). 전일 마감이
// 다루는 범위(최대 금/토/일=3일 전)보다 넉넉하게 잡아둔다.
var DASHBOARD_RECENT_WINDOW_DAYS = 7;

// renderRecentDashboard()가 "여기까지는 이미 확정해서 그려둔 과거 블록"이라고 기억해두는
// 경계값을 Document Properties에 저장할 때 쓰는 키. nextRow는 확정 블록 바로 다음(새 블록이
// 시작될) 행 번호, lastDate는 마지막으로 확정한 날짜(yyyy-MM-dd)다.
var DASH_ARCHIVE_NEXT_ROW_PROP = "DASH_ARCHIVE_NEXT_ROW";
var DASH_ARCHIVE_LAST_DATE_PROP = "DASH_ARCHIVE_LAST_DATE";

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

  var tStart = Date.now();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName(DASH_SHEET_NAME);
  var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);

  var settings = settingSheet.getDataRange().getValues();
  var lastSettingRow = settingSheet.getLastRow();
  var logs = settingSheet.getRange(1, 12, lastSettingRow, 6).getValues();
  _perfLog("설정/로그 읽기 (로그 " + logs.length + "행)", tStart);

  var meta = getMediaAndProducts_(settings);
  var mediaOrder = meta.mediaOrder;
  var activeProducts = meta.activeProducts;

  var tDateMap = Date.now();
  var dateMap = {};

  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    var dateKey = Utilities.formatDate(logDate, TIMEZONE, "yyyy-MM-dd");
    if (!dateMap[dateKey]) dateMap[dateKey] = [];
    dateMap[dateKey].push(logs[i]);
  }

  var dateKeys = Object.keys(dateMap).sort(); // 오래된 -> 최신
  _perfLog("날짜별 로그 그룹핑 (" + dateKeys.length + "일)", tDateMap);

  var cutoff = computeRecentCutoff_();
  var archiveDateKeys = dateKeys.filter(function(dk) { return dk < cutoff; });
  var recentDateKeys = dateKeys.filter(function(dk) { return dk >= cutoff; });

  var tClear = Date.now();
  var lastRowNum = dashboardSheet.getLastRow();
  if (lastRowNum >= DASHBOARD_DETAIL_BASE_ROW) {
    var maxCol = dashboardSheet.getMaxColumns();
    dashboardSheet.getRange(DASHBOARD_DETAIL_BASE_ROW, 1, lastRowNum - DASHBOARD_DETAIL_BASE_ROW + 1, maxCol).clear();
  }
  _perfLog("기존 상세 영역 지우기 (" + lastRowNum + "행)", tClear);

  var tRender = Date.now();

  // 과거(archiveDateKeys)와 최근(recentDateKeys)을 나눠서 그리는 것 자체는 renderFullDashboard()
  // 입장에서는 전체를 다 그리는 것과 결과가 똑같지만, 이렇게 나눠 그려야 그 경계(archive가 끝나는
  // 행/날짜)를 renderRecentDashboard()가 이어받을 수 있도록 저장해둘 수 있다. 즉 "전체 기간
  // 재검증"을 수동 실행한 뒤에도 다음 번 자동 갱신(renderRecentDashboard)이 정확한 경계에서부터
  // 이어서 동작하게 하기 위함이다.
  var cursor = renderDateBlocks_(dashboardSheet, DASHBOARD_DETAIL_BASE_ROW, archiveDateKeys, dateMap, mediaOrder, activeProducts);
  saveArchiveState_(cursor, archiveDateKeys.length > 0 ? archiveDateKeys[archiveDateKeys.length - 1] : "");
  renderDateBlocks_(dashboardSheet, cursor, recentDateKeys, dateMap, mediaOrder, activeProducts);

  _perfLog("날짜 블록 " + dateKeys.length + "개 렌더링", tRender);

  var tWidths = Date.now();
  applyColumnWidths_(dashboardSheet);
  _perfLog("컬럼 너비 적용", tWidths);

  Logger.log("renderFullDashboard 전체: " + (Date.now() - tStart) + "ms");
}

// updateDAReport()/updateDAReport_final()(당일 현황 업데이트 / 전일 마감 버튼)에서 매번 호출되는
// 빠른 갱신 경로. renderFullDashboard()처럼 로그 전체를 매번 처음부터 다시 그리는 대신, 이미
// DASHBOARD_RECENT_WINDOW_DAYS일보다 오래돼 확정된 과거 블록은 그대로 두고, 최근 구간만 지우고
// 다시 그린다. 전일 마감은 최대 "금/토/일"(3일 전)까지만 다루므로 7일 윈도우면 여유 있게 커버된다.
//
// 과거에 이미 확정된(archiveState.nextRow 이전) 블록은 절대 다시 지우거나 옮기지 않는다 — 예전에
// "최근 32일만 갱신" 방식에서 블록 높이가 실행마다 달라지며 위치가 밀려 중복 표시로 이어졌던
// 문제는, 매번 위치를 다시 "추정"했기 때문이었다. 여기서는 위치를 추정하지 않고, 실제로 그 위치에
// 그려 넣은 결과(cursor)를 그대로 저장해두므로 밀림이 누적될 수 없다.
//
// 주의: 과거 로그(7일보다 오래된 날짜)를 직접 수정한 경우 이 경로로는 반영되지 않는다. 그런
// 경우엔 메뉴의 "전체 기간 재검증"(renderFullDashboard)을 수동으로 한 번 실행해야 한다.
function renderRecentDashboard() {

  var tStart = Date.now();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName(DASH_SHEET_NAME);
  var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);

  var settings = settingSheet.getDataRange().getValues();
  var lastSettingRow = settingSheet.getLastRow();
  var logs = settingSheet.getRange(1, 12, lastSettingRow, 6).getValues();
  _perfLog("설정/로그 읽기 (로그 " + logs.length + "행)", tStart);

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

  var dateKeys = Object.keys(dateMap).sort();

  var cutoff = computeRecentCutoff_();
  var archiveState = getArchiveState_();

  // 지난 실행 때 이미 확정한(archiveState.lastDate 이하) 날짜는 다시 손대지 않고, 그 사이 7일
  // 밖으로 새로 밀려난 날짜만 이번에 한 번 더 그려서 확정한다.
  var newlyArchivedDateKeys = dateKeys.filter(function(dk) {
    return dk < cutoff && (!archiveState.lastDate || dk > archiveState.lastDate);
  });
  var recentDateKeys = dateKeys.filter(function(dk) { return dk >= cutoff; });

  var tClear = Date.now();
  var lastRowNum = dashboardSheet.getLastRow();
  if (lastRowNum >= archiveState.nextRow) {
    var maxCol = dashboardSheet.getMaxColumns();
    dashboardSheet.getRange(archiveState.nextRow, 1, lastRowNum - archiveState.nextRow + 1, maxCol).clear();
  }
  _perfLog("최근 구간만 지우기 (" + Math.max(lastRowNum - archiveState.nextRow + 1, 0) + "행)", tClear);

  var tRender = Date.now();
  var cursor = archiveState.nextRow;

  if (newlyArchivedDateKeys.length > 0) {
    cursor = renderDateBlocks_(dashboardSheet, cursor, newlyArchivedDateKeys, dateMap, mediaOrder, activeProducts);
    saveArchiveState_(cursor, newlyArchivedDateKeys[newlyArchivedDateKeys.length - 1]);
  }

  renderDateBlocks_(dashboardSheet, cursor, recentDateKeys, dateMap, mediaOrder, activeProducts);
  _perfLog("최근 " + recentDateKeys.length + "일(+ 신규 확정 " + newlyArchivedDateKeys.length + "일) 렌더링", tRender);

  applyColumnWidths_(dashboardSheet);

  Logger.log("renderRecentDashboard 전체: " + (Date.now() - tStart) + "ms");
}

// "오늘" 기준으로 최근 DASHBOARD_RECENT_WINDOW_DAYS일의 시작 날짜(포함)를 구한다. 이 날짜보다
// 이전(<)은 "확정 가능한 과거", 이후(>=)는 "매번 다시 그리는 최근"으로 취급한다.
function computeRecentCutoff_() {
  var today = new Date();
  var cutoffDate = new Date(today);
  cutoffDate.setDate(cutoffDate.getDate() - (DASHBOARD_RECENT_WINDOW_DAYS - 1));
  return Utilities.formatDate(cutoffDate, TIMEZONE, "yyyy-MM-dd");
}

// renderRecentDashboard()가 기억해둔 "여기까지는 이미 확정된 과거 블록" 경계를 읽어온다. 저장된
// 값이 없으면(최초 실행, 혹은 이 기능 도입 전 데이터) 대시보드가 비어있는 것으로 취급해 처음부터
// 다시 그리게 한다.
function getArchiveState_() {
  var props = PropertiesService.getDocumentProperties();
  var nextRowStr = props.getProperty(DASH_ARCHIVE_NEXT_ROW_PROP);
  var lastDate = props.getProperty(DASH_ARCHIVE_LAST_DATE_PROP);

  return {
    nextRow: nextRowStr ? Number(nextRowStr) : DASHBOARD_DETAIL_BASE_ROW,
    lastDate: lastDate || null
  };
}

function saveArchiveState_(nextRow, lastDate) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty(DASH_ARCHIVE_NEXT_ROW_PROP, String(nextRow));
  props.setProperty(DASH_ARCHIVE_LAST_DATE_PROP, lastDate || "");
}

// dateKeys를 startRow부터 순서대로 렌더링하고, 다음 블록이 이어질 행(마지막 블록 뒤 여백 포함)을
// 반환한다. renderFullDashboard()/renderRecentDashboard()가 공유해서 쓴다.
function renderDateBlocks_(dashboardSheet, startRow, dateKeys, dateMap, mediaOrder, activeProducts) {
  var row = startRow;

  dateKeys.forEach(function(dateKey) {
    var endRow = renderDateBlock_(dashboardSheet, row, dateKey, dateMap[dateKey], mediaOrder, activeProducts);

    dashboardSheet.getRange(row, 1, endRow - row + 1, 1)
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    row = endRow + 3;
  });

  return row;
}

// 날짜 하나에 대한 상세 표(날짜 헤더 + 시간대별 매체/보종 비용·DB·단가 + 소계 + 전체합계)를
// startRow부터 그리고, 이 블록이 차지한 마지막 행(전체합계 행) 번호를 반환한다.
//
// 예전에는 셀/행 단위로 setValues·setBackground·setNumberFormat을 그때그때 호출했는데(매체 x
// 보종 x 시간대만큼, 로그가 쌓일수록 수백~수천 번), Apps Script는 스프레드시트 호출 1건마다
// 고정 오버헤드가 커서 이게 renderFullDashboard() 전체 실행 시간의 대부분을 차지했다. 그래서
// 값/배경색/서식을 블록 전체 크기의 2차원 배열로 메모리에서 다 구성한 뒤, 블록 하나당 딱 한 번씩만
// setValues/setBackgrounds/setNumberFormats/setFontWeights/setFontColors를 호출하도록 바꿨다.
// 화면에 그려지는 결과(값, 배경색, 굵기, 병합, 테두리)는 이전과 동일하다.
function renderDateBlock_(dashboardSheet, startRow, dateKey, dayLogs, mediaOrder, activeProducts) {

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
    header1.push(label, "", "");
    header2.push("비용", "DB", "단가");
  });

  var numCols = header1.length;

  function blankArray(fillValue) { return new Array(numCols).fill(fillValue); }

  // 블록 전체를 담을 2차원 배열들. 실제 시트에는 맨 마지막에 블록당 한 번씩만 써 넣는다.
  var values = [];
  var backgrounds = [];
  var formats = [];
  var fontWeights = [];
  var fontColors = [];

  // 날짜 헤더 행 (dateKey는 "yyyy-MM-dd" 문자열인데, "General" 서식으로 두면 시트가 이를
  // 날짜로 자동 인식해 46213 같은 시리얼 넘버로 표시해버려서, header1의 시간 라벨처럼
  // 텍스트 서식("@")으로 고정해 문자열 그대로 보이게 한다)
  var dateRow = blankArray("");
  dateRow[0] = dateKey;
  values.push(dateRow);
  backgrounds.push(blankArray(null));
  backgrounds[0][0] = "#444444";
  formats.push(blankArray("@"));
  fontWeights.push(blankArray("normal"));
  fontWeights[0][0] = "bold";
  fontColors.push(blankArray(null));
  fontColors[0][0] = "white";

  // header1 행 (시간 라벨, 텍스트 서식으로 고정해 "9:00" 같은 값이 시간으로 자동 변환되지 않게 함)
  values.push(header1.slice());
  backgrounds.push(blankArray(null));
  formats.push(blankArray("@"));
  fontWeights.push(blankArray("normal"));
  fontColors.push(blankArray(null));

  // header2 행
  values.push(header2.slice());
  backgrounds.push(blankArray("#EFEFEF"));
  formats.push(blankArray("General"));
  fontWeights.push(blankArray("bold"));
  fontColors.push(blankArray(null));

  var grandCost = {};
  var grandDb = {};

  times.forEach(function(t) {
    grandCost[t] = 0;
    grandDb[t] = 0;
  });

  var mediaMergeRanges = []; // { rowOffset, rowCount } - 매체명 칸 세로 병합 대상
  var catalogMergeRanges = []; // { rowOffset, rowCount, costCol, cpaCol } - 카탈로그 비용/단가 칸 세로 병합 대상

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
    var mediaStartOffset = values.length;

    mediaRows.forEach(function(item) {

      var rowData = [item.media, item.product];
      var rowBg = blankArray(null);
      var rowFmt = blankArray("General");

      var dataCol = 2; // 0-based, 이 시간대의 "비용" 칸부터

      times.forEach(function(t) {

        var key = t + "_" + item.media + "_" + item.product;
        var found = logMap[key];

        var costVal = "", dbVal = "", cpaVal = "";

        if (found) {
          var db = Number(found[4]) || 0;

          // 카탈로그형 매체(CATALOG_MEDIA)는 보종별 비용을 나눌 수 없으므로, 보종별 행은
          // 일단 비용/단가를 비워두고 DB만 채운다. 매체 전체 비용/CPA는 이 매체 블록을 다 그린
          // 뒤 CATALOG_TOTAL_PRODUCT 로그를 이용해 비용/단가 칸을 세로 병합해서 채워 넣는다
          // (아래 "카탈로그형 매체는 보종별 행에..." 블록 참고).
          if (isCatalogMedia) {
            dbVal = db;
            mediaDb[t] += db;
            grandDb[t] += db;
          } else {
            var cost = Number(found[3]) || 0;
            var cpa = Number(found[5]) || 0;

            costVal = cost;
            dbVal = db;
            cpaVal = cpa;

            mediaCost[t] += cost;
            mediaDb[t] += db;
            grandCost[t] += cost;
            grandDb[t] += db;
          }
        }

        rowData.push(costVal, dbVal, cpaVal);
        rowFmt[dataCol] = "#,##0";
        rowFmt[dataCol + 1] = "#,##0";
        rowFmt[dataCol + 2] = "#,##0";

        if (typeof cpaVal === "number" && cpaVal > 0) {
          if (cpaVal <= item.targetCPA) {
            rowBg[dataCol + 2] = "#B6D7A8";
          } else if (cpaVal <= item.targetCPA * 1.2) {
            rowBg[dataCol + 2] = "#FFD966";
          } else {
            rowBg[dataCol + 2] = "#EA9999";
          }
        }

        dataCol += 3;
      });

      values.push(rowData);
      backgrounds.push(rowBg);
      formats.push(rowFmt);
      fontWeights.push(blankArray("normal"));
      fontColors.push(blankArray(null));

    });

    if (mediaRows.length > 1) {
      mediaMergeRanges.push({ rowOffset: mediaStartOffset, rowCount: mediaRows.length });
    }

    // 카탈로그형 매체는 보종별 행에 비용을 넣지 않았으므로, CATALOG_TOTAL_PRODUCT로 별도
    // 적재해둔 매체 전체 비용을 여기서 매체 소계에 더한다 (DB는 위에서 이미 보종별로 합산됨).
    // 아울러 보종별 행의 비용/단가 칸을 매체명 칸처럼 세로 병합해서, 빈칸으로 휑하게 두는 대신
    // 그 병합된 자리에 매체 전체 비용/CPA를 한 번만 표시한다.
    if (isCatalogMedia) {
      var catalogDataCol = 2; // 0-based, 첫 시간대의 "비용" 칸부터

      times.forEach(function(t) {
        var totalFound = logMap[t + "_" + media + "_" + CATALOG_TOTAL_PRODUCT];

        if (totalFound) {
          var totalCost = Number(totalFound[3]) || 0;
          var totalCpa = Number(totalFound[5]) || 0;
          var costCol = catalogDataCol;
          var cpaCol = catalogDataCol + 2;

          for (var r = 0; r < mediaRows.length; r++) {
            values[mediaStartOffset + r][costCol] = totalCost;
            values[mediaStartOffset + r][cpaCol] = totalCpa;
          }

          // 보종별 행 루프에서는 이 시점에 비용/단가가 아직 비어 있어(위쪽 참고) 목표CPA 배경색을
          // 못 칠했으므로, 실제 값이 채워진 지금 여기서 매체 전체 CPA 기준으로 칠한다. 보종마다
          // 목표CPA가 다를 수 있지만 대표로 첫 보종의 목표CPA를 기준으로 삼는다.
          if (totalCpa > 0 && mediaRows.length > 0) {
            var catalogTargetCPA = mediaRows[0].targetCPA;
            var bgColor = null;

            if (catalogTargetCPA > 0) {
              bgColor = totalCpa <= catalogTargetCPA
                ? "#B6D7A8"
                : (totalCpa <= catalogTargetCPA * 1.2 ? "#FFD966" : "#EA9999");
            }

            if (bgColor) {
              for (var r2 = 0; r2 < mediaRows.length; r2++) {
                backgrounds[mediaStartOffset + r2][cpaCol] = bgColor;
              }
            }
          }

          if (mediaRows.length > 1) {
            catalogMergeRanges.push({ rowOffset: mediaStartOffset, rowCount: mediaRows.length, costCol: costCol, cpaCol: cpaCol });
          }

          mediaCost[t] += totalCost;
          grandCost[t] += totalCost;
        }

        catalogDataCol += 3;
      });
    }

    var subtotal = [media + " 소계", ""];
    var subtotalFmt = blankArray("General");

    times.forEach(function(t, idx) {
      var cost = mediaCost[t];
      var db = mediaDb[t];
      subtotal.push(cost, db, db > 0 ? Math.round(cost / db) : "");

      var c = 2 + idx * 3;
      subtotalFmt[c] = "#,##0";
      subtotalFmt[c + 1] = "#,##0";
      subtotalFmt[c + 2] = "#,##0";
    });

    values.push(subtotal);
    backgrounds.push(blankArray("#EAEAEA"));
    formats.push(subtotalFmt);
    fontWeights.push(blankArray("bold"));
    fontColors.push(blankArray(null));

  });

  var totalRow = ["전체합계", ""];
  var totalFmt = blankArray("General");

  times.forEach(function(t, idx) {
    var cost = grandCost[t];
    var db = grandDb[t];
    totalRow.push(cost, db, db > 0 ? Math.round(cost / db) : "");

    var c = 2 + idx * 3;
    totalFmt[c] = "#,##0";
    totalFmt[c + 1] = "#,##0";
    totalFmt[c + 2] = "#,##0";
  });

  values.push(totalRow);
  backgrounds.push(blankArray("#D9EAD3"));
  formats.push(totalFmt);
  fontWeights.push(blankArray("bold"));
  fontColors.push(blankArray(null));

  var numRows = values.length;
  var endRow = startRow + numRows - 1;

  // 블록 전체를 값/배경색/서식/굵기/글자색 각각 딱 한 번씩만 써서 반영한다.
  var blockRange = dashboardSheet.getRange(startRow, 1, numRows, numCols);
  blockRange.setValues(values);
  blockRange.setBackgrounds(backgrounds);
  blockRange.setNumberFormats(formats);
  blockRange.setFontWeights(fontWeights);
  blockRange.setFontColors(fontColors);

  // header2 행만 가운데 정렬 (기존과 동일)
  dashboardSheet.getRange(startRow + 2, 1, 1, numCols).setHorizontalAlignment("center");

  // 매체명 칸 세로 병합 (매체당 보종이 2개 이상일 때만)
  mediaMergeRanges.forEach(function(m) {
    dashboardSheet.getRange(startRow + m.rowOffset, 1, m.rowCount, 1)
      .mergeVertically()
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("center")
      .setFontWeight("bold");
  });

  // 카탈로그형 매체 비용/단가 칸 세로 병합
  catalogMergeRanges.forEach(function(m) {
    dashboardSheet.getRange(startRow + m.rowOffset, m.costCol + 1, m.rowCount, 1)
      .mergeVertically()
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("right");
    dashboardSheet.getRange(startRow + m.rowOffset, m.cpaCol + 1, m.rowCount, 1)
      .mergeVertically()
      .setVerticalAlignment("middle")
      .setHorizontalAlignment("right");
  });

  // 날짜 표시 행(제목 줄)은 제외하고, 그 아래 실제 표 부분에만 옅은 회색 격자 테두리를 그린다.
  dashboardSheet.getRange(startRow + 1, 1, numRows - 1, numCols)
    .setBorder(true, true, true, true, true, true, "#D9D9D9", SpreadsheetApp.BorderStyle.SOLID);

  return endRow; // 이 블록의 마지막 행(전체합계 행)
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
