var TREND_DASHBOARD_DEFAULT_DAYS = 14; // 팝업을 열었을 때 기본으로 보여줄 일수
var TREND_DASHBOARD_MAX_DAYS = 60; // 팝업 안에서 직접 입력해 늘려볼 수 있는 최대 일수
var TREND_COMPARE_WINDOW_DEFAULT_DAYS = 3; // 조정사항 분석 탭에서 기본으로 비교할 전/후 일수

// 최근 N일(조회일 기준, [최종마감] 00:00 데이터)의 매체/보종별 비용·DB·단가 추이를
// HTML 팝업으로 보여준다. updateDAReport()/updateDAReport_final() 끝에 자동으로 호출되고,
// 메뉴에서 수동으로도 열어볼 수 있다. 서버에서는 TREND_DASHBOARD_MAX_DAYS만큼 넉넉히 미리
// 계산해서 보내고, 실제 몇 일치를 볼지는 팝업 안의 입력창에서 즉시(재요청 없이) 조절한다.
function showRecentTrendDashboard() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);

  var settings = settingSheet.getDataRange().getValues();
  var lastSettingRow = settingSheet.getLastRow();
  var logs = settingSheet.getRange(1, 12, lastSettingRow, 6).getValues();

  var meta = getMediaAndProducts_(settings);
  var mediaOrder = meta.mediaOrder;
  var activeProducts = meta.activeProducts;

  var refDate = getTrendDashboardRefDate_(settingSheet);

  var days = [];
  for (var d = TREND_DASHBOARD_MAX_DAYS; d >= 1; d--) {
    var dt = new Date(refDate);
    dt.setDate(dt.getDate() - d);
    days.push(Utilities.formatDate(dt, TIMEZONE, "yyyy-MM-dd"));
  }

  // 로그 중 [최종마감](00:00) 항목만, 위 기간에 해당하는 것만 인덱싱한다.
  var finalMap = {}; // "날짜_매체_보종" -> {cost, db}

  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    var t = Utilities.formatDate(logDate, TIMEZONE, "HH:mm");
    if (t !== "00:00") continue;

    var dateKey = Utilities.formatDate(logDate, TIMEZONE, "yyyy-MM-dd");
    if (days.indexOf(dateKey) === -1) continue;

    var media = String(logs[i][1]).trim();
    var product = String(logs[i][2]).trim();

    finalMap[dateKey + "_" + media + "_" + product] = {
      cost: Number(logs[i][3]) || 0,
      db: Number(logs[i][4]) || 0
    };
  }

  // 매체별로, 기간 중 데이터가 한 번이라도 있었던 보종만 골라서 날짜별 값을 붙인다.
  var mediaGroups = mediaOrder
    .map(function(media) {
      var items = activeProducts
        .filter(function(x) {
          if (x.media !== media) return false;
          return days.some(function(d) { return !!finalMap[d + "_" + x.media + "_" + x.product]; });
        })
        .map(function(item) {
          return {
            product: item.product,
            targetCPA: item.targetCPA,
            byDate: days.map(function(d) {
              var found = finalMap[d + "_" + item.media + "_" + item.product];
              return found ? { cost: found.cost, db: found.db } : null;
            })
          };
        });
      return { media: media, items: items };
    })
    .filter(function(g) { return g.items.length > 0; });

  // 조정사항 인덱스(setupAdjustmentLogColumns 실행 전이면 빈 배열)도 같은 기간만 골라 함께 보낸다.
  var adjustments = getAdjustmentLog_(settingSheet).filter(function(a) {
    return days.indexOf(a.date) !== -1;
  });

  // "조정사항 분석" 탭의 시간대별 현황 서브탭 + 당일 전/후 분할 계산을 위해, 최종마감(00:00)뿐
  // 아니라 그 날 찍힌 모든 당일현황 스냅샷을 매체/보종/날짜별로 모아 시간순으로 함께 보낸다.
  var daySnapshots = buildDaySnapshots_(logs, days);

  var payload = {
    days: days,
    mediaGroups: mediaGroups,
    defaultDays: TREND_DASHBOARD_DEFAULT_DAYS,
    adjustments: adjustments,
    defaultCompareWindow: TREND_COMPARE_WINDOW_DEFAULT_DAYS,
    daySnapshots: daySnapshots
  };

  var template = HtmlService.createTemplateFromFile('TrendDashboard');
  template.dataJson = JSON.stringify(payload);

  var html = template.evaluate().setWidth(1100).setHeight(750);
  SpreadsheetApp.getUi().showModalDialog(html, "최종마감 추이 대시보드");
}

// DA운영설정 시트의 "조회일" 헤더 밑 값을 읽어온다. 값이 없거나 날짜가 아니면 오늘 날짜로 대체한다.
function getTrendDashboardRefDate_(settingSheet) {
  var headerRow = settingSheet.getRange(1, 1, 1, settingSheet.getLastColumn()).getValues()[0];
  var col = headerRow.indexOf("조회일");

  if (col === -1) {
    throw new Error("DA운영설정 시트에서 '조회일' 헤더를 찾을 수 없습니다.");
  }

  var value = settingSheet.getRange(2, col + 1).getValue();
  return (value instanceof Date) ? value : new Date();
}

// DA운영설정 로그(L~Q, 최종마감 00:00 포함 전체)를 매체/보종/날짜별로 묶어 그 날 찍힌
// 스냅샷들을 시간순으로 정렬해 반환한다. 최종마감(00:00)은 실제로는 다음날 새벽에 눌러
// 마무리하는 값이라, 문자열로는 가장 앞이지만 그 날짜의 마지막 시점으로 취급해 맨 뒤로 보낸다.
function buildDaySnapshots_(logs, days) {

  var result = {}; // "매체||보종" -> "yyyy-MM-dd" -> [{time, cost, db}, ...]

  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    var dateKey = Utilities.formatDate(logDate, TIMEZONE, "yyyy-MM-dd");
    if (days.indexOf(dateKey) === -1) continue;

    var media = String(logs[i][1]).trim();
    var product = String(logs[i][2]).trim();
    var itemKey = media + "||" + product;

    if (!result[itemKey]) result[itemKey] = {};
    if (!result[itemKey][dateKey]) result[itemKey][dateKey] = [];

    result[itemKey][dateKey].push({
      time: Utilities.formatDate(logDate, TIMEZONE, "HH:mm"),
      cost: Number(logs[i][3]) || 0,
      db: Number(logs[i][4]) || 0
    });
  }

  Object.keys(result).forEach(function(itemKey) {
    Object.keys(result[itemKey]).forEach(function(dateKey) {
      result[itemKey][dateKey].sort(function(a, b) {
        if (a.time === "00:00" && b.time !== "00:00") return 1;
        if (b.time === "00:00" && a.time !== "00:00") return -1;
        return a.time < b.time ? -1 : (a.time > b.time ? 1 : 0);
      });
    });
  });

  return result;
}

// setupAdjustmentLogColumns()가 마련한 조정사항 인덱스(조정일자/조정시간/매체/보종/카테고리/
// 세부내용)를 읽어온다. 아직 설정 전이라 헤더가 없으면 빈 배열을 반환한다(추이 대시보드는
// 그대로 동작). 조정시간은 드롭다운으로 "14:00" 같은 문자열을 고르게 돼 있지만, 혹시 시트가
// 이를 시간 값(Date)으로 자동 인식해버린 경우도 대비해 둘 다 처리한다.
function getAdjustmentLog_(settingSheet) {

  var headerRow = settingSheet.getRange(1, 1, 1, settingSheet.getLastColumn()).getValues()[0];
  var col = headerRow.indexOf("조정일자");
  if (col === -1) return [];

  var lastRow = settingSheet.getLastRow();
  if (lastRow < 2) return [];

  var data = settingSheet.getRange(2, col + 1, lastRow - 1, 6).getValues();
  var records = [];

  data.forEach(function(row) {
    var dateCell = row[0];
    if (!(dateCell instanceof Date)) return;

    var timeCell = row[1];
    var timeStr = (timeCell instanceof Date)
      ? Utilities.formatDate(timeCell, TIMEZONE, "HH:mm")
      : String(timeCell).trim();
    if (!timeStr) return; // 조정시간이 비어 있으면 전/후를 가를 기준이 없으므로 건너뛴다.

    records.push({
      date: Utilities.formatDate(dateCell, TIMEZONE, "yyyy-MM-dd"),
      time: timeStr,
      media: String(row[2]).trim(),
      product: String(row[3]).trim(),
      category: String(row[4]).trim(),
      detail: String(row[5]).trim()
    });
  });

  return records;
}
