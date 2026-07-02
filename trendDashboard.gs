var TREND_DASHBOARD_DEFAULT_DAYS = 14; // 팝업을 열었을 때 기본으로 보여줄 일수
var TREND_DASHBOARD_MAX_DAYS = 60; // 팝업 안에서 직접 입력해 늘려볼 수 있는 최대 일수

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

  var payload = { days: days, mediaGroups: mediaGroups, defaultDays: TREND_DASHBOARD_DEFAULT_DAYS };

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
