var TREND_DASHBOARD_DEFAULT_DAYS = 14; // 팝업을 열었을 때 기본으로 보여줄 일수
var TREND_DASHBOARD_MAX_DAYS = 60; // 팝업 안에서 직접 입력해 늘려볼 수 있는 최대 일수
var TREND_COMPARE_WINDOW_DEFAULT_DAYS = 3; // 조정사항 분석 탭에서 기본으로 비교할 전/후 일수

// 최근 N일(조회일 기준, [최종마감] 00:00 데이터)의 매체/보종별 비용·DB·단가 추이를
// HTML 팝업으로 보여준다. updateDAReport()/updateDAReport_final() 끝에 자동으로 호출되고,
// 메뉴에서 수동으로도 열어볼 수 있다. 서버에서는 TREND_DASHBOARD_MAX_DAYS만큼 넉넉히 미리
// 계산해서 보내고, 실제 몇 일치를 볼지는 팝업 안의 입력창에서 즉시(재요청 없이) 조절한다.
function showRecentTrendDashboard() {

  var tStart = Date.now();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settingSheet = ss.getSheetByName(SETTING_SHEET_NAME);

  // 시트 전체를 한 번만 읽고(getDataRange), 로그(L~Q, 0-based 11~16)는 그 안에서 잘라 쓴다.
  // 예전에는 이 로그를 별도 getRange(...).getValues() 호출로 다시 읽었는데, settings에 이미
  // 포함된 데이터라 그 호출은 불필요한 스프레드시트 왕복이었다.
  var settings = settingSheet.getDataRange().getValues();
  var logs = settings.map(function(row) { return row.slice(11, 17); });
  _perfLog("설정 읽기 (" + settings.length + "행)", tStart);

  var meta = getMediaAndProducts_(settings);
  var mediaOrder = meta.mediaOrder;
  var activeProducts = meta.activeProducts;

  var refDate = getTrendDashboardRefDate_(settings);

  var tDays = Date.now();
  var days = [];
  var daysSet = {}; // days.indexOf(...)를 반복 호출하면 로그 행 수 x 기간일수만큼 선형 비교가
                     // 쌓이므로, 로그가 쌓일수록 느려지는 걸 막기 위해 해시셋으로 O(1) 조회한다.
  for (var d = TREND_DASHBOARD_MAX_DAYS; d >= 1; d--) {
    var dt = new Date(refDate);
    dt.setDate(dt.getDate() - d);
    var dateStr = Utilities.formatDate(dt, TIMEZONE, "yyyy-MM-dd");
    days.push(dateStr);
    daysSet[dateStr] = true;
  }

  // 로그 중 [최종마감](00:00) 항목만, 위 기간에 해당하는 것만 인덱싱한다.
  var finalMap = {}; // "날짜_매체_보종" -> {cost, db}

  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    var t = Utilities.formatDate(logDate, TIMEZONE, "HH:mm");
    if (t !== "00:00") continue;

    var dateKey = Utilities.formatDate(logDate, TIMEZONE, "yyyy-MM-dd");
    if (!daysSet[dateKey]) continue;

    var media = String(logs[i][1]).trim();
    var product = String(logs[i][2]).trim();

    finalMap[dateKey + "_" + media + "_" + product] = {
      cost: Number(logs[i][3]) || 0,
      db: Number(logs[i][4]) || 0
    };
  }
  _perfLog("최근 " + TREND_DASHBOARD_MAX_DAYS + "일 최종마감 인덱싱 (로그 " + logs.length + "행)", tDays);

  var tGroups = Date.now();

  // 매체별로, 기간 중 데이터가 한 번이라도 있었던 보종만 골라서 날짜별 값을 붙인다.
  //
  // CATALOG_MEDIA(카탈로그형 매체, 예: 크리테오 다이나믹)는 보종별 로그에 비용이 항상 비어
  // 있어(DA운영설정.gs의 buildSnapshotResults_ 참고) 보종별 항목만으로는 비용을 알 수 없다.
  // 그 매체 전체 비용은 CATALOG_TOTAL_PRODUCT라는 이름으로 로그에 한 번만 별도 적재돼 있으므로,
  // 그 값을 finalMap에서 그대로 꺼내 group.total로 함께 보낸다. 프론트(TrendDashboard.html)는
  // 보종 필터가 없을 때(=매체 전체를 볼 때)만 이 total을 쓰고, 특정 보종을 고르면 지금처럼
  // 그 보종의 DB만(비용 0) 보여준다.
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

      var isCatalog = CATALOG_MEDIA.indexOf(media) !== -1;
      var total = null;

      if (isCatalog) {
        total = {
          byDate: days.map(function(d) {
            var found = finalMap[d + "_" + media + "_" + CATALOG_TOTAL_PRODUCT];
            return found ? { cost: found.cost, db: found.db } : null;
          })
        };
      }

      return { media: media, items: items, isCatalog: isCatalog, total: total };
    })
    .filter(function(g) { return g.items.length > 0; });
  _perfLog("매체/보종별 mediaGroups 구성", tGroups);

  // 조정사항 인덱스(setupAdjustmentLogColumns 실행 전이면 빈 배열)도 같은 기간만 골라 함께 보낸다.
  // settings는 이미 위에서 한 번 읽어둔 데이터라, 조정일자 헤더/행 수를 다시 시트에서 읽지
  // 않고 그대로 재사용한다(불필요한 스프레드시트 호출 왕복을 줄이기 위함). 다만 getAdjustmentLog_
  // 내부에서 조정시간/매체/보종/카테고리/세부내용은 getDisplayValues()로 settings 전체 행 수만큼
  // 별도로 다시 읽으므로(주석 참고), 로그가 많이 쌓이면 이 호출도 눈에 띄게 느려질 수 있다.
  var tAdjust = Date.now();
  var adjustments = getAdjustmentLog_(settingSheet, settings).filter(function(a) {
    return daysSet[a.date];
  });
  _perfLog("조정사항 인덱스 읽기(getDisplayValues, " + settings.length + "행)", tAdjust);

  // "조정사항 분석" 탭의 시간대별 현황 서브탭 + "시간대별 현황" 탭(선택한 아무 날짜나 볼 수 있음)
  // 에서 함께 쓰기 위해, 최종마감(00:00)뿐 아니라 그 날 찍힌 모든 당일현황 스냅샷을 매체/보종/
  // 날짜별로 모아 시간순으로 함께 보낸다. 조회일(=아직 최종마감 전인 당일)도 시간대별로 볼 수
  // 있어야 하므로, 최종마감 추이/조정사항 분석용 daysSet(조회일 제외)과는 별도로 조회일을 더한
  // daysSet을 만들어 넘긴다 — daysSet 자체에 조회일을 섞으면 조정사항 분석 탭(days 인덱스를
  // 기준으로 전/후를 계산)이 깨지기 때문에, 원본 daysSet은 그대로 두고 복사본만 확장한다.
  var todayDateStr = Utilities.formatDate(refDate, TIMEZONE, "yyyy-MM-dd");

  var tSnapshots = Date.now();
  var snapshotDaysSet = {};
  Object.keys(daysSet).forEach(function(d) { snapshotDaysSet[d] = true; });
  snapshotDaysSet[todayDateStr] = true;
  var daySnapshots = buildDaySnapshots_(logs, snapshotDaysSet);
  _perfLog("daySnapshots 구성", tSnapshots);

  var payload = {
    days: days,
    mediaGroups: mediaGroups,
    defaultDays: TREND_DASHBOARD_DEFAULT_DAYS,
    adjustments: adjustments,
    defaultCompareWindow: TREND_COMPARE_WINDOW_DEFAULT_DAYS,
    daySnapshots: daySnapshots,
    todayDate: todayDateStr,
    catalogTotalProduct: CATALOG_TOTAL_PRODUCT
  };

  var tTemplate = Date.now();
  var template = HtmlService.createTemplateFromFile('TrendDashboard');
  template.dataJson = JSON.stringify(payload);

  var html = template.evaluate().setWidth(1100).setHeight(750);
  _perfLog("HTML 템플릿 구성", tTemplate);

  SpreadsheetApp.getUi().showModalDialog(html, "최종마감 추이 대시보드");
  Logger.log("showRecentTrendDashboard 전체(다이얼로그 렌더링 제외): " + (Date.now() - tStart) + "ms");
}

// DA운영설정 시트의 "조회일" 헤더 밑 값을 읽어온다. 값이 없거나 날짜가 아니면 오늘 날짜로 대체한다.
// settings는 호출 쪽이 이미 getDataRange().getValues()로 읽어둔 전체 시트 데이터라, 헤더 조회와
// 값 조회 모두 시트에 다시 요청하지 않고 그 안에서 바로 찾는다.
function getTrendDashboardRefDate_(settings) {
  var headerRow = settings[0];
  var col = headerRow.indexOf("조회일");

  if (col === -1) {
    throw new Error("DA운영설정 시트에서 '조회일' 헤더를 찾을 수 없습니다.");
  }

  var value = settings.length > 1 ? settings[1][col] : null;
  return (value instanceof Date) ? value : new Date();
}

// DA운영설정 로그(L~Q, 최종마감 00:00 포함 전체)를 매체/보종/날짜별로 묶어 그 날 찍힌
// 스냅샷들을 시간순으로 정렬해 반환한다. 최종마감(00:00)은 실제로는 다음날 새벽에 눌러
// 마무리하는 값이라, 문자열로는 가장 앞이지만 그 날짜의 마지막 시점으로 취급해 맨 뒤로 보낸다.
// daysSet은 "yyyy-MM-dd" -> true 형태의 해시셋(showRecentTrendDashboard에서 만든 것)이다.
function buildDaySnapshots_(logs, daysSet) {

  var result = {}; // "매체||보종" -> "yyyy-MM-dd" -> [{time, cost, db}, ...]

  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    var dateKey = Utilities.formatDate(logDate, TIMEZONE, "yyyy-MM-dd");
    if (!daysSet[dateKey]) continue;

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
// 그대로 동작).
//
// settings는 호출 쪽(showRecentTrendDashboard)이 이미 getDataRange().getValues()로 읽어둔
// 전체 시트 데이터를 그대로 넘겨받은 것이다 — 헤더 행 조회, 마지막 행 수, 조정일자(Date 값)
// 모두 이 안에 이미 들어있으므로 시트에 다시 요청하지 않는다. 다만 조정시간/매체/보종/
// 카테고리/세부내용은 getDisplayValues()로 별도 한 번 더 읽는다 — 예를 들어 "10%"처럼
// 타이핑하면 시트가 자동으로 숫자(0.1, 퍼센트 서식)로 바꿔버리는데, settings(getValues() 결과)
// 그대로 쓰면 "0.1"이 되지만 getDisplayValues()는 셀에 실제 보이는 그대로("10%")를 돌려주므로
// 이 컬럼들만큼은 어쩔 수 없이 추가 호출이 필요하다.
function getAdjustmentLog_(settingSheet, settings) {

  var headerRow = settings[0];
  var col = headerRow.indexOf("조정일자");
  if (col === -1) return [];

  if (settings.length < 2) return [];

  var displayValues = settingSheet.getRange(2, col + 1, settings.length - 1, 6).getDisplayValues();
  var records = [];

  for (var i = 1; i < settings.length; i++) {
    var dateCell = settings[i][col];
    if (!(dateCell instanceof Date)) continue;

    var disp = displayValues[i - 1];
    var timeStr = String(disp[1]).trim();
    if (!timeStr) continue; // 조정시간이 비어 있으면 전/후를 가를 기준이 없으므로 건너뛴다.

    records.push({
      date: Utilities.formatDate(dateCell, TIMEZONE, "yyyy-MM-dd"),
      time: timeStr,
      media: String(disp[2]).trim(),
      product: String(disp[3]).trim(),
      category: String(disp[4]).trim(),
      detail: String(disp[5]).trim()
    });
  }

  return records;
}
