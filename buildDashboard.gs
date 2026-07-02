function buildDashboard_v2() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboardSheet = ss.getSheetByName("DA운영현황");
  var settingSheet = ss.getSheetByName("DA운영설정");

  dashboardSheet.clear();

  var settings = settingSheet.getDataRange().getValues();

  var lastRow = settingSheet.getLastRow();
  var logs = settingSheet.getRange(1, 12, lastRow, 6).getValues();

  var sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 32); // 33일 기준

  var activeProducts = [];
  var mediaOrder = [];

  for (var i = 1; i < settings.length; i++) {
    var sortNo = settings[i][8];
    var mediaName = settings[i][9];

    if (sortNo !== "" && mediaName !== "") {
      mediaOrder.push({
        order: Number(sortNo),
        media: String(mediaName).trim()
      });
    }
  }

  mediaOrder.sort(function(a, b) { return a.order - b.order; });
  mediaOrder = mediaOrder.map(function(x) { return x.media; });

  // 사용여부 관계없이 모든 매체/보종 로드 (targetCPA 참조용)
  for (var i = 1; i < settings.length; i++) {
    if (settings[i][0] === "" || settings[i][1] === "") continue;
    activeProducts.push({
      media: String(settings[i][0]).trim(),
      product: String(settings[i][1]).trim(),
      targetCPA: Number(settings[i][3]) || 0
    });
  }

  var dateMap = {};

  for (var i = 1; i < logs.length; i++) {
    var logDate = logs[i][0];
    if (!(logDate instanceof Date)) continue;

    var logDateOnly = new Date(
      Utilities.formatDate(logDate, "Asia/Seoul", "yyyy-MM-dd")
    );

    if (logDateOnly < sevenDaysAgo) continue;

    var dateKey = Utilities.formatDate(logDate, "Asia/Seoul", "yyyy-MM-dd");

    if (!dateMap[dateKey]) dateMap[dateKey] = [];
    dateMap[dateKey].push(logs[i]);
  }

  var row = 1;

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
      var t = Utilities.formatDate(log[0], "Asia/Seoul", "HH:mm");
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
      var t = Utilities.formatDate(log[0], "Asia/Seoul", "HH:mm");
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

  // A열 전체 볼드 + 가운데 정렬
  dashboardSheet.getRange(1, 1, dashboardSheet.getLastRow(), 1)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // 열 너비
  dashboardSheet.setColumnWidth(1, 120);
  dashboardSheet.setColumnWidth(2, 130);

  var lastCol = dashboardSheet.getLastColumn();
  if (lastCol >= 3) {
    dashboardSheet.setColumnWidths(3, lastCol - 2, 80);
  }

}
