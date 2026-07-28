/**
 * Feeding Log - Google Apps Script Backend
 *
 * Setup:
 *   1. Open the Google Sheet, Extensions > Apps Script
 *   2. Paste this entire file, save
 *   3. Deploy > New Deployment > Web App
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   4. Copy the URL into index.html as SYNC_SERVER
 *   5. To update code: Deploy > Manage Deployments > Edit > New version > Deploy
 */

var DEFAULT_TZ = "America/Los_Angeles";
var HEADERS = ["Date 日期", "Time 时间", "Event 事件", "Volume 奶量", "Notes 备注"];

function doPost(e) {
  try {
    // Read raw POST body (sendBeacon sends JSON as text/plain)
    var raw = "";
    if (e.postData && e.postData.contents) {
      raw = e.postData.contents;
    } else if (e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    }

    if (!raw) {
      return output({
        success: false,
        error: "No data received",
        debug: {
          hasParam: !!e.parameter,
          hasPostData: !!e.postData,
          postDataType: e.postData ? e.postData.type : "n/a"
        }
      });
    }

    var data = JSON.parse(raw);
    if (!data || !data.event) {
      return output({ success: false, error: "Missing required field: event" });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    ensureHeaders_(sheet);
    var nextRow = Math.max(2, sheet.getLastRow() + 1);

    sheet.getRange(nextRow, 1, 1, 5).setValues([[
      data.date || "",
      data.time || "",
      eventDisplay_(data.event),
      data.volume ? String(data.volume) : "",
      data.notes || ""
    ]]);

    return output({ success: true, row: nextRow });
  } catch (err) {
    return output({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  // Test mode: ?test=1 writes a dummy row
  if (e.parameter.test === "1") {
    try {
      var testSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      ensureHeaders_(testSheet);
      var testRow = Math.max(2, testSheet.getLastRow() + 1);
      testSheet.getRange(testRow, 1, 1, 5).setValues([[
        formatDateInTz_(new Date(), DEFAULT_TZ),
        Utilities.formatDate(new Date(), DEFAULT_TZ, "HH:mm"),
        "Feeding 喂奶",
        "999",
        "GET test row"
      ]]);
      return output({ success: true, test: true, row: testRow });
    } catch (err) {
      return output({ success: false, error: err.toString() });
    }
  }

  // Sync via GET: ?sync=1&d=<JSON>
  if (e.parameter.sync === "1" && e.parameter.d) {
    try {
      var syncData = JSON.parse(e.parameter.d);
      if (!syncData || !syncData.event) {
        return output({ success: false, error: "Missing event" });
      }
      var syncSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      ensureHeaders_(syncSheet);
      var syncRow = Math.max(2, syncSheet.getLastRow() + 1);
      syncSheet.getRange(syncRow, 1, 1, 5).setValues([[
        syncData.date || "",
        syncData.time || "",
        eventDisplay_(syncData.event),
        syncData.volume ? String(syncData.volume) : "",
        syncData.notes || ""
      ]]);
      return output({ success: true, row: syncRow });
    } catch (err2) {
      return output({ success: false, error: err2.toString() });
    }
  }

  // Cross-device summary via JSONP: ?summary=1&callback=fn&tz=America/Los_Angeles
  if (e.parameter.summary === "1") {
    try {
      var tz = e.parameter.tz || DEFAULT_TZ;
      var summarySheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      var summary = buildTodaySummary_(summarySheet, tz);
      var payload = {
        success: true,
        summary: summary
      };

      if (e.parameter.callback) {
        return outputJsonp_(e.parameter.callback, payload);
      }
      return output(payload);
    } catch (err3) {
      var errPayload = { success: false, error: err3.toString() };
      if (e.parameter.callback) {
        return outputJsonp_(e.parameter.callback, errPayload);
      }
      return output(errPayload);
    }
  }

  // Status
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  return output({
    status: "ok",
    sheet: sheet.getName(),
    rows: lastRow > 0 ? lastRow - 1 : 0
  });
}

function ensureHeaders_(sheet) {
  var firstRow = sheet.getRange(1, 1, 1, 5).getValues()[0];
  if (String(firstRow[0]).trim() !== "Date 日期") {
    sheet.getRange(1, 1, 1, 5).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  }
}

function eventDisplay_(eventKey) {
  var eventMap = {
    "Feeding": "Feeding 喂奶",
    "Diaper": "Diaper 拉粑粑",
    "Other": "Other 其他"
  };
  return eventMap[eventKey] || eventKey;
}

function eventKey_(eventRaw) {
  var s = String(eventRaw || "").trim();
  var low = s.toLowerCase();
  if (low.indexOf("feeding") !== -1 || s.indexOf("喂奶") !== -1) return "Feeding";
  if (low.indexOf("diaper") !== -1 || s.indexOf("拉") !== -1 || s.indexOf("便") !== -1 || s.indexOf("尿") !== -1) return "Diaper";
  if (low.indexOf("other") !== -1 || s.indexOf("其他") !== -1) return "Other";
  return s || "Other";
}

function formatDateInTz_(dateObj, tz) {
  return Utilities.formatDate(dateObj, tz || DEFAULT_TZ, "yyyy-MM-dd");
}

function normalizeDateCell_(cell, tz) {
  if (Object.prototype.toString.call(cell) === "[object Date]" && !isNaN(cell.getTime())) {
    return formatDateInTz_(cell, tz);
  }
  var s = String(cell || "").trim();
  if (!s) return "";

  // Try strict yyyy-mm-dd / yyyy/mm/dd parsing
  var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})$/);
  if (m) {
    var y = m[1];
    var mo = ("0" + m[2]).slice(-2);
    var d = ("0" + m[3]).slice(-2);
    return y + "-" + mo + "-" + d;
  }

  // Fallback: keep only first 10 chars for already-normalized content
  return s.slice(0, 10);
}

function normalizeTimeCell_(cell) {
  var s = String(cell || "").trim();
  var m = s.match(/^(\d{1,2}):(\d{1,2})/);
  if (!m) return s;
  return ("0" + m[1]).slice(-2) + ":" + ("0" + m[2]).slice(-2);
}

function buildTodaySummary_(sheet, tz) {
  ensureHeaders_(sheet);

  var today = formatDateInTz_(new Date(), tz);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return {
      date: today,
      timezone: tz,
      feedingCount: 0,
      diaperCount: 0,
      totalVolume: 0,
      records: []
    };
  }

  var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var rows = [];
  for (var i = 0; i < values.length; i++) {
    var r = values[i];
    var dateStr = normalizeDateCell_(r[0], tz);
    if (dateStr !== today) continue;

    var event = eventKey_(r[2]);
    var volumeNum = parseFloat(String(r[3] || "").replace(/[^\d.]/g, ""));
    rows.push({
      date: dateStr,
      time: normalizeTimeCell_(r[1]),
      event: event,
      volume: isNaN(volumeNum) ? "" : String(Math.round(volumeNum)),
      notes: String(r[4] || "")
    });
  }

  rows.sort(function(a, b) {
    return String(b.time || "").localeCompare(String(a.time || ""));
  });

  var feedingCount = 0;
  var diaperCount = 0;
  var totalVolume = 0;

  for (var j = 0; j < rows.length; j++) {
    if (rows[j].event === "Feeding") {
      feedingCount++;
      totalVolume += parseFloat(rows[j].volume || "0") || 0;
    } else if (rows[j].event === "Diaper") {
      diaperCount++;
    }
  }

  return {
    date: today,
    timezone: tz,
    feedingCount: feedingCount,
    diaperCount: diaperCount,
    totalVolume: Math.round(totalVolume),
    records: rows.slice(0, 20)
  };
}

function outputJsonp_(callbackName, payload) {
  var cb = String(callbackName || "").trim();
  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]{0,120}$/.test(cb)) {
    cb = "feedingLogCallback";
  }
  var js = cb + "(" + JSON.stringify(payload) + ");";
  return ContentService.createTextOutput(js)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function output(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
