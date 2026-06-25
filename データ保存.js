// ========================================
// 騒音測定 — データ保存・読込
// ========================================

var NOISE_RESULT_SHEET = '騒音測定実績';

function detectNoiseResultFormat_(headerRow) {
  var h = (headerRow || []).map(function(c) { return String(c || ''); });
  if (h.indexOf('測定区分') >= 0) return 'v2';
  return 'legacy';
}

function ensureNoiseResultSheet_() {
  var ss = getNoiseSpreadsheet_();
  var sheet = ss.getSheetByName(NOISE_RESULT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(NOISE_RESULT_SHEET);
  }

  var headers = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), NOISE_RESULT_HEADERS.length)).getValues()[0]
    : [];
  if (detectNoiseResultFormat_(headers) !== 'v2') {
    sheet.getRange(1, 1, 1, NOISE_RESULT_HEADERS.length).setValues([NOISE_RESULT_HEADERS]);
  }
  sheet.getRange(1, 1, 1, NOISE_RESULT_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#e2e8f0');
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureNoiseSessionSheet_() {
  var ss = getNoiseSpreadsheet_();
  var sheet = ss.getSheetByName(NOISE_SESSION_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(NOISE_SESSION_SHEET);
    sheet.getRange(1, 1, 1, NOISE_SESSION_HEADERS.length).setValues([NOISE_SESSION_HEADERS]);
    sheet.getRange(1, 1, 1, NOISE_SESSION_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#dbeafe');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function parseNoiseResultRow_(row, format) {
  if (format === 'v2') {
    var pointNo = Number(row[3]);
    if (!pointNo) return null;
    return {
      pointNo: pointNo,
      category: String(row[1] || ''),
      pointLabel: String(row[2] || ''),
      floor: String(row[4] || ''),
      db: row[5] !== '' && row[5] != null ? Number(row[5]) : null,
      recordedAt: row[6] instanceof Date
        ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
        : String(row[6] || ''),
      note: String(row[7] || '')
    };
  }

  var legacyNo = Number(row[1]);
  if (!legacyNo) return null;
  var legacyLabel = '';
  if (legacyNo >= NOISE_LINE_POINT_NO_OFFSET + 1) {
    legacyLabel = String(row[2] || '');
  } else {
    var g = pointNoToGrid_(legacyNo);
    legacyLabel = g.label;
  }
  return {
    pointNo: legacyNo,
    category: legacyNo >= NOISE_LINE_POINT_NO_OFFSET + 1 ? 'B' : 'A',
    pointLabel: legacyLabel,
    floor: String(row[2] || ''),
    db: row[3] !== '' && row[3] != null ? Number(row[3]) : null,
    recordedAt: row[5] instanceof Date
      ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
      : String(row[5] || ''),
    note: String(row[6] || '')
  };
}

/**
 * 指定日の測定値（測定点No → レコード）
 */
function loadNoiseResultsForDate(workDate) {
  var dateKey = normalizeWorkDate_(workDate);
  var sheet = ensureNoiseResultSheet_();
  var lastRow = sheet.getLastRow();
  var byPoint = {};

  if (lastRow < 2) return { workDate: dateKey, byPoint: byPoint };

  var lastCol = Math.max(sheet.getLastColumn(), NOISE_RESULT_HEADERS.length);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var format = detectNoiseResultFormat_(headerRow);
  var values = sheet.getRange(2, 1, lastRow, lastCol).getValues();

  for (var i = values.length - 1; i >= 0; i--) {
    var rowDate = normalizeWorkDate_(values[i][0]);
    if (rowDate !== dateKey) continue;
    var rec = parseNoiseResultRow_(values[i], format);
    if (!rec || !rec.pointNo || byPoint[rec.pointNo]) continue;
    byPoint[rec.pointNo] = rec;
  }

  return { workDate: dateKey, byPoint: byPoint };
}

function parseSessionMethodConditions_(row, headerRow) {
  var h = (headerRow || []).map(function(c) { return String(c || ''); });
  if (h.indexOf('測定条件および測定方法') >= 0) {
    return String(row[4] || NOISE_DEFAULT_METHOD_CONDITIONS).trim();
  }
  if (h.indexOf('②測定方法') >= 0 || h.indexOf('④測定条件') >= 0) {
    var method = String(row[4] || '').trim();
    var cond = String(row[5] || '').trim();
    if (method && cond) return method + ' / ' + cond;
    return method || cond || NOISE_DEFAULT_METHOD_CONDITIONS;
  }
  return String(row[4] || NOISE_DEFAULT_METHOD_CONDITIONS).trim();
}

function loadNoiseSessionForDate_(workDate) {
  var dateKey = normalizeWorkDate_(workDate);
  var empty = {
    workDate: dateKey,
    measuredTime: '',
    employeeId: '',
    employeeName: '',
    measurementMethodConditions: NOISE_DEFAULT_METHOD_CONDITIONS,
    actionSummary: ''
  };
  var sheet = ensureNoiseSessionSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return empty;

  var lastCol = Math.max(sheet.getLastColumn(), NOISE_SESSION_HEADERS.length);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var isLegacySession = headerRow.map(String).join('|').indexOf('測定条件および測定方法') < 0
    && headerRow.map(String).join('|').indexOf('②測定方法') >= 0;
  var values = sheet.getRange(2, 1, lastRow, lastCol).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var rowDate = normalizeWorkDate_(values[i][0]);
    if (rowDate !== dateKey) continue;
    var timeVal = values[i][1];
    var measuredTime = '';
    if (timeVal instanceof Date) {
      measuredTime = Utilities.formatDate(timeVal, Session.getScriptTimeZone(), 'HH:mm');
    } else {
      measuredTime = String(timeVal || '').trim();
    }
    var actionIdx = isLegacySession ? 6 : 5;
    return {
      workDate: dateKey,
      measuredTime: measuredTime,
      employeeId: String(values[i][2] || '').trim(),
      employeeName: String(values[i][3] || '').trim(),
      measurementMethodConditions: parseSessionMethodConditions_(values[i], headerRow),
      actionSummary: String(values[i][actionIdx] || '').trim()
    };
  }

  return empty;
}

function saveNoiseSession(payload) {
  payload = payload || {};
  var dateKey = normalizeWorkDate_(payload.workDate);
  var measuredTime = String(payload.measuredTime || '').trim();
  var employeeId = String(payload.employeeId || '').trim();
  var employeeName = String(payload.employeeName || '').trim();

  if (!employeeId) throw new Error('測定者を選択してください');
  if (!measuredTime) throw new Error('測定時刻を入力してください');

  var sheet = ensureNoiseSessionSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var rows = sheet.getRange(2, 1, lastRow, NOISE_SESSION_HEADERS.length).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      if (normalizeWorkDate_(rows[i][0]) === dateKey) {
        sheet.deleteRow(i + 2);
        lastRow = sheet.getLastRow();
      }
    }
  }

  var timeParts = measuredTime.match(/^(\d{1,2}):(\d{2})/);
  var timeObj = new Date();
  if (timeParts) {
    timeObj.setHours(parseInt(timeParts[1], 10), parseInt(timeParts[2], 10), 0, 0);
  }

  sheet.getRange(lastRow + 1, 1, 1, NOISE_SESSION_HEADERS.length).setValues([[
    dateKey,
    timeObj,
    employeeId,
    employeeName,
    String(payload.measurementMethodConditions || NOISE_DEFAULT_METHOD_CONDITIONS).trim(),
    String(payload.actionSummary || '').trim(),
    new Date()
  ]]);

  return {
    success: true,
    workDate: dateKey,
    measuredTime: measuredTime,
    employeeId: employeeId,
    employeeName: employeeName,
    measurementMethodConditions: String(payload.measurementMethodConditions || NOISE_DEFAULT_METHOD_CONDITIONS).trim(),
    actionSummary: String(payload.actionSummary || '').trim()
  };
}

function saveNoiseSessionFromPayload_(payload) {
  if (!payload || !payload.employeeId || !payload.measuredTime) return null;
  return saveNoiseSession(payload);
}

function findNoiseResultRowForPoint_(sheet, dateKey, pointNo) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  var headerRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), NOISE_RESULT_HEADERS.length)).getValues()[0];
  var format = detectNoiseResultFormat_(headerRow);
  var pointCol = format === 'v2' ? 4 : 2;

  var dates = sheet.getRange(2, 1, lastRow, 1).getValues();
  var points = sheet.getRange(2, pointCol, lastRow, pointCol).getValues();
  for (var i = dates.length - 1; i >= 0; i--) {
    if (normalizeWorkDate_(dates[i][0]) === dateKey && Number(points[i][0]) === pointNo) {
      return i + 2;
    }
  }
  return -1;
}

/**
 * 測定点の dB を保存（同日・同点は上書き）
 */
function saveNoiseMeasurement(payload) {
  payload = payload || {};
  var dateKey = normalizeWorkDate_(payload.workDate);
  var pointNo = Number(payload.pointNo);
  var dbVal = payload.db;

  if (!pointNo) throw new Error('測定点Noが未指定です');
  if (dbVal === '' || dbVal == null || isNaN(Number(dbVal))) {
    throw new Error('測定値(dB)を数値で入力してください');
  }

  var point = findNoisePointByNo_(pointNo);
  if (!point) throw new Error('測定点No ' + pointNo + ' はマスタにありません');

  var category = point.category || 'A';
  var pointLabel = point.label || String(pointNo);
  var sheet = ensureNoiseResultSheet_();

  var existingRow = findNoiseResultRowForPoint_(sheet, dateKey, pointNo);
  if (existingRow > 0) {
    sheet.deleteRow(existingRow);
  }

  var now = new Date();
  var targetRow = sheet.getLastRow() + 1;
  sheet.getRange(targetRow, 1, 1, NOISE_RESULT_HEADERS.length).setValues([[
    dateKey,
    category,
    pointLabel,
    pointNo,
    NOISE_FLOOR_LABEL,
    Number(dbVal),
    now,
    String(payload.note || '').trim()
  ]]);

  return {
    success: true,
    pointNo: pointNo,
    category: category,
    db: Number(dbVal),
    workDate: dateKey
  };
}

/**
 * 一覧（管理用）：直近の測定日サマリ
 */
function listNoiseMeasurementDates(limit) {
  limit = limit || 30;
  var sheet = ensureNoiseResultSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var values = sheet.getRange(2, 1, lastRow, 1).getValues();
  var seen = {};
  var dates = [];
  for (var i = values.length - 1; i >= 0; i--) {
    var d = normalizeWorkDate_(values[i][0]);
    if (!d || seen[d]) continue;
    seen[d] = true;
    dates.push(d);
    if (dates.length >= limit) break;
  }
  return dates;
}
