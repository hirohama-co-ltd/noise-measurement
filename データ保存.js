// ========================================
// 騒音測定 — データ保存・読込
// ========================================

var NOISE_RESULT_SHEET = '騒音測定実績';

function detectNoiseResultFormat_(headerRow) {
  var h = (headerRow || []).map(function(c) { return String(c || ''); });
  if (h.indexOf('測定タイミング') >= 0) return 'v4';
  if (h.indexOf('測定者ID') >= 0) return 'v3';
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
  if (detectNoiseResultFormat_(headers) !== 'v4') {
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
  }
  var headers = sheet.getLastRow() >= 1
    ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), NOISE_SESSION_HEADERS.length)).getValues()[0]
    : [];
  if (String(headers[1] || '').indexOf('タイミング') < 0) {
    sheet.getRange(1, 1, 1, NOISE_SESSION_HEADERS.length).setValues([NOISE_SESSION_HEADERS]);
  }
  sheet.getRange(1, 1, 1, NOISE_SESSION_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#dbeafe');
  sheet.setFrozenRows(1);
  return sheet;
}

function parseNoiseResultRow_(row, format) {
  if (format === 'v4') {
    var pointNoV4 = Number(row[3]);
    if (!pointNoV4) return null;
    var recordedAtV4 = '';
    if (row[6] instanceof Date) {
      recordedAtV4 = Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    } else {
      recordedAtV4 = String(row[6] || '');
    }
    return {
      pointNo: pointNoV4,
      category: String(row[1] || ''),
      pointLabel: String(row[2] || ''),
      floor: String(row[4] || ''),
      db: row[5] !== '' && row[5] != null ? Number(row[5]) : null,
      recordedAt: recordedAtV4,
      employeeId: String(row[7] || '').trim(),
      employeeName: String(row[8] || '').trim(),
      measurementTiming: normalizeMeasurementTiming_(row[9]),
      note: String(row[10] || '')
    };
  }
  if (format === 'v3' || format === 'v2') {
    var pointNo = Number(row[3]);
    if (!pointNo) return null;
    var recordedAt = '';
    if (row[6] instanceof Date) {
      recordedAt = Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    } else {
      recordedAt = String(row[6] || '');
    }
    var noteIdx = format === 'v3' ? 9 : 7;
    return {
      pointNo: pointNo,
      category: String(row[1] || ''),
      pointLabel: String(row[2] || ''),
      floor: String(row[4] || ''),
      db: row[5] !== '' && row[5] != null ? Number(row[5]) : null,
      recordedAt: recordedAt,
      employeeId: format === 'v3' ? String(row[7] || '').trim() : '',
      employeeName: format === 'v3' ? String(row[8] || '').trim() : '',
      measurementTiming: '定期',
      note: String(row[noteIdx] || '')
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
    rec.workDate = rowDate;
    byPoint[rec.pointNo] = rec;
  }

  return { workDate: dateKey, byPoint: byPoint };
}

/**
 * 指定月の測定値（測定点No → 当月最新レコード）
 */
function loadNoiseResultsForMonth_(workMonth, measurementTiming) {
  var monthKey = normalizeWorkMonth_(workMonth);
  var timingKey = normalizeMeasurementTiming_(measurementTiming);
  var sheet = ensureNoiseResultSheet_();
  var lastRow = sheet.getLastRow();
  var byPoint = {};

  if (lastRow < 2) return { workMonth: monthKey, measurementTiming: timingKey, byPoint: byPoint };

  var lastCol = Math.max(sheet.getLastColumn(), NOISE_RESULT_HEADERS.length);
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var format = detectNoiseResultFormat_(headerRow);
  var values = sheet.getRange(2, 1, lastRow, lastCol).getValues();

  for (var i = values.length - 1; i >= 0; i--) {
    var rowDate = normalizeWorkDate_(values[i][0]);
    if (!dateBelongsToWorkMonth_(rowDate, monthKey)) continue;
    var rec = parseNoiseResultRow_(values[i], format);
    if (!rec || !rec.pointNo || byPoint[rec.pointNo]) continue;
    if (normalizeMeasurementTiming_(rec.measurementTiming) !== timingKey) continue;
    rec.workDate = rowDate;
    byPoint[rec.pointNo] = rec;
  }

  return { workMonth: monthKey, measurementTiming: timingKey, byPoint: byPoint };
}

function parseMeasuredAt_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  var s = String(value || '').trim();
  if (!s) return new Date();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})/);
  if (m) {
    return new Date(
      parseInt(m[1], 10),
      parseInt(m[2], 10) - 1,
      parseInt(m[3], 10),
      parseInt(m[4], 10),
      parseInt(m[5], 10),
      0, 0
    );
  }
  var d = new Date(s);
  return isNaN(d.getTime()) ? new Date() : d;
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

function loadNoiseSessionForMonth_(workMonth, measurementTiming) {
  var monthKey = normalizeWorkMonth_(workMonth);
  var timingKey = normalizeMeasurementTiming_(measurementTiming);
  var sessionDateKey = workMonthToSessionDateKey_(monthKey);
  var empty = {
    workMonth: monthKey,
    workDate: sessionDateKey,
    measurementTiming: timingKey,
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
  var timingCol = String(headerRow[1] || '').indexOf('タイミング') >= 0 ? 1 : 1;
  var values = sheet.getRange(2, 1, lastRow, lastCol).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var rowDate = normalizeWorkDate_(values[i][0]);
    if (!dateBelongsToWorkMonth_(rowDate, monthKey)) continue;
    var rowTiming = normalizeMeasurementTiming_(values[i][timingCol]);
    if (rowTiming !== timingKey) continue;
    var actionIdx = isLegacySession ? 6 : 5;
    return {
      workMonth: monthKey,
      workDate: sessionDateKey,
      measurementTiming: rowTiming,
      employeeId: String(values[i][2] || '').trim(),
      employeeName: String(values[i][3] || '').trim(),
      measurementMethodConditions: parseSessionMethodConditions_(values[i], headerRow),
      actionSummary: String(values[i][actionIdx] || '').trim()
    };
  }

  return empty;
}

function loadNoiseSessionForDate_(workDate) {
  return loadNoiseSessionForMonth_(normalizeWorkDate_(workDate));
}

function saveNoiseSession(payload) {
  payload = payload || {};
  var monthKey = normalizeWorkMonth_(payload.workMonth || payload.workDate);
  var timingKey = normalizeMeasurementTiming_(payload.measurementTiming);
  var dateKey = workMonthToSessionDateKey_(monthKey);
  var employeeId = String(payload.employeeId || '').trim();
  var employeeName = String(payload.employeeName || '').trim();

  if (!employeeId) throw new Error('測定者を選択してください');

  var sheet = ensureNoiseSessionSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var rows = sheet.getRange(2, 1, lastRow, NOISE_SESSION_HEADERS.length).getValues();
    for (var i = rows.length - 1; i >= 0; i--) {
      if (!dateBelongsToWorkMonth_(normalizeWorkDate_(rows[i][0]), monthKey)) continue;
      if (normalizeMeasurementTiming_(rows[i][1]) !== timingKey) continue;
      sheet.deleteRow(i + 2);
      lastRow = sheet.getLastRow();
    }
  }

  sheet.getRange(lastRow + 1, 1, 1, NOISE_SESSION_HEADERS.length).setValues([[
    dateKey,
    timingKey,
    employeeId,
    employeeName,
    String(payload.measurementMethodConditions || NOISE_DEFAULT_METHOD_CONDITIONS).trim(),
    String(payload.actionSummary || '').trim(),
    new Date()
  ]]);

  return {
    success: true,
    workMonth: monthKey,
    workDate: dateKey,
    measurementTiming: timingKey,
    employeeId: employeeId,
    employeeName: employeeName,
    measurementMethodConditions: String(payload.measurementMethodConditions || NOISE_DEFAULT_METHOD_CONDITIONS).trim(),
    actionSummary: String(payload.actionSummary || '').trim()
  };
}

function saveNoiseSessionFromPayload_(payload) {
  if (!payload || !payload.employeeId) return null;
  return saveNoiseSession(payload);
}

function findNoiseResultRowForPoint_(sheet, dateKey, pointNo) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  var headerRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), NOISE_RESULT_HEADERS.length)).getValues()[0];
  var format = detectNoiseResultFormat_(headerRow);
  var pointCol = format === 'legacy' ? 2 : 4;

  var dates = sheet.getRange(2, 1, lastRow, 1).getValues();
  var points = sheet.getRange(2, pointCol, lastRow, pointCol).getValues();
  for (var i = dates.length - 1; i >= 0; i--) {
    if (normalizeWorkDate_(dates[i][0]) === dateKey && Number(points[i][0]) === pointNo) {
      return i + 2;
    }
  }
  return -1;
}

function deleteNoiseMeasurementForPointInMonth_(workMonth, pointNo, workDate, measurementTiming) {
  var monthKey = normalizeWorkMonth_(workMonth);
  var timingKey = measurementTiming ? normalizeMeasurementTiming_(measurementTiming) : null;
  var sheet = ensureNoiseResultSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  var headerRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), NOISE_RESULT_HEADERS.length)).getValues()[0];
  var format = detectNoiseResultFormat_(headerRow);
  var pointCol = format === 'legacy' ? 2 : 4;
  var lastCol = Math.max(sheet.getLastColumn(), NOISE_RESULT_HEADERS.length);

  if (workDate) {
    var row = findNoiseResultRowForPoint_(sheet, normalizeWorkDate_(workDate), pointNo);
    if (row > 0) {
      if (timingKey && format === 'v4') {
        var rowData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
        var rec = parseNoiseResultRow_(rowData, format);
        if (rec && normalizeMeasurementTiming_(rec.measurementTiming) !== timingKey) {
          row = -1;
        }
      }
      if (row > 0) {
        sheet.deleteRow(row);
        return true;
      }
    }
  }

  var values = sheet.getRange(2, 1, lastRow, lastCol).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var rowDate = normalizeWorkDate_(values[i][0]);
    if (!dateBelongsToWorkMonth_(rowDate, monthKey)) continue;
    var rec = parseNoiseResultRow_(values[i], format);
    if (!rec || rec.pointNo !== pointNo) continue;
    if (timingKey && normalizeMeasurementTiming_(rec.measurementTiming) !== timingKey) continue;
    sheet.deleteRow(i + 2);
    return true;
  }
  return false;
}

function isNoiseDbEmpty_(dbVal) {
  return dbVal === '' || dbVal == null || (typeof dbVal === 'string' && dbVal.trim() === '');
}

/**
 * 測定点の dB を保存（同日・同点は上書き）
 */
function saveNoiseMeasurement(payload) {
  payload = payload || {};
  var pointNo = Number(payload.pointNo);
  var dbVal = payload.db;
  var workMonth = normalizeWorkMonth_(payload.workMonth || payload.workDate);
  var timingKey = normalizeMeasurementTiming_(payload.measurementTiming);
  var measuredAt = parseMeasuredAt_(payload.measuredAt);
  var dateKey = normalizeWorkDate_(payload.workDate || measuredAt);
  var employeeId = String(payload.employeeId || '').trim();
  var employeeName = String(payload.employeeName || '').trim();
  var noteText = String(payload.note || '').trim();

  if (!pointNo) throw new Error('測定点Noが未指定です');

  var point = findNoisePointByNo_(pointNo);
  if (!point) throw new Error('測定点No ' + pointNo + ' はマスタにありません');
  if (point.category === 'A' && point.enabled === false) {
    throw new Error('この測定点は不要として設定されています');
  }

  var isBoundary = point.category === '境界';

  if (isNoiseDbEmpty_(dbVal)) {
    if (isBoundary && noteText) {
      if (!employeeId) throw new Error('測定者を選択してください');
    } else {
      var clearDate = payload.clearWorkDate ? normalizeWorkDate_(payload.clearWorkDate) : dateKey;
      deleteNoiseMeasurementForPointInMonth_(workMonth, pointNo, clearDate, timingKey);
      return {
        success: true,
        cleared: true,
        pointNo: pointNo,
        workMonth: workMonth,
        measurementTiming: timingKey,
        workDate: clearDate
      };
    }
  } else if (isNaN(Number(dbVal))) {
    throw new Error('測定値(dB)を数値で入力してください');
  }

  if (!employeeId) throw new Error('測定者を選択してください');

  var category = point.category || 'A';
  var pointLabel = point.label || String(pointNo);
  var sheet = ensureNoiseResultSheet_();

  var existingRow = findNoiseResultRowForPoint_(sheet, dateKey, pointNo);
  if (existingRow > 0) {
    sheet.deleteRow(existingRow);
  }

  var dbCell = isNoiseDbEmpty_(dbVal) ? '' : Number(dbVal);
  var targetRow = sheet.getLastRow() + 1;
  sheet.getRange(targetRow, 1, 1, NOISE_RESULT_HEADERS.length).setValues([[
    dateKey,
    category,
    pointLabel,
    pointNo,
    NOISE_FLOOR_LABEL,
    dbCell,
    measuredAt,
    employeeId,
    employeeName,
    timingKey,
    noteText
  ]]);

  return {
    success: true,
    cleared: false,
    pointNo: pointNo,
    category: category,
    db: isNoiseDbEmpty_(dbVal) ? null : Number(dbVal),
    workDate: dateKey,
    workMonth: workMonth,
    measurementTiming: timingKey,
    measuredAt: Utilities.formatDate(measuredAt, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'),
    employeeId: employeeId,
    employeeName: employeeName,
    note: noteText
  };
}

/**
 * 一覧：保存済み実施月のサマリ（測定者・件数・評価記録）
 */
function listNoiseMeasurementHistory_(limit) {
  limit = limit || 60;
  var byMonth = {};

  function ensureEntry_(monthKey, timingKey) {
    var listKey = monthKey + '|' + timingKey;
    if (!byMonth[listKey]) {
      byMonth[listKey] = {
        workMonth: monthKey,
        measurementTiming: timingKey,
        employeeName: '',
        aCount: 0,
        bCount: 0,
        pointCount: 0,
        hasEval: false,
        evalSummary: ''
      };
    }
    return byMonth[listKey];
  }

  var resultSheet = ensureNoiseResultSheet_();
  var resultLastRow = resultSheet.getLastRow();
  if (resultLastRow >= 2) {
    var resultLastCol = Math.max(resultSheet.getLastColumn(), NOISE_RESULT_HEADERS.length);
    var resultHeader = resultSheet.getRange(1, 1, 1, resultLastCol).getValues()[0];
    var resultFormat = detectNoiseResultFormat_(resultHeader);
    var resultValues = resultSheet.getRange(2, 1, resultLastRow, resultLastCol).getValues();
    var pointSeen = {};
    for (var i = resultValues.length - 1; i >= 0; i--) {
      var dateKey = normalizeWorkDate_(resultValues[i][0]);
      if (!dateKey) continue;
      var monthKey = normalizeWorkMonth_(dateKey);
      var rec = parseNoiseResultRow_(resultValues[i], resultFormat);
      if (!rec || !rec.pointNo) continue;
      var timingKey = normalizeMeasurementTiming_(rec.measurementTiming);
      var pointKey = monthKey + '|' + timingKey + '|' + rec.pointNo;
      if (pointSeen[pointKey]) continue;
      pointSeen[pointKey] = true;
      var entry = ensureEntry_(monthKey, timingKey);
      entry.pointCount++;
      if (rec.category === '境界' || rec.pointNo >= NOISE_BOUNDARY_POINT_NO_OFFSET + 1) {
        // 境界測定は別集計
      } else if (rec.category === 'B' || rec.pointNo >= NOISE_LINE_POINT_NO_OFFSET + 1) {
        entry.bCount++;
      } else {
        entry.aCount++;
      }
      if (!entry.employeeName && rec.employeeName) {
        entry.employeeName = rec.employeeName;
      }
    }
  }

  var sessionSheet = ensureNoiseSessionSheet_();
  var sessionLastRow = sessionSheet.getLastRow();
  if (sessionLastRow >= 2) {
    var sessionLastCol = Math.max(sessionSheet.getLastColumn(), NOISE_SESSION_HEADERS.length);
    var sessionValues = sessionSheet.getRange(2, 1, sessionLastRow, sessionLastCol).getValues();
    var sessionSeen = {};
    for (var j = sessionValues.length - 1; j >= 0; j--) {
      var sessionDate = normalizeWorkDate_(sessionValues[j][0]);
      if (!sessionDate) continue;
      var sessionMonth = normalizeWorkMonth_(sessionDate);
      var sessionTiming = normalizeMeasurementTiming_(sessionValues[j][1]);
      var sessionKey = sessionMonth + '|' + sessionTiming;
      if (sessionSeen[sessionKey]) continue;
      sessionSeen[sessionKey] = true;
      var sessionEntry = ensureEntry_(sessionMonth, sessionTiming);
      sessionEntry.employeeName = String(sessionValues[j][3] || '').trim() || sessionEntry.employeeName;
    }
  }

  var evalSheet = ensureNoiseEvalRecordSheet_();
  var evalLastRow = evalSheet.getLastRow();
  if (evalLastRow >= 2) {
    var evalValues = evalSheet.getRange(2, 1, evalLastRow, NOISE_EVAL_RECORD_HEADERS.length).getValues();
    var evalSeen = {};
    var evalResults = {};
    for (var k = evalValues.length - 1; k >= 0; k--) {
      var evalDate = normalizeWorkDate_(evalValues[k][0]);
      if (!evalDate) continue;
      var evalMonth = normalizeWorkMonth_(evalDate);
      var evalTiming = evalDate.endsWith('-02') ? '臨時' : '定期';
      ensureEntry_(evalMonth, evalTiming);
      var evalKey = evalMonth + '|' + evalTiming;
      if (!evalSeen[evalKey]) {
        evalSeen[evalKey] = true;
        byMonth[evalKey].hasEval = true;
        evalResults[evalKey] = [];
      }
      var evalText = String(evalValues[k][8] || '').trim();
      if (evalText && evalResults[evalKey].indexOf(evalText) < 0) {
        evalResults[evalKey].push(evalText);
      }
    }
    Object.keys(evalResults).forEach(function(k) {
      if (byMonth[k]) byMonth[k].evalSummary = evalResults[k].join(' / ');
    });
  }

  return Object.keys(byMonth).sort().reverse().slice(0, limit).map(function(m) {
    return byMonth[m];
  });
}

function deleteRowsMatchingMonth_(sheet, monthKey, colIndex) {
  colIndex = colIndex || 1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  var values = sheet.getRange(2, colIndex, lastRow, colIndex).getValues();
  var deleted = 0;
  for (var i = values.length - 1; i >= 0; i--) {
    var rowDate = normalizeWorkDate_(values[i][0]);
    if (!dateBelongsToWorkMonth_(rowDate, monthKey)) continue;
    sheet.deleteRow(i + 2);
    deleted++;
  }
  return deleted;
}

function deleteNoiseMeasurementMonth_(workMonth) {
  var monthKey = normalizeWorkMonth_(workMonth);
  if (!monthKey || !/^\d{4}-\d{2}$/.test(monthKey)) {
    throw new Error('実施月が不正です');
  }

  var resultDeleted = deleteRowsMatchingMonth_(ensureNoiseResultSheet_(), monthKey);
  var sessionDeleted = deleteRowsMatchingMonth_(ensureNoiseSessionSheet_(), monthKey);
  var evalDeleted = deleteRowsMatchingMonth_(ensureNoiseEvalRecordSheet_(), monthKey);

  return {
    success: true,
    workMonth: monthKey,
    resultDeleted: resultDeleted,
    sessionDeleted: sessionDeleted,
    evalDeleted: evalDeleted,
    totalDeleted: resultDeleted + sessionDeleted + evalDeleted
  };
}

/** @deprecated listNoiseMeasurementHistory_ を使用 */
function listNoiseMeasurementDates(limit) {
  return listNoiseMeasurementHistory_(limit).map(function(item) { return item.workMonth; });
}
