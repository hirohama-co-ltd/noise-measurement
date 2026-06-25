// ========================================
// 騒音測定 — 評価・管理区分判定（作業環境測定基準 別表1）
// 参考: https://www.jaish.gr.jp/horei/hor1-64/hor1-64-17-1-6.html
// ========================================

/**
 * エネルギー平均（80dB以下は除外）
 */
function calculateNoiseEnergyAverage_(dbValues) {
  var filtered = [];
  (dbValues || []).forEach(function(v) {
    var n = Number(v);
    if (!isNaN(n) && n > 80) filtered.push(n);
  });
  if (!filtered.length) return null;
  var sum = 0;
  for (var i = 0; i < filtered.length; i++) {
    sum += Math.pow(10, filtered[i] / 10);
  }
  return Math.round(10 * Math.log10(sum / filtered.length) * 10) / 10;
}

function classifyNoiseDbBand_(db) {
  if (db == null || db === '' || isNaN(Number(db))) return null;
  var v = Number(db);
  if (v < 85) return 'le85';
  if (v < 90) return '85to90';
  return 'gt90';
}

function bandLabel_(band) {
  if (band === 'le85') return '85dB未満';
  if (band === '85to90') return '85dB以上90dB未満';
  if (band === 'gt90') return '90dB以上';
  return '—';
}

/**
 * 別表1 評価表による管理区分判定（第Ⅰ〜第Ⅲ）
 * B未測定時は注3によりB=85dB未満として評価
 */
function evaluateManagementClass_(aAvg, bDb, bMeasured) {
  if (aAvg == null || isNaN(Number(aAvg))) {
    return { code: '', label: '—', reason: 'A測定平均が算出できません（80dB超の測定値が必要）' };
  }

  var aBand = classifyNoiseDbBand_(aAvg);
  var bBand = bMeasured ? classifyNoiseDbBand_(bDb) : 'le85';

  var table = {
    le85: { le85: 1, '85to90': 2, gt90: 3 },
    '85to90': { le85: 2, '85to90': 2, gt90: 3 },
    gt90: { le85: 3, '85to90': 3, gt90: 3 }
  };

  var code = table[aBand] && table[aBand][bBand];
  if (!code) {
    return {
      code: '',
      label: '—',
      aBand: aBand,
      bBand: bBand,
      reason: 'A平均' + bandLabel_(aBand) + '・B測定' + bandLabel_(bBand)
    };
  }

  return {
    code: String(code),
    label: NOISE_CLASS_LABELS[code],
    aBand: aBand,
    bBand: bBand,
    reason: 'A平均' + bandLabel_(aBand) + '・B測定' + bandLabel_(bBand)
  };
}

function collectNoiseDbValuesForDate_(workDate) {
  var loaded = loadNoiseResultsForDate(workDate);
  var byPoint = loaded.byPoint;
  var aValues = [];
  var bItems = [];

  var aPoints = getNoiseMeasurementPoints_();
  for (var i = 0; i < aPoints.length; i++) {
    var rec = byPoint[aPoints[i].no];
    if (rec && rec.db != null) aValues.push(rec.db);
  }

  var bPoints = buildLineMeasurementPoints_();
  for (var j = 0; j < bPoints.length; j++) {
    var bRec = byPoint[bPoints[j].no];
    if (bRec && bRec.db != null) {
      bItems.push({
        pointNo: bPoints[j].no,
        label: bPoints[j].label,
        db: bRec.db
      });
    }
  }

  return {
    aValues: aValues,
    bItems: bItems,
    aAvg: calculateNoiseEnergyAverage_(aValues)
  };
}

function buildNoiseEvaluationPreview_(workDate) {
  var dateKey = normalizeWorkDate_(workDate);
  var session = loadNoiseSessionForDate_(dateKey);
  var collected = collectNoiseDbValuesForDate_(dateKey);
  var aAvg = collected.aAvg;
  var evaluations = [];

  if (collected.bItems.length) {
    collected.bItems.forEach(function(b) {
      var ev = evaluateManagementClass_(aAvg, b.db, true);
      evaluations.push({
        unitName: b.label,
        location: b.label,
        aAvg: aAvg,
        bDb: b.db,
        bMeasured: true,
        managementClass: ev.label,
        managementCode: ev.code,
        detail: ev.reason
      });
    });
  } else if (aAvg != null) {
    var evOnlyA = evaluateManagementClass_(aAvg, null, false);
    evaluations.push({
      unitName: NOISE_FLOOR_LABEL + ' A測定',
      location: NOISE_FLOOR_LABEL + ' A測定マップ',
      aAvg: aAvg,
      bDb: null,
      bMeasured: false,
      managementClass: evOnlyA.label,
      managementCode: evOnlyA.code,
      detail: evOnlyA.reason + '（A測定のみ・Bは85dB未満として評価）'
    });
  }

  return {
    workDate: dateKey,
    session: session,
    aAvg: aAvg,
    aCount: collected.aValues.length,
    bCount: collected.bItems.length,
    evaluations: evaluations
  };
}

function formatMeasuredDateTime_(workDate, measuredTime) {
  var dateKey = normalizeWorkDate_(workDate);
  var time = String(measuredTime || '').trim();
  if (!time) time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
  return dateKey + ' ' + time;
}

function formatEvaluationResultText_(ev) {
  var parts = [];
  if (ev.aAvg != null) parts.push('A測定平均 ' + ev.aAvg + 'dB');
  if (ev.bMeasured && ev.bDb != null) parts.push('B測定 ' + ev.bDb + 'dB');
  else if (!ev.bMeasured) parts.push('B測定 未実施（85dB未満として評価）');
  return parts.join(' / ');
}

function ensureNoiseEvalRecordSheet_() {
  var ss = getNoiseSpreadsheet_();
  var sheet = ss.getSheetByName(NOISE_EVAL_RECORD_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(NOISE_EVAL_RECORD_SHEET);
    sheet.getRange(1, 1, 1, NOISE_EVAL_RECORD_HEADERS.length).setValues([NOISE_EVAL_RECORD_HEADERS]);
  }
  sheet.getRange(1, 1, 1, NOISE_EVAL_RECORD_HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#fef3c7');
  sheet.setFrozenRows(1);
  return sheet;
}

function loadNoiseEvalRecordsForDate_(workDate) {
  var dateKey = normalizeWorkDate_(workDate);
  var sheet = ensureNoiseEvalRecordSheet_();
  var lastRow = sheet.getLastRow();
  var list = [];
  if (lastRow < 2) return list;

  var values = sheet.getRange(2, 1, lastRow, NOISE_EVAL_RECORD_HEADERS.length).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (normalizeWorkDate_(values[i][0]) !== dateKey) continue;
    list.push({
      workDate: dateKey,
      unitName: String(values[i][1] || ''),
      measuredAt: String(values[i][2] || ''),
      methodConditions: String(values[i][3] || ''),
      measureLocation: String(values[i][4] || ''),
      measureResult: String(values[i][5] || ''),
      evaluatedAt: values[i][6] instanceof Date
        ? Utilities.formatDate(values[i][6], Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
        : String(values[i][6] || ''),
      evalLocation: String(values[i][7] || ''),
      evalResult: String(values[i][8] || ''),
      evaluatorName: String(values[i][9] || ''),
      actionSummary: String(values[i][10] || ''),
      aAvg: values[i][11] !== '' ? Number(values[i][11]) : null,
      bDb: values[i][12] !== '' && values[i][12] != null ? Number(values[i][12]) : null
    });
  }
  return list;
}

function runNoiseEvaluationAndRecord_(payload) {
  payload = payload || {};
  var dateKey = normalizeWorkDate_(payload.workDate);
  var sessionPayload = {
    workDate: dateKey,
    measuredTime: payload.measuredTime,
    employeeId: payload.employeeId,
    employeeName: payload.employeeName,
    measurementMethodConditions: payload.measurementMethodConditions,
    actionSummary: payload.actionSummary
  };
  saveNoiseSession(sessionPayload);

  var preview = buildNoiseEvaluationPreview_(dateKey);
  if (!preview.evaluations.length) {
    throw new Error('評価に必要なA測定データがありません（80dB超の測定値を入力してください）');
  }

  var methodConditions = String(payload.measurementMethodConditions || NOISE_DEFAULT_METHOD_CONDITIONS).trim();
  var actionSummary = String(payload.actionSummary || '').trim();
  var evaluatorName = String(payload.employeeName || '').trim();
  var measuredAtText = formatMeasuredDateTime_(dateKey, payload.measuredTime);
  var evaluatedAt = new Date();

  var sheet = ensureNoiseEvalRecordSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var existing = sheet.getRange(2, 1, lastRow, 1).getValues();
    for (var i = existing.length - 1; i >= 0; i--) {
      if (normalizeWorkDate_(existing[i][0]) === dateKey) {
        sheet.deleteRow(i + 2);
        lastRow = sheet.getLastRow();
      }
    }
  }

  var rows = [];
  preview.evaluations.forEach(function(ev) {
    rows.push([
      dateKey,
      ev.unitName,
      measuredAtText,
      methodConditions,
      ev.location,
      formatEvaluationResultText_(ev),
      evaluatedAt,
      ev.location,
      ev.managementClass,
      evaluatorName,
      actionSummary,
      ev.aAvg != null ? ev.aAvg : '',
      ev.bMeasured && ev.bDb != null ? ev.bDb : '',
      evaluatedAt
    ]);
  });

  if (rows.length) {
    sheet.getRange(lastRow + 1, 1, rows.length, NOISE_EVAL_RECORD_HEADERS.length).setValues(rows);
  }

  return {
    success: true,
    workDate: dateKey,
    aAvg: preview.aAvg,
    evaluations: preview.evaluations,
    records: loadNoiseEvalRecordsForDate_(dateKey)
  };
}

function getNoiseEvaluationData(workDate) {
  var dateKey = normalizeWorkDate_(workDate);
  return {
    preview: buildNoiseEvaluationPreview_(dateKey),
    records: loadNoiseEvalRecordsForDate_(dateKey)
  };
}
