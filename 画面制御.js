// ========================================
// 騒音測定 — 画面制御
// ========================================

function doGet(e) {
  e = e || {};
  if (e.parameter && e.parameter.page === 'map') {
    return serveNoiseMapImage_();
  }

  var tpl = HtmlService.createTemplateFromFile('index');
  tpl.mapImageUrl = getNoiseMapImageUrl_();
  return tpl.evaluate()
    .setTitle('騒音測定記録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getNoiseAppInitialData(workMonth, measurementTiming) {
  try {
    var monthKey = normalizeWorkMonth_(workMonth);
    var timingKey = normalizeMeasurementTiming_(measurementTiming);
    var loaded = loadNoiseResultsForMonth_(monthKey, timingKey);
    var mapStatus = getNoiseMapStatus_();
    var aPoints = getNoiseMeasurementPoints_();
    var bPoints = buildLineMeasurementPoints_();
    var boundaryPoints = buildBoundaryMeasurementPoints_();
    var evaluation = getNoiseEvaluationData(monthKey, timingKey);
    var session = loadNoiseSessionForMonth_(monthKey, timingKey);
    return {
      workMonth: monthKey,
      measurementTiming: timingKey,
      workDate: workMonthToRecordDateKey_(monthKey, timingKey),
      isCurrentMonth: monthKey === normalizeWorkMonth_(new Date()),
      aPoints: aPoints,
      bPoints: bPoints,
      boundaryPoints: boundaryPoints,
      points: aPoints.concat(bPoints).concat(boundaryPoints),
      employees: loadEmployeesForNoise_(),
      session: session,
      results: loaded.byPoint,
      evaluation: evaluation,
      gridCols: NOISE_GRID_COLS,
      gridRows: NOISE_GRID_ROWS,
      measurementTimingOptions: NOISE_MEASUREMENT_TIMING_OPTIONS,
      disabledPointNos: listNoiseDisabledPointNos_(),
      mapImageUrl: getNoiseMapImageUrlForApi_(),
      mapStatus: mapStatus,
      mapTitle: '騒音等級評価図 A測定マップ',
      defaultMeasurementMethodConditions: NOISE_DEFAULT_METHOD_CONDITIONS
    };
  } catch (e) {
    Logger.log('getNoiseAppInitialData: ' + e.message + '\n' + e.stack);
    throw new Error('初期データの取得に失敗しました: ' + e.message);
  }
}

function checkNoiseMapImageForMenu() {
  return getNoiseMapStatus_();
}

function saveNoisePointUsageConfig(payload) {
  return saveNoisePointUsageConfig_(payload);
}

function runNoiseEvaluationAndRecord(payload) {
  return runNoiseEvaluationAndRecord_(payload);
}

function getNoiseMeasurementHistory(limit) {
  return listNoiseMeasurementHistory_(limit || 60);
}

function deleteNoiseMeasurementMonth(workMonth) {
  return deleteNoiseMeasurementMonth_(workMonth);
}
