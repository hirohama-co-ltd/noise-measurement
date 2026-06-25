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

function getNoiseAppInitialData(workDate) {
  try {
    var dateKey = normalizeWorkDate_(workDate);
    var loaded = loadNoiseResultsForDate(dateKey);
    var mapStatus = getNoiseMapStatus_();
    var aPoints = getNoiseMeasurementPoints_();
    var bPoints = buildLineMeasurementPoints_();
    var evaluation = getNoiseEvaluationData(dateKey);
    return {
      workDate: dateKey,
      isToday: dateKey === normalizeWorkDate_(new Date()),
      aPoints: aPoints,
      bPoints: bPoints,
      points: aPoints.concat(bPoints),
      employees: loadEmployeesForNoise_(),
      session: loadNoiseSessionForDate_(dateKey),
      results: loaded.byPoint,
      evaluation: evaluation,
      gridCols: NOISE_GRID_COLS,
      gridRows: NOISE_GRID_ROWS,
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
