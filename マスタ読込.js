// ========================================
// 騒音測定 — 共通マスタ読込（ラインマスタ・社員マスタ）
// ========================================

function getCachedJson_(key) {
  try {
    var raw = CacheService.getScriptCache().get(key);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function putCachedJson_(key, value, ttlSec) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), ttlSec || MASTER_CACHE_TTL_SEC);
  } catch (e) { Logger.log(e); }
}

function loadLineMasterFromSheet_() {
  var defaultLines = ['1号ライン', '2号ライン', '3号ライン'];
  var cacheKey = 'noise_lineMaster_' + (MASTER_SS_ID || 'default');
  var cached = getCachedJson_(cacheKey);
  if (cached && cached.length) return cached;

  if (!MASTER_SS_ID) return defaultLines;
  try {
    var masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    var sheet = masterSs.getSheetByName(LINE_MASTER_SHEET_NAME);
    if (!sheet) return defaultLines;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return defaultLines;

    var values = sheet.getRange(2, 1, lastRow, 1).getValues();
    var lineList = [];
    for (var i = 0; i < values.length; i++) {
      var lineName = String(values[i][0] || '').trim();
      if (lineName) lineList.push(lineName);
    }
    var result = lineList.length > 0 ? lineList : defaultLines;
    putCachedJson_(cacheKey, result);
    return result;
  } catch (e) {
    Logger.log('loadLineMasterFromSheet_: ' + e.message);
    return defaultLines;
  }
}

function findNoiseMasterColumnIndex_(headers, aliases) {
  for (var j = 0; j < aliases.length; j++) {
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i] || '').replace(/\s/g, '');
      if (h === aliases[j]) return i;
    }
  }
  for (var k = 0; k < aliases.length; k++) {
    for (var m = 0; m < headers.length; m++) {
      var h2 = String(headers[m] || '').replace(/\s/g, '');
      if (h2.indexOf(aliases[k]) >= 0) return m;
    }
  }
  return -1;
}

function isActiveEmployeeRow_(value) {
  var s = String(value || '').trim();
  if (!s) return true;
  return s !== '無効' && s !== '0' && s.toLowerCase() !== 'false';
}

function normalizeEmployeeCode_(value) {
  var s = String(value || '').trim().toUpperCase();
  if (!s) return '';
  var m = s.match(/^EMP(\d+)$/);
  if (m) return 'EMP' + String(parseInt(m[1], 10)).padStart(3, '0');
  return s;
}

function loadEmployeesForNoise_() {
  var cacheKey = 'noise_employees_' + (MASTER_SS_ID || 'default');
  var cached = getCachedJson_(cacheKey);
  if (cached) return cached;

  if (!MASTER_SS_ID) return [];
  try {
    var masterSs = SpreadsheetApp.openById(MASTER_SS_ID);
    var sheet = masterSs.getSheetByName(EMPLOYEE_MASTER_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) {
      putCachedJson_(cacheKey, []);
      return [];
    }

    var lastCol = Math.max(sheet.getLastColumn(), 8);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(function(h) { return String(h || '').trim(); });
    var idCol = findNoiseMasterColumnIndex_(headers, ['社員ID', 'QRコード', 'ID']);
    var nameCol = findNoiseMasterColumnIndex_(headers, ['氏名', '作業者名', '名前']);
    var activeCol = findNoiseMasterColumnIndex_(headers, ['有効']);
    if (idCol < 0) idCol = 0;
    if (nameCol < 0) nameCol = 1;

    var values = sheet.getRange(2, 1, sheet.getLastRow(), lastCol).getValues();
    var list = [];
    for (var i = 0; i < values.length; i++) {
      if (activeCol >= 0 && !isActiveEmployeeRow_(values[i][activeCol])) continue;
      var id = normalizeEmployeeCode_(values[i][idCol]);
      var name = String(values[i][nameCol] || '').trim();
      if (!id || !name) continue;
      list.push({ id: id, name: name });
    }

    list.sort(function(a, b) {
      return a.id.localeCompare(b.id, 'ja');
    });
    putCachedJson_(cacheKey, list);
    return list;
  } catch (e) {
    Logger.log('loadEmployeesForNoise_: ' + e.message);
    return [];
  }
}

function buildLineMeasurementPoints_() {
  var lines = loadLineMasterFromSheet_();
  var pts = [];
  for (var i = 0; i < lines.length; i++) {
    pts.push({
      no: NOISE_LINE_POINT_NO_OFFSET + i + 1,
      category: 'B',
      label: lines[i],
      lineName: lines[i]
    });
  }
  return pts;
}

function findNoisePointByNo_(pointNo) {
  pointNo = Number(pointNo);
  if (!pointNo) return null;

  var aPoints = getNoiseMeasurementPoints_();
  for (var i = 0; i < aPoints.length; i++) {
    if (aPoints[i].no === pointNo) {
      return Object.assign({ category: 'A' }, aPoints[i]);
    }
  }

  var bPoints = buildLineMeasurementPoints_();
  for (var j = 0; j < bPoints.length; j++) {
    if (bPoints[j].no === pointNo) return bPoints[j];
  }
  return null;
}
