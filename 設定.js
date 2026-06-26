// ========================================
// 騒音測定 — グリッド測定点定義（列①〜⑦ × 行A〜D = 28点）
// ========================================

var NOISE_RESULT_HEADERS = [
  '測定日', '測定区分', '測定点', '測定点No', '階', '測定値(dB)', '記録日時',
  '測定者ID', '測定者名', '測定タイミング', '備考'
];

var NOISE_SESSION_SHEET = '騒音測定日次';
var NOISE_SESSION_HEADERS = [
  '測定日', '測定タイミング', '測定者ID', '測定者名',
  '測定条件および測定方法', '措置概要', '更新日時'
];

var NOISE_EVAL_RECORD_SHEET = '騒音測定記録';
var NOISE_EVAL_RECORD_HEADERS = [
  '測定日', '評価単位',
  '測定日時', '測定条件および測定方法', '測定箇所', '測定結果',
  '評価日時', '評価箇所', '評価結果', '実施者氏名', '措置概要',
  'A測定平均dB', 'B測定dB', '記録日時'
];

var NOISE_DEFAULT_METHOD_CONDITIONS = [
  '測定器:LA:7200（小野測器）',
  '等価騒音測定:LAeq',
  '動特性:FAST',
  '周波数特性:A特性',
  'レンジ :30dB~120dB',
  '測定時間:10分',
  '測定位置:床上から1.2~1.5m'
].join('\n');

var NOISE_CLASS_LABELS = { 1: '第Ⅰ管理区分', 2: '第Ⅱ管理区分', 3: '第Ⅲ管理区分' };

var NOISE_POINT_MASTER_HEADERS = ['測定点No', '列', '行', '使用'];

/** 騒音測定スプレッドシートID（Webアプリ実行時は openById を使用） */
var NOISE_SPREADSHEET_ID = '1Fx3cG4hlWlhozno66S_jwRiL8UjSz4RlEsuUKDrMDQQ';

/** 共通マスタスプレッドシートID（ラインマスタ・社員マスタ） */
var MASTER_SS_ID = '1FrxPVUeKecY8SXwc5daMxjGT0MzQKZ_toa77PfO4iQo';
var LINE_MASTER_SHEET_NAME = 'ラインマスタ';
var EMPLOYEE_MASTER_SHEET_NAME = '社員マスタ';
var MASTER_CACHE_TTL_SEC = 600;

/** B測定（ライン）の測定点No開始 */
var NOISE_LINE_POINT_NO_OFFSET = 1000;

/** 敷地境界測定（Ⓐ〜Ⓓ）の測定点No開始 */
var NOISE_BOUNDARY_POINT_NO_OFFSET = 2000;

var NOISE_BOUNDARY_LABELS = ['Ⓐ', 'Ⓑ', 'Ⓒ', 'Ⓓ'];

var NOISE_MEASUREMENT_TIMING_OPTIONS = ['定期', '臨時'];

var NOISE_GRID_COLS = ['①', '②', '③', '④', '⑤', '⑥', '⑦'];
var NOISE_GRID_ROWS = ['❶', '❷', '❸', '❹'];

var NOISE_LEGACY_GRID_ROWS = { A: '❶', B: '❷', C: '❸', D: '❹' };

function normalizeMeasurementTiming_(value) {
  var s = String(value || '').trim();
  if (s === '臨時' || s === '临时') return '臨時';
  return '定期';
}

function buildBoundaryMeasurementPoints_() {
  var pts = [];
  for (var i = 0; i < NOISE_BOUNDARY_LABELS.length; i++) {
    pts.push({
      no: NOISE_BOUNDARY_POINT_NO_OFFSET + i + 1,
      category: '境界',
      label: NOISE_BOUNDARY_LABELS[i],
      boundaryLabel: NOISE_BOUNDARY_LABELS[i]
    });
  }
  return pts;
}

/** マップPNG（Google Drive）。空ならスプレッドシートと同じフォルダ内を自動検索 */
var NOISE_MAP_DRIVE_FILE_ID = '';

var NOISE_MAP_IMAGE_NAMES = ['騒音測定A測定マップ.png', '騒音測定A測定マップ.jpg'];

var NOISE_FLOOR_LABEL = '1F';

/** 列優先: ①-A,B,C,D → ②-A… → ⑦-D（計28点） */
function buildAllGridPoints_() {
  var pts = [];
  var no = 1;
  for (var c = 0; c < NOISE_GRID_COLS.length; c++) {
    for (var r = 0; r < NOISE_GRID_ROWS.length; r++) {
      pts.push({
        no: no,
        gridCol: NOISE_GRID_COLS[c],
        gridRow: NOISE_GRID_ROWS[r],
        label: NOISE_GRID_COLS[c] + '-' + NOISE_GRID_ROWS[r]
      });
      no++;
    }
  }
  return pts;
}

var NOISE_MEASUREMENT_POINTS = buildAllGridPoints_();

function pointNoToGrid_(no) {
  var all = buildAllGridPoints_();
  var pt = all[no - 1];
  if (pt) {
    return {
      col: pt.gridCol,
      row: pt.gridRow,
      colIdx: NOISE_GRID_COLS.indexOf(pt.gridCol),
      rowIdx: NOISE_GRID_ROWS.indexOf(pt.gridRow),
      label: pt.label
    };
  }
  return { col: '①', row: 'A', colIdx: 0, rowIdx: 0, label: '①-A' };
}

function parseGridCol_(value) {
  var s = String(value || '').trim();
  for (var i = 0; i < NOISE_GRID_COLS.length; i++) {
    if (s === NOISE_GRID_COLS[i]) return { idx: i, col: NOISE_GRID_COLS[i] };
  }
  var n = parseInt(s, 10);
  if (n >= 1 && n <= NOISE_GRID_COLS.length) {
    return { idx: n - 1, col: NOISE_GRID_COLS[n - 1] };
  }
  return null;
}

function parseGridRow_(value) {
  var s = String(value || '').trim();
  var idx = NOISE_GRID_ROWS.indexOf(s);
  if (idx >= 0) return { idx: idx, row: NOISE_GRID_ROWS[idx] };
  var upper = s.toUpperCase();
  if (NOISE_LEGACY_GRID_ROWS[upper]) {
    var legacy = NOISE_LEGACY_GRID_ROWS[upper];
    return { idx: NOISE_GRID_ROWS.indexOf(legacy), row: legacy };
  }
  return null;
}

function buildPointFromGrid_(no, colVal, rowVal) {
  var col = parseGridCol_(colVal);
  var row = parseGridRow_(rowVal);
  if (!col || !row) {
    var fallback = pointNoToGrid_(no);
    col = { idx: fallback.colIdx, col: fallback.col };
    row = { idx: fallback.rowIdx, row: fallback.row };
  }
  return {
    no: no,
    gridCol: col.col,
    gridRow: row.row,
    label: col.col + '-' + row.row
  };
}

function toNoisePointForApi_(pt) {
  return {
    no: pt.no,
    label: pt.label,
    gridCol: pt.gridCol,
    gridRow: pt.gridRow,
    category: 'A',
    enabled: pt.enabled !== false
  };
}

function isNoisePointEnabled_(value) {
  var s = String(value || '').trim();
  if (!s) return true;
  var lower = s.toLowerCase();
  if (s === '不要' || s === '×' || s === '✕' || s === '非使用' || s === '0') return false;
  if (lower === 'false' || lower === 'no' || lower === 'off' || lower === 'disabled') return false;
  return true;
}

function findPointUsageColIndex_(headerRow) {
  var h = (headerRow || []).map(function(c) { return String(c || '').trim(); });
  var idx = h.indexOf('使用');
  if (idx >= 0) return idx;
  if (h.length >= 4) return 3;
  return -1;
}

function getNoiseDisabledPointNos_() {
  try {
    var raw = PropertiesService.getDocumentProperties().getProperty('NOISE_DISABLED_POINT_NOS');
    if (!raw) return {};
    var arr = JSON.parse(raw);
    var map = {};
    (arr || []).forEach(function(n) {
      n = Number(n);
      if (n > 0) map[n] = true;
    });
    return map;
  } catch (e) {
    Logger.log('getNoiseDisabledPointNos_: ' + e.message);
    return {};
  }
}

function setNoiseDisabledPointNos_(pointNos) {
  var list = (pointNos || []).map(function(n) { return Number(n); }).filter(function(n) { return n > 0; });
  PropertiesService.getDocumentProperties().setProperty(
    'NOISE_DISABLED_POINT_NOS',
    JSON.stringify(list)
  );
  return list;
}

function ensurePointMasterRows_(sheet) {
  if (!sheet) return;
  ensurePointMasterUsageColumn_(sheet);
  var defaults = buildAllGridPoints_();
  var lastRow = sheet.getLastRow();
  if (lastRow >= 1 + defaults.length) return;

  var existing = {};
  if (lastRow >= 2) {
    var readCol = Math.max(sheet.getLastColumn(), 4);
    var values = sheet.getRange(2, 1, lastRow, readCol).getValues();
    for (var i = 0; i < values.length; i++) {
      var no = Number(values[i][0]);
      if (no) existing[no] = values[i];
    }
  }

  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, NOISE_POINT_MASTER_HEADERS.length).setValues([NOISE_POINT_MASTER_HEADERS]);
    sheet.getRange(1, 1, 1, NOISE_POINT_MASTER_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#dbeafe');
    sheet.setFrozenRows(1);
  }

  var disabledMap = getNoiseDisabledPointNos_();
  var rows = defaults.map(function(p) {
    if (existing[p.no]) {
      var row = existing[p.no].slice();
      while (row.length < 4) row.push('使用');
      if (disabledMap[p.no]) row[3] = '不要';
      return row.slice(0, 4);
    }
    return [p.no, p.gridCol, p.gridRow, disabledMap[p.no] ? '不要' : '使用'];
  });
  sheet.getRange(2, 1, rows.length, 4).setValues(rows);
}

function ensurePointMasterUsageColumn_(sheet) {
  if (!sheet) return;
  var lastCol = Math.max(sheet.getLastColumn(), 4);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (String(headers[3] || '').trim() === '使用') return;

  sheet.getRange(1, 4).setValue('使用').setFontWeight('bold').setBackground('#dbeafe');
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var rowCount = lastRow - 1;
    var usage = [];
    for (var i = 0; i < rowCount; i++) usage.push(['使用']);
    sheet.getRange(2, 4, rowCount, 1).setValues(usage);
  }
}

function saveNoisePointUsageConfig_(payload) {
  payload = payload || {};
  var updates = payload.points || [];
  if (!updates.length) throw new Error('保存する測定点がありません');

  var disabledNos = [];
  updates.forEach(function(u) {
    var no = Number(u.no);
    if (!no) return;
    if (u.enabled === false) disabledNos.push(no);
  });
  setNoiseDisabledPointNos_(disabledNos);

  var ss = getNoiseSpreadsheet_();
  var sheet = ss.getSheetByName('測定点マスタ');
  if (!sheet) throw new Error('測定点マスタがありません。メニューから初期化してください。');

  ensurePointMasterRows_(sheet);

  var byNo = {};
  updates.forEach(function(u) {
    var no = Number(u.no);
    if (!no) return;
    byNo[no] = u.enabled !== false;
  });

  var lastRow = sheet.getLastRow();
  var saved = 0;
  if (lastRow >= 2) {
    var noValues = sheet.getRange(2, 1, lastRow, 1).getValues();
    for (var i = 0; i < noValues.length; i++) {
      var pointNo = Number(noValues[i][0]);
      if (!pointNo || byNo[pointNo] === undefined) continue;
      sheet.getRange(i + 2, 4).setValue(byNo[pointNo] ? '使用' : '不要');
      saved++;
    }
  }

  return { success: true, count: saved, disabledPointNos: disabledNos };
}

function parseNoisePointRow_(row, format) {
  var no = Number(row[0]);
  if (!no) return null;

  if (format === 'grid') {
    return buildPointFromGrid_(no, row[1], row[2]);
  }

  if (format === 'badge') {
    var label = String(row[1] || '').trim();
    var g = parseGridLabel_(label);
    if (g) return buildPointFromGrid_(no, g.col, g.row);
    return buildPointFromGrid_(no, null, null);
  }

  if (format === 'legacy6') {
    var g6 = parseGridLabel_(String(row[1] || '').trim());
    if (g6) return buildPointFromGrid_(no, g6.col, g6.row);
    return buildPointFromGrid_(no, null, null);
  }

  var legacyLabel = String(row[2] || '').trim();
  var gl = parseGridLabel_(legacyLabel);
  if (gl) return buildPointFromGrid_(no, gl.col, gl.row);
  return buildPointFromGrid_(no, null, null);
}

function parseGridLabel_(label) {
  label = String(label || '').trim();
  var m = label.match(/^([①②③④⑤⑥⑦])[-－]([A-Da-d❶❷❸❹])$/);
  if (!m) return null;
  var rowParsed = parseGridRow_(m[2]);
  if (!rowParsed) return null;
  return { col: m[1], row: rowParsed.row };
}

function detectNoisePointMasterFormat_(headerRow) {
  var h = (headerRow || []).map(function(c) { return String(c || ''); });
  if (h.indexOf('列') >= 0 && h.indexOf('行') >= 0) return 'grid';
  if (h.indexOf('番号X(%)') >= 0 && h.indexOf('表示名') >= 0) return 'badge';
  if (h.indexOf('番号X(%)') >= 0) return 'gridLegacy';
  if (h.indexOf('dB欄X(%)') >= 0 || h.indexOf('タップX(%)') >= 0) return 'legacy6';
  if (h.indexOf('階') >= 0 || h.indexOf('マップX(%)') >= 0) return 'legacy5';
  return 'grid';
}

function getNoiseSpreadsheet_() {
  if (NOISE_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(NOISE_SPREADSHEET_ID);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('スプレッドシートに接続できません');
  return ss;
}

function getNoiseMeasurementPoints_() {
  var defaults = buildAllGridPoints_();
  var byNo = {};
  var disabledMap = getNoiseDisabledPointNos_();
  defaults.forEach(function(p) {
    byNo[p.no] = Object.assign({}, p, { enabled: !disabledMap[p.no] });
  });

  try {
    var ss = getNoiseSpreadsheet_();
    var sheet = ss.getSheetByName('測定点マスタ');
    if (sheet && sheet.getLastRow() >= 2) {
      var lastCol = Math.max(sheet.getLastColumn(), 4);
      var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var format = detectNoisePointMasterFormat_(headerRow);
      var usageCol = findPointUsageColIndex_(headerRow);
      var colCount = format === 'legacy5' ? 5 : (format === 'legacy6' ? 6 : Math.max(4, lastCol));
      var lastRow = sheet.getLastRow();
      var values = sheet.getRange(2, 1, lastRow, colCount).getValues();
      for (var i = 0; i < values.length; i++) {
        var pt = parseNoisePointRow_(values[i], format);
        if (pt && pt.no >= 1 && pt.no <= defaults.length) {
          var enabled = true;
          if (usageCol >= 0 && values[i].length > usageCol) {
            enabled = isNoisePointEnabled_(values[i][usageCol]);
          } else if (disabledMap[pt.no]) {
            enabled = false;
          }
          byNo[pt.no] = Object.assign({}, pt, { enabled: enabled });
        }
      }
    }
  } catch (e) {
    Logger.log('getNoiseMeasurementPoints_: ' + e.message);
  }

  Object.keys(byNo).forEach(function(k) {
    if (disabledMap[Number(k)]) byNo[k].enabled = false;
  });

  return Object.keys(byNo).sort(function(a, b) { return Number(a) - Number(b); })
    .map(function(k) {
      var pt = byNo[k];
      if (pt.enabled === undefined) pt.enabled = true;
      return toNoisePointForApi_(pt);
    });
}

function listNoiseDisabledPointNos_() {
  var map = getNoiseDisabledPointNos_();
  return Object.keys(map).map(function(k) { return Number(k); }).sort(function(a, b) { return a - b; });
}

function extractDriveFileId_(text) {
  var s = String(text || '').trim();
  if (!s) return '';
  var m = s.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  if (/^[a-zA-Z0-9_-]{20,}$/.test(s)) return s;
  return '';
}

function getNoiseMapFileId_() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('noise_map_file_id');
    if (cached) return cached === 'NONE' ? '' : cached;
  } catch (e) { /* キャッシュ不可 */ }

  var fileId = extractDriveFileId_(NOISE_MAP_DRIVE_FILE_ID);
  if (fileId) {
    putNoiseMapFileIdCache_(fileId);
    return fileId;
  }

  fileId = extractDriveFileId_(
    PropertiesService.getDocumentProperties().getProperty('NOISE_MAP_FILE_ID')
  );
  if (fileId) {
    putNoiseMapFileIdCache_(fileId);
    return fileId;
  }

  fileId = findNoiseMapFileIdFromSheet_();
  if (fileId) {
    putNoiseMapFileIdCache_(fileId);
    return fileId;
  }

  fileId = findNoiseMapFileIdInDrive_();
  putNoiseMapFileIdCache_(fileId || '');
  return fileId;
}

function putNoiseMapFileIdCache_(fileId) {
  try {
    CacheService.getScriptCache().put('noise_map_file_id', fileId || 'NONE', 600);
  } catch (e) { Logger.log(e); }
}

function setNoiseMapFileId_(fileIdOrUrl) {
  var fileId = extractDriveFileId_(fileIdOrUrl);
  if (!fileId) throw new Error('DriveのファイルIDまたはURLを入力してください');
  PropertiesService.getDocumentProperties().setProperty('NOISE_MAP_FILE_ID', fileId);
  putNoiseMapFileIdCache_(fileId);
  syncNoiseMapFileIdToSheet_(fileId, '');
  return fileId;
}

/** Apps Script エディタまたはメニューから1回実行して Drive 権限を承認 */
function authorizeNoiseDriveAccess_() {
  DriveApp.getRootFolder().getName();
  return 'Google Drive へのアクセス権限を確認しました。';
}

function syncNoiseMapFileIdToSheet_(fileId, fileName) {
  try {
    var ss = getNoiseSpreadsheet_();
    var sheet = ss.getSheetByName('セットアップ手順');
    if (!sheet) return;
    var lastRow = Math.max(sheet.getLastRow(), 15);
    var values = sheet.getRange(1, 1, lastRow, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === 'マップファイルID') {
        sheet.getRange(i + 1, 2).setValue(fileId);
        return;
      }
    }
    sheet.getRange(lastRow + 1, 1, 1, 2).setValues([['マップファイルID', fileId]]);
  } catch (e) { Logger.log(e); }
}

function findNoiseMapFileIdFromSheet_() {
  try {
    var ss = getNoiseSpreadsheet_();
    var sheetNames = ['セットアップ手順', 'マップ設定'];
    for (var s = 0; s < sheetNames.length; s++) {
      var sheet = ss.getSheetByName(sheetNames[s]);
      if (!sheet) continue;
      var lastRow = Math.max(sheet.getLastRow(), 20);
      var values = sheet.getRange(1, 1, lastRow, 2).getValues();
      for (var i = 0; i < values.length; i++) {
        var key = String(values[i][0] || '').trim();
        var val = String(values[i][1] || '').trim();
        if (key === 'マップファイルID' || key === 'マップPNGファイルID') {
          var id = extractDriveFileId_(val);
          if (id) return id;
        }
        if (val && (key.indexOf('マップ') >= 0 || key.indexOf('map') >= 0 || key.indexOf('Map') >= 0)) {
          var idFromCell = extractDriveFileId_(val);
          if (idFromCell) return idFromCell;
        }
      }
    }
  } catch (e) { Logger.log(e); }
  return '';
}

function findNoiseMapFileIdInDrive_() {
  try {
    var ss = getNoiseSpreadsheet_();
    var ssFile = DriveApp.getFileById(ss.getId());
    var parents = ssFile.getParents();
    while (parents.hasNext()) {
      var id = findMapInFolderTree_(parents.next(), 0);
      if (id) return id;
    }

    for (var n = 0; n < NOISE_MAP_IMAGE_NAMES.length; n++) {
      var name = NOISE_MAP_IMAGE_NAMES[n];
      var files = DriveApp.searchFiles(
        'title contains "' + name.replace(/"/g, '') + '" and trashed = false'
      );
      if (files.hasNext()) return files.next().getId();
    }

    var queries = [
      'title contains "騒音測定A測定マップ" and trashed = false',
      'title contains "A測定マップ" and trashed = false',
      'fullText contains "騒音測定" and title contains "マップ" and trashed = false'
    ];
    for (var q = 0; q < queries.length; q++) {
      var partial = DriveApp.searchFiles(queries[q]);
      while (partial.hasNext()) {
        var f = partial.next();
        if (isNoiseMapMime_(f.getMimeType())) return f.getId();
      }
    }
  } catch (e) { Logger.log(e); }
  return '';
}

function isNoiseMapMime_(mime) {
  mime = String(mime || '');
  return mime.indexOf('image/') === 0 || mime === 'application/pdf';
}

function findMapInFolderTree_(folder, depth) {
  if (depth > 2) return '';
  var id = findMapInFolder_(folder);
  if (id) return id;
  var subs = folder.getFolders();
  while (subs.hasNext()) {
    id = findMapInFolderTree_(subs.next(), depth + 1);
    if (id) return id;
  }
  return '';
}

function findMapInFolder_(folder) {
  for (var n = 0; n < NOISE_MAP_IMAGE_NAMES.length; n++) {
    var files = folder.getFilesByName(NOISE_MAP_IMAGE_NAMES[n]);
    if (files.hasNext()) return files.next().getId();
  }
  var it = folder.getFiles();
  while (it.hasNext()) {
    var f = it.next();
    var name = f.getName();
    if (name.indexOf('騒音測定') >= 0 && name.indexOf('マップ') >= 0 && isNoiseMapMime_(f.getMimeType())) {
      return f.getId();
    }
  }
  return '';
}

/** 本番WebアプリのデプロイID（マッププロキシURL用） */
var NOISE_WEBAPP_DEPLOY_ID = 'AKfycbyU0LtpMyHXNyR4UtyFYIZABOHWXg7UGZBIkOhpgnwxT7Qq860q3idCBzWVgsIghN34vQ';

/** 画面表示用（doGet）。HTML直接出力なので base64 埋め込み可 */
function getNoiseMapImageUrl_() {
  return buildNoiseMapImageSrc_(3 * 1024 * 1024);
}

/** google.script.run 用。base64/URLは返さない（転送上限・img読込失敗を避ける） */
function getNoiseMapImageUrlForApi_() {
  return '';
}

function buildNoiseMapImageSrc_(maxInlineBytes) {
  maxInlineBytes = maxInlineBytes || 50 * 1024;
  var fileId = getNoiseMapFileId_();
  if (!fileId) return '';
  try {
    var file = DriveApp.getFileById(fileId);
    var mime = file.getMimeType() || '';
    if (mime === 'application/pdf') {
      return getNoiseMapProxyUrl_();
    }
    var blob = file.getBlob();
    var bytes = blob.getBytes();
    if (bytes.length > 0 && bytes.length <= maxInlineBytes) {
      return 'data:' + (blob.getContentType() || 'image/png') + ';base64,'
        + Utilities.base64Encode(bytes);
    }
  } catch (e) {
    Logger.log('buildNoiseMapImageSrc_: ' + e);
  }
  return getNoiseMapProxyUrl_();
}

function getNoiseMapProxyUrl_() {
  if (NOISE_WEBAPP_DEPLOY_ID) {
    return 'https://script.google.com/macros/s/' + NOISE_WEBAPP_DEPLOY_ID + '/exec?page=map';
  }
  try {
    var url = ScriptApp.getService().getUrl();
    if (url) return url + (url.indexOf('?') >= 0 ? '&' : '?') + 'page=map';
  } catch (e) { Logger.log(e); }
  return '';
}

function serveNoiseMapImage_() {
  var fileId = getNoiseMapFileId_();
  if (!fileId) {
    return ContentService.createTextOutput('Map image not found')
      .setMimeType(ContentService.MimeType.TEXT);
  }
  try {
    var file = DriveApp.getFileById(fileId);
    var mime = file.getMimeType() || '';
    if (mime === 'application/pdf') {
      var thumb = UrlFetchApp.fetch(
        'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1600',
        { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }
      ).getBlob();
      return ContentService.create(thumb);
    }
    return ContentService.create(file.getBlob())
      .setMimeType(file.getMimeType() || 'image/png');
  } catch (e) {
    Logger.log('serveNoiseMapImage_: ' + e);
    return ContentService.createTextOutput('Error: ' + e.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function getNoiseMapStatus_() {
  var fileId = getNoiseMapFileId_();
  if (!fileId) {
    var diag = getNoiseMapSearchDiagnostic_();
    return {
      found: false,
      message: 'マップ画像が見つかりません。\n\n'
        + '【対処】メニュー「マップファイルIDを登録」から Drive の PNG を指定してください。\n\n'
        + diag
    };
  }
  try {
    var file = DriveApp.getFileById(fileId);
    return {
      found: true,
      fileId: fileId,
      fileName: file.getName(),
      mimeType: file.getMimeType(),
      message: '検出: ' + file.getName()
    };
  } catch (e) {
    return { found: false, message: 'ファイルIDは設定されていますが読み込めません: ' + e.message };
  }
}

function getNoiseMapSearchDiagnostic_() {
  try {
    var ss = getNoiseSpreadsheet_();
    var ssFile = DriveApp.getFileById(ss.getId());
    var folderNames = [];
    var parents = ssFile.getParents();
    while (parents.hasNext()) {
      folderNames.push(parents.next().getName());
    }
    return 'スプレッドシート: ' + ss.getName() + '\n'
      + '配置フォルダ: ' + (folderNames.length ? folderNames.join(', ') : '（マイドライブ直下）') + '\n'
      + '期待ファイル名: 騒音測定A測定マップ.png';
  } catch (e) {
    return '';
  }
}

function findNoiseMapFileIdInSpreadsheetFolder_() {
  return findNoiseMapFileIdInDrive_();
}

function normalizeWorkDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(value || '').trim();
  if (!s) {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m) {
    return m[1] + '-' + String(parseInt(m[2], 10)).padStart(2, '0') + '-'
      + String(parseInt(m[3], 10)).padStart(2, '0');
  }
  return s;
}

function normalizeWorkMonth_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  var s = String(value || '').trim();
  if (!s) {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  }
  var full = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (full) {
    return full[1] + '-' + String(parseInt(full[2], 10)).padStart(2, '0');
  }
  var monthOnly = s.match(/^(\d{4})[-\/](\d{1,2})$/);
  if (monthOnly) {
    return monthOnly[1] + '-' + String(parseInt(monthOnly[2], 10)).padStart(2, '0');
  }
  return s;
}

function workMonthToSessionDateKey_(workMonth) {
  return normalizeWorkMonth_(workMonth) + '-01';
}

function workMonthToRecordDateKey_(workMonth, measurementTiming) {
  var monthKey = normalizeWorkMonth_(workMonth);
  return normalizeMeasurementTiming_(measurementTiming) === '臨時' ? (monthKey + '-02') : (monthKey + '-01');
}

function dateBelongsToWorkMonth_(dateKey, workMonth) {
  return normalizeWorkDate_(dateKey).indexOf(normalizeWorkMonth_(workMonth)) === 0;
}
