// ========================================

// 騒音測定 — スプレッドシート初期化

// ========================================



function initializeNoiseSpreadsheet(options) {

  options = options || {};

  var forceHeaders = options.forceHeaders === true;

  var ss = SpreadsheetApp.getActiveSpreadsheet();



  ensureNoiseResultSheet_();
  ensureNoiseSessionSheet_();
  ensureNoiseEvalRecordSheet_();



  var masterName = '測定点マスタ';

  var master = ss.getSheetByName(masterName);

  if (!master) master = ss.insertSheet(masterName);



  if (forceHeaders || master.getLastRow() === 0) {

    master.clear();

    var rows = NOISE_MEASUREMENT_POINTS.map(function(p) {

      return [p.no, p.gridCol, p.gridRow];

    });

    master.getRange(1, 1, 1, NOISE_POINT_MASTER_HEADERS.length).setValues([NOISE_POINT_MASTER_HEADERS]);

    if (rows.length) {

      master.getRange(2, 1, 1 + rows.length, NOISE_POINT_MASTER_HEADERS.length).setValues(rows);

    }

    master.getRange(1, 1, 1, NOISE_POINT_MASTER_HEADERS.length)

      .setFontWeight('bold')

      .setBackground('#dbeafe');

    master.setFrozenRows(1);

  }



  writeNoiseSetupGuide_(ss, forceHeaders);



  return '騒音測定シートを初期化しました。\n\n'

    + '・' + NOISE_RESULT_SHEET + '\n'

    + '・' + NOISE_SESSION_SHEET + '\n'

    + '・' + NOISE_EVAL_RECORD_SHEET + '（評価記録）\n'

    + '・測定点マスタ（A測定28セル）\n'

    + '・セットアップ手順\n\n'

    + 'Webアプリ：上部に測定者・時刻、A測定（グリッド）・B測定（ラインマスタ）で dB を入力';

}



function normalizeGuideRowsTo2Cols_(rows) {

  return rows.map(function(row) {

    row = row || [];

    return [

      row[0] != null ? String(row[0]) : '',

      row[1] != null ? String(row[1]) : ''

    ];

  });

}



function writeNoiseSetupGuide_(ss, forceRewrite) {

  var name = 'セットアップ手順';

  var sheet = ss.getSheetByName(name);

  if (!sheet) sheet = ss.insertSheet(name);

  else if (!forceRewrite && sheet.getLastRow() > 3) return;



  sheet.clear();

  var guide = [

    ['騒音測定アプリ｜セットアップ手順'],

    [],

    ['手順', '内容'],

    ['1', 'スプレッドシートに Apps Script を紐づけ、NAS の「騒音測定」フォルダを clasp push'],

    ['2', 'メニュー「騒音測定」→「シート＋測定点マスタを初期化」'],

    ['3', 'Webアプリとしてデプロイ（自分として実行 / アクセスは組織内）'],

    ['4', '現場：測定日・測定時刻・測定者（社員マスタ）を入力し保存'],

    ['4b', 'A測定：グリッド（①-A 等）をタップ → テンキーで dB 入力'],

    ['4c', 'B測定：ラインマスタのライン名をタップ → テンキーで dB 入力'],

    ['4d', '測定完了後「評価を実行して記録保存」→ 別表1で管理区分を自動判定'],

    ['5', 'マップPNG：PDFを画像書き出し→Driveに「騒音測定A測定マップ.png」'],

    ['5b', '見つからない場合：下記「マップファイルID」に Drive のファイルIDを直接記入'],

    ['マップファイルID', '（DriveのPNGファイルID。空欄なら自動検索）'],

    [],

    ['測定点（A測定マップ）', ''],

    ['28セル', 'A測定：グリッド ①-A〜⑦-D（列①〜⑦×行A〜D）'],

    ['B測定', '共通マスタ「ラインマスタ」のライン名（自動読込）'],

    ['測定者', '共通マスタ「社員マスタ」から選択（1日1名）'],

    [],

    ['PDF原図', 'デスクネッツ紙ベースアプリフォーマット\\騒音測定A測定マップ.pdf'],

    [],

    ['シート', '用途'],

    ['騒音測定実績', '日別・測定点別の dB 記録（A/B区分）'],

    ['騒音測定日次', '日別の測定時刻・測定者・測定条件および測定方法・措置概要'],

    ['騒音測定記録', '評価記録（管理区分・3年保存）'],

    ['測定点マスタ', 'A測定28セルの列・行定義'],

  ];



  sheet.getRange(1, 1, guide.length, 2).setValues(normalizeGuideRowsTo2Cols_(guide));

  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setFontSize(12);

  sheet.setColumnWidths(1, 1, 200);

  sheet.setColumnWidths(2, 1, 480);

}



function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu('騒音測定')

    .addItem('シート＋測定点マスタを初期化', 'menuInitializeNoiseSpreadsheet')

    .addItem('測定点マスタを更新（28セル）', 'menuRefreshPointMaster')

    .addItem('マップファイルIDを登録', 'menuRegisterNoiseMapFileId')

    .addItem('Drive権限の承認（初回のみ）', 'menuAuthorizeNoiseDrive')

    .addItem('マップ画像の検出状態を確認', 'menuCheckNoiseMapImage')

    .addToUi();

}



function menuAuthorizeNoiseDrive() {

  var ui = SpreadsheetApp.getUi();

  try {

    authorizeNoiseDriveAccess_();

    ui.alert('完了', 'Google Drive へのアクセス権限が確認できました。\n\n「マップ画像の検出状態を確認」を実行してください。', ui.ButtonSet.OK);

  } catch (e) {

    ui.alert(

      '権限の承認が必要です',

      '自動承認できませんでした。次の手順を実行してください。\n\n'

        + '1. メニュー「拡張機能」→「Apps Script」\n'

        + '2. 関数 authorizeNoiseDriveAccess_ を選択\n'

        + '3. ▶実行 →「権限を確認」→「許可」\n\n'

        + '詳細: ' + e.message,

      ui.ButtonSet.OK

    );

  }

}



function menuRegisterNoiseMapFileId() {

  var ui = SpreadsheetApp.getUi();

  var result = ui.prompt(

    'マップPNGを登録',

    'DriveのファイルID または 共有リンクURL を貼り付けてください。\n\n'

      + '例: https://drive.google.com/file/d/xxxxxxxx/view',

    ui.ButtonSet.OK_CANCEL

  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  try {

    var fileId = setNoiseMapFileId_(result.getResponseText());

    ui.alert(

      'IDを保存しました',

      'ファイルID:\n' + fileId + '\n\n'

        + '次に「Drive権限の承認（初回のみ）」を実行し、\n'

        + 'その後「マップ画像の検出状態を確認」を実行してください。',

      ui.ButtonSet.OK

    );

  } catch (e) {

    ui.alert('登録失敗', e.message, ui.ButtonSet.OK);

  }

}



function menuCheckNoiseMapImage() {

  var status = checkNoiseMapImageForMenu();

  SpreadsheetApp.getUi().alert(

    'マップ画像',

    status.found

      ? ('✅ ' + status.message + '\n\nファイルID:\n' + status.fileId)

      : ('❌ ' + status.message),

    SpreadsheetApp.getUi().ButtonSet.OK

  );

}



function menuRefreshPointMaster() {

  var ui = SpreadsheetApp.getUi();

  if (ui.alert('測定点マスタ更新', '測定点マスタ（28セル・①-A〜⑦-D）を上書きします。\n\n実行しますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) {

    return;

  }

  ui.alert('完了', initializeNoiseSpreadsheet({ forceHeaders: true }), ui.ButtonSet.OK);

}



function menuInitializeNoiseSpreadsheet() {

  var ui = SpreadsheetApp.getUi();

  if (ui.alert('初期化', 'シートと測定点マスタ（28セル）を作成します。実行しますか？', ui.ButtonSet.YES_NO) !== ui.Button.YES) {

    return;

  }

  ui.alert('完了', initializeNoiseSpreadsheet({ forceHeaders: false }), ui.ButtonSet.OK);

}

