// ==================================================
// 設定エリア
// ==================================================
// モデル名を変更しやすいように定数化
const DEFAULT_MODEL_NAME = 'gemini-2.5-flash-lite';
const MODEL_CANDIDATES = ['gemini-2.5-flash-lite', 'gemini-3-flash-preview'];

const SETTINGS_SHEET_NAME = '設定';
const ANALYSIS_SHEET_NAME = '解析シート';
const DESTINATION_SHEET_NAME = '移動先リスト';
const FILE_NAME_RULE_SHEET_NAME = 'ファイル名ルール';
const ACCOUNT_RULE_SHEET_NAME = '勘定科目ルール';
const LOG_SHEET_NAME = 'ログ';
const HELP_URL = 'https://ztbtlab.github.io/receipt-parser-gas/';

const SETTINGS_KEYS = {
  folderUrl: '対象フォルダURL',
  delimiter: '区切り文字',
  fileNameRuleSheet: 'ファイル名ルール参照',
  scanLimit: '解析上限',
  logOutput: 'ログ出力',
  creditAccountCard: '貸方勘定科目(クレカ)',
  creditAccountOther: '貸方勘定科目(それ以外)',
  creditSubAccountCard: '貸方補助科目(クレカ)',
  mfSummaryMode: 'マネフォ摘要欄'
};

const DELIMITER_CANDIDATES = ['-', '_', '｜'];
const LOG_OUTPUT_CANDIDATES = ['ON', 'OFF'];
const CREDIT_SUB_ACCOUNT_CANDIDATES = ['カード情報', '空欄'];
const MF_SUMMARY_MODE_CANDIDATES = ['購入内容', '取引先名', '取引先名＋購入内容'];

const EMPTY_CELL_COLOR = '#fff2cc';
const DUPLICATE_CELL_COLOR = '#f4cccc';
const MIXED_TAX_CELL_COLOR = '#fce5cd';
const POSSIBLE_DUPLICATE_ROW_COLOR = '#fce8b2';

const MF_TAX_CATEGORY_STANDARD = '課仕 10%';
const MF_TAX_CATEGORY_REDUCED = '課仕 8%';
const MF_CARD_SUB_ACCOUNT = '三井住友ゴールドカード';
const MF_PROCESSED_PREFIX = 'CSV済';
const MF_TRADE_PARTNER_SHEET_NAME = '取引先一覧';
const MF_TRADE_PARTNER_HEADERS = [
  'コード',
  '取引先名',
  '検索キー',
  '表示設定',
  '登録番号',
  '法人番号'
];
const ANALYSIS_STATUS_ERROR_PREFIX = 'エラー:';
const ANALYSIS_CLEAR_ERROR_OPTION_VALUE = '__ERROR_PREFIX__';
const ANALYSIS_CLEAR_ERROR_OPTION_LABEL = 'エラー（エラー:〜）';
const ANALYSIS_CLEAR_STATUS_JOB_CACHE_PREFIX = 'analysis_clear_status_job:';
const ANALYSIS_CLEAR_STATUS_JOB_CACHE_TTL_SEC = 60 * 60;
const ANALYSIS_CLEAR_STATUS_SCAN_WINDOW_ROWS = 500;
const ANALYSIS_CLEAR_STATUS_MAX_DELETE_ROWS = 50;
const MF_CSV_HEADERS = [
  '取引No',
  '取引日',
  '借方勘定科目',
  '借方補助科目',
  '借方部門',
  '借方取引先',
  '借方税区分',
  '借方インボイス',
  '借方金額(円)',
  '借方税額',
  '貸方勘定科目',
  '貸方補助科目',
  '貸方部門',
  '貸方取引先',
  '貸方税区分',
  '貸方インボイス',
  '貸方金額(円)',
  '貸方税額',
  '摘要',
  '仕訳メモ',
  'タグ',
  'MF仕訳タイプ',
  '決算整理仕訳',
  '作成日時',
  '作成者',
  '最終更新日時',
  '最終更新者'
];
const MF_ACCOUNT_CANDIDATES = [
  { name: '旅費', example: '新幹線代、飛行機代' },
  { name: '交通費', example: '電車・バス・タクシー代、Suicaチャージ、駐車料金' },
  { name: '車両費', example: 'ガソリン代、洗車代' },
  { name: '賃借料', example: '会議室代' },
  { name: '会議費', example: '打ち合わせに伴う喫茶代、弁当代、飲食代' },
  { name: '新聞図書費', example: '書籍、新聞、有料メルマガ、業界紙の購読料' },
  { name: '運搬費', example: '宅急便、郵送の送料、梱包資材' },
  { name: '租税公課', example: '印紙代' },
  { name: '消耗品費', example: '10万円以下の消耗品、事務用品、文房具' },
  { name: '未確定勘定', example: '上記のいずれにも当てはまらない費用' }
];
const MF_ACCOUNT_CANDIDATE_NAMES = MF_ACCOUNT_CANDIDATES.map((item) => item.name);
const MF_ACCOUNT_CANDIDATE_GUIDE = MF_ACCOUNT_CANDIDATES
  .map((item) => `${item.name}：${item.example}`)
  .join('\n');
const RESULT_SHEET_HEADERS = [
  'ファイルID',
  'リンク',
  '元ファイル名',
  '変更案（修正可）',
  '移動先',
  'ステータス',
  '支払日',
  '支払い方法',
  'カード情報',
  '取引先',
  'インボイス番号',
  '品目（概要）',
  '金額',
  '解析メモ',
  '消費税区分'
];

// ※APIキーは「スクリプトプロパティ」から、フォルダURLは「設定」シートから読み込みます

// ==================================================
// メニュー作成 (スプレッドシートを開いた時に実行)
// ==================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const currentModel = getModelName_();
  const settingsMenu = ui.createMenu('⚙️ 設定')
    .addItem('APIキーの設定', 'setApiKey')
    .addItem(`モデル切替（現在: ${currentModel}）`, 'setModelName')
    .addItem('設定シートの初期化', 'initializeConfigSheets');

  ui.createMenu('🧾 レシート解析')
    .addItem('1. レシート解析', 'scanToSheet')
    .addItem('変更案を再生成（全行）', 'regenerateAllNameCandidates')
    .addItem('解析シートをクリア', 'clearAnalysisSheet')
    .addItem('2. ファイル名反映', 'applyRenames')
    .addItem('3. フォルダ移動', 'moveFilesToSpecifiedFolder')
    .addItem('4. 取引先一覧更新', 'updateTradePartnersFromAnalysis')
    .addItem('5. マネフォ用解析', 'analyzeMoneyForward')
    .addItem('6. CSVダウンロード', 'openCsvDownloadSelector')
    .addSeparator()
    .addItem('画像プレビューを開く', 'showImageSidebar')
    .addItem('ヘルプを見る', 'showHelp')
    .addSeparator()
    .addSubMenu(settingsMenu)
    .addToUi();
}

// ==================================================
// 解析結果シートの列構成を整える
// A:ファイルID, B:リンク, C:元ファイル名, D:変更案, E:移動先, F:ステータス,
// G:支払日, H:支払い方法, I:カード情報, J:取引先, K:インボイス番号, L:品目（概要）, M:金額, N:解析メモ, O:消費税区分
// ==================================================
function ensureResultSheetLayout_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(RESULT_SHEET_HEADERS);
    sheet.setRowHeight(1, 30);
    sheet.setColumnWidth(2, 60);
    sheet.setColumnWidth(4, 300);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 110);
    sheet.setColumnWidth(8, 110);
    sheet.setColumnWidth(9, 220);
    sheet.setColumnWidth(10, 190);
    sheet.setColumnWidth(11, 140);
    sheet.setColumnWidth(12, 240);
    sheet.setColumnWidth(13, 100);
    sheet.setColumnWidth(14, 260);
    sheet.setColumnWidth(15, 120);
    sheet.setFrozenRows(1);
    return;
  }

  const headerRow = sheet
    .getRange(1, 1, 1, Math.max(RESULT_SHEET_HEADERS.length, sheet.getLastColumn()))
    .getValues()[0];
  const destinationCol = headerRow.indexOf('移動先') + 1;

  // 旧フォーマット（E列=ステータス）からの移行: D列の後ろに「移動先」を挿入
  if (destinationCol === 0) {
    sheet.insertColumnAfter(4);
    sheet.getRange(1, 5).setValue('移動先');
  }

  // 「ステータス」が無い場合は末尾に追加（通常は移動先追加でFに移動済み）
  const headerRowAfter = sheet
    .getRange(1, 1, 1, Math.max(RESULT_SHEET_HEADERS.length, sheet.getLastColumn()))
    .getValues()[0];
  const statusColAfter = headerRowAfter.indexOf('ステータス') + 1;
  if (statusColAfter === 0) {
    sheet.insertColumnAfter(5);
    sheet.getRange(1, 6).setValue('ステータス');
  }

  const headerRowFinal = sheet
    .getRange(1, 1, 1, Math.max(RESULT_SHEET_HEADERS.length, sheet.getLastColumn()))
    .getValues()[0];
  const statusColFinal = headerRowFinal.indexOf('ステータス') + 1;
  if (statusColFinal > 0) {
    const extractHeaders = [
      '支払日',
      '支払い方法',
      'カード情報',
      '取引先',
      'インボイス番号',
      '品目（概要）',
      '金額',
      '解析メモ',
      '消費税区分'
    ];
    let insertAfter = statusColFinal;
    for (const header of extractHeaders) {
      const currentHeaderRow = sheet
        .getRange(1, 1, 1, Math.max(RESULT_SHEET_HEADERS.length, sheet.getLastColumn()))
        .getValues()[0];
      if (currentHeaderRow.indexOf(header) === -1) {
        sheet.insertColumnAfter(insertAfter);
        sheet.getRange(1, insertAfter + 1).setValue(header);
        insertAfter += 1;
      } else {
        insertAfter = currentHeaderRow.indexOf(header) + 1;
      }
    }
  }

  sheet.setRowHeight(1, 30);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(9, 220);
  sheet.setColumnWidth(10, 190);
  sheet.setColumnWidth(11, 140);
  sheet.setColumnWidth(12, 240);
  sheet.setColumnWidth(13, 100);
  sheet.setColumnWidth(14, 260);
  sheet.setColumnWidth(15, 120);
  sheet.setFrozenRows(1);
}

// ==================================================
// APIキーをスクリプトプロパティに保存する関数
// ==================================================
function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    'Gemini APIキー設定',
    'Gemini APIキーを入力してください：\n（以前のキーは上書きされます）',
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() == ui.Button.OK) {
    const key = result.getResponseText().trim();
    if (key) {
      // スクリプトプロパティに保存
      PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', key);
      ui.alert('APIキーを安全に保存しました。\nこれ以降、シート上にキーを記載する必要はありません。');
    } else {
      ui.alert('キーが空のため保存しませんでした。');
    }
  }
}

// ==================================================
// Geminiモデルをスクリプトプロパティに保存する関数
// ==================================================
function setModelName() {
  const ui = SpreadsheetApp.getUi();
  const currentModel = getModelName_();
  const result = ui.prompt(
    'Geminiモデル切替',
    `使用するモデルを入力してください（現在: ${currentModel}）\n` +
      `候補例: ${MODEL_CANDIDATES.join(', ')}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const model = result.getResponseText().trim();
  if (!model) {
    ui.alert('モデル名が空のため保存しませんでした。');
    return;
  }

  PropertiesService.getScriptProperties().setProperty('GEMINI_MODEL_NAME', model);
  ui.alert(`モデルを「${model}」に切り替えました。`);
}

function getModelName_() {
  const stored = PropertiesService.getScriptProperties().getProperty('GEMINI_MODEL_NAME');
  if (stored && String(stored).trim()) return String(stored).trim();
  return DEFAULT_MODEL_NAME;
}

// ==================================================
// 設定シート初期化
// ==================================================
function initializeConfigSheets() {
  const ui = SpreadsheetApp.getUi();
  const createdSheets = [];

  const settingsInfo = resetSettingsSheetLayout_();
  if (settingsInfo.created) createdSheets.push(SETTINGS_SHEET_NAME);

  const fileNameRuleSheetName =
    settingsInfo.settingsMap[SETTINGS_KEYS.fileNameRuleSheet] || FILE_NAME_RULE_SHEET_NAME;
  if (ensureFileNameRuleSheet_(fileNameRuleSheetName).created) createdSheets.push(fileNameRuleSheetName);
  if (ensureDestinationSheet_().created) createdSheets.push(DESTINATION_SHEET_NAME);
  if (ensureAccountRuleSheet_().created) createdSheets.push(ACCOUNT_RULE_SHEET_NAME);

  const message = createdSheets.length > 0
    ? `設定シートを初期化しました。\n作成: ${createdSheets.join(' / ')}`
    : '設定シートを初期化しました。';
  ui.alert(message);
}

function resetSettingsSheetLayout_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  let created = false;

  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    created = true;
  }

  const existingMap = getSettingsMap_(sheet);
  sheet.clear();

  sheet.getRange('A1:B1').setValues([['項目名', '設定値']]);
  sheet.getRange('A1:B1').setBackground('#efefef').setFontWeight('bold');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 420);

  const definitions = [
    { key: SETTINGS_KEYS.folderUrl, defaultValue: '' },
    { key: SETTINGS_KEYS.delimiter, defaultValue: DELIMITER_CANDIDATES[0] },
    { key: SETTINGS_KEYS.fileNameRuleSheet, defaultValue: FILE_NAME_RULE_SHEET_NAME },
    { key: SETTINGS_KEYS.scanLimit, defaultValue: 50 },
    { key: SETTINGS_KEYS.logOutput, defaultValue: 'ON' },
    { key: SETTINGS_KEYS.creditAccountCard, defaultValue: '未払金' },
    { key: SETTINGS_KEYS.creditAccountOther, defaultValue: '役員借入金' },
    { key: SETTINGS_KEYS.creditSubAccountCard, defaultValue: CREDIT_SUB_ACCOUNT_CANDIDATES[0] },
    { key: SETTINGS_KEYS.mfSummaryMode, defaultValue: MF_SUMMARY_MODE_CANDIDATES[0] }
  ];

  const rows = definitions.map((def) => [
    def.key,
    existingMap.hasOwnProperty(def.key) && existingMap[def.key] !== ''
      ? existingMap[def.key]
      : def.defaultValue
  ]);
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);

  const keyRows = getSettingsKeyRowMap_(sheet);
  const delimiterRow = keyRows[SETTINGS_KEYS.delimiter];
  if (delimiterRow) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(DELIMITER_CANDIDATES, true)
      .build();
    sheet.getRange(delimiterRow, 2).setDataValidation(rule);
  }

  const logRow = keyRows[SETTINGS_KEYS.logOutput];
  if (logRow) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(LOG_OUTPUT_CANDIDATES, true)
      .build();
    sheet.getRange(logRow, 2).setDataValidation(rule);
  }

  const creditSubAccountRow = keyRows[SETTINGS_KEYS.creditSubAccountCard];
  if (creditSubAccountRow) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CREDIT_SUB_ACCOUNT_CANDIDATES, true)
      .build();
    sheet.getRange(creditSubAccountRow, 2).setDataValidation(rule);
  }

  const mfSummaryModeRow = keyRows[SETTINGS_KEYS.mfSummaryMode];
  if (mfSummaryModeRow) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(MF_SUMMARY_MODE_CANDIDATES, true)
      .build();
    sheet.getRange(mfSummaryModeRow, 2).setDataValidation(rule);
  }

  const folderRow = keyRows[SETTINGS_KEYS.folderUrl];
  if (folderRow) {
    sheet
      .getRange(folderRow, 2)
      .setNote('例: https://drive.google.com/drive/folders/xxxxxxxx\n※IDのみの入力は不可');
  }

  return { sheet: sheet, created: created, settingsMap: getSettingsMap_(sheet) };
}

// ==================================================
// 設定値を取得する関数
// ==================================================
function getSettingsOrAlert_(options) {
  const ui = SpreadsheetApp.getUi();
  const settingsInfo = ensureSettingsSheet_();
  const settingsMap = settingsInfo.settingsMap;

  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const folderUrlRaw = normalizeText_(settingsMap[SETTINGS_KEYS.folderUrl]);
  const delimiterRaw = normalizeText_(settingsMap[SETTINGS_KEYS.delimiter]) || DELIMITER_CANDIDATES[0];
  const fileNameRuleSheetName =
    normalizeText_(settingsMap[SETTINGS_KEYS.fileNameRuleSheet]) || FILE_NAME_RULE_SHEET_NAME;
  const scanLimitRaw = settingsMap[SETTINGS_KEYS.scanLimit];
  const logOutputRaw = normalizeText_(settingsMap[SETTINGS_KEYS.logOutput]) || 'ON';
  const creditAccountCard = normalizeText_(settingsMap[SETTINGS_KEYS.creditAccountCard]) || '未払金';
  const creditAccountOther = normalizeText_(settingsMap[SETTINGS_KEYS.creditAccountOther]) || '役員借入金';
  const creditSubAccountCard = normalizeText_(settingsMap[SETTINGS_KEYS.creditSubAccountCard]) ||
    CREDIT_SUB_ACCOUNT_CANDIDATES[0];
  const mfSummaryModeRaw = normalizeText_(settingsMap[SETTINGS_KEYS.mfSummaryMode]);
  const mfSummaryMode = MF_SUMMARY_MODE_CANDIDATES.includes(mfSummaryModeRaw)
    ? mfSummaryModeRaw
    : MF_SUMMARY_MODE_CANDIDATES[0];

  const errors = [];
  let folderId = '';
  if (options && options.requireFolder) {
    if (!isFolderUrl_(folderUrlRaw)) {
      errors.push(`「設定」シートの${SETTINGS_KEYS.folderUrl}にフォルダURLを入力してください（IDのみは不可）。`);
    } else {
      folderId = extractFolderIdFromUrl_(folderUrlRaw);
      if (!folderId) {
        errors.push(`「設定」シートの${SETTINGS_KEYS.folderUrl}が無効です。フォルダURLを確認してください。`);
      }
    }
  } else if (isFolderUrl_(folderUrlRaw)) {
    folderId = extractFolderIdFromUrl_(folderUrlRaw);
  }

  if (options && options.requireApiKey && !apiKey) {
    errors.push('Gemini APIキーが設定されていません。\nメニューの「APIキーの設定」からキーを登録してください。');
  }

  if (!DELIMITER_CANDIDATES.includes(delimiterRaw)) {
    errors.push(`区切り文字は ${DELIMITER_CANDIDATES.join(' / ')} から選択してください。`);
  }

  let scanLimit = parseInt(scanLimitRaw, 10);
  if (!scanLimit || scanLimit < 1) scanLimit = 50;

  const logOutput = String(logOutputRaw).toUpperCase() !== 'OFF';

  if (errors.length > 0) {
    ui.alert(errors.join('\n'));
    return null;
  }

  ensureFileNameRuleSheet_(fileNameRuleSheetName);

  return {
    apiKey: apiKey || '',
    folderUrl: folderUrlRaw,
    folderId: folderId,
    delimiter: delimiterRaw,
    fileNameRuleSheetName: fileNameRuleSheetName,
    scanLimit: scanLimit,
    logOutput: logOutput,
    creditAccountCard: creditAccountCard,
    creditAccountOther: creditAccountOther,
    creditSubAccountCard: creditSubAccountCard,
    mfSummaryMode: mfSummaryMode
  };
}

function ensureSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  let created = false;

  if (!sheet) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    created = true;
  }

  sheet.getRange('A1:B1').setValues([['項目名', '設定値']]);
  sheet.getRange('A1:B1').setBackground('#efefef').setFontWeight('bold');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 420);

  const definitions = [
    { key: SETTINGS_KEYS.folderUrl, defaultValue: '' },
    { key: SETTINGS_KEYS.delimiter, defaultValue: DELIMITER_CANDIDATES[0] },
    { key: SETTINGS_KEYS.fileNameRuleSheet, defaultValue: FILE_NAME_RULE_SHEET_NAME },
    { key: SETTINGS_KEYS.scanLimit, defaultValue: 50 },
    { key: SETTINGS_KEYS.logOutput, defaultValue: 'ON' },
    { key: SETTINGS_KEYS.creditAccountCard, defaultValue: '未払金' },
    { key: SETTINGS_KEYS.creditAccountOther, defaultValue: '役員借入金' },
    { key: SETTINGS_KEYS.creditSubAccountCard, defaultValue: CREDIT_SUB_ACCOUNT_CANDIDATES[0] },
    { key: SETTINGS_KEYS.mfSummaryMode, defaultValue: MF_SUMMARY_MODE_CANDIDATES[0] }
  ];

  const settingsMap = getSettingsMap_(sheet);
  let lastRow = sheet.getLastRow();
  for (const def of definitions) {
    if (settingsMap.hasOwnProperty(def.key)) continue;
    lastRow += 1;
    sheet.getRange(lastRow, 1, 1, 2).setValues([[def.key, def.defaultValue]]);
    settingsMap[def.key] = def.defaultValue;
  }

  const keyRows = getSettingsKeyRowMap_(sheet);
  const delimiterRow = keyRows[SETTINGS_KEYS.delimiter];
  if (delimiterRow) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(DELIMITER_CANDIDATES, true)
      .build();
    sheet.getRange(delimiterRow, 2).setDataValidation(rule);
  }

  const logRow = keyRows[SETTINGS_KEYS.logOutput];
  if (logRow) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(LOG_OUTPUT_CANDIDATES, true)
      .build();
    sheet.getRange(logRow, 2).setDataValidation(rule);
  }

  const creditSubAccountRow = keyRows[SETTINGS_KEYS.creditSubAccountCard];
  if (creditSubAccountRow) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CREDIT_SUB_ACCOUNT_CANDIDATES, true)
      .build();
    sheet.getRange(creditSubAccountRow, 2).setDataValidation(rule);
  }

  const mfSummaryModeRow = keyRows[SETTINGS_KEYS.mfSummaryMode];
  if (mfSummaryModeRow) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(MF_SUMMARY_MODE_CANDIDATES, true)
      .build();
    sheet.getRange(mfSummaryModeRow, 2).setDataValidation(rule);
  }

  const folderRow = keyRows[SETTINGS_KEYS.folderUrl];
  if (folderRow) {
    sheet
      .getRange(folderRow, 2)
      .setNote('例: https://drive.google.com/drive/folders/xxxxxxxx\n※IDのみの入力は不可');
  }

  return { sheet: sheet, created: created, settingsMap: settingsMap };
}

function getSettingsMap_(sheet) {
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (const [key, value] of values) {
    if (!key) continue;
    map[String(key).trim()] = value;
  }
  return map;
}

function getSettingsKeyRowMap_(sheet) {
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const key = values[i][0];
    if (!key) continue;
    map[String(key).trim()] = i + 2;
  }
  return map;
}

function getOrCreateAnalysisSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ANALYSIS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ANALYSIS_SHEET_NAME);
  }
  return sheet;
}

// ==================================================
// 各設定シートの初期化
// ==================================================
function ensureFileNameRuleSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = sheetName || FILE_NAME_RULE_SHEET_NAME;
  let sheet = ss.getSheetByName(name);
  let created = false;

  if (!sheet) {
    sheet = ss.insertSheet(name);
    created = true;
  }

  sheet.getRange('A1:C1').setValues([['項目名', '順番', '使用可否']]);
  sheet.getRange('A1:C1').setBackground('#ddebf7').setFontWeight('bold');
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    const items = [
      '支払日',
      '支払い方法',
      '取引先',
      'インボイス番号',
      '品目（概要）',
      '金額'
    ];
    const rows = items.map((item, index) => [item, index + 1, true]);
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  const orderRule = SpreadsheetApp.newDataValidation().requireNumberBetween(1, 20).build();
  sheet.getRange('B2:B').setDataValidation(orderRule);
  sheet.getRange('C2:C').insertCheckboxes();

  return { sheet: sheet, created: created };
}

function ensureDestinationSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DESTINATION_SHEET_NAME);
  let created = false;

  if (!sheet) {
    sheet = ss.insertSheet(DESTINATION_SHEET_NAME);
    created = true;
  }

  sheet.getRange('A1:B1').setValues([['キーワード', 'フォルダURL']]);
  sheet.getRange('A1:B1').setBackground('#fff2cc').setFontWeight('bold');
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 420);

  return { sheet: sheet, created: created };
}

function ensureAccountRuleSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ACCOUNT_RULE_SHEET_NAME);
  let created = false;

  if (!sheet) {
    sheet = ss.insertSheet(ACCOUNT_RULE_SHEET_NAME);
    created = true;
  }

  sheet.getRange('A1:B1').setValues([['勘定科目', 'ルール']]);
  sheet.getRange('A1:B1').setBackground('#d9ead3').setFontWeight('bold');
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 360);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    const rows = MF_ACCOUNT_CANDIDATES.map((item) => [item.name, item.example]);
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  return { sheet: sheet, created: created };
}

function ensureLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  let created = false;

  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    created = true;
  }

  sheet.getRange('A1:E1').setValues([['日時', '機能', 'ファイルID', 'ファイル名', '内容']]);
  sheet.getRange('A1:E1').setBackground('#efefef').setFontWeight('bold');
  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 240);
  sheet.setColumnWidth(4, 260);
  sheet.setColumnWidth(5, 420);

  return { sheet: sheet, created: created };
}

// ==================================================
// ファイル名ルール / 解析補助
// ==================================================
function getFileNameRules_(sheetName) {
  const sheet = ensureFileNameRuleSheet_(sheetName).sheet;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  return values
    .map((row) => ({
      label: normalizeText_(row[0]),
      order: parseInt(row[1], 10),
      enabled: row[2] === true || String(row[2]).toUpperCase() === 'TRUE'
    }))
    .filter((rule) => rule.label && rule.enabled && !isNaN(rule.order))
    .sort((a, b) => a.order - b.order);
}

function buildFileNameCandidate_(analysis, rules, delimiter) {
  if (!analysis || !rules || rules.length === 0) return '';

  let amountText = '';
  if (analysis.amount !== '' && analysis.amount !== null && analysis.amount !== undefined) {
    amountText = String(analysis.amount).replace(/円$/, '');
    amountText = `${amountText}円`;
  }

  const fieldMap = {
    '支払日': normalizeDate_(analysis.paymentDate),
    '支払い方法': normalizeText_(analysis.paymentMethod),
    '取引先': normalizeText_(analysis.vendorName),
    'インボイス番号': normalizeText_(analysis.invoiceNumber),
    '品目（概要）': normalizeText_(analysis.summary),
    '金額': amountText
  };

  const parts = rules.map((rule) => sanitizeFileNamePart_(fieldMap[rule.label] ?? ''));
  return parts.join(delimiter);
}

function sanitizeFileNamePart_(value) {
  if (!value) return '';
  return String(value).replace(/[\\/:*?\"<>|]/g, '-');
}

function extractFileExtension_(fileName) {
  if (!fileName) return '';
  const match = String(fileName).trim().match(/(\.[^.\s]+)$/);
  return match ? match[1] : '';
}

function ensureFileExtension_(baseName, extension) {
  if (!baseName) return '';
  if (!extension) return baseName;
  if (String(baseName).toLowerCase().endsWith(String(extension).toLowerCase())) {
    return baseName;
  }
  return `${baseName}${extension}`;
}

function resolveDuplicateCandidate_(candidateBase, originalName, extension, nameCounts, existingFullNames) {
  if (!candidateBase) return { base: '', full: '', duplicate: false };

  const originalFull = originalName || '';
  const candidateFull = ensureFileExtension_(candidateBase, extension);
  const count = nameCounts[candidateFull] || 0;
  const existsInFolder = count > 0 && !(candidateFull === originalFull && count === 1);
  const existsInSheet = existingFullNames.has(candidateFull);
  const isDuplicate = existsInFolder || existsInSheet;

  if (!isDuplicate) {
    return { base: candidateBase, full: candidateFull, duplicate: false };
  }

  let index = 1;
  let resolvedBase = `${candidateBase}(${index})`;
  let resolvedFull = ensureFileExtension_(resolvedBase, extension);
  while (
    (nameCounts[resolvedFull] && !(resolvedFull === originalFull && nameCounts[resolvedFull] === 1)) ||
    existingFullNames.has(resolvedFull)
  ) {
    index += 1;
    resolvedBase = `${candidateBase}(${index})`;
    resolvedFull = ensureFileExtension_(resolvedBase, extension);
  }

  return { base: resolvedBase, full: resolvedFull, duplicate: true };
}

function getExistingFileIds_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(Boolean);
}

function getExistingFullNameCandidates_(sheet) {
  const lastRow = sheet.getLastRow();
  const names = new Set();
  if (lastRow < 2) return names;
  const values = sheet.getRange(2, 3, lastRow - 1, 2).getValues();
  for (const [originalName, candidateName] of values) {
    const base = normalizeText_(candidateName);
    if (!base) continue;
    const ext = extractFileExtension_(originalName);
    const fullName = ensureFileExtension_(base, ext);
    if (fullName) names.add(fullName);
  }
  return names;
}

function listTargetFiles_(folder, existingIdSet) {
  const targets = [];
  const nameCounts = {};
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const id = file.getId();
    const name = file.getName();
    const mimeType = file.getMimeType();

    nameCounts[name] = (nameCounts[name] || 0) + 1;

    if (existingIdSet && existingIdSet.has(id)) continue;
    if (!isAllowedReceiptMimeType_(mimeType)) continue;
    if (name.match(/^202\\d{5}_/)) continue;
    targets.push({ id: id, name: name, mimeType: mimeType });
  }

  return { targets: targets, nameCounts: nameCounts };
}

function applyDestinationValidation_(sheet, rowIndex) {
  const listSheet = ensureDestinationSheet_().sheet;
  const range = listSheet.getRange('A2:A');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(range, true)
    .build();
  sheet.getRange(rowIndex, 5).setDataValidation(rule);
}

function highlightEmptyExtractionCells_(sheet, rowIndex) {
  const range = sheet.getRange(rowIndex, 7, 1, 7);
  const values = range.getValues()[0];
  const backgrounds = values.map((value) => (isBlankCell_(value) ? EMPTY_CELL_COLOR : null));
  range.setBackgrounds([backgrounds]);
}

function highlightTaxCategoryCell_(sheet, rowIndex, taxCategory) {
  const normalized = normalizeTaxCategory_(taxCategory);
  const color = normalized === '混在あり' ? MIXED_TAX_CELL_COLOR : null;
  sheet.getRange(rowIndex, 15).setBackground(color);
}

function highlightPossibleDuplicateRowsByDateAndAmount_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const lastColumn = Math.max(RESULT_SHEET_HEADERS.length, sheet.getLastColumn());
  const header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const paymentDateIndex = header.indexOf('支払日');
  const amountIndex = header.indexOf('金額');
  if (paymentDateIndex === -1 || amountIndex === -1) return 0;

  const range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  const values = range.getValues();
  const backgrounds = range.getBackgrounds();
  const duplicateMap = {};
  const duplicateRows = new Set();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const paymentDate = normalizeDate_(row[paymentDateIndex]);
    const amount = normalizeAmount_(row[amountIndex]);
    if (!paymentDate || amount <= 0) continue;

    const key = `${paymentDate}__${amount}`;
    if (!duplicateMap[key]) duplicateMap[key] = [];
    duplicateMap[key].push(i);
  }

  Object.keys(duplicateMap).forEach((key) => {
    if (duplicateMap[key].length > 1) {
      duplicateMap[key].forEach((rowIndex) => duplicateRows.add(rowIndex));
    }
  });

  for (let i = 0; i < backgrounds.length; i++) {
    for (let j = 0; j < backgrounds[i].length; j++) {
      if (backgrounds[i][j] === POSSIBLE_DUPLICATE_ROW_COLOR) {
        backgrounds[i][j] = null;
      }
    }
  }

  duplicateRows.forEach((rowIndex) => {
    const row = values[rowIndex];
    let lastDataColumn = 0;
    for (let j = row.length - 1; j >= 0; j--) {
      if (!isBlankCell_(row[j])) {
        lastDataColumn = j + 1;
        break;
      }
    }
    if (lastDataColumn === 0) return;

    for (let j = 0; j < lastDataColumn; j++) {
      backgrounds[rowIndex][j] = POSSIBLE_DUPLICATE_ROW_COLOR;
    }
  });

  range.setBackgrounds(backgrounds);
  return duplicateRows.size;
}

function collectMissingExtractionLabels_(analysis) {
  const missing = [];
  if (!analysis.paymentDate) missing.push('支払日');
  if (!analysis.paymentMethod) missing.push('支払い方法');
  if (!analysis.cardInfo || analysis.cardInfo === 'カード(不明)') missing.push('カード情報');
  if (!analysis.vendorName) missing.push('取引先');
  if (!analysis.invoiceNumber) missing.push('インボイス番号');
  if (!analysis.summary) missing.push('品目（概要）');
  if (!analysis.amount) missing.push('金額');
  return missing;
}

function appendMemo_(base, addition) {
  if (!addition) return base || '';
  if (!base) return addition;
  if (base.includes(addition)) return base;
  return `${base} / ${addition}`;
}

function validateRenameTargets_(sheet, data) {
  const duplicateMap = {};
  const missingRows = [];
  const duplicateRows = new Set();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const fileId = row[0];
    const newNameBase = normalizeText_(row[3]);
    const status = row[5];
    const originalName = row[2];
    const extension = extractFileExtension_(originalName);
    const newName = ensureFileExtension_(newNameBase, extension);

    if (status !== '未処理') continue;
    if (!fileId || !newNameBase) {
      missingRows.push(i + 2);
      continue;
    }

    if (!duplicateMap[newName]) duplicateMap[newName] = [];
    duplicateMap[newName].push(i + 2);
  }

  Object.keys(duplicateMap).forEach((name) => {
    if (duplicateMap[name].length > 1) {
      duplicateMap[name].forEach((rowIndex) => duplicateRows.add(rowIndex));
    }
  });

  const driveDuplicateRows = new Set();
  if (duplicateRows.size === 0 && missingRows.length === 0) {
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const fileId = row[0];
      const newNameBase = normalizeText_(row[3]);
      const status = row[5];
      const originalName = row[2];
      const extension = extractFileExtension_(originalName);
      const newName = ensureFileExtension_(newNameBase, extension);

      if (status !== '未処理' || !fileId || !newNameBase) continue;
      if (hasDuplicateNameInParents_(fileId, newName)) {
        driveDuplicateRows.add(i + 2);
      }
    }
  }

  if (missingRows.length > 0) {
    missingRows.forEach((rowIndex) => sheet.getRange(rowIndex, 4).setBackground(DUPLICATE_CELL_COLOR));
    return {
      ok: false,
      message: `変更案が未入力の行があります。\n行: ${missingRows.slice(0, 10).join(', ')}`
    };
  }

  if (duplicateRows.size > 0) {
    duplicateRows.forEach((rowIndex) => sheet.getRange(rowIndex, 4).setBackground(DUPLICATE_CELL_COLOR));
    return {
      ok: false,
      message: `変更案が重複しています。\n行: ${Array.from(duplicateRows).slice(0, 10).join(', ')}`
    };
  }

  if (driveDuplicateRows.size > 0) {
    driveDuplicateRows.forEach((rowIndex) => sheet.getRange(rowIndex, 4).setBackground(DUPLICATE_CELL_COLOR));
    return {
      ok: false,
      message: `対象フォルダ内に同名ファイルがあります。\n行: ${Array.from(driveDuplicateRows).slice(0, 10).join(', ')}`
    };
  }

  return { ok: true };
}

function hasDuplicateNameInParents_(fileId, newName) {
  try {
    const file = DriveApp.getFileById(fileId);
    const parents = file.getParents();
    while (parents.hasNext()) {
      const folder = parents.next();
      const filesByName = folder.getFilesByName(newName);
      while (filesByName.hasNext()) {
        const candidate = filesByName.next();
        if (candidate.getId() !== fileId) return true;
      }
    }
  } catch (e) {
    console.error(e);
  }
  return false;
}

function logError_(action, fileId, fileName, message) {
  if (!message) return;
  if (!isLogOutputEnabled_()) return;
  const sheet = ensureLogSheet_().sheet;
  sheet.appendRow([
    new Date(),
    action,
    fileId || '',
    fileName || '',
    message
  ]);
}

function isLogOutputEnabled_() {
  const settingsInfo = ensureSettingsSheet_();
  const value = normalizeText_(settingsInfo.settingsMap[SETTINGS_KEYS.logOutput]) || 'ON';
  return String(value).toUpperCase() !== 'OFF';
}

function isFolderUrl_(value) {
  if (!value) return false;
  return /^https?:\/\//i.test(String(value).trim());
}

function extractFolderIdFromUrl_(url) {
  if (!url) return '';
  const text = String(url).trim();
  if (!/^https?:\/\//i.test(text)) return '';
  const match =
    text.match(/\/folders\/([a-zA-Z0-9_-]+)/) ||
    text.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : '';
}

// ==================================================
// 機能1: ドライブをスキャンしてシートに書き出す
// ==================================================
function scanToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateAnalysisSheet_();
  const ui = SpreadsheetApp.getUi();

  // 設定を取得
  const settings = getSettingsOrAlert_({ requireApiKey: true, requireFolder: true });
  if (!settings) return;

  ss.setActiveSheet(sheet);

  const fileNameRules = getFileNameRules_(settings.fileNameRuleSheetName);

  ensureResultSheetLayout_(sheet);
  ensureDestinationSheet_();

  // 既存のファイルIDを取得
  const existingIds = getExistingFileIds_(sheet);
  const existingIdSet = new Set(existingIds);
  const existingNameCandidates = getExistingFullNameCandidates_(sheet);

  let folder;
  try {
    folder = DriveApp.getFolderById(settings.folderId);
  } catch (e) {
    ui.alert('フォルダが見つかりません。URLが正しいか確認してください。\n' + e.message);
    return;
  }

  const list = listTargetFiles_(folder, existingIdSet);
  const targetFiles = list.targets;
  const nameCounts = list.nameCounts;

  const maxCount = settings.scanLimit || 50;
  let processCount = 0;
  let errorCount = 0;
  let duplicateCount = 0;

  for (const target of targetFiles) {
    if (processCount >= maxCount) break;

    const id = target.id;
    const fileName = target.name;
    const mimeType = target.mimeType;

    try {
      const file = DriveApp.getFileById(id);
      const blob = file.getBlob();
      const base64Data = Utilities.base64Encode(blob.getBytes());

      // HYPERLINK関数でリンクを作成
      const thumbnailFormula = `=HYPERLINK("https://drive.google.com/file/d/${id}/view", "開く")`;

      // Gemini API呼び出し
      const analysis = callGeminiApi(base64Data, mimeType, settings.apiKey);

      let newNameCandidate = '';
      let status = '解析失敗';
      let paymentDate = '';
      let paymentMethod = '';
      let cardInfo = '';
      let vendorName = '';
      let invoiceNumber = '';
      let summary = '';
      let amount = '';
      let taxCategory = '';
      let memo = '';
      let duplicateFlag = false;

      if (analysis) {
        paymentDate = analysis.paymentDate || '';
        paymentMethod = analysis.paymentMethod || '';
        cardInfo = normalizeCardInfo_(analysis.cardInfo, paymentMethod);
        vendorName = analysis.vendorName || '';
        invoiceNumber = analysis.invoiceNumber || '';
        summary = normalizeText_(analysis.summary || '');
        amount = analysis.amount ? analysis.amount : '';
        taxCategory = normalizeTaxCategory_(analysis.taxCategory);
        status = '未処理';

        const candidate = buildFileNameCandidate_(
          {
            paymentDate: paymentDate,
            paymentMethod: paymentMethod,
            cardInfo: cardInfo,
            vendorName: vendorName,
            invoiceNumber: invoiceNumber,
            summary: summary,
            amount: amount
          },
          fileNameRules,
          settings.delimiter
        );

        const extension = extractFileExtension_(fileName);
        const resolved = resolveDuplicateCandidate_(
          candidate,
          fileName,
          extension,
          nameCounts,
          existingNameCandidates
        );
        newNameCandidate = resolved.full;
        duplicateFlag = resolved.duplicate;
        if (duplicateFlag) {
          duplicateCount++;
          memo = appendMemo_(memo, '重複あり');
        }
      } else {
        memo = appendMemo_(memo, '解析失敗');
        logError_('レシート解析', id, fileName, 'Gemini解析に失敗しました');
      }

      sheet.appendRow([
        id,
        thumbnailFormula,
        fileName,
        newNameCandidate,
        '',
        status,
        paymentDate,
        paymentMethod,
        cardInfo,
        vendorName,
        invoiceNumber,
        summary,
        amount,
        memo,
        taxCategory
      ]);

      const rowIndex = sheet.getLastRow();
      sheet.setRowHeight(rowIndex, 30);

      applyDestinationValidation_(sheet, rowIndex);
      highlightEmptyExtractionCells_(sheet, rowIndex);
      highlightTaxCategoryCell_(sheet, rowIndex, taxCategory);
      if (duplicateFlag) {
        sheet.getRange(rowIndex, 4).setBackground(DUPLICATE_CELL_COLOR);
      }

      if (analysis) {
        const missingLabels = collectMissingExtractionLabels_({
          paymentDate: paymentDate,
          paymentMethod: paymentMethod,
          cardInfo: cardInfo,
          vendorName: vendorName,
          invoiceNumber: invoiceNumber,
          summary: summary,
          amount: amount
        });
        if (missingLabels.length > 0) {
          const memoText = `未取得: ${missingLabels.join(' / ')}`;
          sheet.getRange(rowIndex, 14).setValue(appendMemo_(memo, memoText));
        }
      }

      if (newNameCandidate) existingNameCandidates.add(newNameCandidate);
      processCount++;
    } catch (e) {
      console.error(e);
      errorCount++;
      sheet.appendRow([id, '', fileName, '', '', 'エラー: ' + e.toString(), '', '', '', '', '', '', '', '', '']);
      logError_('レシート解析', id, fileName, e.message);
    }
  }

  const limitReached = targetFiles.length > processCount;
  const possibleDuplicateRowCount = highlightPossibleDuplicateRowsByDateAndAmount_(sheet);
  const messages = [];
  if (processCount === 0) {
    messages.push('新しいファイルは見つかりませんでした。');
  } else {
    messages.push(`${processCount} 件のファイルをスキャンしました。`);
  }
  if (duplicateCount > 0) messages.push(`重複候補: ${duplicateCount} 件（提案名に連番を付与済み）`);
  if (possibleDuplicateRowCount > 0) {
    messages.push(`日付+金額が一致する重複候補: ${possibleDuplicateRowCount} 行（行全体を色付け）`);
  }
  if (errorCount > 0) messages.push(`エラー: ${errorCount} 件`);
  if (limitReached) messages.push(`解析上限 ${maxCount} 件に達しました。再実行で続きが追加されます。`);

  ui.alert(messages.join('\n'));
}

// ==================================================
// 機能2: シートの内容をファイル名に反映する
// ==================================================
function applyRenames() {
  const sheet = getOrCreateAnalysisSheet_();
  const ui = SpreadsheetApp.getUi();

  ensureResultSheetLayout_(sheet);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('データがありません。');
    return;
  }

  const range = sheet.getRange(2, 1, lastRow - 1, 6);
  const data = range.getValues();

  const validation = validateRenameTargets_(sheet, data);
  if (!validation.ok) {
    ui.alert(validation.message);
    return;
  }

  let successCount = 0;
  let errorCount = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const fileId = row[0];
    const newName = row[3];
    const status = row[5];

    if (status === '未処理' && newName !== '' && fileId !== '') {
      try {
        const file = DriveApp.getFileById(fileId);
        const oldName = file.getName();
        const finalName = ensureFileExtension_(normalizeText_(newName), extractFileExtension_(oldName));
        
        if (oldName !== finalName) {
          file.setName(finalName);
          sheet.getRange(i + 2, 6).setValue('完了');
          sheet.getRange(i + 2, 3).setValue(finalName);
          successCount++;
        } else {
          sheet.getRange(i + 2, 6).setValue('変更なし');
        }
      } catch (e) {
        sheet.getRange(i + 2, 6).setValue('エラー: ' + e.message);
        logError_('ファイル名反映', fileId, '', e.message);
        errorCount++;
      }
    }
  }

  const messages = [`${successCount} 件のファイル名を変更しました。`];
  if (errorCount > 0) messages.push(`エラー: ${errorCount} 件`);
  ui.alert(messages.join('\n'));
}

// ==================================================
// 解析シートのクリア
// ==================================================
function clearAnalysisSheet() {
  const sheet = getOrCreateAnalysisSheet_();
  const ui = SpreadsheetApp.getUi();

  ensureResultSheetLayout_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('解析シートにデータがありません。');
    return;
  }

  const options = getClearAnalysisStatusOptions_(sheet);
  if (options.length === 0) {
    ui.alert('クリア対象のステータスがありません。');
    return;
  }

  const html = HtmlService.createHtmlOutput(buildClearAnalysisStatusDialogHtml_(options))
    .setWidth(520)
    .setHeight(620);
  ui.showModalDialog(html, '解析シートのクリア');
}

function getClearAnalysisStatusOptions_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const header = sheet
    .getRange(1, 1, 1, Math.max(RESULT_SHEET_HEADERS.length, sheet.getLastColumn()))
    .getValues()[0];
  const statusCol = header.indexOf('ステータス') + 1;
  if (statusCol === 0) return [];

  const values = sheet.getRange(2, statusCol, lastRow - 1, 1).getValues();
  const unique = new Set();
  let hasError = false;
  for (const [value] of values) {
    const status = normalizeText_(value);
    if (!status) continue;
    if (status.startsWith(ANALYSIS_STATUS_ERROR_PREFIX)) {
      hasError = true;
      continue;
    }
    unique.add(status);
  }

  const options = [];
  if (hasError) {
    options.push({
      value: ANALYSIS_CLEAR_ERROR_OPTION_VALUE,
      label: ANALYSIS_CLEAR_ERROR_OPTION_LABEL
    });
  }

  const sorted = Array.from(unique);
  sorted.sort();
  for (const status of sorted) {
    options.push({ value: status, label: status });
  }
  return options;
}

function buildClearAnalysisStatusDialogHtml_(options) {
  const optionsJson = JSON.stringify(options || []);
  return `
<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <style>
      html,
      body {
        font-family: "Noto Sans JP", Arial, sans-serif;
        margin: 0;
        padding: 0;
      }
      .wrap {
        padding: 16px;
      }
      .title {
        font-size: 14px;
        font-weight: 700;
        margin: 0 0 8px;
      }
      .desc {
        color: #555;
        font-size: 12px;
        line-height: 1.5;
        margin: 0 0 12px;
      }
      .list {
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        padding: 10px;
        max-height: 420px;
        overflow: auto;
      }
      .item {
        align-items: center;
        display: flex;
        gap: 8px;
        padding: 6px 0;
      }
      .item label {
        cursor: pointer;
        font-size: 13px;
      }
      .actions {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        margin-top: 12px;
      }
      button {
        background: #f3f3f3;
        border: 1px solid #d0d0d0;
        border-radius: 8px;
        cursor: pointer;
        font-size: 12px;
        padding: 8px 12px;
      }
      button.primary {
        background: #1a73e8;
        border-color: #1a73e8;
        color: #fff;
      }
      button:disabled {
        cursor: not-allowed;
        opacity: 0.6;
      }
      .msg {
        color: #666;
        font-size: 12px;
        margin-top: 10px;
        min-height: 16px;
      }
      .log {
        background: #fafafa;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        color: #444;
        font-size: 11px;
        line-height: 1.4;
        margin-top: 8px;
        max-height: 120px;
        overflow: auto;
        padding: 8px;
        white-space: pre-wrap;
      }
      .warn {
        color: #b45309;
      }
      .muted {
        color: #888;
      }
    </style>
  </head>
  <body>
    <div class="wrap">
      <div class="title">削除するステータスを選択</div>
      <p class="desc">
        チェックしたステータスの行を削除し、行を詰めます。<span class="warn">削除は元に戻せません。</span>
      </p>
      <div id="list" class="list"></div>
      <div class="actions">
        <button type="button" id="selectAll">全選択</button>
        <button type="button" id="clearAll">全解除</button>
        <button type="button" id="run" class="primary">実行</button>
        <button type="button" id="cancel">キャンセル</button>
      </div>
      <div id="msg" class="msg muted"></div>
      <div id="log" class="log"></div>
    </div>

    <script>
      const options = ${optionsJson};

      const list = document.getElementById('list');
      const msg = document.getElementById('msg');
      const log = document.getElementById('log');
      const btnSelectAll = document.getElementById('selectAll');
      const btnClearAll = document.getElementById('clearAll');
      const btnRun = document.getElementById('run');
      const btnCancel = document.getElementById('cancel');

      const startedAt = Date.now();
      let heartbeatTimer = null;
      let waitLogTimer = null;
      let activeJobId = '';
      let totalRows = 0;
      let deletedRows = 0;

      function logLine(text) {
        const sec = ((Date.now() - startedAt) / 1000).toFixed(1);
        log.textContent += '[+' + sec + 's] ' + text + '\\n';
        log.scrollTop = log.scrollHeight;
      }

      function startWaitLog(label) {
        const start = Date.now();
        stopWaitLog();
        waitLogTimer = setInterval(() => {
          const sec = Math.floor((Date.now() - start) / 1000);
          logLine(label + ' 応答待ち... ' + sec + '秒');
        }, 5000);
      }

      function stopWaitLog() {
        if (waitLogTimer) clearInterval(waitLogTimer);
        waitLogTimer = null;
      }

      function setBusy(busy) {
        btnSelectAll.disabled = busy;
        btnClearAll.disabled = busy;
        btnRun.disabled = busy;
        list.querySelectorAll('input[type="checkbox"]').forEach((el) => { el.disabled = busy; });
      }

      function ensureGoogleScriptRun_() {
        try {
          if (google && google.script && google.script.run) return true;
        } catch (e) {
          // ignore
        }
        stopHeartbeat();
        stopWaitLog();
        setBusy(false);
        const message = 'google.script.run が利用できません（ダイアログを閉じて再度お試しください）';
        setMsg('エラー: ' + message, false);
        logLine('失敗: ' + message);
        return false;
      }

      function setMsg(text, muted) {
        msg.textContent = text || '';
        msg.className = muted ? 'msg muted' : 'msg';
      }

      function startHeartbeat(prefix) {
        const start = Date.now();
        stopHeartbeat();
        heartbeatTimer = setInterval(() => {
          const sec = Math.floor((Date.now() - start) / 1000);
          setMsg(prefix + '（サーバー応答待ち: ' + sec + '秒）', true);
        }, 1000);
      }

      function stopHeartbeat() {
        if (heartbeatTimer) clearInterval(heartbeatTimer);
        heartbeatTimer = null;
      }

      function render() {
        list.innerHTML = '';
        if (!options || options.length === 0) {
          const p = document.createElement('div');
          p.className = 'muted';
          p.textContent = 'ステータスが見つかりません。';
          list.appendChild(p);
          btnRun.disabled = true;
          return;
        }

        options.forEach((opt, idx) => {
          const id = 'status_' + idx;
          const row = document.createElement('div');
          row.className = 'item';

          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.id = id;
          checkbox.value = opt.value;

          const label = document.createElement('label');
          label.htmlFor = id;
          label.textContent = opt.label;

          row.appendChild(checkbox);
          row.appendChild(label);
          list.appendChild(row);
        });
      }

      function getSelected() {
        return Array.from(list.querySelectorAll('input[type=\"checkbox\"]:checked')).map((el) => el.value);
      }

      btnSelectAll.addEventListener('click', () => {
        list.querySelectorAll('input[type=\"checkbox\"]').forEach((el) => { el.checked = true; });
      });

      btnClearAll.addEventListener('click', () => {
        list.querySelectorAll('input[type=\"checkbox\"]').forEach((el) => { el.checked = false; });
      });

      btnCancel.addEventListener('click', () => {
        logLine('キャンセルしました（ダイアログを閉じます）');
        google.script.host.close();
      });

      btnRun.addEventListener('click', () => {
        const selected = getSelected();
        if (selected.length === 0) {
          setMsg('ステータスを1つ以上選択してください。', false);
          return;
        }
        setBusy(true);
        setMsg('準備中...', true);
        log.textContent = '';
        logLine('実行開始');
        logLine('選択: ' + selected.join(', '));
        startHeartbeat('準備中...');
        if (!ensureGoogleScriptRun_()) return;
        logLine('サーバー呼び出し: startClearAnalysisStatusJob');
        startWaitLog('startClearAnalysisStatusJob');

        try {
          google.script.run
            .withSuccessHandler((job) => {
              stopHeartbeat();
              stopWaitLog();
              if (!job || !job.jobId) {
                setBusy(false);
                setMsg('対象となる行はありませんでした。', false);
                logLine('対象0件のため終了');
                return;
              }
              activeJobId = job.jobId;
              totalRows = job.totalRows || 0;
              deletedRows = 0;
              logLine('削除ジョブ作成: ' + activeJobId);
              logLine('削除対象: ' + totalRows + '行');
              runStep();
            })
            .withFailureHandler((err) => {
              stopHeartbeat();
              stopWaitLog();
              setBusy(false);
              const message = err && err.message ? err.message : String(err);
              setMsg('エラー: ' + message, false);
              logLine('失敗: ' + message);
            })
            .startClearAnalysisStatusJob(selected);
        } catch (e) {
          stopHeartbeat();
          stopWaitLog();
          setBusy(false);
          const message = e && e.message ? e.message : String(e);
          setMsg('エラー: ' + message, false);
          logLine('失敗: ' + message);
        }
      });

      function runStep() {
        if (!activeJobId) {
          setBusy(false);
          setMsg('エラー: ジョブIDがありません。', false);
          return;
        }

        const prefix = totalRows > 0
          ? ('削除中... ' + deletedRows + '/' + totalRows + '行')
          : ('削除中... ' + deletedRows + '行');
        setMsg(prefix, true);
        startHeartbeat(prefix);
        if (!ensureGoogleScriptRun_()) return;
        logLine('サーバー呼び出し: processClearAnalysisStatusJob');
        startWaitLog('processClearAnalysisStatusJob');

        try {
          google.script.run
            .withSuccessHandler((res) => {
              stopHeartbeat();
              stopWaitLog();
              if (!res) {
                setBusy(false);
                setMsg('エラー: サーバー応答が不正です。', false);
                logLine('失敗: 不正な応答');
                return;
              }

              deletedRows = typeof res.deletedRows === 'number' ? res.deletedRows : deletedRows;
              const processedRows = typeof res.processedRows === 'number' ? res.processedRows : 0;
              const scannedRows = typeof res.scannedRows === 'number' ? res.scannedRows : 0;
              const cursorRow = typeof res.cursorRow === 'number' ? res.cursorRow : 0;

              const totalPart = totalRows > 0 ? ('/' + totalRows) : '';
              logLine('削除ステップ完了: +' + processedRows + '行（累計 ' + deletedRows + totalPart + '）');
              if (scannedRows > 0) {
                logLine('スキャン: ' + scannedRows + '行（次の走査行: ' + cursorRow + '）');
              }

              if (res.done) {
                setMsg('完了しました。', false);
                logLine('完了');
                google.script.host.close();
                return;
              }

              const nextMsg = totalRows > 0
                ? ('削除中... ' + deletedRows + '/' + totalRows + '行（次の走査行: ' + cursorRow + '）')
                : ('削除中... ' + deletedRows + '行（次の走査行: ' + cursorRow + '）');
              setMsg(nextMsg, true);
              setTimeout(runStep, 50);
            })
            .withFailureHandler((err) => {
              stopHeartbeat();
              stopWaitLog();
              setBusy(false);
              const message = err && err.message ? err.message : String(err);
              setMsg('エラー: ' + message, false);
              logLine('失敗: ' + message);
            })
            .processClearAnalysisStatusJob(activeJobId);
        } catch (e) {
          stopHeartbeat();
          stopWaitLog();
          setBusy(false);
          const message = e && e.message ? e.message : String(e);
          setMsg('エラー: ' + message, false);
          logLine('失敗: ' + message);
        }
      }

      render();
      setMsg('', true);
      logLine('ダイアログを開きました');
      // 疎通確認（google.script.run が有効か／サーバーが応答するか）をログに出す
      (function pingServer_() {
        if (!ensureGoogleScriptRun_()) return;
        const start = Date.now();
        logLine('サーバー疎通確認中...');
        startWaitLog('pingClearAnalysisStatusDialog');
        try {
          google.script.run
            .withSuccessHandler((res) => {
              stopWaitLog();
              const ms = Date.now() - start;
              if (res && res.ok) {
                logLine('サーバー疎通: OK (' + ms + 'ms)');
              } else {
                logLine('サーバー疎通: 応答あり (' + ms + 'ms)');
              }
            })
            .withFailureHandler((err) => {
              stopWaitLog();
              const message = err && err.message ? err.message : String(err);
              logLine('サーバー疎通: NG (' + message + ')');
            })
            .pingClearAnalysisStatusDialog();
        } catch (e) {
          stopWaitLog();
          const message = e && e.message ? e.message : String(e);
          logLine('サーバー疎通: NG (' + message + ')');
        }
      })();
    </script>
  </body>
</html>
  `;
}

function pingClearAnalysisStatusDialog() {
  return pingClearAnalysisStatusDialog_();
}

function pingClearAnalysisStatusDialog_() {
  return { ok: true, now: Date.now() };
}

function startClearAnalysisStatusJob(selectedValues) {
  return startClearAnalysisStatusJob_(selectedValues);
}

function startClearAnalysisStatusJob_(selectedValues) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5 * 1000)) {
    throw new Error('スプレッドシートが他の処理中です。しばらくして再度お試しください。');
  }

  try {
    const sheet = getOrCreateAnalysisSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('解析シートにデータがありません。');

    const selected = Array.isArray(selectedValues)
      ? selectedValues.map((value) => normalizeText_(value)).filter((value) => value)
      : [];
    if (selected.length === 0) throw new Error('ステータスが選択されていません。');

    const header = sheet
      .getRange(1, 1, 1, Math.max(RESULT_SHEET_HEADERS.length, sheet.getLastColumn()))
      .getValues()[0];
    const statusCol = header.indexOf('ステータス') + 1;
    if (statusCol === 0) throw new Error('解析シートに「ステータス」列が見つかりません。');

    const jobId = Utilities.getUuid();
    const includeError = selected.includes(ANALYSIS_CLEAR_ERROR_OPTION_VALUE);
    const state = {
      selected: selected,
      includeError: includeError,
      statusCol: statusCol,
      cursorRow: lastRow,
      deletedRows: 0,
      createdAt: Date.now()
    };
    CacheService.getUserCache().put(
      getClearAnalysisStatusJobCacheKey_(jobId),
      JSON.stringify(state),
      ANALYSIS_CLEAR_STATUS_JOB_CACHE_TTL_SEC
    );
    return { jobId: jobId, totalRows: 0 };
  } finally {
    lock.releaseLock();
  }
}

function processClearAnalysisStatusJob(jobId) {
  return processClearAnalysisStatusJob_(jobId);
}

function processClearAnalysisStatusJob_(jobId) {
  const cache = CacheService.getUserCache();
  const cacheKey = getClearAnalysisStatusJobCacheKey_(normalizeText_(jobId));
  const text = cache.get(cacheKey);
  if (!text) {
    throw new Error('処理状態が見つかりませんでした。もう一度お試しください。');
  }

  const state = JSON.parse(text);
  if (!state || !Array.isArray(state.selected)) {
    throw new Error('処理状態が不正です。もう一度お試しください。');
  }

  const sheet = getOrCreateAnalysisSheet_();

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(20 * 1000)) {
    throw new Error('スプレッドシートが他の処理中です。しばらくして再度お試しください。');
  }

  const header = sheet
    .getRange(1, 1, 1, Math.max(RESULT_SHEET_HEADERS.length, sheet.getLastColumn()))
    .getValues()[0];
  const statusCol = header.indexOf('ステータス') + 1;
  if (statusCol === 0) {
    lock.releaseLock();
    throw new Error('解析シートに「ステータス」列が見つかりません。');
  }

  const selectedSet = new Set(
    state.selected
      .map((value) => normalizeText_(value))
      .filter((value) => value && value !== ANALYSIS_CLEAR_ERROR_OPTION_VALUE)
  );
  const includeError = !!state.includeError;

  let cursorRow = Math.min(
    typeof state.cursorRow === 'number' ? state.cursorRow : sheet.getLastRow(),
    sheet.getLastRow()
  );
  let scannedRows = 0;
  let processedRows = 0;
  try {
    const startTime = Date.now();
    const maxWindows = 5;
    let windows = 0;

    while (
      cursorRow >= 2 &&
      processedRows < ANALYSIS_CLEAR_STATUS_MAX_DELETE_ROWS &&
      windows < maxWindows &&
      Date.now() - startTime < 20 * 1000
    ) {
      const windowEnd = cursorRow;
      const windowStart = Math.max(2, windowEnd - ANALYSIS_CLEAR_STATUS_SCAN_WINDOW_ROWS + 1);
      const windowValues = sheet.getRange(windowStart, statusCol, windowEnd - windowStart + 1, 1).getValues();
      scannedRows += windowValues.length;
      windows++;

      let i;
      for (i = windowValues.length - 1; i >= 0 && processedRows < ANALYSIS_CLEAR_STATUS_MAX_DELETE_ROWS; i--) {
        const status = normalizeText_(windowValues[i][0]);
        if (!status) continue;

        const isMatch =
          (includeError && status.startsWith(ANALYSIS_STATUS_ERROR_PREFIX)) ||
          selectedSet.has(status);
        if (!isMatch) continue;

        const rowNumber = windowStart + i;
        sheet.deleteRow(rowNumber);
        processedRows += 1;
        state.deletedRows += 1;
      }

      if (i >= 0 && processedRows >= ANALYSIS_CLEAR_STATUS_MAX_DELETE_ROWS) {
        cursorRow = windowStart + i;
        break;
      }
      cursorRow = windowStart - 1;
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  const done = cursorRow < 2;
  state.cursorRow = cursorRow;
  if (done) {
    cache.remove(cacheKey);
  } else {
    cache.put(cacheKey, JSON.stringify(state), ANALYSIS_CLEAR_STATUS_JOB_CACHE_TTL_SEC);
  }

  return {
    done: done,
    deletedRows: state.deletedRows,
    processedRows: processedRows,
    scannedRows: scannedRows,
    cursorRow: cursorRow
  };
}

function getClearAnalysisStatusJobCacheKey_(jobId) {
  return `${ANALYSIS_CLEAR_STATUS_JOB_CACHE_PREFIX}${jobId}`;
}

function clearAnalysisRowsByStatusSelection_(selectedValues) {
  const sheet = getOrCreateAnalysisSheet_();
  ensureResultSheetLayout_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('解析シートにデータがありません。');
  }

  const selected = Array.isArray(selectedValues)
    ? selectedValues.map((value) => normalizeText_(value)).filter((value) => value)
    : [];
  if (selected.length === 0) {
    throw new Error('ステータスが選択されていません。');
  }

  const header = sheet
    .getRange(1, 1, 1, Math.max(RESULT_SHEET_HEADERS.length, sheet.getLastColumn()))
    .getValues()[0];
  const statusCol = header.indexOf('ステータス') + 1;
  if (statusCol === 0) {
    throw new Error('解析シートに「ステータス」列が見つかりません。');
  }

  const values = sheet.getRange(2, statusCol, lastRow - 1, 1).getValues();
  const selectedSet = new Set(selected);
  const includeError = selectedSet.has(ANALYSIS_CLEAR_ERROR_OPTION_VALUE);
  if (includeError) selectedSet.delete(ANALYSIS_CLEAR_ERROR_OPTION_VALUE);

  const rowsToDelete = [];
  const breakdown = {};
  for (let i = 0; i < values.length; i++) {
    const status = normalizeText_(values[i][0]);
    if (!status) continue;

    let matchedKey = '';
    if (includeError && status.startsWith(ANALYSIS_STATUS_ERROR_PREFIX)) {
      matchedKey = ANALYSIS_CLEAR_ERROR_OPTION_LABEL;
    } else if (selectedSet.has(status)) {
      matchedKey = status;
    }

    if (matchedKey) {
      rowsToDelete.push(i + 2);
      breakdown[matchedKey] = (breakdown[matchedKey] || 0) + 1;
    }
  }

  if (rowsToDelete.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast('対象となる行はありませんでした。', '解析シートのクリア', 5);
    return { deletedCount: 0, breakdown: breakdown };
  }

  rowsToDelete.sort((a, b) => b - a);
  let blockEnd = rowsToDelete[0];
  let blockStart = blockEnd;
  for (let idx = 1; idx < rowsToDelete.length; idx++) {
    const row = rowsToDelete[idx];
    if (row === blockStart - 1) {
      blockStart = row;
      continue;
    }
    sheet.deleteRows(blockStart, blockEnd - blockStart + 1);
    blockEnd = row;
    blockStart = row;
  }
  sheet.deleteRows(blockStart, blockEnd - blockStart + 1);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `削除しました: ${rowsToDelete.length}件`,
    '解析シートのクリア',
    5
  );
  return { deletedCount: rowsToDelete.length, breakdown: breakdown };
}

// ==================================================
// 変更案の再生成（全行）
// ==================================================
function regenerateAllNameCandidates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateAnalysisSheet_();
  const ui = SpreadsheetApp.getUi();

  const settings = getSettingsOrAlert_({ requireFolder: true });
  if (!settings) return;

  ensureResultSheetLayout_(sheet);
  ss.setActiveSheet(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('解析シートにデータがありません。');
    return;
  }

  const fileNameRules = getFileNameRules_(settings.fileNameRuleSheetName);
  if (fileNameRules.length === 0) {
    ui.alert('ファイル名ルールが設定されていません。');
    return;
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(settings.folderId);
  } catch (e) {
    ui.alert('フォルダが見つかりません。URLが正しいか確認してください。\n' + e.message);
    return;
  }

  const list = listTargetFiles_(folder, new Set());
  const nameCounts = list.nameCounts;

  const data = sheet.getRange(2, 1, lastRow - 1, RESULT_SHEET_HEADERS.length).getValues();
  const existingFullNames = new Set();
  let updatedCount = 0;
  let duplicateCount = 0;
  let errorCount = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowIndex = i + 2;
    const fileId = row[0];
    if (!fileId) continue;

    try {
      const originalName = row[2];
      const extension = extractFileExtension_(originalName);
      const candidateBase = buildFileNameCandidate_(
        {
          paymentDate: row[6],
          paymentMethod: row[7],
          cardInfo: row[8],
          vendorName: row[9],
          invoiceNumber: row[10],
          summary: row[11],
          amount: row[12]
        },
        fileNameRules,
        settings.delimiter
      );

      if (!candidateBase) {
        sheet.getRange(rowIndex, 4).setValue('');
        sheet.getRange(rowIndex, 4).setBackground(null);
        highlightEmptyExtractionCells_(sheet, rowIndex);
        highlightTaxCategoryCell_(sheet, rowIndex, row[14]);
        continue;
      }

      const resolved = resolveDuplicateCandidate_(
        candidateBase,
        originalName,
        extension,
        nameCounts,
        existingFullNames
      );

      sheet.getRange(rowIndex, 4).setValue(resolved.full);
      sheet.getRange(rowIndex, 4).setBackground(resolved.duplicate ? DUPLICATE_CELL_COLOR : null);
      highlightEmptyExtractionCells_(sheet, rowIndex);
      highlightTaxCategoryCell_(sheet, rowIndex, row[14]);

      if (resolved.full) existingFullNames.add(resolved.full);
      if (resolved.duplicate) duplicateCount++;
      updatedCount++;
    } catch (e) {
      console.error(e);
      errorCount++;
      logError_('変更案再生成', fileId, row[2], e.message);
    }
  }

  const messages = [`${updatedCount} 件の変更案を再生成しました。`];
  if (duplicateCount > 0) messages.push(`重複候補: ${duplicateCount} 件`);
  if (errorCount > 0) messages.push(`エラー: ${errorCount} 件`);
  ui.alert(messages.join('\n'));
}

// ==================================================
// 移動先リストシートから [キーワード -> フォルダURL] を取得
// ==================================================
function getDestinationFolderIdByKeyword_(keyword) {
  if (!keyword) return null;

  const sheet = ensureDestinationSheet_().sheet;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (const [key, folderUrl] of values) {
    if (!key || !folderUrl) continue;
    if (String(key).trim() === String(keyword).trim()) {
      if (!isFolderUrl_(folderUrl)) return null;
      return extractFolderIdFromUrl_(folderUrl);
    }
  }
  return null;
}

// ==================================================
// 指定フォルダに移動（E列=キーワード, F列=ステータス）
// ==================================================
function moveFilesToSpecifiedFolder() {
  const sheet = getOrCreateAnalysisSheet_();
  const ui = SpreadsheetApp.getUi();

  ensureResultSheetLayout_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('データがありません。');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  let movedCount = 0;
  let skippedCount = 0;
  let unmatchedCount = 0;
  let errorCount = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const fileId = row[0];
    const destinationKeyword = row[4];
    const status = row[5];

    if (!fileId || !destinationKeyword || status === '移動済み') {
      skippedCount++;
      continue;
    }

    const destinationFolderId = getDestinationFolderIdByKeyword_(destinationKeyword);
    if (!destinationFolderId) {
      sheet.getRange(i + 2, 6).setValue('移動先なし');
      unmatchedCount++;
      continue;
    }

    try {
      const file = DriveApp.getFileById(fileId);
      const folder = DriveApp.getFolderById(destinationFolderId);
      file.moveTo(folder);
      sheet.getRange(i + 2, 6).setValue('移動済み');
      movedCount++;
    } catch (e) {
      console.error(e);
      sheet.getRange(i + 2, 6).setValue('移動失敗');
      logError_('フォルダ移動', fileId, '', e.message);
      errorCount++;
    }
  }

  ui.alert(
    `移動処理が完了しました。\n移動済み: ${movedCount}\n未移動: ${skippedCount}\n移動先なし: ${unmatchedCount}\nエラー: ${errorCount}`
  );
}

// ==================================================
// マネフォ用解析: 解析結果シート -> マネフォ用シート
// ==================================================
function analyzeMoneyForward() {
  const ui = SpreadsheetApp.getUi();

  const settings = getSettingsOrAlert_({ requireApiKey: true });
  if (!settings) return;

  const sheet = getOrCreateAnalysisSheet_();
  ensureResultSheetLayout_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('解析結果がありません。');
    return;
  }

  const partnerMap = getTradePartnerMap_(ui);
  const accountRules = getAccountRules_();
  const rows = [];
  const mixedTaxFlags = [];
  const errors = [];
  let transactionNo = 1;

  const headerWidth = Math.max(RESULT_SHEET_HEADERS.length, sheet.getLastColumn());
  const headerRow = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
  const headerIndexMap = buildHeaderIndexMap_(headerRow);
  const data = sheet.getRange(2, 1, lastRow - 1, headerWidth).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const fileId = getRowValueByHeader_(
      row,
      headerIndexMap,
      'ファイルID',
      RESULT_SHEET_HEADERS.indexOf('ファイルID')
    );
    if (!fileId) continue;

    try {
      const file = DriveApp.getFileById(fileId);
      const fileName = file.getName();
      if (isCsvMarkedFile_(fileName)) continue;

      const paymentDate = normalizeDate_(getRowValueByHeader_(
        row,
        headerIndexMap,
        '支払日',
        RESULT_SHEET_HEADERS.indexOf('支払日')
      )) || formatDate_(file.getDateCreated());
      const paymentMethod = normalizePaymentMethod_(getRowValueByHeader_(
        row,
        headerIndexMap,
        '支払い方法',
        RESULT_SHEET_HEADERS.indexOf('支払い方法')
      ));
      const cardInfo = normalizeCardInfo_(getRowValueByHeader_(
        row,
        headerIndexMap,
        'カード情報',
        RESULT_SHEET_HEADERS.indexOf('カード情報')
      ), paymentMethod);
      const vendorName = normalizeText_(getRowValueByHeader_(
        row,
        headerIndexMap,
        '取引先',
        RESULT_SHEET_HEADERS.indexOf('取引先')
      ));
      const invoiceNumber = normalizeInvoiceNumber_(getRowValueByHeader_(
        row,
        headerIndexMap,
        'インボイス番号',
        RESULT_SHEET_HEADERS.indexOf('インボイス番号')
      ));
      const summary = normalizeText_(getRowValueByHeader_(
        row,
        headerIndexMap,
        '品目（概要）',
        RESULT_SHEET_HEADERS.indexOf('品目（概要）')
      ));
      const amount = normalizeAmount_(getRowValueByHeader_(
        row,
        headerIndexMap,
        '金額',
        RESULT_SHEET_HEADERS.indexOf('金額')
      ));
      const rawTaxCategory = getRowValueByHeader_(
        row,
        headerIndexMap,
        '消費税区分',
        RESULT_SHEET_HEADERS.indexOf('消費税区分')
      );
      const taxCategory = normalizeTaxCategory_(rawTaxCategory);

      const accountTitle = accountRules.length > 0
        ? callGeminiApiForAccountTitle_(
          {
            paymentDate: paymentDate,
            paymentMethod: paymentMethod,
            cardInfo: cardInfo,
            vendorName: vendorName,
            invoiceNumber: invoiceNumber,
            summary: summary,
            amount: amount
          },
          accountRules,
          settings.apiKey
        )
        : '未確定勘定';

      const partnerName = resolvePartnerName_(invoiceNumber, partnerMap, vendorName);
      const merged = {
        date: paymentDate,
        amount: amount,
        invoiceNumber: invoiceNumber,
        vendorName: vendorName,
        summary: summary,
        paymentMethod: paymentMethod,
        cardInfo: cardInfo,
        taxCategory: rawTaxCategory,
        accountTitle: accountTitle
      };
      const rowData = buildMoneyForwardRow_(
        transactionNo,
        merged,
        partnerName,
        file.getUrl(),
        settings
      );
      rows.push(rowData);
      mixedTaxFlags.push(taxCategory === '混在あり');
      transactionNo++;
    } catch (e) {
      console.error(e);
      errors.push(`行${i + 2}: ${e.message}`);
      logError_('マネフォ用解析', fileId, '', e.message);
    }
  }

  const mfSheet = getMoneyForwardSheet_(ui);
  resetMoneyForwardSheet_(mfSheet);

  if (rows.length === 0) {
    ui.alert('対象ファイルがありませんでした。');
    return;
  }

  mfSheet.getRange(2, 1, rows.length, MF_CSV_HEADERS.length).setValues(rows);
  applyMoneyForwardAccountTitleValidation_(mfSheet);
  applyMoneyForwardTaxCategoryHighlights_(mfSheet, mixedTaxFlags);
  const messages = [`${rows.length} 件のレシートを「マネフォ用」シートに更新しました。`];
  if (errors.length > 0) {
    messages.push(`エラー: ${errors.length} 件`);
  }
  ui.alert(messages.join('\n'));
}

function buildHeaderIndexMap_(headerRow) {
  const map = {};
  if (!headerRow || !Array.isArray(headerRow)) return map;
  for (let i = 0; i < headerRow.length; i++) {
    const key = normalizeText_(headerRow[i]);
    if (!key || Object.prototype.hasOwnProperty.call(map, key)) continue;
    map[key] = i;
  }
  return map;
}

function getRowValueByHeader_(row, headerIndexMap, headerName, fallbackIndex) {
  if (!row || !Array.isArray(row)) return '';
  const mappedIndex = headerIndexMap ? headerIndexMap[headerName] : undefined;
  if (typeof mappedIndex === 'number' && mappedIndex >= 0 && mappedIndex < row.length) {
    return row[mappedIndex];
  }
  if (typeof fallbackIndex === 'number' && fallbackIndex >= 0 && fallbackIndex < row.length) {
    return row[fallbackIndex];
  }
  return '';
}

function getMoneyForwardSheet_(ui) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('マネフォ用');

  if (!sheet) {
    sheet = ss.insertSheet('マネフォ用');
    ui.alert('「マネフォ用」シートを作成しました。');
  }

  return sheet;
}

function resetMoneyForwardSheet_(sheet) {
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.clearContents();
  sheet.getRange(1, 1, 1, MF_CSV_HEADERS.length).setValues([MF_CSV_HEADERS]);
  sheet.setFrozenRows(1);
  applyMoneyForwardAccountTitleValidation_(sheet);
}

function applyMoneyForwardAccountTitleValidation_(sheet) {
  if (!sheet) return;

  const accountCol = MF_CSV_HEADERS.indexOf('借方勘定科目') + 1;
  if (accountCol <= 0) return;

  const accountRuleSheet = ensureAccountRuleSheet_().sheet;
  const ruleLastRow = accountRuleSheet.getLastRow();
  const targetRows = Math.max(sheet.getMaxRows() - 1, 1);
  const targetRange = sheet.getRange(2, accountCol, targetRows, 1);

  if (ruleLastRow < 2) {
    targetRange.clearDataValidations();
    return;
  }

  const sourceRange = accountRuleSheet.getRange(2, 1, ruleLastRow - 1, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(true)
    .build();
  targetRange.setDataValidation(rule);
}

function openCsvDownloadSelector() {
  const html = HtmlService.createHtmlOutput(buildCsvDownloadSelectorDialogHtml_())
    .setWidth(440)
    .setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(html, 'CSVダウンロード');
}

function buildCsvDownloadSelectorDialogHtml_() {
  return `
<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <style>
      html,
      body {
        font-family: "Noto Sans JP", Arial, sans-serif;
        margin: 0;
        padding: 0;
      }
      .wrap {
        padding: 16px;
      }
      .title {
        font-size: 14px;
        font-weight: 700;
        margin: 0 0 8px;
      }
      .desc {
        color: #555;
        font-size: 12px;
        margin: 0 0 12px;
      }
      .item {
        align-items: center;
        display: flex;
        gap: 8px;
        margin: 8px 0;
      }
      .item label {
        cursor: pointer;
        font-size: 13px;
      }
      .actions {
        display: flex;
        gap: 8px;
        margin-top: 12px;
      }
      button {
        background: #f3f3f3;
        border: 1px solid #d0d0d0;
        border-radius: 8px;
        cursor: pointer;
        font-size: 12px;
        padding: 8px 12px;
      }
      button.primary {
        background: #1a73e8;
        border-color: #1a73e8;
        color: #fff;
      }
      button:disabled {
        cursor: not-allowed;
        opacity: 0.6;
      }
      .msg {
        color: #555;
        font-size: 12px;
        line-height: 1.5;
        margin-top: 10px;
        min-height: 16px;
        white-space: pre-wrap;
      }
      .error {
        color: #b91c1c;
      }
    </style>
  </head>
  <body>
    <div class="wrap">
      <div class="title">ダウンロード対象を選択</div>
      <p class="desc">チェックしたCSVを出力します（複数選択可）。</p>
      <div class="item">
        <input type="checkbox" id="moneyForward" checked>
        <label for="moneyForward">マネフォ用CSV</label>
      </div>
      <div class="item">
        <input type="checkbox" id="tradePartners" checked>
        <label for="tradePartners">取引先CSV</label>
      </div>
      <div class="actions">
        <button type="button" id="run" class="primary">ダウンロード</button>
        <button type="button" id="cancel">キャンセル</button>
      </div>
      <div id="msg" class="msg"></div>
    </div>
    <script>
      const runButton = document.getElementById('run');
      const cancelButton = document.getElementById('cancel');
      const msg = document.getElementById('msg');
      const moneyForward = document.getElementById('moneyForward');
      const tradePartners = document.getElementById('tradePartners');

      function ensureGoogleScriptRun_() {
        try {
          return !!(google && google.script && google.script.run);
        } catch (e) {
          return false;
        }
      }

      function setBusy(busy) {
        runButton.disabled = busy;
        cancelButton.disabled = busy;
        moneyForward.disabled = busy;
        tradePartners.disabled = busy;
      }

      function setMessage(text, isError) {
        msg.textContent = text || '';
        if (isError) {
          msg.classList.add('error');
        } else {
          msg.classList.remove('error');
        }
      }

      function triggerDownloads(downloads) {
        if (!downloads || downloads.length === 0) {
          setBusy(false);
          const currentMessage = (msg.textContent || '').trim();
          if (!currentMessage || currentMessage === 'CSVを準備しています...') {
            setMessage('ダウンロード対象がありません。', true);
          }
          return;
        }

        let index = 0;
        const runNext = () => {
          if (index >= downloads.length) {
            setMessage('ダウンロードを開始しました。', false);
            setTimeout(() => google.script.host.close(), 700);
            return;
          }

          const target = downloads[index++];
          const blob = new Blob([target.csvContent || ''], { type: 'text/csv;charset=utf-8' });
          const url = URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url;
          link.download = target.filename || 'download.csv';
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          setTimeout(() => {
            URL.revokeObjectURL(url);
            runNext();
          }, 300);
        };

        runNext();
      }

      runButton.addEventListener('click', () => {
        if (!ensureGoogleScriptRun_()) {
          setMessage('google.script.run が利用できません（ダイアログを閉じて再度お試しください）', true);
          return;
        }

        const options = {
          moneyForward: moneyForward.checked,
          tradePartners: tradePartners.checked
        };
        if (!options.moneyForward && !options.tradePartners) {
          setMessage('少なくとも1つ選択してください。', true);
          return;
        }

        setBusy(true);
        setMessage('CSVを準備しています...', false);
        try {
          google.script.run
            .withSuccessHandler((result) => {
              const messages = result && result.messages ? result.messages.filter(Boolean) : [];
              if (messages.length > 0) {
                setMessage(messages.join('\\n'), false);
              }
              triggerDownloads(result && result.downloads ? result.downloads : []);
            })
            .withFailureHandler((error) => {
              setBusy(false);
              const message = error && error.message ? error.message : 'CSVの準備に失敗しました。';
              setMessage(message, true);
            })
            .prepareCsvDownloads(options);
        } catch (error) {
          setBusy(false);
          const message = error && error.message ? error.message : 'CSVの準備に失敗しました。';
          setMessage(message, true);
        }
      });

      cancelButton.addEventListener('click', () => google.script.host.close());

      (function pingServer_() {
        if (!ensureGoogleScriptRun_()) return;
        try {
          google.script.run
            .withSuccessHandler(() => {})
            .withFailureHandler((error) => {
              const message = error && error.message ? error.message : 'サーバー疎通に失敗しました。';
              setMessage(message, true);
            })
            .pingCsvDownloadDialog();
        } catch (error) {
          const message = error && error.message ? error.message : 'サーバー疎通に失敗しました。';
          setMessage(message, true);
        }
      })();
    </script>
  </body>
</html>
  `;
}

function prepareCsvDownloads(options) {
  return prepareCsvDownloads_(options);
}

function prepareCsvDownloads_(options) {
  const wantMoneyForward = !!(options && options.moneyForward);
  const wantTradePartners = !!(options && options.tradePartners);
  const downloads = [];
  const messages = [];
  const errors = [];

  if (wantMoneyForward) {
    try {
      const payload = buildMoneyForwardCsvPayload_();
      downloads.push({ filename: payload.filename, csvContent: payload.csvContent });
      if (payload.markedCount > 0) {
        messages.push(`マネフォ用: ${payload.markedCount} 件のファイルに「CSV済」を付与しました。`);
      }
    } catch (e) {
      errors.push(String(e && e.message ? e.message : e));
    }
  }

  if (wantTradePartners) {
    try {
      const payload = buildTradePartnersCsvPayload_();
      downloads.push({ filename: payload.filename, csvContent: payload.csvContent });
      if (payload.missingNames > 0) {
        messages.push(`取引先CSV: 取引先名が空欄の行 ${payload.missingNames} 件を含めて出力します。`);
      }
    } catch (e) {
      errors.push(String(e && e.message ? e.message : e));
    }
  }

  if (downloads.length === 0 && errors.length > 0) {
    throw new Error(errors.join('\n'));
  }

  if (errors.length > 0) {
    messages.push(...errors.map((message) => `未出力: ${message}`));
  }

  return { downloads: downloads, messages: messages };
}

function pingCsvDownloadDialog() {
  return { ok: true, timestamp: Date.now() };
}

function buildMoneyForwardCsvPayload_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('マネフォ用');
  if (!sheet) {
    throw new Error('「マネフォ用」シートが見つかりません。先に「5. マネフォ用解析」を実行してください。');
  }

  const filter = sheet.getFilter();
  if (filter) filter.remove();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('「マネフォ用」シートにデータがありません。');
  }

  const range = sheet.getRange(2, 1, lastRow - 1, MF_CSV_HEADERS.length);
  const values = range.getValues();
  const rows = values.filter((row) => row.some((cell) => !isBlankCell_(cell)));
  if (rows.length === 0) {
    throw new Error('「マネフォ用」シートに出力対象の行がありません。');
  }

  return {
    filename: buildMoneyForwardFilename_(),
    csvContent: buildCsvContent_(MF_CSV_HEADERS, rows),
    markedCount: markCsvProcessedFiles_(rows)
  };
}

function downloadMoneyForwardCsv() {
  const ui = SpreadsheetApp.getUi();
  try {
    const payload = buildMoneyForwardCsvPayload_();
    showDownloadDialog_(payload.filename, payload.csvContent);

    const messages = ['CSVをダウンロードしました。'];
    if (payload.markedCount > 0) {
      messages.push(`${payload.markedCount} 件のファイルに「CSV済」を付与しました。`);
    }
    ui.alert(messages.join('\n'));
  } catch (e) {
    ui.alert(String(e && e.message ? e.message : e));
  }
}

function isAllowedReceiptMimeType_(mimeType) {
  const allowedTypes = [
    MimeType.JPEG,
    MimeType.PNG,
    MimeType.PDF,
    'image/heic',
    'image/heif'
  ];
  return allowedTypes.includes(mimeType);
}

function parseReceiptFilename_(fileName) {
  const baseName = fileName.replace(/\.[^.]+$/, '');
  const parts = baseName.split('｜');
  if (parts.length < 4) return {};

  const paymentMethod = normalizePaymentMethod_(parts[0]);
  const rawDate = normalizeText_(parts[1]);
  const date =
    /^\d{8}$/.test(rawDate) || /^\d{4}[-/]\d{2}[-/]\d{2}$/.test(rawDate)
      ? normalizeDate_(rawDate)
      : '';
  const invoiceNumber = normalizeInvoiceNumber_(parts[2]);
  const amountText = normalizeText_(parts[3]);
  const hasAmountField = parts.length >= 5 && /^\d[\d,]*$/.test(amountText);
  const summaryStartIndex = hasAmountField ? 4 : 3;
  const summary = normalizeText_(parts.slice(summaryStartIndex).join('｜'));

  return {
    paymentMethod: paymentMethod,
    date: date,
    invoiceNumber: invoiceNumber,
    summary: summary
  };
}

function getAccountRules_() {
  const sheet = ensureAccountRuleSheet_().sheet;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const rules = [];
  for (const [name, rule] of values) {
    const accountName = normalizeText_(name);
    if (!accountName) continue;
    rules.push({
      name: accountName,
      rule: normalizeText_(rule)
    });
  }
  return rules;
}

function callGeminiApiForAccountTitle_(data, rules, apiKey) {
  const modelName = getModelName_();
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
  const prompt = buildAccountTitlePrompt_(data, rules);

  const payload = {
    "contents": [{
      "parts": [{ "text": prompt }]
    }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const json = JSON.parse(response.getContentText());
    if (json.error) {
      Logger.log(`API Error: ${JSON.stringify(json.error)}`);
  return '未確定勘定';
}
    const text = json?.candidates?.[0]?.content?.parts?.[0]?.text;
    const parsed = extractJsonFromText_(text);
    const accountTitle = normalizeText_(parsed?.accountTitle);
    const candidates = rules.map((rule) => rule.name);
    if (candidates.includes(accountTitle)) return accountTitle;
  } catch (e) {
    Logger.log(`API Error: ${e.message}`);
  }
  return '未確定勘定';
}

function buildAccountTitlePrompt_(data, rules) {
  const ruleLines = rules
    .map((rule) => rule.rule ? `${rule.name}：${rule.rule}` : rule.name)
    .join('\n');

  return `
次のレシート情報から、勘定科目を1つ選んでください。
出力はJSONのみです（前後の説明は禁止）。
候補にない勘定科目は出力しないでください。

出力形式:
{"accountTitle":"勘定科目名"}

勘定科目の候補とルール:
${ruleLines}

入力:
- 支払日: ${data.paymentDate || ''}
- 支払い方法: ${data.paymentMethod || ''}
- 取引先: ${data.vendorName || ''}
- インボイス番号: ${data.invoiceNumber || ''}
- 摘要: ${data.summary || ''}
- 金額: ${data.amount || ''}
  `;
}

function extractJsonFromText_(text) {
  if (!text) return null;
  const cleaned = text.replace(/```(?:json)?/g, '').replace(/```/g, '').trim();
  const match = cleaned.match(/\{[\s\S]*\}/);
  if (!match) return null;
  try {
    return JSON.parse(match[0]);
  } catch (e) {
    console.error(e);
    return null;
  }
}

function normalizeReceiptData_(data) {
  if (!data) return null;
  return {
    date: normalizeDate_(data.date),
    amount: normalizeAmount_(data.amount),
    invoiceNumber: normalizeInvoiceNumber_(data.invoiceNumber),
    vendorName: normalizeText_(data.vendorName),
    summary: normalizeText_(data.summary),
    paymentMethod: normalizePaymentMethod_(data.paymentMethod),
    taxCategory: normalizeTaxCategory_(data.taxCategory),
    accountTitle: normalizeAccountTitle_(data.accountTitle)
  };
}

function normalizePaymentMethod_(value) {
  const text = normalizeText_(value);
  if (!text) return '現金';

  const compact = text.replace(/\s+/g, '');
  const lower = text.toLowerCase();
  const lowerCompact = lower.replace(/\s+/g, '');

  if (lowerCompact.includes('paypay') || compact.includes('ペイペイ')) return 'PayPay';

  const isId = /(^|\W)i\s*d(\W|$)/i.test(text) || compact.includes('ｉｄ');
  const isQuicPay =
    lowerCompact.includes('quicpay') ||
    lowerCompact.includes('quickpay') ||
    lowerCompact.includes('quiqpay') ||
    lowerCompact.includes('qpay') ||
    compact.includes('クイックペイ');
  if (
    isId ||
    isQuicPay ||
    text.includes('クレカ') ||
    text.includes('クレジット') ||
    text.includes('カード') ||
    lower.includes('visa') ||
    lower.includes('master') ||
    lower.includes('jcb') ||
    lower.includes('amex') ||
    lower.includes('diners')
  ) {
    return 'クレカ';
  }

  if (
    compact.includes('銀行振込') ||
    compact.includes('振込') ||
    compact.includes('振り込み') ||
    compact.includes('振替') ||
    lower.includes('bank transfer')
  ) {
    return '銀行振込';
  }

  if (
    text.includes('電子') ||
    text.includes('交通系') ||
    text.includes('楽天Edy') ||
    text.includes('WAON') ||
    text.includes('nanaco') ||
    text.includes('はやかけん') ||
    text.includes('SUGOCA') ||
    text.includes('nimoca') ||
    text.includes('manaca') ||
    text.includes('TOICA') ||
    text.includes('ICOCA') ||
    text.includes('PASMO') ||
    text.includes('Suica') ||
    text.includes('IC') ||
    lower.includes('suica') ||
    lower.includes('pasmo') ||
    lower.includes('icoca') ||
    lower.includes('toica') ||
    lower.includes('manaca') ||
    lower.includes('hayakaken') ||
    lower.includes('nimoca') ||
    lower.includes('sugoca') ||
    lower.includes('edy') ||
    lower.includes('waon') ||
    lower.includes('nanaco')
  ) {
    return '電子マネー';
  }

  if (text.includes('現金')) return '現金';
  return '現金';
}

function normalizeTaxCategory_(value, taxHints) {
  const hintSource = {};
  if (value && typeof value === 'object') Object.assign(hintSource, value);
  if (taxHints && typeof taxHints === 'object') Object.assign(hintSource, taxHints);

  const categorySource = value && typeof value === 'object' ? value.taxCategory : value;
  const rawCategory = normalizeTaxEvalText_(categorySource || hintSource.taxCategory);
  const compactCategory = rawCategory.replace(/\s+/g, '');
  if (compactCategory.includes('混在あり') || compactCategory.includes('混在')) {
    return '混在あり';
  }

  const parsedRate = parseTaxRatePercent_(categorySource || hintSource.taxCategory);
  if (parsedRate === 8) return '8%';
  if (parsedRate === 10) return '10%';

  const textParts = [];
  if (typeof value === 'string') textParts.push(value);
  if (value && typeof value === 'object' && value.evidence) {
    textParts.push(String(value.evidence));
  }
  if (hintSource.taxCategory) textParts.push(String(hintSource.taxCategory));
  if (hintSource.evidence) textParts.push(String(hintSource.evidence));
  const evalText = normalizeTaxEvalText_(textParts.join('\n'));

  const amountSignals = buildTaxAmountSignals_(evalText, hintSource);
  if (amountSignals.hasTax8Amount || amountSignals.hasTax10Amount) {
    const tax8Amount = amountSignals.hasTax8Amount ? amountSignals.tax8Amount : 0;
    const tax10Amount = amountSignals.hasTax10Amount ? amountSignals.tax10Amount : 0;

    if (tax8Amount > 0 && tax10Amount > 0) return '混在あり';
    if (tax8Amount === 0 && tax10Amount > 0) return '10%';
    if (tax8Amount > 0 && tax10Amount === 0) return '8%';
    if (
      amountSignals.hasTax8Amount &&
      amountSignals.hasTax10Amount &&
      tax8Amount === 0 &&
      tax10Amount === 0
    ) {
      return '10%';
    }
  }

  const mentions = extractTaxMentions_(evalText);
  if (mentions.has8Target && mentions.has10Target) return '混在あり';
  if (mentions.has8Target && !mentions.has10Target) return '8%';
  if (mentions.has10Target && !mentions.has8Target) return '10%';
  if (mentions.has8Mention && mentions.has10Mention) return '混在あり';
  if (mentions.has8Mention) return '8%';

  const isLabelLike = compactCategory.length <= 12;
  if (isLabelLike && compactCategory.includes('8%')) return '8%';
  if (isLabelLike && compactCategory.includes('10%')) return '10%';
  return '10%';
}

function buildTaxAmountSignals_(text, hints) {
  const hint8 = readTaxAmountHint_(hints, 'tax8Amount', 'hasTax8Amount');
  const hint10 = readTaxAmountHint_(hints, 'tax10Amount', 'hasTax10Amount');
  const tax8Amounts = extractTaxAmountsByRate_(text, 8);
  const tax10Amounts = extractTaxAmountsByRate_(text, 10);

  return {
    hasTax8Amount: hint8.found || tax8Amounts.length > 0,
    hasTax10Amount: hint10.found || tax10Amounts.length > 0,
    tax8Amount: hint8.found ? hint8.amount : chooseTaxAmount_(tax8Amounts),
    tax10Amount: hint10.found ? hint10.amount : chooseTaxAmount_(tax10Amounts)
  };
}

function readTaxAmountHint_(hints, key, foundFlagKey) {
  if (!hints || typeof hints !== 'object') {
    return { found: false, amount: 0 };
  }

  if (Object.prototype.hasOwnProperty.call(hints, key)) {
    const raw = hints[key];
    if (raw === '' || raw === null || raw === undefined) {
      return { found: false, amount: 0 };
    }
    return { found: true, amount: normalizeAmount_(raw) };
  }

  if (hints[foundFlagKey] === true) {
    return { found: true, amount: 0 };
  }
  return { found: false, amount: 0 };
}

function chooseTaxAmount_(amounts) {
  if (!amounts || amounts.length === 0) return 0;
  return amounts.reduce((max, value) => (value > max ? value : max), 0);
}

function extractTaxMentions_(text) {
  const lines = normalizeTaxEvalText_(text).split(/\r?\n/);
  let has8Mention = false;
  let has10Mention = false;
  let has8Target = false;
  let has10Target = false;

  for (const rawLine of lines) {
    const compact = String(rawLine || '').replace(/\s+/g, '');
    if (!compact || isTaxLegendLine_(compact)) continue;

    if (/(?:8%\s*対象|対象\s*8%|軽減税率[^0-9]*対象)/.test(compact)) {
      has8Target = true;
    }
    if (/(?:10%\s*対象|対象\s*10%)/.test(compact)) {
      has10Target = true;
    }
    if (compact.includes('軽減税率') || /(^|[^0-9])8%(?!\d)/.test(compact)) {
      has8Mention = true;
    }
    if (/(^|[^0-9])10%(?!\d)/.test(compact)) {
      has10Mention = true;
    }
  }

  return {
    has8Mention: has8Mention,
    has10Mention: has10Mention,
    has8Target: has8Target,
    has10Target: has10Target
  };
}

function isTaxLegendLine_(compactLine) {
  if (!compactLine) return false;
  const withoutRate = compactLine.replace(/\d+%/g, '');
  const hasAmount = /\d/.test(withoutRate);
  if (hasAmount) return false;
  return /軽減税率.*適用|適用商品|対象商品|凡例|★印|※|注記|注:/.test(compactLine);
}

function extractTaxAmountsByRate_(text, rate) {
  const lines = normalizeTaxEvalText_(text).split(/\r?\n/);
  const amounts = [];
  const ratePattern = new RegExp(`(?:^|[^0-9])${rate}\\s*%(?!\\d)`);
  const targetPattern = new RegExp(`(?:${rate}\\s*%\\s*対象|対象\\s*${rate}\\s*%)`);
  const taxKeywordPattern = /(内消費税(?:等)?|消費税(?:額)?|内税(?:額)?|税額|外税)/;

  for (let i = 0; i < lines.length; i++) {
    const line = String(lines[i] || '');
    const compact = line.replace(/\s+/g, '');
    if (!compact || isTaxLegendLine_(compact)) continue;

    if (ratePattern.test(compact) && taxKeywordPattern.test(line)) {
      const lineAmounts = extractAmountsFromLine_(line);
      if (lineAmounts.length > 0) {
        amounts.push(lineAmounts[lineAmounts.length - 1]);
      }
    }

    if (!targetPattern.test(compact)) continue;

    const baseAmount = extractBaseAmountByRate_(line, rate);
    if (baseAmount !== null) {
      amounts.push(baseAmount);
    }

    for (let offset = -1; offset <= 2; offset++) {
      const index = i + offset;
      if (index < 0 || index >= lines.length) continue;
      const nearLine = String(lines[index] || '');
      const nearCompact = nearLine.replace(/\s+/g, '');
      if (!nearCompact || isTaxLegendLine_(nearCompact)) continue;
      if (!taxKeywordPattern.test(nearLine)) continue;
      const nearAmounts = extractAmountsFromLine_(nearLine);
      if (nearAmounts.length > 0) {
        amounts.push(nearAmounts[nearAmounts.length - 1]);
      }
    }
  }

  return amounts.filter((amount) => !isNaN(amount));
}

function extractAmountsFromLine_(line) {
  const normalized = normalizeTaxEvalText_(line).replace(/,/g, '');
  const matches = normalized.match(/\d+/g);
  if (!matches) return [];
  return matches
    .map((value) => parseInt(value, 10))
    .filter((value) => !isNaN(value));
}

function extractBaseAmountByRate_(line, rate) {
  const normalized = normalizeTaxEvalText_(line).replace(/,/g, '');
  const pattern = new RegExp(`(?:${rate}\\s*%\\s*対象|対象\\s*${rate}\\s*%)[^0-9]{0,8}([0-9][0-9]*)`);
  const match = normalized.match(pattern);
  if (!match) return null;
  const amount = parseInt(match[1], 10);
  return isNaN(amount) ? null : amount;
}

function normalizeTaxEvalText_(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/[％﹪]/g, '%')
    .replace(/[￥¥]/g, '¥')
    .replace(/[，]/g, ',')
    .replace(/（/g, '(')
    .replace(/）/g, ')');
}

function parseTaxRatePercent_(value) {
  if (value === null || value === undefined) return null;

  let numeric = null;
  if (typeof value === 'number' && isFinite(value)) {
    numeric = value;
  } else {
    const text = normalizeTaxEvalText_(value).replace(/\s+/g, '');
    if (!text) return null;

    if (text.includes('8%') && !text.includes('10%')) return 8;
    if (text.includes('10%') && !text.includes('8%')) return 10;

    const numericText = text
      .replace(/%/g, '')
      .replace(/[^0-9.+-]/g, '');
    if (!numericText) return null;

    const parsed = Number(numericText);
    if (!isFinite(parsed)) return null;
    numeric = parsed;
  }

  const percent = Math.abs(numeric) <= 1 ? numeric * 100 : numeric;
  if (Math.abs(percent - 8) < 0.01) return 8;
  if (Math.abs(percent - 10) < 0.01) return 10;
  return null;
}

function normalizeCardInfo_(value, paymentMethod) {
  const method = normalizePaymentMethod_(paymentMethod);
  if (method !== 'クレカ') return '-';
  const text = normalizeText_(value);
  if (!text) return 'カード(不明)';

  const digits = text.match(/\d{4}/g);
  if (digits && digits.length > 0) {
    const last4 = digits[digits.length - 1];
    return `カード(${last4})`;
  }

  return 'カード(不明)';
}

function resolveMoneyForwardTaxCategory_(value) {
  const parsedRate = parseTaxRatePercent_(value);
  if (parsedRate === 8) return MF_TAX_CATEGORY_REDUCED;
  if (parsedRate === 10) return MF_TAX_CATEGORY_STANDARD;

  const compact = normalizeTaxEvalText_(value).replace(/\s+/g, '');
  if (compact.includes('混在')) return MF_TAX_CATEGORY_STANDARD;
  if (compact.includes('8%') && !compact.includes('10%')) return MF_TAX_CATEGORY_REDUCED;
  if (compact.includes('10%') && !compact.includes('8%')) return MF_TAX_CATEGORY_STANDARD;
  return normalizeTaxCategory_(value) === '8%'
    ? MF_TAX_CATEGORY_REDUCED
    : MF_TAX_CATEGORY_STANDARD;
}

function normalizeInvoiceNumber_(value) {
  const text = normalizeText_(value);
  if (!text) return '';
  const normalized = text
    .replace(/[Ｔｔ]/g, 'T')
    .replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/\s+/g, '');
  const match = normalized.match(/T\d{13}/);
  return match ? match[0] : '';
}

function normalizeDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return formatDate_(value);
  }

  const text = normalizeText_(value);
  if (/^\d{8}$/.test(text)) {
    return `${text.slice(0, 4)}/${text.slice(4, 6)}/${text.slice(6, 8)}`;
  }
  if (/^\d{4}[-/]\d{2}[-/]\d{2}$/.test(text)) {
    return text.replace(/-/g, '/');
  }
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return formatDate_(parsed);
  }
  return '';
}

function normalizeAmount_(value) {
  if (value === 0) return 0;
  if (!value) return 0;
  const text = String(value).replace(/[^\d]/g, '');
  return text ? parseInt(text, 10) : 0;
}

function normalizeText_(value) {
  if (!value) return '';
  return String(value).replace(/\s+/g, ' ').trim();
}

function isBlankCell_(value) {
  if (value === null || value === undefined) return true;
  if (typeof value === 'string') return value.trim() === '';
  return false;
}

function isCsvMarkedFile_(fileName) {
  if (!fileName) return false;
  return String(fileName).startsWith(MF_PROCESSED_PREFIX);
}

function markCsvProcessedFiles_(rows) {
  const memoIndex = MF_CSV_HEADERS.indexOf('仕訳メモ');
  if (memoIndex === -1) return 0;

  const settings = getSettingsOrAlert_({});
  const delimiter = settings && DELIMITER_CANDIDATES.includes(settings.delimiter)
    ? settings.delimiter
    : DELIMITER_CANDIDATES[0];

  let marked = 0;
  for (const row of rows) {
    const memo = row[memoIndex];
    const fileId = extractDriveFileIdFromUrl_(memo);
    if (!fileId) continue;
    try {
      const file = DriveApp.getFileById(fileId);
      const name = file.getName();
      if (isCsvMarkedFile_(name)) continue;
      file.setName(`${MF_PROCESSED_PREFIX}${delimiter}${name}`);
      marked++;
    } catch (e) {
      console.error(e);
    }
  }
  return marked;
}

function extractDriveFileIdFromUrl_(url) {
  if (!url) return '';
  const text = String(url);
  const match = text.match(/\/d\/([^/]+)\//);
  return match ? match[1] : '';
}

function normalizeAccountTitle_(value) {
  const text = normalizeText_(value);
  if (MF_ACCOUNT_CANDIDATE_NAMES.includes(text)) return text;
  return '未確定勘定';
}

function mergeReceiptData_(analysis, nameInfo, file) {
  const date = analysis.date || nameInfo.date || formatDate_(file.getDateCreated());
  const paymentMethod = nameInfo.paymentMethod || analysis.paymentMethod || '現金';
  const invoiceNumber = nameInfo.invoiceNumber || analysis.invoiceNumber || '';
  const summary = analysis.summary || nameInfo.summary || '';

  return {
    date: date,
    amount: analysis.amount,
    invoiceNumber: invoiceNumber,
    vendorName: analysis.vendorName,
    summary: summary,
    paymentMethod: paymentMethod,
    accountTitle: analysis.accountTitle
  };
}

function getTradePartnerMap_(ui) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MF_TRADE_PARTNER_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(MF_TRADE_PARTNER_SHEET_NAME);
    initializeTradePartnerSheet_(sheet);
    ui.alert(
      '「取引先一覧」シートを作成しました。\n' +
      'マネーフォワードの取引先インポート仕様に合わせています。\n' +
      '- B列: 取引先名（必須）\n' +
      '- E列: 登録番号（T+13桁）'
    );
    return {};
  }

  ensureTradePartnerSheetLayout_(sheet, ui);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const invoiceIndex = header.indexOf('登録番号');
  const nameIndex = header.indexOf('取引先名');

  if (invoiceIndex === -1 || nameIndex === -1) {
    ui.alert('「取引先一覧」シートに「登録番号」「取引先名」の列が必要です。');
    return {};
  }

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const map = {};
  for (const row of values) {
    const invoiceNumber = normalizeInvoiceNumber_(row[invoiceIndex]);
    const partnerName = normalizeText_(row[nameIndex]);
    if (invoiceNumber && partnerName) {
      map[invoiceNumber] = partnerName;
    }
  }
  return map;
}

function initializeTradePartnerSheet_(sheet) {
  sheet.getRange(1, 1, 1, MF_TRADE_PARTNER_HEADERS.length).setValues([MF_TRADE_PARTNER_HEADERS]);
  sheet.getRange(1, 1, 1, MF_TRADE_PARTNER_HEADERS.length).setBackground('#d9ead3').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 120); // コード
  sheet.setColumnWidth(2, 240); // 取引先名
  sheet.setColumnWidth(3, 200); // 検索キー
  sheet.setColumnWidth(4, 110); // 表示設定
  sheet.setColumnWidth(5, 180); // 登録番号
  sheet.setColumnWidth(6, 160); // 法人番号
}

function ensureTradePartnerSheetLayout_(sheet, ui) {
  const lastColumn = sheet.getLastColumn();
  const headerWidth = Math.max(MF_TRADE_PARTNER_HEADERS.length, lastColumn || 0, 2);
  const header = sheet.getRange(1, 1, 1, headerWidth).getValues()[0];
  const a1 = normalizeText_(header[0]);
  const b1 = normalizeText_(header[1]);

  // 空シートの場合はヘッダーを初期化
  if (!a1 && !b1 && sheet.getLastRow() === 0) {
    initializeTradePartnerSheet_(sheet);
    return;
  }

  // 旧フォーマット（A:登録番号 / B:取引先名）からの移行
  if (a1 === '登録番号' && b1 === '取引先名') {
    const needColumns = MF_TRADE_PARTNER_HEADERS.length;
    const currentColumns = Math.max(lastColumn || 0, 2);
    if (currentColumns < needColumns) {
      sheet.insertColumnsAfter(currentColumns, needColumns - currentColumns);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const invoiceValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      sheet.getRange(2, 5, lastRow - 1, 1).setValues(invoiceValues);
      sheet.getRange(2, 1, lastRow - 1, 1).clearContent();
    }

    initializeTradePartnerSheet_(sheet);
    if (ui) {
      ui.alert(
        '「取引先一覧」シートを旧フォーマット（登録番号/取引先名の2列）から移行しました。\n' +
        '登録番号はE列、取引先名はB列を使用してください。'
      );
    }
    return;
  }

  // 既にマネフォ仕様のヘッダーなら、見た目だけ整える
  const isMfHeader =
    normalizeText_(header[0]) === MF_TRADE_PARTNER_HEADERS[0] &&
    normalizeText_(header[1]) === MF_TRADE_PARTNER_HEADERS[1] &&
    normalizeText_(header[2]) === MF_TRADE_PARTNER_HEADERS[2] &&
    normalizeText_(header[3]) === MF_TRADE_PARTNER_HEADERS[3] &&
    normalizeText_(header[4]) === MF_TRADE_PARTNER_HEADERS[4] &&
    normalizeText_(header[5]) === MF_TRADE_PARTNER_HEADERS[5];
  if (isMfHeader) {
    initializeTradePartnerSheet_(sheet);
    return;
  }
}

function buildTradePartnerFilename_() {
  const now = new Date();
  const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  return `mf_trade_partners_${stamp}.csv`;
}

function buildTradePartnersCsvPayload_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MF_TRADE_PARTNER_SHEET_NAME);
  if (!sheet) {
    throw new Error('「取引先一覧」シートが見つかりません。先に「4. 取引先一覧更新」を実行してください。');
  }

  ensureTradePartnerSheetLayout_(sheet, null);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    throw new Error('「取引先一覧」シートにデータがありません。');
  }

  const headers = MF_TRADE_PARTNER_HEADERS;
  const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  if (rows.length === 0) {
    throw new Error('「取引先一覧」シートにデータがありません。');
  }

  const missingNames = rows.filter((row) => isBlankCell_(row[1])).length;
  return {
    filename: buildTradePartnerFilename_(),
    csvContent: buildCsvContent_(headers, rows),
    missingNames: missingNames
  };
}

function downloadTradePartnersCsv() {
  const ui = SpreadsheetApp.getUi();
  try {
    const payload = buildTradePartnersCsvPayload_();
    if (payload.missingNames > 0) {
      ui.alert(
        `取引先名が空欄の行が ${payload.missingNames} 件あります。\n` +
        'マネーフォワードの仕様では「取引先名」は必須です。\n' +
        '空欄のままでも全行をCSV出力します。'
      );
    }
    showDownloadDialog_(payload.filename, payload.csvContent);
  } catch (e) {
    ui.alert(String(e && e.message ? e.message : e));
  }
}

function updateTradePartnersFromAnalysis() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analysisSheet = getOrCreateAnalysisSheet_();
  ensureResultSheetLayout_(analysisSheet);

  const lastRow = analysisSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('解析シートにデータがありません。');
    return;
  }

  const header = analysisSheet
    .getRange(1, 1, 1, Math.max(RESULT_SHEET_HEADERS.length, analysisSheet.getLastColumn()))
    .getValues()[0];
  const vendorIndex = header.indexOf('取引先');
  const invoiceIndex = header.indexOf('インボイス番号');

  if (vendorIndex === -1 || invoiceIndex === -1) {
    ui.alert('解析シートに「取引先」「インボイス番号」の列が見つかりません。');
    return;
  }

  let partnerSheet = ss.getSheetByName(MF_TRADE_PARTNER_SHEET_NAME);
  if (!partnerSheet) {
    partnerSheet = ss.insertSheet(MF_TRADE_PARTNER_SHEET_NAME);
    initializeTradePartnerSheet_(partnerSheet);
  }
  ensureTradePartnerSheetLayout_(partnerSheet, ui);

  const existingInvoices = getTradePartnerInvoiceSet_(partnerSheet);
  const seenInvoices = new Set();
  const toAppend = [];
  let skippedBlankInvoice = 0;
  let skippedBlankName = 0;
  let skippedExisting = 0;
  let skippedDuplicate = 0;

  const values = analysisSheet.getRange(2, 1, lastRow - 1, header.length).getValues();
  for (const row of values) {
    const invoiceNumber = normalizeInvoiceNumber_(row[invoiceIndex]);
    if (!invoiceNumber) {
      skippedBlankInvoice++;
      continue;
    }
    if (seenInvoices.has(invoiceNumber)) {
      skippedDuplicate++;
      continue;
    }
    seenInvoices.add(invoiceNumber);

    if (existingInvoices.has(invoiceNumber)) {
      skippedExisting++;
      continue;
    }

    const partnerName = normalizeText_(row[vendorIndex]);
    if (!partnerName) {
      skippedBlankName++;
      continue;
    }

    toAppend.push(['', partnerName, '', '', invoiceNumber, '']);
    existingInvoices.add(invoiceNumber);
  }

  if (toAppend.length > 0) {
    const startRow = Math.max(partnerSheet.getLastRow() + 1, 2);
    partnerSheet.getRange(startRow, 1, toAppend.length, MF_TRADE_PARTNER_HEADERS.length).setValues(toAppend);
  }

  ui.alert(
    '取引先一覧を更新しました。\n' +
    `追加: ${toAppend.length}\n` +
    `既存（登録番号重複のためスキップ）: ${skippedExisting}\n` +
    `登録番号なし（スキップ）: ${skippedBlankInvoice}\n` +
    `取引先名なし（スキップ）: ${skippedBlankName}\n` +
    `解析シート内の重複（スキップ）: ${skippedDuplicate}`
  );
}

function getTradePartnerInvoiceSet_(sheet) {
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const invoiceIndex = header.indexOf('登録番号');
  const lastRow = sheet.getLastRow();
  const set = new Set();
  if (invoiceIndex === -1 || lastRow < 2) return set;

  const values = sheet.getRange(2, invoiceIndex + 1, lastRow - 1, 1).getValues();
  for (const [value] of values) {
    const invoiceNumber = normalizeInvoiceNumber_(value);
    if (invoiceNumber) set.add(invoiceNumber);
  }
  return set;
}

function resolvePartnerName_(invoiceNumber, partnerMap, vendorName) {
  if (invoiceNumber && partnerMap[invoiceNumber]) return partnerMap[invoiceNumber];
  if (vendorName) return vendorName;
  return '';
}

function buildMoneyForwardRow_(transactionNo, data, partnerName, fileUrl, settings) {
  const memoUrl = fileUrl || '';
  const amount = data.amount || 0;
  const csvDate = normalizeDate_(data.date);
  const taxCategory = resolveMoneyForwardTaxCategory_(data.taxCategory);
  const creditAccount =
    data.paymentMethod === 'クレカ'
      ? (settings?.creditAccountCard || '未払金')
      : (settings?.creditAccountOther || '役員借入金');
  const creditTaxCategory = '対象外';
  const creditSubAccount =
    data.paymentMethod === 'クレカ'
      ? resolveCreditSubAccount_(settings, data)
      : '';
  const summary = buildMoneyForwardSummary_(settings, data.summary, partnerName);

  return [
    transactionNo,
    csvDate,
    data.accountTitle,
    '',
    '',
    partnerName,
    taxCategory,
    data.invoiceNumber,
    amount,
    0,
    creditAccount,
    creditSubAccount,
    '',
    '',
    creditTaxCategory,
    '',
    amount,
    0,
    summary,
    memoUrl,
    '',
    '',
    '',
    '',
    '',
    '',
    ''
  ];
}

function buildMoneyForwardSummary_(settings, summary, partnerName) {
  const summaryText = normalizeText_(summary);
  const partnerText = normalizeText_(partnerName);
  const modeRaw = normalizeText_(settings?.mfSummaryMode);
  const mode = MF_SUMMARY_MODE_CANDIDATES.includes(modeRaw)
    ? modeRaw
    : MF_SUMMARY_MODE_CANDIDATES[0];

  if (mode === '取引先名') return partnerText;
  if (mode === '取引先名＋購入内容') {
    if (partnerText && summaryText) {
      if (partnerText === summaryText) return partnerText;
      return `${partnerText} ${summaryText}`;
    }
    return partnerText || summaryText;
  }

  // 既存実装（購入内容優先。空なら取引先名）
  return summaryText || partnerText;
}

function applyMoneyForwardTaxCategoryHighlights_(sheet, mixedTaxFlags) {
  if (!sheet || !mixedTaxFlags || mixedTaxFlags.length === 0) return;
  const taxCategoryCol = MF_CSV_HEADERS.indexOf('借方税区分') + 1;
  if (taxCategoryCol <= 0) return;

  const backgrounds = mixedTaxFlags.map((isMixed) => [isMixed ? MIXED_TAX_CELL_COLOR : null]);
  sheet.getRange(2, taxCategoryCol, mixedTaxFlags.length, 1).setBackgrounds(backgrounds);
}

function resolveCreditSubAccount_(settings, data) {
  const mode = normalizeText_(settings?.creditSubAccountCard) || CREDIT_SUB_ACCOUNT_CANDIDATES[0];
  if (mode === '空欄') return '';
  return normalizeCardInfo_(data.cardInfo, data.paymentMethod);
}

function buildMoneyForwardFilename_() {
  const now = new Date();
  const stamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  return `mf_journal_${stamp}.csv`;
}

function formatDate_(dateValue) {
  return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function buildCsvContent_(headers, rows) {
  const lines = [];
  lines.push(headers.map(escapeCsvValue_).join(','));
  for (const row of rows) {
    lines.push(row.map(escapeCsvValue_).join(','));
  }
  return `\uFEFF${lines.join('\n')}`;
}

function escapeCsvValue_(value) {
  if (value === null || value === undefined) return '';
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return formatDate_(value);
  }
  const text = String(value);
  if (text.includes('"')) {
    const escaped = text.replace(/"/g, '""');
    return `"${escaped}"`;
  }
  if (text.includes(',') || text.includes('\n')) {
    return `"${text}"`;
  }
  return text;
}

function showDownloadDialog_(filename, csvContent) {
  const html = `
    <html>
      <body>
        <script>
          const csv = ${JSON.stringify(csvContent)};
          const filename = ${JSON.stringify(filename)};
          const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
          const url = URL.createObjectURL(blob);
          const link = document.createElement('a');
          link.href = url;
          link.download = filename;
          document.body.appendChild(link);
          link.click();
          setTimeout(() => {
            URL.revokeObjectURL(url);
            google.script.host.close();
          }, 100);
        </script>
      </body>
    </html>
  `;
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(200).setHeight(80),
    'CSVをダウンロード'
  );
}

// ==================================================
// Gemini API 呼び出し関数（インボイス・iD・PDF・HEIC対応）
// ==================================================
function callGeminiApi(base64Data, mimeType, apiKey) {
  const modelName = getModelName_();
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${modelName}:generateContent?key=${apiKey}`;
  const taxAnalysis = callGeminiTaxCategoryApi_(endpoint, base64Data, mimeType);
  const taxHintSection = buildTaxCategoryHintSection_(taxAnalysis);

  const prompt = `
このファイル（画像またはPDF）はレシート／領収書／請求書です。
内容から必要情報を抽出し、JSONオブジェクト1個だけを返してください。

返答はJSONのみ（前後の説明、コードフェンス、Markdown、箇条書き、コメントは禁止）。
キー順は以下のスキーマ例と完全一致させること（順序違いは不可）。
余計なキーは追加しない。
文字列は必ずダブルクォート。
数値は必ず数値型（"123"ではなく 123）。
不明な項目は空文字 "" または 0 を入れる（null / undefined は使用禁止）。

出力スキーマ（キー順固定）:
{"paymentDate":"YYYY/MM/DD","paymentMethod":"現金|クレカ|PayPay|電子マネー|銀行振込","cardInfo":"カード(1234)","vendorName":"取引先名","invoiceNumber":"T1234567890123","summary":"品目（概要）","taxCategory":"10%|8%|混在あり","amount":12345}

────────────────
抽出ルール
────────────────

1) paymentDate（支払日）
- レシート/領収書: 「日付」「取引日」「発行日」「購入日」等のうち最も妥当な日付。
- 請求書: 支払日が明確ならそれ、無ければ発行日。支払期限/入金期限は使わない。
- 年が読み取れない場合は、現在に近い年を推定。
- 「25-06-03」のような「年-月-日」形式や2桁の年は、和暦ではなく西暦の下2桁（20XX年）として優先的に解釈する。
- 解釈結果が著しく未来の日付（例：令和25年＝2043年など）になる場合は、和暦ではなく西暦とみなす。
- YYYY/MM/DD 形式に整形できない場合は ""。

2) paymentMethod（支払方法）
- 候補は必ず次の5つから1つだけ:
  「現金」「クレカ」「PayPay」「電子マネー」「銀行振込」

- 判定前に、画像から読めた文字列を正規化して扱う（出力には出さない）:
  - 全角/半角の統一（英数・記号・スペース）
  - 半角カナの統一（例: ｸﾚｼﾞｯﾄ→クレジット、ﾃﾞﾝｼﾏﾈｰ→電子マネー）
  - 連続スペースや改行の正規化

- 支払方法は、以下の優先順位で確定する（上から順に最初に成立したもの）:

- 「PayPay」が「お支払い」「お支払方法」「決済」「支払」等の支払欄付近にあれば「PayPay」

- 「振込」「銀行振込」「口座」「振込先」「お振込」等があり、支払手段として読める場合は「銀行振込」

- 次のいずれかが支払欄付近にあれば「クレカ」:
  - 「クレジット」「ｸﾚｼﾞｯﾄ」「ｸﾚｼﾞﾂﾄ」「CREDIT」
  - 「カード」「ｶｰﾄﾞ」
  - 「VISA」「MASTER」「MASTERCARD」「JCB」「AMEX」「AMERICAN EXPRESS」「DINERS」
  - 「iD」「QUICPay」「QUIC PAY」
  ※iD / QUICPay は「電子マネー」ではなく必ず「クレカ」扱いで固定

- ただし「ID（大文字）」のみは会員ID/伝票IDの可能性があるため、
  支払欄付近で見つかった場合のみ iD とみなす（それ以外の場所の ID は支払方法判定に使わない）

- 次のいずれかが支払欄付近にあれば「電子マネー」:
  - 「電子マネー」「交通系IC」
  - 「Suica」「PASMO」「ICOCA」「TOICA」「manaca」「はやかけん」「nimoca」「SUGOCA」
  - 「楽天Edy」「Edy」「WAON」「nanaco」
  ※「電子マネー」表記があっても、同時にクレカ条件（特に iD/QUICPay/クレジット/カード/国際ブランド）が成立する場合は「クレカ」を優先

- 「現計」「お預り」「お釣り」「釣銭」「おつり」等が支払欄付近にあり、
  かつ上記の「PayPay」「銀行振込」「クレカ」「電子マネー」のいずれも成立しない場合は「現金」

- 判別できない場合は「現金」

2-2) cardInfo（カード情報）
- paymentMethod が「クレカ」の場合のみ設定
- カード番号の下4桁を抽出し「カード(1234)」形式
- 抽出できない場合は「カード(不明)」
- クレカ以外の場合は "" とする

3) vendorName（取引先名）
- 店名 / 会社名 / 発行者名 / 請求元名から最も適切な名称を短く抽出
- 住所・電話番号・FAX・登録番号は含めない
- 「株式会社 / 有限会社 / 合同会社」等を含む正式名称が読める場合はそれを優先
- 店舗名のみの場合は店舗名で可
- 不明なら ""

4) invoiceNumber（登録番号）
- 「登録番号」「適格請求書発行事業者登録番号」「インボイス番号」に続く文字列から抽出
- 形式は T + 13桁
- 全角T / 全角数字 / 空白混入は正規化して半角にする
- 伝票番号・顧客番号等は入れない
- 不明なら ""

5) summary（概要）
- 15文字程度までを目安に簡潔に
- レシート: 主な購入内容または用途カテゴリを優先
- 請求書: 請求内容の要約を優先
- vendorName と同一文字列のみになるのは避ける
- 内容が取れない場合のみ vendorName を使ってよい
- 不明なら ""

6) taxCategory（消費税区分）

【最優先ルール】
- 税率の文字（8% / 10%）の出現だけでは判定しない。
- 「内消費税(8%)」「内消費税(10%)」「内消費税等」「消費税」「消費税額」「内税」「外税」「税額」等の“税額”を根拠に判定する。
- 注意書き・凡例（例：「★印は軽減税率(8%)適用の商品です」）は判定に使わない。
- 「8%対象」「10%対象」「外税(10%対象)」等の税率別集計ブロックは根拠に使う（同一行または前後2行以内の税額を同一ブロックとして扱う）。

【内部的に行う抽出（出力しない）】
- tax8Amount：8%の税額
- tax10Amount：10%の税額
- 「¥0」「￥0」「0」「0円」「消費税額:¥0」等は必ず 0 として扱う
- 税額行が存在しない場合は「税額が見つからない」と判断する（0とは別）

【判定】
- tax8Amount > 0 かつ tax10Amount > 0 → 「混在あり」
- tax8Amount == 0 かつ tax10Amount > 0 → 「10%」
- tax8Amount > 0 かつ tax10Amount == 0 → 「8%」
- tax8Amount == 0 かつ tax10Amount == 0 → 「10%」
- 税額が見つからない場合:
  - 8%対象と10%対象の両方がある → 「混在あり」
  - 8%対象のみある → 「8%」
  - 10%対象のみある → 「10%」
  - 8%と10%の両方の税率表記がある → 「混在あり」
  - 10%のみ明確 → 「10%」
  - 8%のみ明確 → 「8%」
  - それ以外 → 「10%」

${taxHintSection}

7) amount（税込合計）
- 支払総額（税込）の整数
- 小数点以下は切り捨て（四捨五入しない）
- 「合計」「総計」「お支払金額」「ご請求金額」等を優先
- 数値は必ず整数型で出力
- 不明なら 0

出力例（JSONのみ）:
{"paymentDate":"2026/01/18","paymentMethod":"クレカ","cardInfo":"カード(2235)","vendorName":"ENEOS","invoiceNumber":"T1234567890123","summary":"ガソリン代","taxCategory":"10%","amount":4500}
  `;

  const response = callGeminiWithPrompt_(endpoint, prompt, base64Data, mimeType, 'レシート抽出');
  if (!response) return null;

  const normalized = normalizeReceiptExtraction_(response.parsed, taxAnalysis);
  if (normalized && taxAnalysis && taxAnalysis.taxCategory) {
    normalized.taxCategory = normalizeTaxCategory_(taxAnalysis.taxCategory, taxAnalysis);
  }
  if (normalized && !normalized.invoiceNumber) {
    normalized.invoiceNumber = extractInvoiceNumberFromText_(response.text);
  }
  return normalized;
}

function callGeminiTaxCategoryApi_(endpoint, base64Data, mimeType) {
  const prompt = buildTaxCategoryPrompt_();
  const response = callGeminiWithPrompt_(endpoint, prompt, base64Data, mimeType, '税率判定');
  if (!response || !response.parsed) return null;
  return normalizeTaxAnalysisResult_(response.parsed);
}

function callGeminiWithPrompt_(endpoint, prompt, base64Data, mimeType, contextLabel) {
  const payload = {
    "contents": [{
      "parts": [
        { "text": prompt },
        { "inline_data": { "mime_type": mimeType, "data": base64Data } }
      ]
    }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(endpoint, options);
    const json = JSON.parse(response.getContentText());
    if (json.error) {
      Logger.log(`${contextLabel} API Error: ${JSON.stringify(json.error)}`);
      return null;
    }
    const text = json?.candidates?.[0]?.content?.parts?.[0]?.text;
    if (!text) return null;
    return {
      text: text,
      parsed: extractJsonFromText_(text)
    };
  } catch (e) {
    Logger.log(`${contextLabel} API Error: ${e.message}`);
    return null;
  }
}

function buildTaxCategoryPrompt_() {
  return `
このファイル（画像またはPDF）はレシート／領収書／請求書です。
税率判定専用タスクとして、消費税区分だけを抽出してください。
返答はJSONのみ。前後の説明は禁止。キー順は固定。余計なキーは禁止。

出力スキーマ（キー順固定）:
{"taxCategory":"10%|8%|混在あり","tax8Amount":0,"tax10Amount":0,"tax8Base":0,"tax10Base":0,"evidence":"","confidence":0}

判定ルール:
- 税率文字（8%/10%）だけでは判定しない。税額根拠を優先する。
- 税額候補: 「内消費税(8%)」「内消費税(10%)」「内消費税等」「消費税」「消費税額」「内税」「外税」「税額」。
- 税率別集計ブロック候補: 「8%対象」「10%対象」「軽」「外税(10%対象)」。
- 税率別ブロックは、同一行または前後2行以内を同一ブロックとして読み、税額・対象額を対応付ける。
- 注意書き・凡例（例: 「★印は軽減税率(8%)適用の商品です」）は根拠に使わない。
- tax8Amount/tax10Amount は税額。見つからない場合は 0。
- tax8Base/tax10Base は税率別の課税対象額。見つからない場合は 0。
- taxCategory 判定:
  - tax8Amount > 0 かつ tax10Amount > 0 → 混在あり
  - tax8Amount == 0 かつ tax10Amount > 0 → 10%
  - tax8Amount > 0 かつ tax10Amount == 0 → 8%
  - 税額が両方 0 の場合:
    - 8%対象と10%対象の両方がある → 混在あり
    - 8%対象のみある → 8%
    - 10%対象のみある → 10%
    - 8%と10%の両方表記のみある → 混在あり
    - 8%のみ表記がある → 8%
    - それ以外 → 10%
- evidence には、実際に根拠として使った行を最大2行だけ短く入れる。
- confidence は 0〜1 の数値。
`;
}

function normalizeTaxAnalysisResult_(data) {
  if (!data || typeof data !== 'object') return null;

  const hasTax8Amount = hasTaxFieldValue_(data, 'tax8Amount');
  const hasTax10Amount = hasTaxFieldValue_(data, 'tax10Amount');
  const tax8Amount = hasTax8Amount ? normalizeAmount_(data.tax8Amount) : 0;
  const tax10Amount = hasTax10Amount ? normalizeAmount_(data.tax10Amount) : 0;

  const base8Found = hasTaxFieldValue_(data, 'tax8Base');
  const base10Found = hasTaxFieldValue_(data, 'tax10Base');
  const tax8Base = base8Found ? normalizeAmount_(data.tax8Base) : 0;
  const tax10Base = base10Found ? normalizeAmount_(data.tax10Base) : 0;
  const evidence = normalizeText_(data.evidence);

  const normalized = {
    tax8Amount: tax8Amount,
    tax10Amount: tax10Amount,
    tax8Base: tax8Base,
    tax10Base: tax10Base,
    hasTax8Amount: hasTax8Amount,
    hasTax10Amount: hasTax10Amount,
    evidence: evidence,
    confidence: normalizeTaxConfidence_(data.confidence)
  };
  normalized.taxCategory = normalizeTaxCategory_(data.taxCategory, normalized);
  return normalized;
}

function hasTaxFieldValue_(data, key) {
  if (!data || typeof data !== 'object') return false;
  if (!Object.prototype.hasOwnProperty.call(data, key)) return false;
  return data[key] !== '' && data[key] !== null && data[key] !== undefined;
}

function normalizeTaxConfidence_(value) {
  if (value === '' || value === null || value === undefined) return 0;
  const numeric = Number(String(value).replace(/[^0-9.]/g, ''));
  if (!isFinite(numeric)) return 0;
  const scaled = numeric > 1 ? numeric / 100 : numeric;
  const clamped = Math.max(0, Math.min(1, scaled));
  return Math.round(clamped * 100) / 100;
}

function buildTaxCategoryHintSection_(taxAnalysis) {
  if (!taxAnalysis) {
    return `
【taxCategory追加ルール】
- 先行税率判定結果が取得できない場合は、上記ルールのみで判定する。`;
  }

  const evidence = taxAnalysis.evidence || '不明';
  const tax8Amount = taxAnalysis.hasTax8Amount ? taxAnalysis.tax8Amount : '未検出';
  const tax10Amount = taxAnalysis.hasTax10Amount ? taxAnalysis.tax10Amount : '未検出';

  return `
【taxCategory追加ルール（先行税率判定を最優先）】
- 先行判定結果:
  - taxCategory: ${taxAnalysis.taxCategory}
  - tax8Amount: ${tax8Amount}
  - tax10Amount: ${tax10Amount}
  - tax8Base: ${taxAnalysis.tax8Base}
  - tax10Base: ${taxAnalysis.tax10Base}
  - evidence: ${evidence}
  - confidence: ${taxAnalysis.confidence}
- taxCategory は先行判定結果を最優先で採用する。
- 先行判定が不明な場合のみ通常ルールで最終判定する。`;
}

function normalizeReceiptExtraction_(data, taxHints) {
  if (!data) return null;
  const amountRaw = data.amount;
  const hasAmount =
    amountRaw !== undefined && amountRaw !== null && String(amountRaw).trim() !== '';
  return {
    paymentDate: normalizeDate_(data.paymentDate || data.date),
    paymentMethod: normalizePaymentMethod_(data.paymentMethod),
    cardInfo: normalizeCardInfo_(data.cardInfo, data.paymentMethod),
    vendorName: normalizeText_(data.vendorName),
    invoiceNumber: normalizeInvoiceNumber_(data.invoiceNumber),
    summary: normalizeText_(data.summary),
    taxCategory: normalizeTaxCategory_(data.taxCategory, taxHints),
    amount: hasAmount ? normalizeAmount_(amountRaw) : 0
  };
}

function extractInvoiceNumberFromText_(text) {
  if (!text) return '';
  const normalized = String(text)
    .replace(/[Ｔｔ]/g, 'T')
    .replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
    .replace(/\s+/g, '');
  const match = normalized.match(/T\d{13}/);
  return match ? match[0] : '';
}

// ==================================================
// 画像プレビューサイドバー
// ==================================================
function showImageSidebar() {
  const html = HtmlService.createHtmlOutput(buildImageSidebarHtml_())
    .setTitle('画像プレビュー');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showImagePreviewDialog() {
  const initialInfo = getSelectedRowPreviewInfo();
  const html = HtmlService.createHtmlOutput(buildImagePreviewDialogHtml_(initialInfo))
    .setWidth(900)
    .setHeight(650);
  SpreadsheetApp.getUi().showModelessDialog(html, '画像プレビュー（拡大）');
}

function showHelp() {
  const html = HtmlService.createHtmlOutput(buildHelpDialogHtml_())
    .setTitle('ヘルプ')
    .setWidth(1000)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'ヘルプ');
}

function buildHelpDialogHtml_() {
  const helpUrl = HELP_URL;
  return `
<!DOCTYPE html>
<html>
  <head>
    <meta charset="utf-8">
    <style>
      html,
      body {
        font-family: "Noto Sans JP", Arial, sans-serif;
        height: 100%;
        margin: 0;
      }
      .header {
        background: #f8f9fa;
        border-bottom: 1px solid #e0e0e0;
        padding: 12px;
      }
      .title {
        color: #555;
        font-size: 12px;
        margin-bottom: 6px;
      }
      .link {
        font-size: 13px;
        text-decoration: none;
      }
      .content {
        height: calc(100% - 64px);
      }
      iframe {
        border: 0;
        height: 100%;
        width: 100%;
      }
    </style>
  </head>
  <body>
    <div class="header">
      <div class="title">表示されない場合はこちらから開いてください。</div>
      <a class="link" href="${helpUrl}" target="_blank" rel="noopener">ヘルプを別タブで開く</a>
    </div>
    <div class="content">
      <iframe src="${helpUrl}" referrerpolicy="no-referrer"></iframe>
    </div>
  </body>
</html>
  `;
}

function getSelectedRowPreviewInfo() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getActiveRange();
  if (!range) return { message: '行を選択してください。' };

  const row = range.getRow();
  if (row < 2) return { message: 'データ行を選択してください。' };

  const fileId = resolvePreviewFileId_(sheet, row);
  if (!fileId) return { message: 'ファイルIDがありません。' };

  try {
    const file = DriveApp.getFileById(fileId);
    return {
      fileId: fileId,
      fileName: file.getName(),
      mimeType: file.getMimeType(),
      previewUrl: `https://drive.google.com/file/d/${fileId}/preview`,
      thumbnailUrl: `https://drive.google.com/thumbnail?id=${fileId}`,
      fileUrl: file.getUrl()
    };
  } catch (e) {
    return { message: `ファイルを取得できません: ${e.message}` };
  }
}

function resolvePreviewFileId_(sheet, row) {
  const sheetName = sheet.getName();
  if (sheetName === ANALYSIS_SHEET_NAME) {
    return sheet.getRange(row, 1).getValue();
  }
  if (sheetName === 'マネフォ用') {
    const memoCol = MF_CSV_HEADERS.indexOf('仕訳メモ') + 1;
    if (memoCol <= 0) return '';
    const memoValue = sheet.getRange(row, memoCol).getValue();
    return extractDriveFileIdFromUrl_(memoValue);
  }
  return '';
}

function buildImageSidebarHtml_() {
  return `
    <html>
      <head>
        <style>
          body { font-family: Arial, sans-serif; font-size: 12px; }
          .title { font-weight: bold; margin-bottom: 6px; }
          .meta { color: #666; margin-bottom: 8px; }
          .actions { display: flex; gap: 8px; margin: 8px 0; }
          .preview { width: 100%; height: 320px; border: 1px solid #ddd; }
          .thumb { width: 100%; border: 1px solid #ddd; margin-top: 8px; }
          .note { color: #888; margin-top: 6px; font-size: 11px; }
        </style>
      </head>
      <body>
        <div class="title">画像プレビュー</div>
        <div id="meta" class="meta">行を選択してください。</div>
        <div class="actions">
          <button onclick="refresh()">更新</button>
          <button id="openPopup" onclick="openPopup()" disabled>拡大表示</button>
        </div>
        <iframe id="preview" class="preview" src=""></iframe>
        <img id="thumb" class="thumb" src="" />
        <div class="note">プレビューが表示されない場合はサムネイルを確認してください。</div>
        <script>
          let lastFileId = '';
          function openPopup() {
            google.script.run.showImagePreviewDialog();
          }
          function refresh() {
            google.script.run.withSuccessHandler(render).getSelectedRowPreviewInfo();
          }
          function render(data) {
            if (!data) return;
            const button = document.getElementById('openPopup');
            if (data.message) {
              lastFileId = '';
              document.getElementById('meta').textContent = data.message;
              document.getElementById('preview').src = '';
              document.getElementById('thumb').src = '';
              button.disabled = true;
              return;
            }
            if (data.fileId && data.fileId === lastFileId) return;
            lastFileId = data.fileId || '';
            document.getElementById('meta').textContent = data.fileName || '';
            document.getElementById('preview').src = data.previewUrl || '';
            document.getElementById('thumb').src = data.thumbnailUrl || '';
            button.disabled = !Boolean(data.previewUrl);
          }
          refresh();
          setInterval(refresh, 3000);
        </script>
      </body>
    </html>
  `;
}

function escapeJsonForInlineScript_(value) {
  return JSON.stringify(value || {})
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');
}

function buildImagePreviewDialogHtml_(initialInfo) {
  const initialDataJson = escapeJsonForInlineScript_(initialInfo);
  return `
    <html>
      <head>
        <style>
          html, body { height: 100%; margin: 0; font-family: Arial, sans-serif; font-size: 12px; }
          .root { display: flex; flex-direction: column; height: 100%; padding: 12px; box-sizing: border-box; }
          .meta { color: #666; margin-bottom: 8px; min-height: 18px; }
          .actions { display: flex; gap: 8px; margin-bottom: 8px; }
          .preview { flex: 1; width: 100%; border: 1px solid #ddd; min-height: 320px; }
          .thumb { width: 100%; border: 1px solid #ddd; margin-top: 8px; max-height: 140px; object-fit: contain; }
          .note { color: #888; margin-top: 6px; font-size: 11px; }
        </style>
      </head>
      <body>
        <div class="root">
          <div id="meta" class="meta">行を選択してください。</div>
          <div class="actions">
            <button onclick="refresh()">更新</button>
            <button onclick="closeDialog()">閉じる</button>
          </div>
          <iframe id="preview" class="preview" src=""></iframe>
          <img id="thumb" class="thumb" src="" />
          <div class="note">選択行が変わると、3秒ごとにプレビューが自動更新されます。</div>
        </div>
        <script>
          const initialData = ${initialDataJson};
          let lastFileId = '';
          function closeDialog() {
            google.script.host.close();
          }
          function refresh() {
            google.script.run.withSuccessHandler(render).getSelectedRowPreviewInfo();
          }
          function render(data) {
            if (!data) return;
            if (data.message) {
              lastFileId = '';
              document.getElementById('meta').textContent = data.message;
              document.getElementById('preview').src = '';
              document.getElementById('thumb').src = '';
              return;
            }
            if (data.fileId && data.fileId === lastFileId) return;
            lastFileId = data.fileId || '';
            document.getElementById('meta').textContent = data.fileName || '';
            document.getElementById('preview').src = data.previewUrl || '';
            document.getElementById('thumb').src = data.thumbnailUrl || '';
          }
          render(initialData);
          setInterval(refresh, 3000);
        </script>
      </body>
    </html>
  `;
}
