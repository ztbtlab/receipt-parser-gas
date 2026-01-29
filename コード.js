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

const SETTINGS_KEYS = {
  folderUrl: '対象フォルダURL',
  delimiter: '区切り文字',
  fileNameRuleSheet: 'ファイル名ルール参照',
  scanLimit: '解析上限',
  logOutput: 'ログ出力',
  creditAccountCard: '貸方勘定科目(クレカ)',
  creditAccountOther: '貸方勘定科目(それ以外)',
  creditSubAccountCard: '貸方補助科目(クレカ)'
};

const DELIMITER_CANDIDATES = ['-', '_', '｜'];
const LOG_OUTPUT_CANDIDATES = ['ON', 'OFF'];
const CREDIT_SUB_ACCOUNT_CANDIDATES = ['カード情報', '空欄'];

const EMPTY_CELL_COLOR = '#fff2cc';
const DUPLICATE_CELL_COLOR = '#f4cccc';

const MF_TAX_CATEGORY = '課税仕入 10%';
const MF_CARD_SUB_ACCOUNT = '三井住友ゴールドカード';
const MF_PROCESSED_PREFIX = 'CSV済';
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
  '解析メモ'
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
    .addItem('4. マネフォ用解析', 'analyzeMoneyForward')
    .addItem('5. マネフォCSVダウンロード', 'downloadMoneyForwardCsv')
    .addSeparator()
    .addItem('画像プレビューを開く', 'showImageSidebar')
    .addSeparator()
    .addSubMenu(settingsMenu)
    .addToUi();
}

// ==================================================
// 解析結果シートの列構成を整える
// A:ファイルID, B:リンク, C:元ファイル名, D:変更案, E:移動先, F:ステータス,
// G:支払日, H:支払い方法, I:取引先, J:インボイス番号, K:品目（概要）, L:金額, M:解析メモ
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
      '解析メモ'
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
    { key: SETTINGS_KEYS.creditSubAccountCard, defaultValue: CREDIT_SUB_ACCOUNT_CANDIDATES[0] }
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
    creditSubAccountCard: creditSubAccountCard
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
    { key: SETTINGS_KEYS.logOutput, defaultValue: 'ON' }
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
        memo
      ]);

      const rowIndex = sheet.getLastRow();
      sheet.setRowHeight(rowIndex, 30);

      applyDestinationValidation_(sheet, rowIndex);
      highlightEmptyExtractionCells_(sheet, rowIndex);
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
      sheet.appendRow([id, '', fileName, '', '', 'エラー: ' + e.toString(), '', '', '', '', '', '', '', '']);
      logError_('レシート解析', id, fileName, e.message);
    }
  }

  const limitReached = targetFiles.length > processCount;
  const messages = [];
  if (processCount === 0) {
    messages.push('新しいファイルは見つかりませんでした。');
  } else {
    messages.push(`${processCount} 件のファイルをスキャンしました。`);
  }
  if (duplicateCount > 0) messages.push(`重複候補: ${duplicateCount} 件（提案名に連番を付与済み）`);
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

  const range = sheet.getRange(2, 1, lastRow - 1, RESULT_SHEET_HEADERS.length);
  range.clearContent();
  range.setBackground(null);
  ui.alert('解析シートのデータをクリアしました。');
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
  const errors = [];
  let transactionNo = 1;

  const data = sheet.getRange(2, 1, lastRow - 1, RESULT_SHEET_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const fileId = row[0];
    if (!fileId) continue;

    try {
      const file = DriveApp.getFileById(fileId);
      const fileName = file.getName();
      if (isCsvMarkedFile_(fileName)) continue;

      const paymentDate = normalizeDate_(row[6]) || formatDate_(file.getDateCreated());
      const paymentMethod = normalizePaymentMethod_(row[7]);
      const cardInfo = normalizeCardInfo_(row[8], paymentMethod);
      const vendorName = normalizeText_(row[9]);
      const invoiceNumber = normalizeInvoiceNumber_(row[10]);
      const summary = normalizeText_(row[11]);
      const amount = normalizeAmount_(row[12]);

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
        : '未入力';

      const partnerName = resolvePartnerName_(invoiceNumber, partnerMap, vendorName);
      const merged = {
        date: paymentDate,
        amount: amount,
        invoiceNumber: invoiceNumber,
        vendorName: vendorName,
        summary: summary,
        paymentMethod: paymentMethod,
        cardInfo: cardInfo,
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
  const messages = [`${rows.length} 件のレシートを「マネフォ用」シートに更新しました。`];
  if (errors.length > 0) {
    messages.push(`エラー: ${errors.length} 件`);
  }
  ui.alert(messages.join('\n'));
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
}

function downloadMoneyForwardCsv() {
  const ui = SpreadsheetApp.getUi();
  const sheet = getMoneyForwardSheet_(ui);
  const filter = sheet.getFilter();
  if (filter) filter.remove();

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('「マネフォ用」シートにデータがありません。');
    return;
  }

  const range = sheet.getRange(2, 1, lastRow - 1, MF_CSV_HEADERS.length);
  const values = range.getValues();
  const rows = values.filter((row) => row.some((cell) => !isBlankCell_(cell)));

  if (rows.length === 0) {
    ui.alert('「マネフォ用」シートに出力対象の行がありません。');
    return;
  }

  const csvContent = buildCsvContent_(MF_CSV_HEADERS, rows);
  const filename = buildMoneyForwardFilename_();
  showDownloadDialog_(filename, csvContent);

  const markedCount = markCsvProcessedFiles_(rows);
  const messages = ['CSVをダウンロードしました。'];
  if (markedCount > 0) {
    messages.push(`${markedCount} 件のファイルに「CSV済」を付与しました。`);
  }
  ui.alert(messages.join('\n'));
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
    accountTitle: normalizeAccountTitle_(data.accountTitle)
  };
}

function normalizePaymentMethod_(value) {
  const text = normalizeText_(value);
  if (!text) return '不明';

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
  if (text.includes('不明')) return '不明';
  return '不明';
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
  let sheet = ss.getSheetByName('取引先一覧');

  if (!sheet) {
    sheet = ss.insertSheet('取引先一覧');
    sheet.getRange('A1:B1').setValues([['登録番号', '取引先名']]);
    sheet.getRange('A1:B1').setBackground('#d9ead3').setFontWeight('bold');
    sheet.setColumnWidth(1, 180);
    sheet.setColumnWidth(2, 240);
    ui.alert('「取引先一覧」シートを作成しました。\nA列に登録番号、B列に取引先名を入力してください。');
    return {};
  }

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

function resolvePartnerName_(invoiceNumber, partnerMap, vendorName) {
  if (invoiceNumber && partnerMap[invoiceNumber]) return partnerMap[invoiceNumber];
  if (vendorName) return vendorName;
  return '';
}

function buildMoneyForwardRow_(transactionNo, data, partnerName, fileUrl, settings) {
  const memoUrl = fileUrl || '';
  const amount = data.amount || 0;
  const csvDate = normalizeDate_(data.date);
  const creditAccount =
    data.paymentMethod === 'クレカ'
      ? (settings?.creditAccountCard || '未払金')
      : (settings?.creditAccountOther || '役員借入金');
  const creditSubAccount =
    data.paymentMethod === 'クレカ'
      ? resolveCreditSubAccount_(settings, data)
      : '';

  return [
    transactionNo,
    csvDate,
    data.accountTitle,
    '',
    '',
    partnerName,
    MF_TAX_CATEGORY,
    data.invoiceNumber,
    amount,
    0,
    creditAccount,
    creditSubAccount,
    '',
    '',
    MF_TAX_CATEGORY,
    '',
    amount,
    0,
    data.summary || partnerName,
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
          const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
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

  const prompt = `
    このファイル（画像またはPDF）はレシート／領収書／請求書です。
    内容から必要情報を抽出し、JSONオブジェクト1個だけを返してください。
    返答はJSONのみ（前後の説明、改行以外の文字、コードフェンス、Markdown、箇条書き、コメントは禁止）。
    余計なキーは追加しない。
    文字列は必ずダブルクォート、数値は数値型（"123"ではなく123）。
    不明な項目は空文字 "" または 0 を入れる（null/undefinedは使わない）。

    出力スキーマ（キー順固定）:
    {"paymentDate":"YYYY/MM/DD","paymentMethod":"現金|クレカ|PayPay|電子マネー|銀行振込|不明","cardInfo":"カード(1234)","vendorName":"取引先名","invoiceNumber":"T1234567890123","summary":"品目（概要）","amount":12345}

    抽出ルール:
    1) paymentDate（支払日）
    - レシート/領収書: 「日付」「取引日」「発行日」「購入日」などの最も妥当な日付。
    - 請求書: 支払日が明確ならそれ、無ければ発行日。支払期限/入金期限は使わない。
    - 西暦が無い場合は現在に近い年を推定。整形できなければ ""。

    2) paymentMethod（支払方法）
    - 候補は必ずこの5つから1つだけ: 「現金」「クレカ」「PayPay」「電子マネー」「銀行振込」
    - iD/QUICPay/クレジット/カード/Visa/Master/JCB/Amex/Diners は「クレカ」
    - Suica/PASMO/ICOCA/TOICA/manaca/はやかけん/nimoca/SUGOCA/楽天Edy/WAON/nanaco/交通系IC は「電子マネー」
    - 「PayPay」表記があれば「PayPay」
    - 「振込」「銀行」「口座」「振込先」等があり、支払方法が振込と読める場合は「銀行振込」
    - 判別できない場合は "不明"

    2-2) cardInfo（カード情報）
    - 支払い方法がクレカの場合、カード番号の下4桁を抽出して「カード(1234)」形式で出力
    - 抽出できない場合は「カード(不明)」

    3) vendorName（取引先名）
    - 店名/会社名/発行者名/請求元名から最も適切な名称を短く抽出（住所や電話番号は含めない）
    - 不明なら ""

    4) invoiceNumber（登録番号）
    - 「登録番号」「適格請求書発行事業者登録番号」「インボイス番号」に続く文字列から抽出
    - 形式は T + 13桁（全角T/全角数字/空白混入も可）
    - 似た番号（伝票番号等）は入れない

    5) summary（概要）
    - 15文字程度までを目安に短く
    - レシート: 主な購入内容または用途カテゴリを優先
    - 請求書: 請求内容の要約を優先
    - vendorName と同じ文字列だけになるのは避ける（内容が取れない場合は vendorName で可）
    - 不明なら ""

    6) amount（税込合計）
    - 支払総額（税込）の整数。小数は四捨五入せず小数点以下を無視
    - 「合計」「総計」「お支払金額」「ご請求金額」等を優先
    - 不明なら 0

    出力例（JSONのみ）:
    {"paymentDate":"2026/01/18","paymentMethod":"クレカ","cardInfo":"カード(2235)","vendorName":"ENEOS","invoiceNumber":"T1234567890123","summary":"ガソリン代","amount":4500}
  `;

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

  const response = UrlFetchApp.fetch(endpoint, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    Logger.log(`API Error: ${JSON.stringify(json.error)}`);
    return null;
  }

  if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts) {
    const text = json.candidates[0].content.parts[0].text;
    const parsed = extractJsonFromText_(text);
    const normalized = normalizeReceiptExtraction_(parsed);
    if (normalized && !normalized.invoiceNumber) {
      normalized.invoiceNumber = extractInvoiceNumberFromText_(text);
    }
    return normalized;
  }

  return null;
}

function normalizeReceiptExtraction_(data) {
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
          .preview { width: 100%; height: 320px; border: 1px solid #ddd; }
          .thumb { width: 100%; border: 1px solid #ddd; margin-top: 8px; }
          .note { color: #888; margin-top: 6px; font-size: 11px; }
          .button { margin: 8px 0; }
        </style>
      </head>
      <body>
        <div class="title">画像プレビュー</div>
        <div id="meta" class="meta">行を選択してください。</div>
        <button class="button" onclick="refresh()">更新</button>
        <iframe id="preview" class="preview" src=""></iframe>
        <img id="thumb" class="thumb" src="" />
        <div class="note">プレビューが表示されない場合はサムネイルを確認してください。</div>
        <script>
          let lastFileId = '';
          function refresh() {
            google.script.run.withSuccessHandler(render).getSelectedRowPreviewInfo();
          }
          function render(data) {
            if (!data) return;
            if (data.message) {
              document.getElementById('meta').textContent = data.message;
              return;
            }
            if (data.fileId && data.fileId === lastFileId) return;
            lastFileId = data.fileId || '';
            document.getElementById('meta').textContent = data.fileName || '';
            document.getElementById('preview').src = data.previewUrl || '';
            document.getElementById('thumb').src = data.thumbnailUrl || '';
          }
          refresh();
          setInterval(refresh, 3000);
        </script>
      </body>
    </html>
  `;
}
