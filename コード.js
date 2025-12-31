// ==================================================
// 設定エリア
// ==================================================
// モデル名を変更しやすいように定数化
const MODEL_NAME = 'gemini-2.5-flash-lite'; 

const MF_TAX_CATEGORY = '課税仕入 10%';
const MF_CARD_SUB_ACCOUNT = '三井住友ゴールドカード';
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
  { name: '未入力', example: '上記のいずれにも当てはまらない少額の費用' }
];
const MF_ACCOUNT_CANDIDATE_NAMES = MF_ACCOUNT_CANDIDATES.map((item) => item.name);
const MF_ACCOUNT_CANDIDATE_GUIDE = MF_ACCOUNT_CANDIDATES
  .map((item) => `${item.name}：${item.example}`)
  .join('\n');

// ※APIキーは「スクリプトプロパティ」から、フォルダIDは「設定」シートから読み込みます

// ==================================================
// メニュー作成 (スプレッドシートを開いた時に実行)
// ==================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🧾 レシート解析')
    .addItem('1. ドライブをスキャンして案を出す', 'scanToSheet')
    .addSeparator()
    .addItem('2. 記入された名前を反映する', 'applyRenames')
    .addSeparator()
    .addItem('3. マネフォ用解析', 'analyzeMoneyForward')
    .addSeparator()
    .addItem('4. マネフォCSVダウンロード', 'downloadMoneyForwardCsv')
    .addSeparator()
    .addItem('指定フォルダに移動', 'moveFilesToSpecifiedFolder')
    .addSeparator()
    .addItem('⚙️ APIキー設定', 'setApiKey')
    .addToUi();
}

// ==================================================
// 解析結果シートの列構成を整える
// A:ファイルID, B:リンク, C:元ファイル名, D:変更案, E:移動先, F:ステータス
// ==================================================
function ensureResultSheetLayout_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ファイルID', 'リンク', '元ファイル名', '変更案（修正可）', '移動先', 'ステータス']);
    sheet.setRowHeight(1, 30);
    sheet.setColumnWidth(2, 60);
    sheet.setColumnWidth(4, 300);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 120);
    sheet.setFrozenRows(1);
    return;
  }

  const headerRow = sheet.getRange(1, 1, 1, Math.max(6, sheet.getLastColumn())).getValues()[0];
  const destinationCol = headerRow.indexOf('移動先') + 1;

  // 旧フォーマット（E列=ステータス）からの移行: D列の後ろに「移動先」を挿入
  if (destinationCol === 0) {
    sheet.insertColumnAfter(4);
    sheet.getRange(1, 5).setValue('移動先');
  }

  // 「ステータス」が無い場合は末尾に追加（通常は移動先追加でFに移動済み）
  const headerRowAfter = sheet.getRange(1, 1, 1, Math.max(6, sheet.getLastColumn())).getValues()[0];
  const statusColAfter = headerRowAfter.indexOf('ステータス') + 1;
  if (statusColAfter === 0) {
    sheet.insertColumnAfter(5);
    sheet.getRange(1, 6).setValue('ステータス');
  }

  sheet.setRowHeight(1, 30);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 120);
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
// 設定値を取得する関数
// ==================================================
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName('設定');
  
  // 設定シートがない場合は作成
  if (!configSheet) {
    configSheet = ss.insertSheet('設定');
    // ヘッダーと初期値を設定
    configSheet.getRange('A1:B1').setValues([['項目名', '設定値']]);
    configSheet.getRange('A2:B2').setValues([
      ['対象フォルダID', '']
    ]);
    
    // 見た目を整える
    configSheet.getRange('A1:B1').setBackground('#efefef').setFontWeight('bold');
    configSheet.setColumnWidth(1, 150);
    configSheet.setColumnWidth(2, 400);
    
    SpreadsheetApp.getUi().alert('「設定」シートを作成しました。\nB2セルに「フォルダID」を入力してください。\nAPIキーはメニューの「⚙️ APIキー設定」から登録してください。');
    return null;
  }
  
  // フォルダIDはシートから取得
  const folderId = configSheet.getRange('B2').getValue();
  
  // APIキーはスクリプトプロパティから取得
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  // バリデーション
  let errorMsg = [];
  if (!folderId) errorMsg.push('「設定」シートのB2セルに、対象のフォルダIDを入力してください。');
  if (!apiKey) errorMsg.push('Gemini APIキーが設定されていません。\nメニューの「⚙️ APIキー設定」からキーを登録してください。');
  
  if (errorMsg.length > 0) {
    SpreadsheetApp.getUi().alert(errorMsg.join('\n'));
    return null;
  }
  
  return { folderId: folderId, apiKey: apiKey };
}

// ==================================================
// 置換ルールシートからルールを取得する関数
// ==================================================
function getReplacementRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('置換ルール');
  
  // シートがない場合は作成
  if (!sheet) {
    sheet = ss.insertSheet('置換ルール');
    sheet.getRange('A1:B1').setValues([['検索キーワード（これを含んでいたら）', '置換後の概要（これにする）']]);
    
    // サンプルデータ
    sheet.getRange('A2:B4').setValues([
      ['ピカピカ', 'ガソリン'],
      ['ENEOS', 'ガソリン'],
      ['セブンイレブン', '食費']
    ]);
    
    sheet.getRange('A1:B1').setBackground('#d9ead3').setFontWeight('bold');
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(2, 200);
    SpreadsheetApp.getUi().alert('「置換ルール」シートを作成しました。\nこのシートに変換ルールを登録すると、AIの抽出結果を自動で書き換えます。');
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  // A列とB列の値を取得して返す
  return sheet.getRange(2, 1, lastRow - 1, 2).getValues(); 
}

// ==================================================
// 置換ロジックを適用する関数（インボイス番号対応版）
// ==================================================
function applyReplacement(nameText, rules) {
  if (!nameText || !rules || rules.length === 0) return nameText;

  // 全角の「｜」で分割
  const parts = nameText.split('｜');
  
  // フォーマットが「支払方法｜日付｜インボイス｜概要」の4要素でない場合は何もしない
  if (parts.length < 4) return nameText;
  
  // [0]:支払方法, [1]:日付, [2]:インボイス番号, [3]:概要
  let summary = parts[3]; // 概要部分
  
  // ルール表を上から順に走査
  for (const rule of rules) {
    const keyword = rule[0];      // A列：検索キーワード
    const replacement = rule[1];  // B列：置換後の文字
    
    // キーワードが空でなく、概要にそのキーワードが含まれていれば置換
    if (keyword && String(summary).includes(keyword)) {
      summary = replacement;
      // 1つヒットしたら終了（上にあるルールが優先）
      break; 
    }
  }
  
  // 再結合して返す
  return `${parts[0]}｜${parts[1]}｜${parts[2]}｜${summary}`;
}

// ==================================================
// 機能1: ドライブをスキャンしてシートに書き出す
// ==================================================
function scanToSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  // 設定を取得
  const settings = getSettings();
  if (!settings) return; 

  const currentSheetName = sheet.getName();
  if (currentSheetName === '設定' || currentSheetName === '置換ルール' || currentSheetName === '移動先リスト') {
    ui.alert('解析結果を出力したいシート（「シート1」など）を開いてから実行してください。');
    return;
  }

  // 置換ルールを読み込む
  const replacementRules = getReplacementRules();

  ensureResultSheetLayout_(sheet);

  // 既存のファイルIDを取得
  const lastRow = sheet.getLastRow();
  let existingIds = [];
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    existingIds = data.flat();
  }

  let folder;
  try {
    folder = DriveApp.getFolderById(settings.folderId);
  } catch (e) {
    ui.alert('フォルダが見つかりません。IDが正しいか確認してください。\n' + e.message);
    return;
  }

  const files = folder.getFiles();
  let processCount = 0;

  while (files.hasNext()) {
    const file = files.next();
    const id = file.getId();
    const fileName = file.getName();
    const mimeType = file.getMimeType();

    if (existingIds.includes(id)) continue;
    
    if (!isAllowedReceiptMimeType_(mimeType)) continue;
    
    if (fileName.match(/^202\d{5}_/)) continue;

    try {
      const blob = file.getBlob();
      const base64Data = Utilities.base64Encode(blob.getBytes());
      
      // HYPERLINK関数でリンクを作成
      const thumbnailFormula = `=HYPERLINK("https://drive.google.com/file/d/${id}/view", "開く")`;

      // Gemini API呼び出し
      let aiSuggestedName = callGeminiApi(base64Data, mimeType, settings.apiKey);
      
      let newNameCandidate = "";
      let status = "解析失敗";

      if (aiSuggestedName) {
        // AIの結果に対して置換ルールを適用
        aiSuggestedName = applyReplacement(aiSuggestedName, replacementRules);

        const extension = fileName.substring(fileName.lastIndexOf('.'));
        newNameCandidate = aiSuggestedName + extension;
        status = "未処理";
      }

      sheet.appendRow([id, thumbnailFormula, fileName, newNameCandidate, "", status]);
      sheet.setRowHeight(sheet.getLastRow(), 30);
      processCount++;

    } catch (e) {
      console.error(e);
      sheet.appendRow([id, "", fileName, "エラー発生", "", e.toString()]);
    }
  }

  if (processCount === 0) {
    ui.alert('新しいファイルは見つかりませんでした。');
  } else {
    ui.alert(`${processCount} 件のファイルをスキャンしました。\n「変更案」列を確認・修正してから、メニューの「2. 反映する」を実行してください。`);
  }
}

// ==================================================
// 機能2: シートの内容をファイル名に反映する
// ==================================================
function applyRenames() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  ensureResultSheetLayout_(sheet);
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('データがありません。');
    return;
  }

  const range = sheet.getRange(2, 1, lastRow - 1, 6);
  const data = range.getValues();
  
  let successCount = 0;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const fileId = row[0];
    const newName = row[3];
    const status = row[5];

    if (status === "未処理" && newName !== "" && fileId !== "") {
      try {
        const file = DriveApp.getFileById(fileId);
        const oldName = file.getName();
        
        if (oldName !== newName) {
          file.setName(newName);
          sheet.getRange(i + 2, 6).setValue("完了");
          sheet.getRange(i + 2, 3).setValue(newName);
          successCount++;
        } else {
          sheet.getRange(i + 2, 6).setValue("変更なし");
        }

      } catch (e) {
        sheet.getRange(i + 2, 6).setValue("エラー: " + e.message);
      }
    }
  }

  ui.alert(`${successCount} 件のファイル名を変更しました。`);
}

// ==================================================
// 移動先リストシートから [キーワード -> フォルダID] を取得
// ==================================================
function getDestinationFolderIdByKeyword_(keyword) {
  if (!keyword) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('移動先リスト');

  if (!sheet) {
    sheet = ss.insertSheet('移動先リスト');
    sheet.getRange('A1:B1').setValues([['キーワード', 'フォルダID']]);
    sheet.getRange('A1:B1').setBackground('#fff2cc').setFontWeight('bold');
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 420);
    SpreadsheetApp.getUi().alert('「移動先リスト」シートを作成しました。\nA列にキーワード、B列に移動先フォルダIDを設定してください。');
    return null;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (const [key, folderId] of values) {
    if (!key || !folderId) continue;
    if (String(key).trim() === String(keyword).trim()) return String(folderId).trim();
  }
  return null;
}

// ==================================================
// 指定フォルダに移動（E列=キーワード, F列=ステータス）
// ==================================================
function moveFilesToSpecifiedFolder() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const currentSheetName = sheet.getName();
  if (currentSheetName === '設定' || currentSheetName === '置換ルール' || currentSheetName === '移動先リスト') {
    ui.alert('解析結果のシート（「シート1」など）を開いてから実行してください。');
    return;
  }

  ensureResultSheetLayout_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('データがありません。');
    return;
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  let movedCount = 0;
  let skippedCount = 0;
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
      skippedCount++;
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
      errorCount++;
    }
  }

  ui.alert(`移動処理が完了しました。\n移動済み: ${movedCount}\n未移動（条件不一致）: ${skippedCount}\nエラー: ${errorCount}`);
}

// ==================================================
// マネフォ用解析: CSV出力
// ==================================================
function analyzeMoneyForward() {
  const ui = SpreadsheetApp.getUi();

  const settings = getSettings();
  if (!settings) return;

  let folder;
  try {
    folder = DriveApp.getFolderById(settings.folderId);
  } catch (e) {
    ui.alert('フォルダが見つかりません。IDが正しいか確認してください。\n' + e.message);
    return;
  }

  const partnerMap = getTradePartnerMap_(ui);
  const files = folder.getFiles();
  const rows = [];
  const errors = [];
  let transactionNo = 1;

  while (files.hasNext()) {
    const file = files.next();
    const mimeType = file.getMimeType();
    if (!isAllowedReceiptMimeType_(mimeType)) continue;

    try {
      const nameInfo = parseReceiptFilename_(file.getName());
      const analysis = analyzeReceiptForMoneyForward_(file, mimeType, settings.apiKey);
      if (!analysis) {
        errors.push(`${file.getName()}: 解析失敗`);
        continue;
      }

      const merged = mergeReceiptData_(analysis, nameInfo, file);
      const partnerName = resolvePartnerName_(
        merged.invoiceNumber,
        partnerMap,
        merged.vendorName
      );
      const row = buildMoneyForwardRow_(transactionNo, merged, partnerName);
      rows.push(row);
      transactionNo++;
    } catch (e) {
      console.error(e);
      errors.push(`${file.getName()}: ${e.message}`);
    }
  }

  const sheet = getMoneyForwardSheet_(ui);
  resetMoneyForwardSheet_(sheet);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, MF_CSV_HEADERS.length).setValues(rows);
  }

  if (rows.length === 0) {
    ui.alert('対象ファイルがありませんでした。');
  } else {
    ui.alert(`${rows.length} 件のレシートを「マネフォ用」シートに更新しました。`);
  }

  if (errors.length > 0) {
    ui.alert(`解析できないファイルがありました。\n${errors.slice(0, 10).join('\n')}`);
  }
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

  const range = sheet.getRange(1, 1, lastRow, MF_CSV_HEADERS.length);
  const values = range.getValues();
  const header = values[0];
  const rows = values
    .slice(1)
    .filter((row) => row.some((cell) => !isBlankCell_(cell)));

  if (rows.length === 0) {
    ui.alert('「マネフォ用」シートに出力対象の行がありません。');
    return;
  }

  const csvContent = buildCsvContent_(header, rows);
  const filename = buildMoneyForwardFilename_();
  showDownloadDialog_(filename, csvContent);
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
  const summary = normalizeText_(parts.slice(3).join('｜'));

  return {
    paymentMethod: paymentMethod,
    date: date,
    invoiceNumber: invoiceNumber,
    summary: summary
  };
}

function analyzeReceiptForMoneyForward_(file, mimeType, apiKey) {
  const blob = file.getBlob();
  const base64Data = Utilities.base64Encode(blob.getBytes());
  return callGeminiApiForMoneyForward_(base64Data, mimeType, apiKey);
}

function callGeminiApiForMoneyForward_(base64Data, mimeType, apiKey) {
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${apiKey}`;
  const prompt = buildMoneyForwardPrompt_();

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

  const text = json?.candidates?.[0]?.content?.parts?.[0]?.text;
  const parsed = extractJsonFromText_(text);
  return normalizeReceiptData_(parsed);
}

function buildMoneyForwardPrompt_() {
  return `
あなたは経理入力補助です。次のレシート画像/PDFから必要な情報を抽出してください。
出力はJSONのみ（前後の説明やマークダウンは不要）です。

出力形式:
{
  "date": "YYYY/MM/DD",
  "amount": 12345,
  "invoiceNumber": "T1234567890123",
  "vendorName": "取引先名",
  "summary": "摘要",
  "paymentMethod": "現金|クレカ|電子マネー",
  "accountTitle": "勘定科目"
}

ルール:
- amount は税込合計の整数。
- invoiceNumber は見つからない場合は空文字。
- paymentMethod は「現金」「クレカ」「電子マネー」のいずれか。
- accountTitle は次の候補から1つだけ選ぶ。

勘定科目候補:
${MF_ACCOUNT_CANDIDATE_GUIDE}
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
  if (!text) return '';
  if (text.includes('クレカ') || text.includes('クレジット') || text.includes('カード')) return 'クレカ';
  if (text.includes('電子') || text.includes('交通系') || text.includes('IC')) return '電子マネー';
  if (text.includes('現金')) return '現金';
  return '';
}

function normalizeInvoiceNumber_(value) {
  const text = normalizeText_(value);
  if (!text) return '';
  const match = text.match(/T\d{13}/);
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

function normalizeAccountTitle_(value) {
  const text = normalizeText_(value);
  if (MF_ACCOUNT_CANDIDATE_NAMES.includes(text)) return text;
  return '未入力';
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

function buildMoneyForwardRow_(transactionNo, data, partnerName) {
  const amount = data.amount || 0;
  const creditAccount =
    data.paymentMethod === 'クレカ' ? '未払金' : '役員借入金';
  const creditSubAccount =
    data.paymentMethod === 'クレカ' ? MF_CARD_SUB_ACCOUNT : '';

  return [
    transactionNo,
    data.date,
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
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    ''
  ];
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
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${MODEL_NAME}:generateContent?key=${apiKey}`;

  const prompt = `
    このファイル（画像またはPDF）はレシート、または領収書です。
    内容を解析して、ファイル名として使うための文字列を作成してください。

    【出力ルール】
    1. フォーマットは「支払方法｜YYYYMMDD｜インボイス番号｜概要」としてください。
       ※区切り文字は全角の縦線「｜」を使用してください。
    2. 日付（YYYYMMDD）はレシートの日付を使用してください。西暦が不明な場合は現在に近い年を推測してください。
    3. 「インボイス番号」は「T」から始まる13桁の数字（登録番号）を抽出してください。
       ※見つからない、または判読できない場合は「T0000000000000」としてください。
    4. 支払方法は「現金」「クレカ」「電子マネー」のいずれかに分類してください。
       ※特に「iD」での支払いは「クレカ」として判定してください。
       ※不明な場合は「現金」としてください。
    5. 「概要」は、レシートの内容から「店名」または「購入した主な商品・サービス（例：ガソリン代、食料品、書籍代）」を短く抽出してください。
    6. 余計な説明やマークダウン記号は一切不要です。ファイル名の文字列のみを返してください。
    
    【例】
    クレカ｜20231126｜T1234567890123｜スーパーの店名
    現金｜20240105｜T0000000000000｜タクシー代
    電子マネー｜20250815｜T9876543210987｜ガソリン代
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
    let text = json.candidates[0].content.parts[0].text;
    text = text.trim().replace(/\n/g, '').replace(/`/g, '');
    return text;
  }
  
  return null;
}
