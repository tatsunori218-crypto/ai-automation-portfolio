/**
 * 請求一覧（スプレッドシート）→ 請求書PDFを自動作成してドライブに保存 → 宛先ごとにGmailの下書きを作成
 *
 * 使い方：
 *   1. setup() を1回だけ実行する（請求一覧・請求書のひな形・保存用フォルダを自動作成。サンプル3件入り）
 *   2. 請求一覧に1行＝1枚の請求書として入力する
 *   3. createInvoices() を実行すると、「状態」が空欄の行だけPDFと下書きが作られ、「作成済み」になる
 *   4. 担当者がGmailの下書きを確認してから送信する（このプログラムは送信しない）
 *
 * 権限は必要最小限：Gmailは「下書きの作成」のみ、ドライブは「このプログラムが作ったファイル」のみ。
 */

// ===== カスタマイズできる設定 =====
const CONFIG = {
  SHEET_FILE_NAME: '請求管理',
  SHEET_NAME: '請求一覧',
  TEMPLATE_NAME: '請求書ひな形',
  FOLDER_NAME: '請求書PDF',
  ISSUER: 'サンプル事務所',            // 発行者名（請求書の右上に入る）
  TAX_RATE: 0.1,                      // 消費税率
  STATUS_DONE: '作成済み',
  MAIL_SUBJECT: '【請求書】{請求番号} {宛先会社名} 御中',
  MAIL_BODY: [
    '{宛先会社名}',
    '{担当者名} 様',
    '',
    'いつもお世話になっております。{発行者}です。',
    '{請求日}付の請求書（{請求番号}）をお送りいたします。',
    'ご請求金額：{合計}円（税込）',
    'お支払期限：{支払期限}',
    '',
    'ご確認のほど、よろしくお願いいたします。',
  ].join('\n'),
};

// 請求一覧の列（この順番で並ぶ）
const HEADERS = ['請求番号', '宛先会社名', '担当者名', 'メールアドレス', '請求日', '支払期限',
  '品目', '数量', '単価', '状態', 'PDF', '作成日時'];
const COL = Object.fromEntries(HEADERS.map((h, i) => [h, i + 1]));
const SAMPLE_ROWS = [
  ['INV-2026-001', '株式会社サンプル商事', '佐藤 花子', 'hanako.sato@example.com', '2026/10/31', '2026/11/30', 'Webサイト保守（10月分）', 1, 30000],
  ['INV-2026-002', 'テスト工業株式会社', '鈴木 一郎', 'ichiro.suzuki@example.com', '2026/10/31', '2026/11/30', '業務自動化ツール作成', 1, 55000],
  ['INV-2026-003', '合同会社エグザンプル', '田中 次郎', 'jiro.tanaka@example.com', '2026/10/31', '2026/11/30', 'データ整理作業（時間単価）', 12, 2500],
];

/**
 * 初期設定：保存用フォルダ・請求書のひな形・請求一覧（サンプル3件）を作る。
 * 2回目以降は、作成済みの情報を表示するだけで何も作らない。
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_ID')) {
    Logger.log('セットアップ済みです。');
    showInfo();
    return;
  }
  // 途中で止まった setup をやり直しても重複しないよう、このプログラムが前に作ったものがあれば使い回す
  const folder = findOwnFile_(CONFIG.FOLDER_NAME, 'application/vnd.google-apps.folder')
    || Drive.Files.create({ name: CONFIG.FOLDER_NAME, mimeType: 'application/vnd.google-apps.folder' });
  const oldTemplate = findOwnFile_(CONFIG.TEMPLATE_NAME, 'application/vnd.google-apps.document', folder.id);
  if (oldTemplate) Drive.Files.update({ name: CONFIG.TEMPLATE_NAME + '（作り直し前）' }, oldTemplate.id);
  const templateId = createTemplate_(folder.id);

  const ss = SpreadsheetApp.create(CONFIG.SHEET_FILE_NAME);
  const sheet = ss.getSheets()[0].setName(CONFIG.SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  sheet.getRange(2, 1, SAMPLE_ROWS.length, SAMPLE_ROWS[0].length).setValues(SAMPLE_ROWS);
  sheet.getRange('E:F').setNumberFormat('yyyy/mm/dd');
  sheet.getRange('I:I').setNumberFormat('#,##0');
  sheet.getRange('L:L').setNumberFormat('yyyy/mm/dd hh:mm');
  sheet.autoResizeColumns(1, HEADERS.length);

  props.setProperties({ SHEET_ID: ss.getId(), FOLDER_ID: folder.id, TEMPLATE_ID: templateId });
  Logger.log('セットアップが完了しました。');
  showInfo();
}

/**
 * 「状態」が空欄の行について、請求書PDFとGmailの下書きを作る。
 * 作成できた行は「作成済み」、入力に不備がある行は「要確認：理由」になる（どちらも次回は飛ばす）。
 */
function createInvoices() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('別の実行が進行中のため、今回はスキップしました。');
    return;
  }
  try {
    const props = PropertiesService.getScriptProperties();
    const sheet = SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getSheetByName(CONFIG.SHEET_NAME);
    const last = sheet.getLastRow();
    if (last < 2) return;
    const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
    let created = 0;
    let skipped = 0;

    rows.forEach((row, i) => {
      const r = Object.fromEntries(HEADERS.map((h, j) => [h, row[j]]));
      if (String(r['状態']).trim() !== '') return; // 作成済み・要確認は飛ばす（二重作成の防止）
      if (!r['請求番号'] && !r['宛先会社名']) return; // 空行
      const rowNumber = i + 2;
      const problem = validate_(r);
      if (problem) {
        sheet.getRange(rowNumber, COL['状態']).setValue('要確認：' + problem);
        skipped++;
        return;
      }
      const data = toInvoiceData_(r);
      const pdf = createPdf_(data, props.getProperty('TEMPLATE_ID'), props.getProperty('FOLDER_ID'));
      createDraft_(data, pdf.blob);
      sheet.getRange(rowNumber, COL['状態'], 1, 3)
        .setValues([[CONFIG.STATUS_DONE, pdf.url, new Date()]]);
      created++;
    });
    Logger.log(`請求書 ${created}件を作成し、下書きを作りました。要確認 ${skipped}件。`);
  } finally {
    lock.releaseLock();
  }
}

/** 作成済みの情報をログに表示する */
function showInfo() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SHEET_ID')) {
    Logger.log('まだ setup() が実行されていません。');
    return;
  }
  Logger.log('請求一覧URL：' + SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getUrl());
  Logger.log('ひな形URL：https://docs.google.com/document/d/' + props.getProperty('TEMPLATE_ID') + '/edit');
  Logger.log('保存用フォルダURL：https://drive.google.com/drive/folders/' + props.getProperty('FOLDER_ID'));
}

// ===== 以下は内部で使う関数 =====

/**
 * このプログラムが作ったファイルを名前で探す（drive.file 権限では、自分が作ったファイルしか見えない）。
 * 見つからなければ null
 */
function findOwnFile_(name, mimeType, parentId) {
  let q = `name = '${name.replace(/'/g, "\\'")}' and mimeType = '${mimeType}' and trashed = false`;
  if (parentId) q += ` and '${parentId}' in parents`;
  const res = Drive.Files.list({ q, fields: 'files(id,name)', pageSize: 1 });
  return res.files && res.files.length ? res.files[0] : null;
}

/** 入力の不備を調べる。問題がなければ空文字 */
function validate_(r) {
  if (!r['請求番号']) return '請求番号が空欄';
  if (!r['宛先会社名']) return '宛先会社名が空欄';
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(r['メールアドレス']).trim())) return 'メールアドレスの形式が不正';
  if (!isDate_(r['請求日'])) return '請求日が日付ではない';
  if (!isDate_(r['支払期限'])) return '支払期限が日付ではない';
  if (!(Number(r['数量']) > 0)) return '数量が正の数ではない';
  if (!(Number(r['単価']) >= 0) || r['単価'] === '') return '単価が数値ではない';
  return '';
}

/** 正しい日付かどうか（スプレッドシートの日付セルは Date として読まれる） */
function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

/** 1行分の入力から、請求書に差し込む値を作る（金額の計算もここで行う） */
function toInvoiceData_(r) {
  const fmtDate = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy年M月d日');
  const yen = n => Math.round(n).toLocaleString('ja-JP');
  const subtotal = Number(r['数量']) * Number(r['単価']);
  const tax = Math.floor(subtotal * CONFIG.TAX_RATE);
  return {
    '請求番号': String(r['請求番号']).trim(),
    '宛先会社名': String(r['宛先会社名']).trim(),
    '担当者名': String(r['担当者名']).trim(),
    'メールアドレス': String(r['メールアドレス']).trim(),
    '請求日': fmtDate(r['請求日']),
    '支払期限': fmtDate(r['支払期限']),
    '品目': String(r['品目']).trim(),
    '数量': String(r['数量']),
    '単価': yen(r['単価']),
    '小計': yen(subtotal),
    '消費税': yen(tax),
    '合計': yen(subtotal + tax),
    '発行者': CONFIG.ISSUER,
  };
}

/** {項目名} を値に置き換える */
function fill_(text, data) {
  return text.replace(/\{([^}]+)\}/g, (m, key) => (key in data ? data[key] : m));
}

/** ひな形をコピーして値を差し込み、PDFにしてフォルダへ保存する */
function createPdf_(data, templateId, folderId) {
  const name = `請求書_${data['請求番号']}_${data['宛先会社名']}`;
  const copy = Drive.Files.copy({ name, parents: [folderId] }, templateId);
  const doc = DocumentApp.openById(copy.id);
  const body = doc.getBody();
  Object.keys(data).forEach(key => body.replaceText(`\\{\\{${key}\\}\\}`, data[key]));
  doc.saveAndClose();

  const blob = DocumentApp.openById(copy.id).getAs('application/pdf').setName(name + '.pdf');
  const pdf = Drive.Files.create({ name: name + '.pdf', parents: [folderId], mimeType: 'application/pdf' }, blob,
    { fields: 'id,webViewLink' });
  return { blob, url: pdf.webViewLink };
}

/** PDFを添付したメールを、Gmailの下書きとして作る（送信はしない） */
function createDraft_(data, pdfBlob) {
  const boundary = 'invoice_' + Utilities.getUuid();
  const b64 = s => Utilities.base64Encode(s, Utilities.Charset.UTF_8);
  const subject = '=?UTF-8?B?' + b64(fill_(CONFIG.MAIL_SUBJECT, data)) + '?=';
  const filename = '=?UTF-8?B?' + b64(pdfBlob.getName()) + '?=';
  const mime = [
    `To: ${data['メールアドレス']}`,
    `Subject: ${subject}`,
    'MIME-Version: 1.0',
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    '',
    `--${boundary}`,
    'Content-Type: text/plain; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
    '',
    b64(fill_(CONFIG.MAIL_BODY, data)),
    `--${boundary}`,
    `Content-Type: application/pdf; name="${filename}"`,
    `Content-Disposition: attachment; filename="${filename}"`,
    'Content-Transfer-Encoding: base64',
    '',
    Utilities.base64Encode(pdfBlob.getBytes()),
    `--${boundary}--`,
  ].join('\r\n');
  Gmail.Users.Drafts.create({ message: { raw: Utilities.base64EncodeWebSafe(mime) } }, 'me');
}

/** 請求書のひな形（Googleドキュメント）を作る。{{項目名}} の部分に値が入る */
function createTemplate_(folderId) {
  const file = Drive.Files.create({ name: CONFIG.TEMPLATE_NAME, mimeType: 'application/vnd.google-apps.document', parents: [folderId] });
  const doc = DocumentApp.openById(file.id);
  const body = doc.getBody();
  body.setMarginTop(48).setMarginBottom(48).setMarginLeft(56).setMarginRight(56);

  const title = body.getParagraphs()[0]; // 新しい文書の最初の空段落を使う
  title.setText('請 求 書');
  title.setHeading(DocumentApp.ParagraphHeading.TITLE).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('請求番号：{{請求番号}}\n請求日：{{請求日}}').setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph('{{宛先会社名}} 御中').editAsText().setFontSize(16).setBold(true);
  body.appendParagraph('{{担当者名}} 様').editAsText().setFontSize(12).setBold(false);
  body.appendParagraph('{{発行者}}').setAlignment(DocumentApp.HorizontalAlignment.RIGHT).editAsText().setFontSize(12);
  body.appendParagraph('下記のとおりご請求申し上げます。').editAsText().setFontSize(11);
  body.appendParagraph('ご請求金額　{{合計}}円（税込）').editAsText().setFontSize(18).setBold(true);
  body.appendParagraph('').editAsText().setFontSize(11).setBold(false);

  const table = body.appendTable([
    ['品目', '数量', '単価（円）', '金額（円）'],
    ['{{品目}}', '{{数量}}', '{{単価}}', '{{小計}}'],
    ['', '', '小計', '{{小計}}'],
    ['', '', `消費税（${Math.round(CONFIG.TAX_RATE * 100)}%）`, '{{消費税}}'],
    ['', '', '合計', '{{合計}}'],
  ]);
  const header = table.getRow(0);
  for (let c = 0; c < header.getNumCells(); c++) header.getCell(c).setBackgroundColor('#e8f0fe').editAsText().setBold(true);

  body.appendParagraph('');
  body.appendParagraph('お支払期限：{{支払期限}}');
  body.appendParagraph('お振込先：（振込先をここに記載してください）');
  doc.saveAndClose();
  return file.id;
}
