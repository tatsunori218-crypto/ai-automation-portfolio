// PC上で、金額計算・入力チェック・下書きメールの中身・二重作成の防止を試す（Googleへの接続は不要）
// 実行：npm test
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const drafts = [];
const files = [];
let rows;
const sheet = {
  getLastRow: () => rows.length + 1,
  getRange: (r, c, nr, nc) => ({
    getValues: () => rows.map(x => x.slice()),
    setValue: v => { rows[r - 2][c - 1] = v; },
    setValues: v => { v[0].forEach((x, k) => { rows[r - 2][c - 1 + k] = x; }); },
  }),
};
const b64 = (s, cs) => Buffer.from(typeof s === 'string' ? s : Buffer.from(s)).toString('base64');
const ctx = {
  console,
  Logger: { log: console.log },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => k + '_ID' }) },
  SpreadsheetApp: { openById: () => ({ getSheetByName: () => sheet }) },
  Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
  Utilities: {
    formatDate: d => `${d.getFullYear()}年${d.getMonth() + 1}月${d.getDate()}日`,
    base64Encode: b64,
    base64EncodeWebSafe: s => Buffer.from(s).toString('base64url'),
    getUuid: () => 'test-uuid',
    Charset: { UTF_8: 'UTF-8' },
  },
  Drive: { Files: {
    copy: (res) => ({ id: 'copy_' + res.name }),
    create: (res) => { files.push(res.name); return { id: 'pdf', webViewLink: 'https://drive.example/' + res.name }; },
  } },
  DocumentApp: { openById: () => ({
    getBody: () => ({ replaceText() {} }), saveAndClose() {},
    getAs: () => ({ setName(n) { this.name = n; return this; }, getName() { return this.name; }, getBytes: () => Buffer.from('%PDF-test') }),
  }) },
  Gmail: { Users: { Drafts: { create: (d) => drafts.push(Buffer.from(d.message.raw, 'base64url').toString()) } } },
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8'), ctx);

// 1) 金額計算
const d = (s) => new Date(s);
const data = ctx.toInvoiceData_({ '請求番号': 'INV-1', '宛先会社名': '株式会社サンプル商事', '担当者名': '佐藤 花子',
  'メールアドレス': 'hanako.sato@example.com', '請求日': d('2026-10-31'), '支払期限': d('2026-11-30'),
  '品目': '保守', '数量': 12, '単価': 2500 });
assert.strictEqual(data['小計'], '30,000');
assert.strictEqual(data['消費税'], '3,000');
assert.strictEqual(data['合計'], '33,000');
assert.strictEqual(data['請求日'], '2026年10月31日');
console.log('OK 金額計算：12 × 2,500 = 小計30,000 / 消費税3,000 / 合計33,000');

// 2) 入力チェック
const base = { '請求番号': 'X', '宛先会社名': 'A社', 'メールアドレス': 'a@example.com', '請求日': d('2026-10-31'), '支払期限': d('2026-11-30'), '数量': 1, '単価': 100 };
assert.strictEqual(ctx.validate_(base), '');
assert.strictEqual(ctx.validate_({ ...base, 'メールアドレス': 'a@example' }), 'メールアドレスの形式が不正');
assert.strictEqual(ctx.validate_({ ...base, '数量': 0 }), '数量が正の数ではない');
assert.strictEqual(ctx.validate_({ ...base, '請求日': '10月31日' }), '請求日が日付ではない');
console.log('OK 入力チェック：メール形式・数量・日付の不備を検出');

// 3) createInvoices の流れ（作成・要確認・二重作成の防止）
rows = [
  ['INV-001', '株式会社サンプル商事', '佐藤 花子', 'hanako.sato@example.com', d('2026-10-31'), d('2026-11-30'), '保守', 1, 30000, '', '', ''],
  ['INV-002', 'テスト工業株式会社', '鈴木 一郎', 'ichiro.suzuki@example', d('2026-10-31'), d('2026-11-30'), '作成', 1, 55000, '', '', ''],
  ['INV-003', '合同会社エグザンプル', '田中 次郎', 'jiro.tanaka@example.com', d('2026-10-31'), d('2026-11-30'), '作業', 12, 2500, '作成済み', 'x', 'y'],
];
ctx.createInvoices();
assert.strictEqual(rows[0][9], '作成済み');
assert.strictEqual(rows[1][9], '要確認：メールアドレスの形式が不正');
assert.strictEqual(drafts.length, 1, '下書きは1件だけ');
assert.deepStrictEqual(files, ['請求書_INV-001_株式会社サンプル商事.pdf']);
ctx.createInvoices(); // もう一度実行しても増えない
assert.strictEqual(drafts.length, 1);
assert.strictEqual(files.length, 1);
console.log('OK 作成済み・要確認の行は二度目の実行で飛ばす（PDF 1件・下書き 1件のまま）');

// 4) 下書きメールの中身
const mime = drafts[0];
const decodeB64Word = s => s.replace(/=\?UTF-8\?B\?([^?]+)\?=/g, (m, x) => Buffer.from(x, 'base64').toString());
assert.ok(mime.includes('To: hanako.sato@example.com'));
assert.ok(decodeB64Word(mime).includes('Subject: 【請求書】INV-001 株式会社サンプル商事 御中'));
const bodyPart = mime.split('\r\n\r\n')[2].split('\r\n')[0];
const body = Buffer.from(bodyPart, 'base64').toString();
assert.ok(body.includes('ご請求金額：33,000円（税込）'), body);
assert.ok(decodeB64Word(mime).includes('filename="請求書_INV-001_株式会社サンプル商事.pdf"'));
console.log('OK 下書きメール：宛先・件名・本文（合計金額）・PDF添付のファイル名');
console.log('すべての試験に合格しました');
