/**
 * ドライブのフォルダに置いたCSV → 列の並びと書式を整えて、スプレッドシートに出力
 *
 * 使い方：
 *   1. Googleドライブに、取り込み用フォルダ（初期設定：「CSV取込」）を作る
 *   2. setup() を1回だけ実行する（出力用スプレッドシートと10分ごとのトリガーを自動作成）
 *   3. フォルダにCSVを置くと、10分以内に整形されて「整形済みデータ」シートに追記される
 *
 * 整形の内容：
 *   - 列の並びを統一（見出しの名前がファイルごとに違っても、同じ列として扱う）
 *   - 日付を yyyy/MM/dd に統一（2026-10-1、2026年10月1日、20261001、R8.10.1 などに対応）
 *   - 全角/半角を統一（英数字・記号は半角、カタカナは全角、空白は1つに）
 *   - 重複行を削除（ファイル内の重複も、過去に取り込んだ行との重複も）
 *
 * ドライブは「読み取り専用」の権限だけで扱う（元のCSVの変更・移動・削除はしない）。
 */

// ===== カスタマイズできる設定 =====
const CONFIG = {
  INPUT_FOLDER_NAME: 'CSV取込',     // 取り込み用フォルダの名前（マイドライブ内）
  OUTPUT_FILE_NAME: 'CSV整形結果',
  DATA_SHEET: '整形済みデータ',
  LOG_SHEET: '処理履歴',
  INTERVAL_MINUTES: 10,             // チェック間隔（1・5・10・15・30分）
  DATE_FORMAT: 'yyyy/MM/dd',
  SHEET_DATE_FORMAT: 'yyyy/mm/dd',
};

// 出力する列の並び。aliases に、取り込むCSVで使われうる見出し名を並べる
const COLUMNS = [
  { key: 'date', title: '注文日', type: 'date', aliases: ['注文日', '日付', '受注日', '購入日', 'date'] },
  { key: 'name', title: '顧客名', type: 'text', aliases: ['顧客名', '氏名', '名前', 'お名前', 'name'] },
  { key: 'email', title: 'メールアドレス', type: 'email', aliases: ['メールアドレス', 'メール', 'email', 'e-mail', 'mail'] },
  { key: 'phone', title: '電話番号', type: 'phone', aliases: ['電話番号', '電話', 'tel', 'phone'] },
  { key: 'code', title: '商品コード', type: 'code', aliases: ['商品コード', '商品cd', '品番', 'sku'] },
  { key: 'qty', title: '数量', type: 'number', aliases: ['数量', '個数', 'qty'] },
  { key: 'amount', title: '金額', type: 'number', aliases: ['金額', '合計', '売上', 'amount'] },
];
// 重複判定に使わない補足列
const EXTRA_HEADERS = ['取込元ファイル', '要確認'];
const HEADERS = COLUMNS.map(c => c.title).concat(EXTRA_HEADERS);
// 列ごとの表示形式（日付・数値以外は文字列として扱う）
const NUMBER_FORMATS = COLUMNS.map(c =>
  c.type === 'date' ? CONFIG.SHEET_DATE_FORMAT : c.type === 'number' ? '#,##0' : '@'
).concat(EXTRA_HEADERS.map(() => '@'));
const LOG_HEADERS = ['処理日時', 'ファイル名', 'ファイルID', '読込行数', '出力行数', '重複削除数', '要確認数', '結果'];
const HANDLER = 'processCsv';

/**
 * 初期設定：出力用スプレッドシートとトリガーを作成し、フォルダ内のCSVを初回処理する。
 * 2回目以降は、作成済みの情報を表示するだけで何も作らない。
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('OUTPUT_ID')) {
    Logger.log('セットアップ済みです。');
    showInfo();
    return;
  }
  const folder = getInputFolder_(); // フォルダが無ければここで止まる

  const ss = SpreadsheetApp.create(CONFIG.OUTPUT_FILE_NAME);
  const data = ss.getSheets()[0].setName(CONFIG.DATA_SHEET);
  const log = ss.insertSheet(CONFIG.LOG_SHEET);
  [[data, HEADERS], [log, LOG_HEADERS]].forEach(([sheet, headers]) => {
    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]).setFontWeight('bold').setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
  });

  props.setProperties({ OUTPUT_ID: ss.getId(), FOLDER_ID: folder.getId() });
  ScriptApp.newTrigger(HANDLER).timeBased().everyMinutes(CONFIG.INTERVAL_MINUTES).create();

  Logger.log('セットアップが完了しました。フォルダ内のCSVを処理します。');
  processCsv();
  showInfo();
}

/**
 * フォルダ内の未処理のCSVを整形して追記する。
 * トリガーにより定期的に自動で実行される（手動で実行してもよい）。
 */
function processCsv() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('別の実行が進行中のため、今回はスキップしました。');
    return;
  }
  try {
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('OUTPUT_ID'));
    const data = ss.getSheetByName(CONFIG.DATA_SHEET);
    const log = ss.getSheetByName(CONFIG.LOG_SHEET);
    const processed = readColumn_(log, LOG_HEADERS.indexOf('ファイルID') + 1);
    const seen = new Set(readRows_(data).map(rowKey_));

    const files = [];
    const it = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('FOLDER_ID')).getFiles();
    while (it.hasNext()) {
      const f = it.next();
      if (/\.csv$/i.test(f.getName()) && !processed.has(f.getId())) files.push(f);
    }
    files.sort((a, b) => a.getDateCreated() - b.getDateCreated());
    if (files.length === 0) {
      Logger.log('新しいCSVはありませんでした。');
      return;
    }

    files.forEach(file => {
      const now = new Date();
      try {
        const result = formatCsv_(readCsvText_(file.getBlob()), seen);
        const rows = result.rows.map(r => r.concat([file.getName(), r.issues]).map(safe_));
        if (rows.length > 0) {
          const range = data.getRange(data.getLastRow() + 1, 1, rows.length, HEADERS.length);
          // 書式を先に決める（「1-2」のような文字が勝手に日付に変わるのを防ぐ）。
          // 日付・数値として読み取れなかった値は、勝手に変換されないよう文字列にする
          range.setNumberFormats(rows.map(r => NUMBER_FORMATS.map((f, i) =>
            i < COLUMNS.length && f !== '@' && typeof r[i] === 'string' ? '@' : f)));
          range.setValues(rows);
        }
        log.appendRow([now, file.getName(), file.getId(), result.read, rows.length,
          result.duplicates, result.issueCount, '成功']);
        Logger.log(`${file.getName()}：${result.read}行を読み込み、${rows.length}行を出力（重複${result.duplicates}行を削除）`);
      } catch (e) {
        log.appendRow([now, file.getName(), file.getId(), '', '', '', '', 'エラー：' + e.message]);
        Logger.log(`${file.getName()}：エラー ${e.message}`);
      }
    });
    log.getRange(2, 1, log.getLastRow() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
  } finally {
    lock.releaseLock();
  }
}

/** 出力先のURLや設定をログに表示する */
function showInfo() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('OUTPUT_ID')) {
    Logger.log('まだ setup() が実行されていません。');
    return;
  }
  Logger.log('取り込み用フォルダURL：' + DriveApp.getFolderById(props.getProperty('FOLDER_ID')).getUrl());
  Logger.log('出力先スプレッドシートURL：' + SpreadsheetApp.openById(props.getProperty('OUTPUT_ID')).getUrl());
  Logger.log('トリガー数：' + ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === HANDLER).length);
}

// ===== 以下は内部で使う関数 =====

function getInputFolder_() {
  const it = DriveApp.getRootFolder().getFoldersByName(CONFIG.INPUT_FOLDER_NAME);
  if (!it.hasNext()) {
    throw new Error(`マイドライブにフォルダ「${CONFIG.INPUT_FOLDER_NAME}」がありません。作成してから実行してください。`);
  }
  return it.next();
}

/** CSVの文字コード（UTF-8 / Shift_JIS）を判定して文字列にする */
function readCsvText_(blob) {
  let text = blob.getDataAsString('UTF-8');
  if (text.indexOf('�') !== -1) text = blob.getDataAsString('Shift_JIS');
  return text.replace(/^﻿/, '');
}

/**
 * CSVの文字列を整形する。seen（重複判定用の集合）は更新される。
 * @return {{rows: Array, read: number, duplicates: number, issueCount: number}}
 */
function formatCsv_(text, seen) {
  const table = Utilities.parseCsv(text).filter(r => r.some(c => String(c).trim() !== ''));
  if (table.length === 0) throw new Error('CSVが空です');

  const index = mapHeaders_(table[0]);
  if (index.date === undefined && index.name === undefined) {
    throw new Error('見出し行に、注文日・顧客名に当たる列が見つかりません');
  }

  const rows = [];
  let duplicates = 0;
  let issueCount = 0;
  table.slice(1).forEach(raw => {
    const issues = [];
    const row = COLUMNS.map(col => {
      const value = index[col.key] === undefined ? '' : raw[index[col.key]];
      const out = normalizeValue_(value, col.type);
      if (out.issue) issues.push(`${col.title}：${out.issue}`);
      return out.value;
    });
    const key = rowKey_(row);
    if (seen.has(key)) {
      duplicates++;
      return;
    }
    seen.add(key);
    if (issues.length) issueCount++;
    row.issues = issues.join(' / ');
    rows.push(row);
  });
  return { rows, read: table.length - 1, duplicates, issueCount };
}

/** 見出し名から、各列が何番目にあるかを調べる */
function mapHeaders_(headerRow) {
  const index = {};
  headerRow.forEach((h, i) => {
    const name = toHalfWidth_(h).toLowerCase().replace(/\s/g, '');
    const col = COLUMNS.find(c => c.aliases.some(a => a.toLowerCase() === name));
    if (col && index[col.key] === undefined) index[col.key] = i;
  });
  return index;
}

/** 列の種類ごとに書式を整える。整えられなかったときは issue に理由を入れる */
function normalizeValue_(value, type) {
  const text = toHalfWidth_(value);
  if (text === '') return { value: '' };
  switch (type) {
    case 'date': {
      const d = parseDate_(text);
      return d ? { value: d } : { value: text, issue: '日付を読み取れません' };
    }
    case 'email':
      return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(text)
        ? { value: text.toLowerCase() }
        : { value: text, issue: 'メールアドレスの形式ではありません' };
    case 'phone': {
      const digits = text.replace(/[-\s()]/g, '');
      const mobile = /^0[5789]0/.test(digits);
      if (!/^0\d{9,10}$/.test(digits) || (mobile && digits.length !== 11)) {
        return { value: text, issue: '電話番号の形式ではありません' };
      }
      // 携帯・IP電話は 3-4-4 に統一。固定電話は地域で区切り方が違うため、入力の区切りを残す
      if (mobile) return { value: digits.replace(/^(\d{3})(\d{4})(\d{4})$/, '$1-$2-$3') };
      return { value: text.replace(/\s/g, '-') };
    }
    case 'code':
      return { value: text.toUpperCase().replace(/\s/g, '') };
    case 'number': {
      const n = Number(text.replace(/[,¥円\\\s]/g, ''));
      return isNaN(n) ? { value: text, issue: '数値ではありません' } : { value: n };
    }
    default:
      return { value: text };
  }
}

/**
 * 全角英数字・記号を半角に、半角カタカナを全角にする（Unicode の NFKC 正規化）。
 * ハイフンに似た記号は「-」に、連続する空白は1つにまとめる。
 */
function toHalfWidth_(value) {
  return String(value == null ? '' : value)
    .normalize('NFKC')
    .replace(/[‐‑‒–—―−ｰー](?=\d)|(?<=\d)[‐‑‒–—―−ｰー]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

/** いろいろな書き方の日付を Date にする。読み取れなければ null */
function parseDate_(text) {
  let y, m, d, match;
  if ((match = text.match(/^(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})日?$/))) {
    [y, m, d] = [match[1], match[2], match[3]].map(Number);
  } else if ((match = text.match(/^(\d{4})(\d{2})(\d{2})$/))) {
    [y, m, d] = [match[1], match[2], match[3]].map(Number);
  } else if ((match = text.match(/^(R|令和|H|平成)\s?(\d{1,2}|元)[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})日?$/i))) {
    const base = /^(R|令和)$/i.test(match[1]) ? 2018 : 1988;
    y = base + (match[2] === '元' ? 1 : Number(match[2]));
    [m, d] = [match[3], match[4]].map(Number);
  } else {
    return null;
  }
  const date = new Date(y, m - 1, d);
  return date.getFullYear() === y && date.getMonth() === m - 1 && date.getDate() === d ? date : null;
}

/** 重複判定用のキー（整形後の値をつなげた文字列） */
function rowKey_(row) {
  return COLUMNS.map((c, i) => {
    const v = row[i];
    return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), CONFIG.DATE_FORMAT) : String(v);
  }).join('\u0001');
}

function readRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, COLUMNS.length).getValues();
}

function readColumn_(sheet, col) {
  const last = sheet.getLastRow();
  if (last < 2) return new Set();
  return new Set(sheet.getRange(2, col, last - 1, 1).getValues().map(r => String(r[0])));
}

/** 「=」などで始まる文字が数式として動かないようにする（外部から届くCSV対策） */
function safe_(value) {
  return typeof value === 'string' && /^[=+\-@]/.test(value) ? "'" + value : value;
}
