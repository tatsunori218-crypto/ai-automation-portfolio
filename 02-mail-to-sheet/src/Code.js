/**
 * Gmailの特定ラベルのメール → スプレッドシートに自動記録（1時間ごと）
 *
 * 使い方：
 *   1. Gmailで、記録したいメールにラベル（初期設定：「請求書」）を付けておく
 *   2. setup() を1回だけ実行する（管理表と1時間ごとのトリガーを自動作成し、初回の記録も行う）
 *   3. 以降は1時間ごとに新しいメールだけが追記される（同じメールは二重に記録しない）
 *
 * Gmailは「読み取り専用」の権限だけで扱う（メールの送信・削除・変更はしない）。
 */

// ===== カスタマイズできる設定 =====
const CONFIG = {
  LABEL_NAME: '請求書',           // 対象のGmailラベル（「経理/請求書」のような階層付きも可）
  SHEET_FILE_NAME: '請求書メール記録',
  SHEET_NAME: 'メール一覧',
  BODY_LENGTH: 100,               // 本文を何文字まで記録するか
  MAX_PER_RUN: 100,               // 1回の実行で記録する最大件数（残りは次回に記録）
  INTERVAL_HOURS: 1,              // チェック間隔（時間）
  SHEET_DATE_FORMAT: 'yyyy/mm/dd hh:mm:ss',
};

const HEADERS = ['差出人', '件名', '受信日時', '本文（冒頭）', 'Gmailで開く', 'メールID'];
const COL = { DATE: 3, ID: 6 };
const HANDLER = 'checkMail';

/**
 * 初期設定：管理表と1時間ごとのトリガーを作成し、初回の記録を行う。
 * 2回目以降は、作成済みの情報を表示するだけで何も作らない。
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_ID')) {
    Logger.log('セットアップ済みです。');
    showInfo();
    return;
  }
  getLabelId_(); // ラベルが無ければここで止まる

  const ss = SpreadsheetApp.create(CONFIG.SHEET_FILE_NAME);
  const sheet = ss.getSheets()[0].setName(CONFIG.SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(HEADERS.indexOf('本文（冒頭）') + 1, 420);
  sheet.setColumnWidth(HEADERS.indexOf('件名') + 1, 260);
  props.setProperty('SHEET_ID', ss.getId());

  ScriptApp.newTrigger(HANDLER).timeBased().everyHours(CONFIG.INTERVAL_HOURS).create();

  Logger.log('セットアップが完了しました。初回の記録を行います。');
  checkMail();
  showInfo();
}

/**
 * ラベルの付いたメールのうち、まだ記録していないものを追記する。
 * トリガーにより1時間ごとに自動で実行される（手動で実行してもよい）。
 */
function checkMail() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) {
    Logger.log('別の実行が進行中のため、今回はスキップしました。');
    return;
  }
  try {
    const sheet = getSheet_();
    const recorded = getRecordedIds_(sheet);
    const newIds = listMessageIds_(getLabelId_()).filter(id => !recorded.has(id));
    if (newIds.length === 0) {
      Logger.log('新しいメールはありませんでした。');
      return;
    }

    // 古い順に記録する。件数が多いときは古いものから MAX_PER_RUN 件だけ処理し、残りは次回
    const messages = newIds.reverse().slice(0, CONFIG.MAX_PER_RUN)
      .map(id => Gmail.Users.Messages.get('me', id, { format: 'full' }))
      .sort((a, b) => Number(a.internalDate) - Number(b.internalDate));

    const rows = messages.map(m => {
      const headers = toHeaderMap_(m.payload.headers);
      return [
        safe_(headers['from'] || ''),
        safe_(headers['subject'] || '（件名なし）'),
        new Date(Number(m.internalDate)),
        safe_(truncate_(extractBody_(m.payload) || m.snippet || '', CONFIG.BODY_LENGTH)),
        'https://mail.google.com/mail/u/0/#all/' + m.id,
        m.id,
      ];
    });

    const start = sheet.getLastRow() + 1;
    sheet.getRange(start, 1, rows.length, HEADERS.length).setValues(rows);
    sheet.getRange(start, COL.DATE, rows.length, 1).setNumberFormat(CONFIG.SHEET_DATE_FORMAT);

    const rest = newIds.length - rows.length;
    Logger.log(`${rows.length}件を記録しました。` + (rest > 0 ? `残り${rest}件は次回記録します。` : ''));
  } finally {
    lock.releaseLock();
  }
}

/** 管理表のURLや設定をログに表示する */
function showInfo() {
  if (!PropertiesService.getScriptProperties().getProperty('SHEET_ID')) {
    Logger.log('まだ setup() が実行されていません。');
    return;
  }
  const sheet = getSheet_();
  Logger.log('スプレッドシートURL：' + sheet.getParent().getUrl());
  Logger.log('対象ラベル：' + CONFIG.LABEL_NAME);
  Logger.log('記録済み件数：' + Math.max(sheet.getLastRow() - 1, 0));
  Logger.log('トリガー数：' + ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === HANDLER).length);
}

// ===== 以下は内部で使う関数 =====

function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return SpreadsheetApp.openById(id).getSheetByName(CONFIG.SHEET_NAME);
}

function getLabelId_() {
  const labels = Gmail.Users.Labels.list('me').labels || [];
  const label = labels.find(l => l.name === CONFIG.LABEL_NAME);
  if (!label) {
    throw new Error(`Gmailにラベル「${CONFIG.LABEL_NAME}」がありません。Gmailでラベルを作成してから実行してください。`);
  }
  return label.id;
}

/** ラベルの付いたメールIDを新しい順にすべて取得する */
function listMessageIds_(labelId) {
  const ids = [];
  let pageToken;
  do {
    const res = Gmail.Users.Messages.list('me', { labelIds: [labelId], maxResults: 500, pageToken });
    (res.messages || []).forEach(m => ids.push(m.id));
    pageToken = res.nextPageToken;
  } while (pageToken);
  return ids;
}

/** 記録済みのメールIDを読み込む（二重記録の防止に使う） */
function getRecordedIds_(sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return new Set();
  return new Set(sheet.getRange(2, COL.ID, last - 1, 1).getValues().map(r => String(r[0])));
}

function toHeaderMap_(headers) {
  const map = {};
  (headers || []).forEach(h => { map[h.name.toLowerCase()] = h.value; });
  return map;
}

/** 本文を取り出す（テキスト形式を優先し、無ければHTMLからタグを除いて使う） */
function extractBody_(payload) {
  const text = findPart_(payload, 'text/plain');
  if (text) return decodePart_(text);
  const html = findPart_(payload, 'text/html');
  if (html) {
    return decodePart_(html)
      .replace(/<(style|script)[\s\S]*?<\/\1>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>');
  }
  return '';
}

function findPart_(part, mimeType) {
  if (part.mimeType === mimeType && part.body && part.body.data) return part;
  for (const child of part.parts || []) {
    const found = findPart_(child, mimeType);
    if (found) return found;
  }
  return null;
}

/** 文字コード（UTF-8 / ISO-2022-JP / Shift_JIS など）に合わせて本文を読む */
function decodePart_(part) {
  const type = toHeaderMap_(part.headers)['content-type'] || '';
  const match = type.match(/charset="?([^";\s]+)"?/i);
  // GASのGmailサービスは本文をバイト列で返す（文字列で来た場合だけBase64を戻す）
  const data = part.body.data;
  const bytes = typeof data === 'string' ? Utilities.base64DecodeWebSafe(data) : data;
  try {
    return Utilities.newBlob(bytes).getDataAsString(match ? match[1] : 'UTF-8');
  } catch (e) {
    return Utilities.newBlob(bytes).getDataAsString('UTF-8');
  }
}

/** 改行や連続した空白を1つの空白にまとめ、指定の文字数で切る */
function truncate_(text, length) {
  const chars = Array.from(text.replace(/\s+/g, ' ').trim());
  return chars.length > length ? chars.slice(0, length).join('') + '…' : chars.join('');
}

/** 「=」などで始まる文字が数式として動かないようにする（外部から届くメール対策） */
function safe_(value) {
  return /^[=+\-@]/.test(value) ? "'" + value : value;
}
