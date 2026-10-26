/**
 * 予約申し込みフォーム → Googleカレンダーに予定を自動登録（同じ時間帯に予定があれば登録せず「要確認」）
 *
 * 使い方：
 *   1. setup() を1回だけ実行する（予約用カレンダー・予約一覧・フォーム・送信時トリガーを自動作成）
 *   2. フォームから予約が届くと、予約用カレンダーに予定が入り、予約一覧に「登録済み」と記録される
 *   3. 同じ時間帯に予定がある・過去の日時・受付時間外のときは登録せず、「要確認」と理由を記録する
 *
 * 予定は、普段使いのカレンダーではなく専用の「予約（デモ）」カレンダーに入る。
 * 申込者を参加者（ゲスト）には追加しないため、招待メールは送られない。
 */

// ===== カスタマイズできる設定 =====
const CONFIG = {
  FORM_TITLE: '予約申し込みフォーム',
  CALENDAR_NAME: '予約（デモ）',
  SHEET_FILE_NAME: '予約管理',
  SHEET_NAME: '予約一覧',
  OPEN_HOUR: 10,                 // 受付開始（時）
  CLOSE_HOUR: 18,                // 受付終了（時）。終了時刻がこれを超える予約は要確認
  SLOT_MINUTES: 30,              // 開始時刻の選択肢の間隔（分）
  DURATIONS: [30, 60, 90],       // 所要時間の選択肢（分）
  // 予約用カレンダーに加えて、普段使いのカレンダーの予定とも重ならないか調べる場合は true
  ALSO_CHECK_DEFAULT_CALENDAR: false,
  EVENT_TITLE: '【予約】{氏名}様',
};

const Q = { NAME: '氏名', EMAIL: 'メールアドレス', DATE: '希望日', TIME: '開始時刻', DURATION: '所要時間', NOTE: 'ご相談内容' };
const HEADERS = ['受付日時', '氏名', 'メールアドレス', '予約日', '開始', '終了', '所要時間', 'ご相談内容', '状態', '理由', '予定ID'];
const HANDLER = 'onFormSubmit';

/**
 * 初期設定：予約用カレンダー・予約一覧・フォーム・トリガーを作る。
 * 2回目以降は、作成済みの情報を表示するだけで何も作らない。
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('FORM_ID')) {
    Logger.log('セットアップ済みです。');
    showInfo();
    return;
  }
  const cal = CalendarApp.createCalendar(CONFIG.CALENDAR_NAME, { timeZone: Session.getScriptTimeZone() });

  const ss = SpreadsheetApp.create(CONFIG.SHEET_FILE_NAME);
  const sheet = ss.getSheets()[0].setName(CONFIG.SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  // 状態が「要確認」の行を目立たせる
  const statusCol = HEADERS.indexOf('状態') + 1;
  const range = sheet.getRange(2, statusCol, sheet.getMaxRows() - 1, 1);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('要確認').setBackground('#fde2e1').setFontColor('#b3261e').setRanges([range]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('登録済み').setBackground('#e6f4ea').setFontColor('#137333').setRanges([range]).build(),
  ]);

  const form = FormApp.create(CONFIG.FORM_TITLE)
    .setDescription(`ご希望の日時を選んで送信してください（受付時間：${CONFIG.OPEN_HOUR}:00〜${CONFIG.CLOSE_HOUR}:00）。`)
    .setConfirmationMessage('お申し込みありがとうございます。予約の確定をご連絡いたします。');
  form.addTextItem().setTitle(Q.NAME).setRequired(true);
  form.addTextItem().setTitle(Q.EMAIL).setRequired(true)
    .setValidation(FormApp.createTextValidation().requireTextIsEmail().setHelpText('メールアドレスの形式で入力してください').build());
  form.addDateItem().setTitle(Q.DATE).setRequired(true);
  form.addListItem().setTitle(Q.TIME).setRequired(true).setChoiceValues(timeSlots_());
  form.addListItem().setTitle(Q.DURATION).setRequired(true).setChoiceValues(CONFIG.DURATIONS.map(m => m + '分'));
  form.addParagraphTextItem().setTitle(Q.NOTE);

  ScriptApp.newTrigger(HANDLER).forForm(form).onFormSubmit().create();
  props.setProperties({ FORM_ID: form.getId(), SHEET_ID: ss.getId(), CALENDAR_ID: cal.getId() });
  Logger.log('セットアップが完了しました。');
  showInfo();
}

/**
 * フォームに予約が届いたときに自動で実行される。
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e
 */
function onFormSubmit(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30 * 1000); // 同時に届いた予約が、同じ時間帯に二重登録されないよう順番に処理する
  try {
    const props = PropertiesService.getScriptProperties();
    const answers = {};
    e.response.getItemResponses().forEach(r => { answers[r.getItem().getTitle()] = r.getResponse(); });
    const req = parseRequest_(answers);
    const cal = CalendarApp.getCalendarById(props.getProperty('CALENDAR_ID'));

    let reason = req.error || checkTime_(req.start, req.end, new Date());
    if (!reason) {
      const conflicts = findConflicts_(cal, req.start, req.end);
      if (conflicts.length) reason = '同じ時間帯に予定あり（' + conflicts.join('、') + '）';
    }

    let eventId = '';
    if (!reason) {
      const event = cal.createEvent(CONFIG.EVENT_TITLE.replace('{氏名}', req.name), req.start, req.end, {
        description: [`氏名：${req.name}`, `メールアドレス：${req.email}`, `ご相談内容：${req.note || '（なし）'}`].join('\n'),
      });
      eventId = event.getId();
    }

    const sheet = SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getSheetByName(CONFIG.SHEET_NAME);
    sheet.appendRow([e.response.getTimestamp(), req.name, req.email, req.start || answers[Q.DATE], req.start || '', req.end || '',
      req.minutes ? req.minutes + '分' : '', req.note, reason ? '要確認' : '登録済み', reason, eventId].map(safe_));
    formatRows_(sheet);
  } finally {
    lock.releaseLock();
  }
}

/** 日時の列の表示形式を整える（行の追加で自動の形式に上書きされるため、追加のたびに付け直す） */
function formatRows_(sheet) {
  const sh = sheet || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')).getSheetByName(CONFIG.SHEET_NAME);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  sh.getRange(2, 1, n, 6).setNumberFormats(Array(n).fill(['yyyy/mm/dd hh:mm', '@', '@', 'yyyy/mm/dd (ddd)', 'hh:mm', 'hh:mm']));
}

/** 予約一覧の日時の表示形式を付け直す（手で実行してもよい） */
function formatSheet() {
  formatRows_();
}

/** 作成済みの情報をログに表示する */
function showInfo() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('FORM_ID')) {
    Logger.log('まだ setup() が実行されていません。');
    return;
  }
  const form = FormApp.openById(props.getProperty('FORM_ID'));
  Logger.log('回答用フォームURL：' + form.getPublishedUrl());
  Logger.log('予約一覧URL：' + SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getUrl());
  Logger.log('予約用カレンダー：' + CONFIG.CALENDAR_NAME);
  Logger.log('トリガー数：' + ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === HANDLER).length);
}

// ===== 以下は内部で使う関数 =====

/** 開始時刻の選択肢（例：10:00, 10:30, … 17:30） */
function timeSlots_() {
  const slots = [];
  for (let m = CONFIG.OPEN_HOUR * 60; m < CONFIG.CLOSE_HOUR * 60; m += CONFIG.SLOT_MINUTES) {
    slots.push(`${Math.floor(m / 60)}:${String(m % 60).padStart(2, '0')}`);
  }
  return slots;
}

/** フォームの回答から、予約の開始・終了日時を作る。読み取れないときは error に理由を入れる */
function parseRequest_(answers) {
  const req = {
    name: String(answers[Q.NAME] || '').trim(),
    email: String(answers[Q.EMAIL] || '').trim(),
    note: String(answers[Q.NOTE] || '').trim(),
  };
  const d = String(answers[Q.DATE] || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);   // 例：2026-11-10
  const t = String(answers[Q.TIME] || '').match(/^(\d{1,2}):(\d{2})$/);             // 例：10:30
  const min = parseInt(String(answers[Q.DURATION] || ''), 10);                      // 例：「60分」→ 60
  if (!d || !t || !(min > 0)) {
    req.error = '希望日・開始時刻・所要時間を読み取れません';
    return req;
  }
  req.minutes = min;
  req.start = new Date(Number(d[1]), Number(d[2]) - 1, Number(d[3]), Number(t[1]), Number(t[2]));
  req.end = new Date(req.start.getTime() + min * 60 * 1000);
  return req;
}

/** 過去の日時・受付時間外を調べる。問題がなければ空文字 */
function checkTime_(start, end, now) {
  if (start.getTime() <= now.getTime()) return '過去の日時';
  const close = new Date(start.getFullYear(), start.getMonth(), start.getDate(), CONFIG.CLOSE_HOUR, 0);
  if (start.getHours() < CONFIG.OPEN_HOUR || end.getTime() > close.getTime()) return '受付時間外';
  return '';
}

/** 同じ時間帯に入っている予定の名前を返す（終日の予定は対象外。ちょうど前後に接する予定は重なりとみなさない） */
function findConflicts_(cal, start, end) {
  const cals = [cal];
  if (CONFIG.ALSO_CHECK_DEFAULT_CALENDAR) cals.push(CalendarApp.getDefaultCalendar());
  const names = [];
  cals.forEach(c => c.getEvents(start, end).forEach(ev => {
    if (ev.isAllDayEvent()) return;
    if (ev.getStartTime() < end && ev.getEndTime() > start) names.push(c === cal ? ev.getTitle() : '別の予定');
  }));
  return names;
}

/** 「=」などで始まる文字が数式として動かないようにする（外部から届く入力への対策） */
function safe_(value) {
  return typeof value === 'string' && /^[=+\-@]/.test(value) ? "'" + value : value;
}
