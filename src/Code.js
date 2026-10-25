/**
 * 問い合わせフォーム → スプレッドシート記録 → Gmail通知
 *
 * 使い方：
 *   1. setup() を1回だけ実行する（フォーム・スプレッドシート・トリガーを自動作成）
 *   2. 実行ログに表示されるフォームURLから回答すると、
 *      スプレッドシートに1行追加され、通知メールが届く
 */

// ===== カスタマイズできる設定 =====
const CONFIG = {
  FORM_TITLE: 'お問い合わせフォーム',
  SHEET_FILE_NAME: '問い合わせ管理表',
  SHEET_NAME: '問い合わせ一覧',
  // 通知先。空欄なら、このスクリプトを実行したGoogleアカウントのアドレスに送る
  NOTIFY_TO: '',
  STATUS_DEFAULT: '未対応',
  STATUS_OPTIONS: ['未対応', '対応中', '対応済み'],
  DATE_FORMAT: 'yyyy/MM/dd HH:mm:ss',       // メール本文の日時形式
  SHEET_DATE_FORMAT: 'yyyy/mm/dd hh:mm:ss', // スプレッドシートの日時形式（書き方がメールと少し違う）
};

// フォームの項目（title は通知メールとシートの見出しにも使う）
const FIELDS = [
  { key: 'company', title: '会社名', type: 'text' },
  { key: 'name', title: '氏名', type: 'text' },
  { key: 'email', title: 'メールアドレス', type: 'email' },
  { key: 'message', title: '問い合わせ内容', type: 'paragraph' },
];

const HEADERS = ['受付日時'].concat(FIELDS.map(f => f.title), ['対応状況']);
const HANDLER = 'onFormSubmit';

/**
 * 初期設定：フォーム・スプレッドシート・送信時トリガーを作成する。
 * 2回目以降は、作成済みの情報を表示するだけで何も作らない。
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('FORM_ID')) {
    Logger.log('セットアップ済みです。');
    showInfo();
    return;
  }

  // 1) スプレッドシート
  const ss = SpreadsheetApp.create(CONFIG.SHEET_FILE_NAME);
  const sheet = ss.getSheets()[0].setName(CONFIG.SHEET_NAME);
  sheet.getRange(1, 1, 1, HEADERS.length)
    .setValues([HEADERS])
    .setFontWeight('bold')
    .setBackground('#e8f0fe');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(HEADERS.indexOf('問い合わせ内容') + 1, 360);

  // 「対応状況」列はプルダウンで選べるようにする
  const statusCol = HEADERS.indexOf('対応状況') + 1;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.STATUS_OPTIONS, true)
    .build();
  sheet.getRange(2, statusCol, sheet.getMaxRows() - 1, 1).setDataValidation(rule);

  // 2) フォーム
  const form = FormApp.create(CONFIG.FORM_TITLE)
    .setDescription('以下の項目をご入力のうえ、送信してください。')
    .setConfirmationMessage('お問い合わせありがとうございます。担当者よりご連絡いたします。');
  FIELDS.forEach(f => {
    let item;
    if (f.type === 'paragraph') {
      item = form.addParagraphTextItem();
    } else {
      item = form.addTextItem();
      if (f.type === 'email') {
        item.setValidation(FormApp.createTextValidation()
          .setHelpText('メールアドレスの形式で入力してください')
          .requireTextIsEmail()
          .build());
      }
    }
    item.setTitle(f.title).setRequired(true);
  });

  // 3) 送信時トリガー（フォームに回答が届いたら onFormSubmit を実行）
  ScriptApp.newTrigger(HANDLER).forForm(form).onFormSubmit().create();

  props.setProperties({ FORM_ID: form.getId(), SHEET_ID: ss.getId() });
  Logger.log('セットアップが完了しました。');
  showInfo();
}

/**
 * フォームに回答が届いたときに自動で実行される。
 * @param {GoogleAppsScript.Events.FormsOnFormSubmit} e
 */
function onFormSubmit(e) {
  const response = e.response;
  const answers = {};
  response.getItemResponses().forEach(r => {
    answers[r.getItem().getTitle()] = String(r.getResponse());
  });
  const timestamp = response.getTimestamp();
  const receivedAt = Utilities.formatDate(
    timestamp, Session.getScriptTimeZone(), CONFIG.DATE_FORMAT);

  // スプレッドシートに1行追加（受付日時は日付として保存し、メールと同じ形式で表示）
  const sheet = getSheet_();
  sheet.appendRow([timestamp].concat(
    FIELDS.map(f => answers[f.title] || ''), [CONFIG.STATUS_DEFAULT]));
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat(CONFIG.SHEET_DATE_FORMAT);

  // 通知メール
  const company = answers['会社名'] || '';
  const name = answers['氏名'] || '';
  const subject = `【新規問い合わせ】${company} ${name}様`;
  const lines = ['新しいお問い合わせが届きました。', '', `受付日時：${receivedAt}`];
  FIELDS.forEach(f => lines.push(`${f.title}：${answers[f.title] || ''}`));
  lines.push('', '▼ 問い合わせ管理表', sheet.getParent().getUrl());

  const options = {};
  const replyTo = answers['メールアドレス'];
  if (replyTo) options.replyTo = replyTo; // 通知メールに「返信」すると問い合わせ者宛てになる
  MailApp.sendEmail(getNotifyTo_(), subject, lines.join('\n'), options);
}

/** 作成済みのフォームURL・スプレッドシートURLをログに表示する */
function showInfo() {
  const props = PropertiesService.getScriptProperties();
  const formId = props.getProperty('FORM_ID');
  if (!formId) {
    Logger.log('まだ setup() が実行されていません。');
    return;
  }
  const form = FormApp.openById(formId);
  Logger.log('回答用フォームURL：' + form.getPublishedUrl());
  Logger.log('フォーム編集URL：' + form.getEditUrl());
  Logger.log('スプレッドシートURL：' + getSheet_().getParent().getUrl());
  Logger.log('通知先：' + getNotifyTo_());
  Logger.log('トリガー数：' + ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === HANDLER).length);
}

function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return SpreadsheetApp.openById(id).getSheetByName(CONFIG.SHEET_NAME);
}

function getNotifyTo_() {
  return CONFIG.NOTIFY_TO || Session.getEffectiveUser().getEmail();
}
