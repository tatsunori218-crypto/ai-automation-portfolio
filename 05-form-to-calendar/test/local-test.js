// PC上で、日時の読み取り・受付時間の判定・予定の重なりの判定・登録の流れを試す（Googleへの接続は不要）
// 実行：npm test
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

const events = [];
const rows = [];
const calendar = {
  getEvents: (s, e) => events.filter(ev => ev.start < e && ev.end > s).map(ev => ({
    isAllDayEvent: () => false, getStartTime: () => ev.start, getEndTime: () => ev.end, getTitle: () => ev.title,
  })),
  createEvent: (title, start, end, opt) => { events.push({ title, start, end, opt }); return { getId: () => 'ev' + events.length }; },
};
const ctx = {
  console,
  Logger: { log: console.log },
  LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
  PropertiesService: { getScriptProperties: () => ({ getProperty: k => k }) },
  CalendarApp: { getCalendarById: () => calendar },
  SpreadsheetApp: { openById: () => ({ getSheetByName: () => ({ appendRow: r => rows.push(r), getLastRow: () => rows.length + 1, getRange: () => ({ setNumberFormats() {} }) }) }) },
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../src/Code.js'), 'utf8'), ctx);
ctx.Date = Date; // 試験側と同じ Date を使う

const submit = (name, date, time, duration) => ctx.onFormSubmit({ response: {
  getTimestamp: () => new Date(),
  getItemResponses: () => [['氏名', name], ['メールアドレス', name + '@example.com'], ['希望日', date], ['開始時刻', time], ['所要時間', duration], ['ご相談内容', '']]
    .map(([t, v]) => ({ getItem: () => ({ getTitle: () => t }), getResponse: () => v })),
} });
const last = () => rows[rows.length - 1];

// 1) 選択肢と日時の読み取り
const slots = ctx.timeSlots_();
assert.strictEqual(slots[0], '10:00');
assert.strictEqual(slots[slots.length - 1], '17:30');
const req = ctx.parseRequest_({ '希望日': '2030-11-10', '開始時刻': '10:30', '所要時間': '60分' });
assert.strictEqual(req.start.getHours(), 10);
assert.strictEqual(req.end.getHours(), 11);
assert.strictEqual(req.end.getMinutes(), 30);
console.log('OK 開始時刻の選択肢（10:00〜17:30）と、日時の読み取り（10:30から60分 → 11:30まで）');

// 2) 登録の流れ
submit('yamada', '2030-11-10', '10:00', '60分');
assert.strictEqual(last()[8], '登録済み');
submit('sato', '2030-11-10', '10:30', '30分');  // 10:00〜11:00 と重なる
assert.strictEqual(last()[8], '要確認');
assert.ok(last()[9].startsWith('同じ時間帯に予定あり'), last()[9]);
submit('suzuki', '2030-11-10', '11:00', '30分'); // ちょうど終わった直後は重ならない
assert.strictEqual(last()[8], '登録済み');
submit('tanaka', '2030-11-10', '17:30', '60分'); // 18:30まで → 受付時間外
assert.strictEqual(last()[9], '受付時間外');
submit('ito', '2020-01-06', '10:00', '30分');    // 過去
assert.strictEqual(last()[9], '過去の日時');
assert.strictEqual(events.length, 2, 'カレンダーに入ったのは2件だけ');
assert.ok(!('guests' in events[0].opt), '参加者（ゲスト）は追加しない');
console.log('OK 重なり・受付時間外・過去の日時は登録せず「要確認」、ちょうど接する時間は登録（カレンダー2件）');
console.log('すべての試験に合格しました');
