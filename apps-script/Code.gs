/**
 * 大会・イベント情報 (eventsData) を Google スプレッドシートに保存/取得するための
 * Google Apps Script（Web App として公開する）
 *
 * セットアップ手順は運用マニュアル.md の「大会情報の保存先（Googleスプレッドシート）」を参照。
 */

// 管理画面のログインパスワードと同じ値にしてあります。
// index.html 側の ADMIN_PASSWORD を変更した場合は、こちらも必ず同じ値に変更してください。
const ADMIN_PASSWORD = 'boundtennis2025';

const SHEET_NAME = 'events';

const COLUMNS = [
  'id', 'date', 'dateEnd', 'dow', 'name', 'venue', 'address',
  'status', 'body', 'youkou', 'draw', 'entry', 'result', 'photo',
  'extra1_label', 'extra1_url', 'extra2_label', 'extra2_url', 'extra3_label', 'extra3_url',
];

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]);
  }
  return sheet;
}

// 初回セットアップ時に一度だけ手動実行してください（ヘッダー行を作成します）。
function setup() {
  getSheet_();
}

function doGet(e) {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonOutput_([]);
  }
  const rows = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
  const events = rows
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => rowToEvent_(row));
  return jsonOutput_(events);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ ok: false, error: 'invalid_json' });
  }

  if (body.password !== ADMIN_PASSWORD) {
    return jsonOutput_({ ok: false, error: 'unauthorized' });
  }
  if (!Array.isArray(body.events)) {
    return jsonOutput_({ ok: false, error: 'invalid_events' });
  }

  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).clearContent();
  }
  if (body.events.length > 0) {
    const rows = body.events.map(eventToRow_);
    sheet.getRange(2, 1, rows.length, COLUMNS.length).setValues(rows);
  }
  return jsonOutput_({ ok: true });
}

function rowToEvent_(row) {
  const obj = {};
  COLUMNS.forEach((key, i) => { obj[key] = row[i]; });

  const extras = [
    { label: obj.extra1_label, url: obj.extra1_url },
    { label: obj.extra2_label, url: obj.extra2_url },
    { label: obj.extra3_label, url: obj.extra3_url },
  ].filter(x => x.label && x.url);

  return {
    id: Number(obj.id),
    date: formatDate_(obj.date),
    dateEnd: formatDate_(obj.dateEnd),
    dow: obj.dow,
    name: obj.name,
    venue: obj.venue,
    address: obj.address,
    status: obj.status,
    body: obj.body,
    youkou: obj.youkou,
    draw: obj.draw,
    entry: obj.entry,
    result: obj.result,
    photo: obj.photo,
    extras: extras,
  };
}

function eventToRow_(ev) {
  const ex = ev.extras || [];
  return [
    ev.id, ev.date, ev.dateEnd || '', ev.dow, ev.name, ev.venue, ev.address || '',
    ev.status, ev.body || '', ev.youkou || '', ev.draw || '', ev.entry || '', ev.result || '', ev.photo || '',
    ex[0] ? ex[0].label : '', ex[0] ? ex[0].url : '',
    ex[1] ? ex[1].label : '', ex[1] ? ex[1].url : '',
    ex[2] ? ex[2].label : '', ex[2] ? ex[2].url : '',
  ];
}

// スプレッドシートの日付セルは Date 型で返ってくることがあるため文字列(YYYY-MM-DD)に揃える
// ※ Apps Script(V8)では getValues() で返る Date が instanceof Date にならない場合があるため
//    Object.prototype.toString で判定する
function formatDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
