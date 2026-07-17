/**
 * 管理画面の各データ（大会・イベント／お知らせ／他県案内）を Google スプレッドシートに
 * 保存/取得するための Google Apps Script（Web App として公開する）
 *
 * セットアップ手順は apps-script/セットアップ手順.md を参照。
 *
 * API仕様:
 *   GET  ?type=events|news|away|content            → 該当データをJSONで返す（認証不要・誰でも閲覧可）
 *   POST { idToken, type, items: [] / content: {} } → 該当データを丸ごと保存（上書き）する
 *     idToken には index.html 側でGoogleサインインして取得したID tokenを渡す。
 *     ALLOWED_EMAILS に含まれるメールアドレスのアカウントのみ書き込みを許可する。
 */

// index.html の GOOGLE_CLIENT_ID と同じ値にしてください（Google Cloud ConsoleのOAuthクライアントID）。
const GOOGLE_CLIENT_ID = '684319382004-0619hecm5kivaefuaaaojneemn3rghqr.apps.googleusercontent.com';

// 管理画面での書き込みを許可するGoogleアカウントのメールアドレス一覧。
const ALLOWED_EMAILS = ['ishikawaboundtennis@gmail.com'];

// Googleが発行したID tokenを検証し、認証済みメールアドレスを返す（不正な場合はnull）。
function verifyIdToken_(idToken) {
  if (!idToken) return null;
  try {
    const res = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    if (res.getResponseCode() !== 200) return null;
    const info = JSON.parse(res.getContentText());
    if (info.aud !== GOOGLE_CLIENT_ID) return null;
    if (info.email_verified !== 'true' && info.email_verified !== true) return null;
    if (ALLOWED_EMAILS.indexOf(info.email) === -1) return null;
    return info.email;
  } catch (e) {
    return null;
  }
}

const EVENT_COLUMNS = [
  'id', 'date', 'dateEnd', 'dow', 'name', 'venue', 'address',
  'status', 'body', 'youkou', 'draw', 'entry', 'entryLabel', 'result', 'photo',
  'extra1_label', 'extra1_url', 'extra2_label', 'extra2_url', 'extra3_label', 'extra3_url',
];

const NEWS_COLUMNS = ['id', 'date', 'category', 'title', 'important', 'body'];

const AWAY_COLUMNS = ['id', 'date', 'dateEnd', 'region', 'name', 'youkou', 'entry', 'entryLabel'];

// type ごとの設定（シート名・カラム・行⇔オブジェクトの変換）
const REGISTRY = {
  events: { sheetName: 'events', columns: EVENT_COLUMNS, rowToObj: rowToEvent_, objToRow: eventToRow_ },
  news:   { sheetName: 'news',   columns: NEWS_COLUMNS,   rowToObj: rowToNews_,  objToRow: newsToRow_ },
  away:   { sheetName: 'away',   columns: AWAY_COLUMNS,   rowToObj: rowToAway_,  objToRow: awayToRow_ },
};

function getSheet_(cfg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.sheetName);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, cfg.columns.length).setValues([cfg.columns]);
  }
  return sheet;
}

// 初回セットアップ時に一度だけ手動実行してください（各シート・ヘッダー行を作成します）。
function setup() {
  Object.values(REGISTRY).forEach(cfg => getSheet_(cfg));
  getContentSheet_();
}

function doGet(e) {
  const type = (e && e.parameter && e.parameter.type) || 'events';

  if (type === 'content') {
    return jsonOutput_(getContent_());
  }

  const cfg = REGISTRY[type];
  if (!cfg) return jsonOutput_({ ok: false, error: 'unknown_type' });

  const sheet = getSheet_(cfg);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonOutput_([]);

  const rows = sheet.getRange(2, 1, lastRow - 1, cfg.columns.length).getValues();
  const items = rows
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => cfg.rowToObj(row));
  return jsonOutput_(items);
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ ok: false, error: 'invalid_json' });
  }

  if (!verifyIdToken_(body.idToken)) {
    return jsonOutput_({ ok: false, error: 'unauthorized' });
  }

  if (body.type === 'content') {
    if (typeof body.content !== 'object' || body.content === null) {
      return jsonOutput_({ ok: false, error: 'invalid_content' });
    }
    setContent_(body.content);
    return jsonOutput_({ ok: true });
  }

  const cfg = REGISTRY[body.type];
  if (!cfg) return jsonOutput_({ ok: false, error: 'unknown_type' });
  if (!Array.isArray(body.items)) {
    return jsonOutput_({ ok: false, error: 'invalid_items' });
  }

  const sheet = getSheet_(cfg);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, cfg.columns.length).clearContent();
  }
  if (body.items.length > 0) {
    const rows = body.items.map(cfg.objToRow);
    sheet.getRange(2, 1, rows.length, cfg.columns.length).setValues(rows);
  }
  return jsonOutput_({ ok: true });
}

// ---- content（ページ本文・情報管理。単一のJSON設定オブジェクトとして1セルに保存） ----

const CONTENT_SHEET_NAME = 'content';

function getContentSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONTENT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONTENT_SHEET_NAME);
    sheet.getRange(1, 1).setValue('content_json');
  }
  return sheet;
}

function getContent_() {
  const sheet = getContentSheet_();
  const v = sheet.getRange(2, 1).getValue();
  if (!v) return {};
  try {
    return JSON.parse(v);
  } catch (e) {
    return {};
  }
}

function setContent_(content) {
  const sheet = getContentSheet_();
  sheet.getRange(2, 1).setValue(JSON.stringify(content));
}

// ---- events ----

function rowToEvent_(row) {
  const obj = {};
  EVENT_COLUMNS.forEach((key, i) => { obj[key] = row[i]; });

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
    entryLabel: obj.entryLabel,
    result: obj.result,
    photo: obj.photo,
    extras: extras,
  };
}

function eventToRow_(ev) {
  const ex = ev.extras || [];
  return [
    ev.id, ev.date, ev.dateEnd || '', ev.dow, ev.name, ev.venue, ev.address || '',
    ev.status, ev.body || '', ev.youkou || '', ev.draw || '', ev.entry || '', ev.entryLabel || '', ev.result || '', ev.photo || '',
    ex[0] ? ex[0].label : '', ex[0] ? ex[0].url : '',
    ex[1] ? ex[1].label : '', ex[1] ? ex[1].url : '',
    ex[2] ? ex[2].label : '', ex[2] ? ex[2].url : '',
  ];
}

// ---- news ----

function rowToNews_(row) {
  const obj = {};
  NEWS_COLUMNS.forEach((key, i) => { obj[key] = row[i]; });
  return {
    id: Number(obj.id),
    date: formatDate_(obj.date),
    category: obj.category,
    title: obj.title,
    important: obj.important === true || obj.important === 'TRUE' || obj.important === 'true',
    body: obj.body,
  };
}

function newsToRow_(n) {
  return [n.id, n.date, n.category, n.title, !!n.important, n.body || ''];
}

// ---- away（他県案内） ----

function rowToAway_(row) {
  const obj = {};
  AWAY_COLUMNS.forEach((key, i) => { obj[key] = row[i]; });
  return {
    id: Number(obj.id),
    date: formatDate_(obj.date),
    dateEnd: formatDate_(obj.dateEnd),
    region: obj.region,
    name: obj.name,
    youkou: obj.youkou,
    entry: obj.entry,
    entryLabel: obj.entryLabel,
  };
}

function awayToRow_(a) {
  return [a.id, a.date, a.dateEnd || '', a.region, a.name, a.youkou || '', a.entry || '', a.entryLabel || ''];
}

// ---- 共通 ----

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
