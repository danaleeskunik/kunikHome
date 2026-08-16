// הדבק/י את כל הקוד הזה ב-Extensions > Apps Script של הגיליון, שמור, ואז:
// Deploy > Manage deployments > עריכה (עיפרון) > Version: New version > Deploy
// Execute as: Me | Who has access: Anyone
// (חובה "Anyone" ולא "Only myself" — אחרת גוגל עצמה חוסמת כל מי שהוא לא
// את, עוד לפני שהקוד כאן בכלל רץ. ההגבלה האמיתית למי מותר היא כאן בקוד,
// דרך ALLOWED_EMAILS למטה — לא דרך הגדרת ה-deployment.)
//
// כדי להפעיל כניסה עם חשבון Google (מומלץ):
// Project Settings (גלגל שיניים) > Script Properties > Add script property
//   GOOGLE_CLIENT_ID  =  ה-Client ID מ-Google Cloud Console
//   ALLOWED_EMAILS    =  dana.lees@gmail.com,alonkunik@gmail.com  (מופרד בפסיקים)
//
// אפשר גם (או במקום זה) סיסמה משותפת:
//   APP_PASSWORD      =  <הסיסמה שתרצי>
// אם אין אף property מהשלושה — האפליקציה עובדת בלי הגנה בכלל, כמו קודם.

var SHEET_NAMES = ['insurance', 'income', 'savings', 'vehicles', 'health', 'payments', 'contacts'];
var META_SHEET = '_meta';

function getAllowedEmails() {
  var raw = PropertiesService.getScriptProperties().getProperty('ALLOWED_EMAILS') || '';
  return raw.split(',').map(function (s) { return s.trim().toLowerCase(); }).filter(function (s) { return s; });
}

// מאמת ID token של Google מול השרתים של גוגל עצמה (חתימה, תוקף, ושהוא
// באמת מיועד לאפליקציה שלנו) ומחזיר את המייל המאומת, או null אם לא תקין.
function verifyGoogleToken(idToken) {
  var clientId = PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID');
  if (!clientId || !idToken) return null;
  try {
    var resp = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    var info = JSON.parse(resp.getContentText());
    if (info.aud !== clientId) return null;
    if (info.email_verified !== 'true' && info.email_verified !== true) return null;
    var email = String(info.email || '').toLowerCase();
    var allowed = getAllowedEmails();
    if (allowed.length && allowed.indexOf(email) === -1) return null;
    return email;
  } catch (e) {
    return null;
  }
}

function checkAuth(p) {
  var password = PropertiesService.getScriptProperties().getProperty('APP_PASSWORD');
  var allowedEmails = getAllowedEmails();
  var googleConfigured = !!PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID') && allowedEmails.length > 0;

  if (!password && !googleConfigured) return true; // שום הגנה לא הוגדרה — לא נועלים בטעות

  if (password && p.password === password) return true;
  if (googleConfigured && p.idToken && verifyGoogleToken(p.idToken)) return true;
  return false;
}

// נועל בזמן קריאה/כתיבה כדי שטעינה לא תתפוס גיליון באמצע כתיבה (clearContents
// ואז שכתוב הן שתי פעולות נפרדות — בלי נעילה קריאה יכולה "לתפוס" את הרגע
// שבין השתיים ולראות דומיין ריק זמנית, גם אם לא באמת נמחק כלום).
function withLock(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { status: 'error', message: 'lock timeout' };
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};

  if (!checkAuth(p)) {
    return reply({ status: 'unauthorized' }, p.callback);
  }

  if (p.action === 'save' && p.payload) {
    var result = withLock(function () { return saveAll(JSON.parse(p.payload)); });
    return reply(result, p.callback);
  }

  return reply(withLock(loadAll), p.callback);
}

function doPost(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  if (!checkAuth(p)) {
    return reply({ status: 'unauthorized' }, null);
  }
  var raw = (e.parameter && e.parameter.payload)
    ? e.parameter.payload
    : (e.postData ? e.postData.contents : '{}');
  return reply(withLock(function () { return saveAll(JSON.parse(raw)); }), null);
}

function saveAll(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var skipped = [];

  SHEET_NAMES.forEach(function (key) {
    var items = (body.data && body.data[key]) ? body.data[key] : [];
    var sheet = ss.getSheetByName(key) || ss.insertSheet(key);

    // Safety guard: never silently wipe a domain that already has rows in
    // the sheet when the incoming payload for it is empty. In normal use
    // the client always sends every domain's full current list, so a
    // domain arriving empty while the sheet already has real data almost
    // always means the client's local state was never actually loaded for
    // that domain (failed/slow fetch, fresh browser profile, etc.) rather
    // than a genuine "delete everything" — so refuse and report it instead
    // of clearing. Tradeoff: intentionally emptying every record in a
    // domain from the app won't take effect on the sheet; do that directly
    // in the sheet if it's ever really needed.
    if (items.length === 0 && sheet.getLastRow() > 1) {
      skipped.push(key);
      return;
    }

    sheet.clearContents();
    if (items.length === 0) return;

    var allKeys = {};
    items.forEach(function (item) {
      Object.keys(item).forEach(function (k) { allKeys[k] = true; });
    });
    var headers = Object.keys(allKeys);
    sheet.appendRow(headers);

    var rows = items.map(function (item) {
      return headers.map(function (h) { return item[h] !== undefined ? String(item[h]) : ''; });
    });
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  });

  if (body.categories) {
    var meta = ss.getSheetByName(META_SHEET) || ss.insertSheet(META_SHEET);
    meta.clearContents();
    meta.appendRow(['categories', JSON.stringify(body.categories)]);
  }

  var result = { status: 'ok', savedAt: new Date().toISOString() };
  if (skipped.length) result.skipped = skipped;
  return result;
}

function loadAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = {};

  SHEET_NAMES.forEach(function (key) {
    data[key] = [];
    var sheet = ss.getSheetByName(key);
    if (!sheet) return;
    var values = sheet.getDataRange().getValues();
    if (values.length < 2) return;
    var headers = values[0];
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (row.join('') === '') continue;
      var item = {};
      headers.forEach(function (h, c) {
        if (!h) return;
        var v = row[c];
        if (v instanceof Date) {
          v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        item[h] = v === '' ? '' : String(v);
      });
      if (!item.id) item.id = 'row' + r;
      data[key].push(item);
    }
  });

  var categories = null;
  var meta = ss.getSheetByName(META_SHEET);
  if (meta) {
    var mv = meta.getDataRange().getValues();
    for (var i = 0; i < mv.length; i++) {
      if (mv[i][0] === 'categories' && mv[i][1]) {
        try { categories = JSON.parse(mv[i][1]); } catch (err) {}
      }
    }
  }

  return { status: 'ok', data: data, categories: categories };
}

function reply(payload, callback) {
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
