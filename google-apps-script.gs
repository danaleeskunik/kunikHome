// הדבק/י את כל הקוד הזה ב-Extensions > Apps Script של הגיליון, שמור, ואז:
// Deploy > Manage deployments > עריכה (עיפרון) > Version: New version > Deploy
// Execute as: Me | Who has access: Anyone
//
// כדי להפעיל הגנת סיסמה (מומלץ, האתר ציבורי ללא login):
// Project Settings (גלגל שיניים) > Script Properties > Add script property
// Property: APP_PASSWORD   Value: <הסיסמה שתרצי>
// כל עוד המאפיין הזה לא קיים, האפליקציה תמשיך לעבוד בלי סיסמה כרגיל.

var SHEET_NAMES = ['insurance', 'income', 'savings', 'vehicles', 'health', 'payments'];
var META_SHEET = '_meta';

function checkAuth(p) {
  var required = PropertiesService.getScriptProperties().getProperty('APP_PASSWORD');
  if (!required) return true;
  return p.password === required;
}

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  var viaPostMessage = p.mode === 'postmessage';

  if (!checkAuth(p)) {
    return reply({ status: 'unauthorized' }, p.callback, viaPostMessage);
  }

  if (p.action === 'save' && p.payload) {
    var result = saveAll(JSON.parse(p.payload));
    return reply(result, p.callback, viaPostMessage);
  }

  return reply(loadAll(), p.callback, viaPostMessage);
}

function doPost(e) {
  var p = (e && e.parameter) ? e.parameter : {};
  if (!checkAuth(p)) {
    return reply({ status: 'unauthorized' }, null, p.mode === 'postmessage');
  }
  var raw = (e.parameter && e.parameter.payload)
    ? e.parameter.payload
    : (e.postData ? e.postData.contents : '{}');
  return reply(saveAll(JSON.parse(raw)), null, p.mode === 'postmessage');
}

function saveAll(body) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  SHEET_NAMES.forEach(function (key) {
    var items = (body.data && body.data[key]) ? body.data[key] : [];
    var sheet = ss.getSheetByName(key) || ss.insertSheet(key);
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

  return { status: 'ok', savedAt: new Date().toISOString() };
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

function reply(payload, callback, viaPostMessage) {
  if (viaPostMessage) {
    // עמוד HTML זעיר שרץ בתוך iframe ומעביר את התשובה להורה דרך postMessage —
    // עוקף מגבלה של WebKit/iOS שחוסמת הפניה (redirect) בין-דומיינים כשטוענים
    // תג <script> ישירות (JSONP), אבל לא חוסמת ניווט iframe.
    // (משתמשים ב-HtmlService ולא ב-ContentService — רק HtmlService מריץ בפועל
    // <script> בתוך הדף; ContentService מיועד לתוכן גולמי בלבד ולא מריץ סקריפטים.)
    var html = '<script>parent.postMessage(' + JSON.stringify(JSON.stringify(payload)) + ', "*");<\/script>';
    return HtmlService.createHtmlOutput(html);
  }
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
