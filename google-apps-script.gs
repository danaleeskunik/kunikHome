// הדבק/י את כל הקוד הזה ב-Extensions > Apps Script של הגיליון, שמור, ואז:
// Deploy > Manage deployments > עריכה (עיפרון) > Version: New version > Deploy
// Execute as: Me | Who has access: Anyone

var SHEET_NAMES = ['insurance', 'income', 'savings', 'vehicles', 'health', 'documents', 'payments'];
var META_SHEET = '_meta';

function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};

  if (p.action === 'save' && p.payload) {
    var result = saveAll(JSON.parse(p.payload));
    return reply(result, p.callback);
  }

  return reply(loadAll(), p.callback);
}

function doPost(e) {
  var raw = (e.parameter && e.parameter.payload)
    ? e.parameter.payload
    : (e.postData ? e.postData.contents : '{}');
  return reply(saveAll(JSON.parse(raw)), null);
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
