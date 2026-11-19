/**
 * 媽媽的照顧日誌 — 資料存放程式（Google Apps Script）
 *
 * 這段程式貼到 script.google.com 的新專案裡，部署成「網頁應用程式」。
 * 第一次執行會在你的 Google 雲端硬碟自動建立一份試算表
 * 「媽媽的照顧日誌－紀錄」，之後所有紀錄都存在那裡。
 */

var TZ = 'Asia/Taipei';
var BOOK_NAME = '媽媽的照顧日誌－紀錄';

function book_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SSID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) {}
  }
  var book = SpreadsheetApp.create(BOOK_NAME);
  props.setProperty('SSID', book.getId());
  return book;
}

function sheet_(name, head) {
  var book = book_();
  var sh = book.getSheetByName(name);
  if (!sh) {
    sh = book.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.getRange(1, 1, sh.getMaxRows(), head.length).setNumberFormat('@');
    sh.setFrozenRows(1);
  }
  return sh;
}

function days_()  { return sheet_('每日紀錄', ['日期', '摘要', '資料']); }
function cfg_()   { return sheet_('設定', ['項目', '資料']); }
function sugs_()  { return sheet_('姐妹建議', ['時間戳', '日期', '誰', '內容']); }
function chgs_()  { return sheet_('改藥紀錄', ['時間戳', '日期', '誰', '醫院醫生', '改了什麼']); }

/** 先執行一次這個，讓 Google 跳出授權畫面，並建立試算表。 */
function setup() {
  days_(); cfg_(); sugs_(); chgs_();
  var book = book_();
  var first = book.getSheetByName('工作表1') || book.getSheetByName('Sheet1');
  if (first && book.getSheets().length > 1) book.deleteSheet(first);
  Logger.log('試算表網址：' + book.getUrl());
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 文字開頭是 = + - @ 時加上單引號，避免被試算表當成公式。 */
function safe_(v) {
  var s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function dateKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '').trim();
}

function rows_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
}

function parse_(s) {
  try { return JSON.parse(s); } catch (e) { return null; }
}

function readAll_() {
  var days = {};
  var rows = rows_(days_());
  var from = Math.max(0, rows.length - 120);
  for (var i = from; i < rows.length; i++) {
    var key = dateKey_(rows[i][0]);
    var data = parse_(rows[i][2]);
    if (key && data) days[key] = data;
  }

  var cfg = null;
  rows_(cfg_()).forEach(function (r) {
    if (String(r[0]) === 'main') cfg = parse_(r[1]);
  });

  var sugs = rows_(sugs_()).map(function (r) {
    return { ts: Number(r[0]) || 0, date: dateKey_(r[1]), by: String(r[2]), text: String(r[3]) };
  }).sort(function (a, b) { return b.ts - a.ts; }).slice(0, 50);

  var chgs = rows_(chgs_()).map(function (r) {
    return { ts: Number(r[0]) || 0, date: dateKey_(r[1]), by: String(r[2]), doc: String(r[3]), what: String(r[4]) };
  }).sort(function (a, b) { return b.ts - a.ts; }).slice(0, 20);

  return { ok: true, days: days, cfg: cfg, sugs: sugs, chgs: chgs };
}

function doGet(e) {
  try {
    return out_(readAll_());
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var body = parse_(e.postData.contents) || {};

    if (body.action === 'saveDay') {
      var date = String(body.date || '');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return out_({ ok: false, error: 'date' });
      var sh = days_();
      var rows = rows_(sh);
      var rowNo = -1, cur = { date: date };
      for (var i = 0; i < rows.length; i++) {
        if (dateKey_(rows[i][0]) === date) { rowNo = i + 2; cur = parse_(rows[i][2]) || cur; break; }
      }
      var patch = body.patch || {};
      Object.keys(patch).forEach(function (k) { cur[k] = patch[k]; });
      cur.date = date;
      var line = [date, safe_(body.summary), JSON.stringify(cur)];
      if (rowNo > 0) sh.getRange(rowNo, 1, 1, 3).setValues([line]);
      else sh.appendRow(line);
      return out_({ ok: true });
    }

    if (body.action === 'setCfg') {
      var cs = cfg_();
      var line2 = ['main', JSON.stringify(body.cfg || {})];
      var crow = rows_(cs), found = -1;
      for (var j = 0; j < crow.length; j++) if (String(crow[j][0]) === 'main') found = j + 2;
      if (found > 0) cs.getRange(found, 1, 1, 2).setValues([line2]);
      else cs.appendRow(line2);
      return out_({ ok: true });
    }

    if (body.action === 'addSug') {
      var s = body.item || {};
      sugs_().appendRow([String(Number(s.ts) || Date.now()), safe_(s.date), safe_(s.by), safe_(s.text)]);
      return out_({ ok: true });
    }

    if (body.action === 'addChg') {
      var c = body.item || {};
      chgs_().appendRow([String(Number(c.ts) || Date.now()), safe_(c.date), safe_(c.by), safe_(c.doc), safe_(c.what)]);
      return out_({ ok: true });
    }

    return out_({ ok: false, error: 'action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
