var SPREADSHEET_ID = "14I4Bgj87pJX7hu6PL4M2DhKysAphvf6ArQMp88SDfA4";

// ---------------------------------------------------------------------------
// Setup (once): Project Settings > Script properties > add API_KEY = a long random string.
// The page sends it as ?key=... on reads and as "key" in the POST body.
// If API_KEY is not set, access stays open (so nothing breaks before you set it).
//
// Cache: the full transaction list is kept in CacheService so most opens skip
// reading the sheet. Any write clears it.
//
// Encoding repair: text that was damaged on import (UTF-8 read as Windows-1252,
// e.g. "â–º" instead of "►") is repaired when served. To fix the sheet itself,
// run previewRepair() (changes nothing) and then applyRepair() from the editor.
// ---------------------------------------------------------------------------

var CACHE_KEY = "tx_v3";
var CACHE_TTL = 21600;   // seconds (6 hours, the maximum)
var CACHE_CHUNK = 30000; // characters per cache entry (limit is 100KB each; Thai text is up to 3 bytes/char)
var MAX_BODY = 10000;    // bytes accepted by doPost for a transaction
var MAX_SETTINGS = 300000; // characters accepted for a settings save (budgets, groups, categories, schedules...)
var SET_CHUNK = 2500;      // characters per script property (limit 9KB each; Thai text is up to 3 bytes/char)

// ---- App settings shared by every device (kept in Script Properties as chunks, not in the sheet) ----
function readSettings_() {
  var p = PropertiesService.getScriptProperties().getProperties();
  var n = parseInt(p.SET_n, 10) || 0, json = "";
  for (var i = 0; i < n; i++) json += p["SET_" + i] || "";
  return { ts: parseInt(p.SET_ts, 10) || 0, json: json };
}

function saveSettings_(ts, json) {
  var props = PropertiesService.getScriptProperties();
  var old = parseInt(props.getProperty("SET_n"), 10) || 0;
  var out = { SET_ts: String(parseInt(ts, 10) || Date.now()) }, n = 0;
  for (var i = 0; i < json.length; i += SET_CHUNK) { out["SET_" + n] = json.slice(i, i + SET_CHUNK); n++; }
  out.SET_n = String(n);
  props.setProperties(out);
  for (var j = n; j < old; j++) props.deleteProperty("SET_" + j);
}

// ---- Repair text whose UTF-8 bytes were read as Windows-1252/Latin-1 (e.g. "â–º" for "►", "à¸à¸¥" for Thai) ----
var CP1252_REV_ = {0x20AC:0x80,0x201A:0x82,0x0192:0x83,0x201E:0x84,0x2026:0x85,0x2020:0x86,0x2021:0x87,0x02C6:0x88,0x2030:0x89,0x0160:0x8A,0x2039:0x8B,0x0152:0x8C,0x017D:0x8E,0x2018:0x91,0x2019:0x92,0x201C:0x93,0x201D:0x94,0x2022:0x95,0x2013:0x96,0x2014:0x97,0x02DC:0x98,0x2122:0x99,0x0161:0x9A,0x203A:0x9B,0x0153:0x9C,0x017E:0x9E,0x0178:0x9F};

function strictUtf8_(b) {
  var out = '', i = 0, n = b.length;
  while (i < n) {
    var c = b[i++];
    if (c < 0x80) { out += String.fromCharCode(c); continue; }
    var need, cp;
    if (c >= 0xC2 && c <= 0xDF) { need = 1; cp = c & 0x1F; }
    else if (c >= 0xE0 && c <= 0xEF) { need = 2; cp = c & 0x0F; }
    else if (c >= 0xF0 && c <= 0xF4) { need = 3; cp = c & 0x07; }
    else return null;
    if (i + need > n) return null;
    for (var k = 0; k < need; k++) {
      var x = b[i++];
      if ((x & 0xC0) !== 0x80) return null;
      cp = (cp << 6) | (x & 0x3F);
    }
    if ((need === 2 && cp < 0x800) || (need === 3 && cp < 0x10000) || cp > 0x10FFFF || (cp >= 0xD800 && cp <= 0xDFFF)) return null;
    out += String.fromCodePoint(cp);
  }
  return out;
}

function fixMojibakeOnce_(s) {
  var bytes = [], high = false;
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i), b;
    if (c < 0x80) b = c;
    else if (c <= 0xFF) { b = c; high = true; }
    else if (CP1252_REV_[c] !== undefined) { b = CP1252_REV_[c]; high = true; }
    else return s; // contains real Thai/emoji/etc: not mojibake
    bytes.push(b);
  }
  if (!high) return s;
  var out = strictUtf8_(bytes);
  return out === null ? s : out;
}

function fixMojibake(s) {
  if (typeof s !== 'string' || !s) return s;
  for (var pass = 0; pass < 3; pass++) {
    var t = fixMojibakeOnce_(s);
    if (t === s) break;
    s = t;
  }
  return s;
}

function safeEq_(a, b) {
  a = String(a || ""); b = String(b || "");
  var r = a.length ^ b.length;
  for (var i = 0; i < Math.max(a.length, b.length); i++) r |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return r === 0;
}

function authOk_(key) {
  var expected = PropertiesService.getScriptProperties().getProperty("API_KEY");
  if (!expected) return true;
  return safeEq_(key, expected);
}

function deny_() {
  return json_({ status: "error", message: "unauthorized" });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
}

function cellText_(v, tz, fmt) {
  // a text column that Sheets turned into a real date (e.g. "1/5") comes back as a Date
  if (v instanceof Date) return Utilities.formatDate(v, tz, fmt || "dd/MM/yyyy");
  return v === null || v === undefined ? "" : fixMojibake(String(v));
}

function toAmount_(v) {
  if (typeof v === "number") return isFinite(v) ? v : 0;
  var n = parseFloat(String(v === null || v === undefined ? "" : v).replace(/[,\s\u202f\u00a0]/g, ""));
  return isFinite(n) ? n : 0;
}

// Read the sheet and build the objects the page receives.
function readTransactions_() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var tz = sheet.getParent().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  var rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues(); // A..J, header skipped
  // Date (A) and Time (B) are read exactly as shown in the sheet. Real Date cells would otherwise be
  // re-formatted here and swap day/month (e.g. shown "12/10/2026" = 12 Oct, but the cell holds 10 Dec),
  // and time-only cells come back as 1899 dates with a timezone offset that shifts the minutes.
  var shown = sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues();
  var transactions = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row[0]) {
      transactions.push({
        date: String(shown[i][0]).trim().split('T')[0],
        time: String(shown[i][1]).trim(),
        account: cellText_(row[2], tz, ""),      // C: Account
        category: cellText_(row[3], tz, ""),     // D: Category
        description: cellText_(row[4], tz, ""),  // E: Description
        payee: cellText_(row[5], tz, ""),        // F: Payee
        transfers: cellText_(row[6], tz, ""),    // G: Transfers
        amount: toAmount_(row[7]),               // H: Amount
        currency: cellText_(row[8], tz, "") || "THB", // I: Currency
        memo: cellText_(row[9], tz, "")          // J: Memo
      });
    }
  }
  return transactions;
}

function cacheGet_() {
  try {
    var cache = CacheService.getScriptCache();
    var n = parseInt(cache.get(CACHE_KEY + "_n"), 10);
    if (!n) return null;
    var keys = [];
    for (var i = 0; i < n; i++) keys.push(CACHE_KEY + "_" + i);
    var parts = cache.getAll(keys);
    var out = [];
    for (var j = 0; j < n; j++) {
      var p = parts[keys[j]];
      if (p === null || p === undefined) return null; // an entry expired or was evicted
      out.push(p);
    }
    return out.join("");
  } catch (err) {
    return null;
  }
}

function cachePut_(jsonString) {
  try {
    var obj = {};
    var n = 0;
    var start = 0;
    while (start < jsonString.length) {
      var end = Math.min(start + CACHE_CHUNK, jsonString.length);
      // never cut between the two halves of an emoji
      if (end < jsonString.length) {
        var c = jsonString.charCodeAt(end - 1);
        if (c >= 0xD800 && c <= 0xDBFF) end--;
      }
      obj[CACHE_KEY + "_" + n] = jsonString.substring(start, end);
      n++;
      start = end;
    }
    obj[CACHE_KEY + "_n"] = String(n);
    CacheService.getScriptCache().putAll(obj, CACHE_TTL);
  } catch (err) {
    // too big or cache unavailable: just run without cache
  }
}

function cacheClear_() {
  try {
    CacheService.getScriptCache().remove(CACHE_KEY + "_n");
  } catch (err) {}
}

// Run this from the editor (select clearCache, press Run) after editing the sheet by hand,
// so the page picks the change up immediately instead of waiting for the 6 h cache to expire.
function clearCache() {
  cacheClear_();
  Logger.log("Cache cleared");
}

function doGet(e) {
  if (!authOk_(e && e.parameter ? e.parameter.key : "")) return deny_();
  if (e && e.parameter && e.parameter.nocache === "1") cacheClear_();   // ...&nocache=1 forces a fresh read

  if (e && e.parameter && e.parameter.settings === "1") return json_({ settings: readSettings_() });

  var jsonString = cacheGet_();

  if (jsonString === null) {
    // Cache miss: rebuild under a lock so a write can't slip in between
    // reading the sheet and storing the cache.
    var lock = LockService.getScriptLock();
    var locked = lock.tryLock(20000);
    try {
      jsonString = locked ? cacheGet_() : null;
      if (jsonString === null) {
        jsonString = JSON.stringify({ transactions: readTransactions_() });
        if (locked) cachePut_(jsonString);
      }
    } finally {
      if (locked) lock.releaseLock();
    }
  }

  var callback = (e && e.parameter) ? e.parameter.callback : null;
  if (callback && /^[A-Za-z_$][A-Za-z0-9_$.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + jsonString + ');')
                         .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonString).setMimeType(ContentService.MimeType.JSON);
}

// Clean text for a cell: repair encoding, cap length. Cells are written as plain text
// (see doPost), so a value like "=IMPORTXML(...)" is stored as text, never run as a formula.
function textField_(v, max) {
  return fixMojibake(String(v === null || v === undefined ? "" : v)).slice(0, max);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    var raw = e && e.postData ? e.postData.contents : "";
    if (!raw || raw.length > MAX_SETTINGS) return json_({ status: "error", message: "bad request" });
    var data = JSON.parse(raw);
    if (!data || typeof data !== "object") return json_({ status: "error", message: "bad request" });
    if (!authOk_(data.key)) return deny_();

    if (data.action === "saveSettings") {
      if (typeof data.json !== "string") return json_({ status: "error", message: "bad request" });
      JSON.parse(data.json);   // must be valid JSON
      if (!lock.tryLock(20000)) return json_({ status: "error", message: "busy, try again" });
      locked = true;
      saveSettings_(data.ts, data.json);
      return json_({ status: "success" });
    }
    if (raw.length > MAX_BODY) return json_({ status: "error", message: "bad request" });

    var amount = toAmount_(data.amount);
    if (Math.abs(amount) > 1e12) return json_({ status: "error", message: "bad amount" });
    var dt = new Date(data.datetime || Date.now());
    if (isNaN(dt.getTime())) return json_({ status: "error", message: "bad date" });

    // แยกวันที่กับเวลาจาก datetime
    var tz = Session.getScriptTimeZone();
    var datePart = Utilities.formatDate(dt, tz, "dd/MM/yyyy");
    var timePart = Utilities.formatDate(dt, tz, "hh:mm a");

    locked = lock.tryLock(20000);
    if (!locked) return json_({ status: "error", message: "busy, try again" });

    var sheet = getSheet_();
    var r = sheet.getLastRow() + 1;
    sheet.getRange(r, 1, 1, 7).setNumberFormat("@");   // A-G as plain text
    sheet.getRange(r, 9, 1, 2).setNumberFormat("@");   // I-J as plain text
    sheet.getRange(r, 1, 1, 10).setValues([[
      datePart,                              // A: Date
      timePart,                              // B: Time
      textField_(data.account, 120),         // C: Account
      textField_(data.category, 120),        // D: Category
      textField_(data.description, 300),     // E: Description
      textField_(data.payee, 120),           // F: Payee
      textField_(data.transfers, 120),       // G: Transfers
      amount,                                // H: Amount
      textField_(data.currency || "THB", 8), // I: Currency
      textField_(data.memo, 500)             // J: Memo
    ]]);

    cacheClear_();
    return json_({ status: "success" });
  } catch (error) {
    cacheClear_();
    return json_({ status: "error", message: "server error" });
  } finally {
    if (locked) lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// One-time repair of text that was damaged when the data was imported.
// 1) File > Make a copy of the spreadsheet first.
// 2) Run previewRepair(): it only logs how many cells would change and shows examples.
// 3) Run applyRepair() to write the fixed text back.
// ---------------------------------------------------------------------------
var REPAIR_COLS_ = [3, 4, 5, 6, 7, 10]; // C, D, E, F, G, J (text columns)

function scanRepair_() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  var changes = [];
  if (lastRow < 2) return { sheet: sheet, changes: changes };
  var rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  for (var i = 0; i < rows.length; i++) {
    for (var k = 0; k < REPAIR_COLS_.length; k++) {
      var c = REPAIR_COLS_[k];
      var v = rows[i][c - 1];
      if (typeof v === "string" && v) {
        var f = fixMojibake(v);
        if (f !== v) changes.push({ row: i + 2, col: c, from: v, to: f });
      }
    }
  }
  return { sheet: sheet, changes: changes };
}

function previewRepair() {
  var res = scanRepair_();
  Logger.log("Cells that would be repaired: " + res.changes.length);
  for (var i = 0; i < Math.min(res.changes.length, 25); i++) {
    var c = res.changes[i];
    Logger.log("R" + c.row + "C" + c.col + ": " + c.from + "  =>  " + c.to);
  }
}

function applyRepair() {
  var res = scanRepair_();
  if (!res.changes.length) { Logger.log("Nothing to repair."); return; }
  var sheet = res.sheet;
  // write each affected column once, as plain text
  var byCol = {};
  res.changes.forEach(function (c) { (byCol[c.col] = byCol[c.col] || []).push(c); });
  var lastRow = sheet.getLastRow();
  Object.keys(byCol).forEach(function (col) {
    var rng = sheet.getRange(2, Number(col), lastRow - 1, 1);
    var vals = rng.getValues();
    byCol[col].forEach(function (c) { vals[c.row - 2][0] = c.to; });
    rng.setNumberFormat("@");
    rng.setValues(vals);
  });
  cacheClear_();
  Logger.log("Repaired cells: " + res.changes.length);
}
