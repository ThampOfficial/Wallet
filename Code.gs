var SPREADSHEET_ID = "14I4Bgj87pJX7hu6PL4M2DhKysAphvf6ArQMp88SDfA4";

// Cache settings: the full transaction list is stored in CacheService so most
// opens skip reading the sheet. Any write (doPost) clears it.
var CACHE_KEY = "tx_v1";
var CACHE_TTL = 21600;   // seconds (6 hours, the maximum)
var CACHE_CHUNK = 30000; // characters per cache entry (limit is 100KB each; Thai text is up to 3 bytes/char)

function getSheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
}

// Read the sheet and build the same objects the page has always received.
function readTransactions_() {
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  // A..J only, header row skipped
  var rows = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  var transactions = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row[0]) {
      // คอลัมน์จริงใน Google Sheet (A=0 ถึง J=9)
      var dateStr = row[0] ? String(row[0]).split('T')[0] : "";
      var timeStr = row[1] ? String(row[1]) : "";

      transactions.push({
        date: dateStr,
        time: timeStr,
        account: row[2] || "",      // C: Account
        category: row[3] || "",     // D: Category
        description: row[4] || "",  // E: Description
        payee: row[5] || "",        // F: Payee
        transfers: row[6] || "",    // G: Transfers
        amount: parseFloat(row[7]) || 0, // H: Amount
        currency: row[8] || "THB",  // I: Currency
        memo: row[9] || ""          // J: Memo
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

function cachePut_(json) {
  try {
    var obj = {};
    var n = 0;
    var start = 0;
    while (start < json.length) {
      var end = Math.min(start + CACHE_CHUNK, json.length);
      // never cut between the two halves of an emoji
      if (end < json.length) {
        var c = json.charCodeAt(end - 1);
        if (c >= 0xD800 && c <= 0xDBFF) end--;
      }
      obj[CACHE_KEY + "_" + n] = json.substring(start, end);
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

function doGet(e) {
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
  } else {
    return ContentService.createTextOutput(jsonString)
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var locked = false;
  try {
    locked = lock.tryLock(20000);
    var sheet = getSheet_();
    var data = JSON.parse(e.postData.contents);

    // แยกวันที่กับเวลาจาก datetime
    var dt = new Date(data.datetime || Date.now());
    var datePart = Utilities.formatDate(dt, Session.getScriptTimeZone(), "dd/MM/yyyy");
    var timePart = Utilities.formatDate(dt, Session.getScriptTimeZone(), "hh:mm a");

    // บันทึกลงแถวใหม่ เรียงตามคอลัมน์ A ถึง J เป๊ะๆ
    sheet.appendRow([
      datePart,             // A: Date
      timePart,             // B: Time
      data.account || "",   // C: Account
      data.category || "",  // D: Category
      data.description || "", // E: Description
      data.payee || "",     // F: Payee
      data.transfers || "", // G: Transfers
      data.amount || 0,     // H: Amount
      data.currency || "THB", // I: Currency
      data.memo || ""       // J: Memo
    ]);

    cacheClear_();

    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    cacheClear_();
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (locked) lock.releaseLock();
  }
}
