/**
 * Chăm Bé Na — nơi lưu chung cho trang lịch biểu.
 *
 * Dán toàn bộ file này vào Apps Script của một Google Sheet rồi Deploy thành
 * Web app. Hướng dẫn từng bước nằm ở README, mục "Dùng chung cho cả nhà".
 *
 * Trang web gọi vào đây hai kiểu:
 *   GET   -> trả về toàn bộ lịch đang có, kèm số phiên bản `rev`
 *   POST  -> gửi lên phần vừa sửa, ghi vào sheet rồi tăng `rev`
 *
 * Dữ liệu nằm thẳng trong sheet dưới dạng bảng, nên mở Google Sheet ra xem
 * hay sửa tay đều được.
 */

var SHEET_LICH = 'Lịch';
var SHEET_TGB = 'Thời gian biểu';
var SHEET_NGUOI = 'Người làm';

var HEADERS = ['Ngày', 'Thứ', 'Sáng', 'Người sáng', 'Trưa', 'Người trưa',
               'Chiều', 'Người chiều', 'Tối', 'Người tối', 'Notes', 'Special Note'];
var SLOTS = ['sang', 'trua', 'chieu', 'toi'];
var DAY_NAMES = ['Chủ Nhật', 'Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7'];

/* ---------- vào / ra ---------- */

/**
 * Trình duyệt nhiều khi chặn cuộc gọi thẳng sang Google vì khác tên miền.
 * Nên doGet nhận thêm hai tham số để trang web đi đường vòng:
 *   callback=tên   -> trả về JavaScript gọi hàm đó thay vì JSON thuần
 *   patch={...}    -> ghi luôn phần vừa sửa, giống hệt doPost
 */
function doGet(e) {
  var params = (e && e.parameter) || {};
  var out;

  if (params.patch) {
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      applyPatch(JSON.parse(params.patch));
      out = { ok: true, rev: bumpRev() };
    } catch (err) {
      out = { ok: false, error: String(err) };
    } finally {
      lock.releaseLock();
    }
  } else {
    out = readAll();
  }

  if (params.callback) {
    return ContentService
      .createTextOutput(params.callback + '(' + JSON.stringify(out) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json(out);
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  // hai người bấm cùng lúc thì xếp hàng, không ghi đè lẫn nhau
  lock.waitLock(20000);
  try {
    var body = JSON.parse(e.postData.contents);
    applyPatch(body.patch || {});
    return json({ ok: true, rev: bumpRev() });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- số phiên bản ---------- */

function getRev() {
  var v = PropertiesService.getDocumentProperties().getProperty('rev');
  return v ? Number(v) : 0;
}

function bumpRev() {
  var next = getRev() + 1;
  PropertiesService.getDocumentProperties().setProperty('rev', String(next));
  return next;
}

/* ---------- sheet ---------- */

function sheetNamed(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (header) {
      sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

// Ngày có thể là chuỗi "2026-08-10" hoặc ô ngày tháng thật của Sheet
function isoOf(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(value).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function readAll() {
  var days = {};
  var sh = sheetNamed(SHEET_LICH, HEADERS);
  var rows = sh.getLastRow() - 1;
  if (rows > 0) {
    var values = sh.getRange(2, 1, rows, HEADERS.length).getValues();
    for (var r = 0; r < values.length; r++) {
      var iso = isoOf(values[r][0]);
      if (!iso) continue;
      var day = {};
      for (var s = 0; s < SLOTS.length; s++) {
        var text = String(values[r][2 + s * 2] || '').trim();
        var who = String(values[r][3 + s * 2] || '').trim();
        if (text) day[SLOTS[s]] = text;
        if (who) day[SLOTS[s] + 'By'] = who;
      }
      var notes = String(values[r][10] || '').trim();
      var special = String(values[r][11] || '').trim();
      if (notes) day.notes = notes;
      if (special) day.special = special;
      if (Object.keys(day).length) days[iso] = day;
    }
  }

  return {
    rev: getRev(),
    state: {
      days: days,
      routine: readColumn(SHEET_TGB, 'Thời gian biểu').join('\n'),
      people: readColumn(SHEET_NGUOI, 'Tên')
    }
  };
}

function readColumn(name, header) {
  var sh = sheetNamed(name, [header]);
  var rows = sh.getLastRow() - 1;
  if (rows <= 0) return [];
  var values = sh.getRange(2, 1, rows, 1).getValues();
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var line = String(values[i][0] || '').trim();
    if (line) out.push(line);
  }
  return out;
}

function writeColumn(name, header, lines) {
  var sh = sheetNamed(name, [header]);
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).clearContent();
  }
  if (!lines.length) return;
  var cells = lines.map(function (line) { return [line]; });
  sh.getRange(2, 1, cells.length, 1).setValues(cells);
}

/* ---------- ghi phần vừa sửa ---------- */

function applyPatch(patch) {
  if (patch.days) writeDays(patch.days);
  if (typeof patch.routine === 'string') {
    writeColumn(SHEET_TGB, 'Thời gian biểu', patch.routine.split('\n'));
  }
  if (patch.people) {
    writeColumn(SHEET_NGUOI, 'Tên', patch.people.filter(function (n) {
      return String(n).trim() !== '';
    }));
  }
}

function writeDays(days) {
  var sh = sheetNamed(SHEET_LICH, HEADERS);
  var lastRow = sh.getLastRow();

  // ngày nào đang ở dòng nào
  var rowOf = {};
  if (lastRow > 1) {
    var col = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < col.length; i++) {
      var iso = isoOf(col[i][0]);
      if (iso) rowOf[iso] = i + 2;
    }
  }

  var appended = [];
  for (var key in days) {
    if (!days.hasOwnProperty(key) || !/^\d{4}-\d{2}-\d{2}$/.test(key)) continue;
    var row = rowFor(key, days[key]);
    if (rowOf[key]) {
      sh.getRange(rowOf[key], 1, 1, HEADERS.length).setValues([row]);
    } else {
      appended.push(row);
    }
  }

  if (appended.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appended.length, HEADERS.length).setValues(appended);
    sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length)
      .sort({ column: 1, ascending: true });
  }
}

function rowFor(iso, day) {
  var parts = iso.split('-');
  var d = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  var row = [iso, DAY_NAMES[d.getDay()]];
  for (var s = 0; s < SLOTS.length; s++) {
    row.push(day[SLOTS[s]] || '');
    row.push(day[SLOTS[s] + 'By'] || '');
  }
  row.push(day.notes || '');
  row.push(day.special || '');
  return row;
}
