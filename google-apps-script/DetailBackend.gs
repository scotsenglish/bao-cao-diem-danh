/**
 * Code.gs — "Backend" cho tab Chi tiết + Tra cứu Học viên của Attendance Dashboard.
 * PHIÊN BẢN 2: tách dữ liệu theo THÁNG (mỗi tháng 1 tab riêng), thay vì gộp
 * chung theo năm như bản đầu — để tốc độ đọc/ghi không chậm dần theo thời
 * gian dù dữ liệu tích luỹ nhiều năm.
 *
 * ============================================================================
 * CÁCH DỰNG (giống bản cũ, không đổi gì về Sheets/deploy):
 * ============================================================================
 * 1. Sheet "Attendance Index" với 2 tab: "YearMap" (Year | SpreadsheetID) và
 *    "StudentIndex" (tự tạo, để trống).
 * 2. Mỗi năm 1 Google Sheet riêng (ví dụ "Attendance Data 2025"), ID điền vào
 *    tab "YearMap". Không cần tạo tab con nào bên trong — script tự tạo tab
 *    theo tháng khi ghi dữ liệu lần đầu (ví dụ "NumberOfStudent_2026-07").
 * 3. Script Properties cần có: WRITE_TOKEN, INDEX_SPREADSHEET_ID.
 * 4. Deploy dạng Web App, Execute as "Me", Who has access "Anyone".
 *
 * FIX (xem actionStudentSessions_ bên dưới): dữ liệu thô trong sheet
 * "ListOfStudent_yyyy-mm" được scrape.js ghi NGUYÊN field name gốc từ LMS
 * (spread "...r" không đổi tên) — mã học viên trong đó tên cột là "ID", KHÔNG
 * phải "StudentID" (khác với sheet StudentIndex, nơi cột thật sự tên
 * "StudentID"). Bản trước lọc theo r.StudentID trên sheet ListOfStudent nên
 * luôn undefined, không bao giờ khớp -> rows luôn trả về rỗng dù học viên có
 * dữ liệu thật. Đã sửa thành r.ID ở dòng filter tương ứng.
 * ============================================================================
 */

function getIndexSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('INDEX_SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error(
      'Không xác định được Spreadsheet "Attendance Index". Vào Project Settings > ' +
      'Script Properties, thêm INDEX_SPREADSHEET_ID = ID của sheet "Attendance Index".'
    );
  }
  return active;
}

const YEAR_MAP_SHEET = 'YearMap';
const STUDENT_INDEX_SHEET = 'StudentIndex';
// Đổi "Years" -> "Months": giờ cần biết chính xác THÁNG NÀO (không chỉ năm nào)
// có dữ liệu của học viên, để tra cứu chỉ mở đúng tab tháng đó thay vì quét
// nguyên năm.
const STUDENT_INDEX_HEADER = ['StudentID', 'Name', 'Branch', 'Program', 'ClassName', 'Months'];

// ----------------------------------------------------------------------------
// TIỆN ÍCH CHUNG
// ----------------------------------------------------------------------------

function getWriteToken_() {
  return PropertiesService.getScriptProperties().getProperty('WRITE_TOKEN');
}

function yearOfMonth_(monthStr) {
  // monthStr dạng "mm/yyyy" (vd "07/2026") -> trả về "2026"
  const parts = String(monthStr || '').split('/');
  return parts.length === 2 ? parts[1] : null;
}

function monthTag_(monthStr) {
  // "07/2026" -> "2026-07" — dùng làm hậu tố tên tab (dễ sắp xếp theo thời gian
  // trên giao diện Google Sheets hơn là giữ nguyên "07/2026").
  const parts = String(monthStr || '').split('/');
  if (parts.length !== 2) return null;
  const mm = parts[0];
  const yyyy = parts[1];
  return `${yyyy}-${mm.padStart(2, '0')}`;
}

function monthTagToSlashFormat_(tag) {
  // "2026-07" -> "07/2026"
  const parts = String(tag || '').split('-');
  if (parts.length !== 2) return null;
  const yyyy = parts[0];
  const mm = parts[1];
  return `${mm}/${yyyy}`;
}

function monthTabName_(sheetBaseName, month) {
  const tag = monthTag_(month);
  return tag ? `${sheetBaseName}_${tag}` : null;
}

function getYearMap_() {
  const sheet = getIndexSpreadsheet_().getSheetByName(YEAR_MAP_SHEET);
  if (!sheet) throw new Error('Không tìm thấy tab "YearMap" trong Sheet "Attendance Index".');
  const data = sheet.getDataRange().getValues();
  const map = {};
  data.forEach((row) => {
    const year = String(row[0] || '').trim();
    const id = String(row[1] || '').trim();
    if (year && id) map[year] = id;
  });
  return map;
}

function getYearSpreadsheet_(year) {
  const map = getYearMap_();
  const id = map[String(year)];
  if (!id) throw new Error(`Chưa có Spreadsheet cho năm ${year} trong tab "YearMap". Vào Sheet "Attendance Index" > tab YearMap, thêm 1 dòng: ${year} | <ID spreadsheet của năm đó>.`);
  return SpreadsheetApp.openById(id);
}

function getOrCreateSheet_(spreadsheet, name) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  return sheet;
}

function getStudentIndexSheet_() {
  let sheet = getIndexSpreadsheet_().getSheetByName(STUDENT_INDEX_SHEET);
  if (!sheet) sheet = getIndexSpreadsheet_().insertSheet(STUDENT_INDEX_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(STUDENT_INDEX_HEADER);
  }
  return sheet;
}

// Đọc toàn bộ 1 sheet thành mảng object (dùng dòng đầu làm tên cột)
function sheetToObjects_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];
  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0];
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every((v) => v === '' || v === null)) continue;
    const obj = {};
    header.forEach((h, idx) => { if (h) obj[h] = row[idx]; });
    out.push(obj);
  }
  return out;
}

function jsonOut_(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(`${callback}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------------------
// doGet — các API đọc dữ liệu, gọi từ dashboard (JSONP, có ?callback=...)
// ----------------------------------------------------------------------------
function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = params.action;
  const callback = params.callback || null;

  try {
    let result;
    switch (action) {
      case 'availableMonths':
        result = actionAvailableMonths_();
        break;
      case 'monthNumber':
        result = actionMonthData_(params.month, 'NumberOfStudent');
        break;
      case 'monthList':
        result = actionMonthData_(params.month, 'ListOfStudent');
        break;
      case 'studentSearch':
        result = actionStudentSearch_(params.q || '');
        break;
      case 'studentSessions':
        result = actionStudentSessions_(params.student_id, params.class_name);
        break;
      default:
        result = {
          error: `Không nhận diện được action: ${action}`,
          debug_parameter: params,
          debug_queryString: (e && e.queryString) || null,
          debug_hasEvent: !!e,
        };
    }
    return jsonOut_(result, callback);
  } catch (err) {
    return jsonOut_({ error: String(err && err.message ? err.message : err) }, callback);
  }
}

// Liệt kê tháng có dữ liệu bằng cách đọc TÊN TAB (không cần quét dữ liệu bên
// trong) — vd tab "NumberOfStudent_2026-07" -> tháng "07/2026". Nhanh vì chỉ
// đọc danh sách tên sheet, không đọc nội dung.
function actionAvailableMonths_() {
  const yearMap = getYearMap_();
  const months = new Set();
  const debugPerYear = {};
  Object.keys(yearMap).forEach((year) => {
    try {
      const ss = SpreadsheetApp.openById(yearMap[year]);
      const sheets = ss.getSheets();
      let countThisYear = 0;
      sheets.forEach((sheet) => {
        const name = sheet.getName();
        const m = name.match(/^NumberOfStudent_(\d{4}-\d{2})$/);
        if (m) {
          const monthStr = monthTagToSlashFormat_(m[1]);
          if (monthStr) { months.add(monthStr); countThisYear++; }
        }
      });
      debugPerYear[year] = `OK — ${countThisYear} tab tháng tìm thấy`;
    } catch (err) {
      debugPerYear[year] = `LỖI: ${String(err && err.message ? err.message : err)}`;
    }
  });
  const sorted = Array.from(months).sort((a, b) => {
    const pa = a.split('/'); const ma = Number(pa[0]); const ya = Number(pa[1]);
    const pb = b.split('/'); const mb = Number(pb[0]); const yb = Number(pb[1]);
    return ya !== yb ? ya - yb : ma - mb;
  });
  return { months: sorted, debug_perYear: debugPerYear, debug_yearMap: yearMap };
}

// Đọc trực tiếp ĐÚNG 1 tab của đúng tháng được yêu cầu — không cần lọc gì cả,
// vì tab đó vốn chỉ chứa dữ liệu của tháng đó.
function actionMonthData_(month, sheetBaseName) {
  if (!month) return { error: 'Thiếu tham số month (định dạng mm/yyyy)' };
  const year = yearOfMonth_(month);
  if (!year) return { error: `Tham số month không hợp lệ: ${month}` };
  let ss;
  try {
    ss = getYearSpreadsheet_(year);
  } catch (err) {
    return { rows: [], note: String(err.message) };
  }
  const tabName = monthTabName_(sheetBaseName, month);
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return { rows: [] };
  const rows = sheetToObjects_(sheet);
  return { rows, total: rows.length };
}

function actionStudentSearch_(q) {
  const term = String(q || '').trim().toLowerCase();
  if (!term) return { results: [] };
  const sheet = getStudentIndexSheet_();
  const all = sheetToObjects_(sheet);
  const results = [];
  for (const r of all) {
    const hay = `${r.Name || ''} ${r.StudentID || ''}`.toLowerCase();
    if (hay.includes(term)) {
      results.push(r);
      if (results.length >= 30) break;
    }
  }
  return { results };
}

// Tra cứu 1 học viên: dùng danh sách "Months" trong StudentIndex để biết CHÍNH
// XÁC cần mở những tab tháng nào (không phải quét cả năm như bản trước).
function actionStudentSessions_(studentId, className) {
  if (!studentId || !className) return { error: 'Thiếu student_id hoặc class_name' };
  const sheet = getStudentIndexSheet_();
  const all = sheetToObjects_(sheet);
  const idxRow = all.find((r) => String(r.StudentID) === String(studentId) && String(r.ClassName) === String(className));
  if (!idxRow) return { rows: [] };
  const months = String(idxRow.Months || '').split(',').map((m) => m.trim()).filter(Boolean);
  let rows = [];
  months.forEach((month) => {
    try {
      const year = yearOfMonth_(month);
      if (!year) return;
      const ss = getYearSpreadsheet_(year);
      const tabName = monthTabName_('ListOfStudent', month);
      const sheetL = ss.getSheetByName(tabName);
      if (!sheetL) return;
      const monthRows = sheetToObjects_(sheetL);
      // FIX: sheet "ListOfStudent_*" chứa dữ liệu thô ghi nguyên field name gốc
      // từ LMS (scrape.js spread "...r" không đổi tên) -> mã học viên ở đây tên
      // cột là "ID", KHÔNG phải "StudentID" (đối chiếu scrape.js: sessionRows
      // map StudentID: r.ID). Trước đây lọc theo r.StudentID nên luôn
      // undefined, không bao giờ khớp -> rows luôn rỗng dù học viên có dữ liệu.
      rows = rows.concat(monthRows.filter((r) => String(r.ID) === String(studentId) && String(r.Class) === String(className)));
    } catch (err) {
      // bỏ qua tháng lỗi, không chặn các tháng khác
    }
  });
  rows.sort((a, b) => String(a.Date).localeCompare(String(b.Date)));
  return { rows, student: idxRow };
}

// ----------------------------------------------------------------------------
// doPost — API ghi dữ liệu, chỉ scrape.js gọi (có kèm token bí mật)
// ----------------------------------------------------------------------------
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.token !== getWriteToken_()) {
      return jsonOut_({ error: 'Sai token, từ chối ghi dữ liệu.' });
    }

    let result;
    switch (body.action) {
      case 'replaceMonth':
        result = actionReplaceMonth_(body);
        break;
      case 'upsertStudentIndex':
        result = actionUpsertStudentIndex_(body);
        break;
      default:
        result = { error: `Không nhận diện được action: ${body.action}` };
    }
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ error: String(err && err.message ? err.message : err) });
  }
}

// body: { token, month: "07/2026", sheetName: 'NumberOfStudent'|'ListOfStudent',
//         replace: true/false, rows: [...] }
// replace=true (lô ĐẦU của tháng này trong lượt push hiện tại): xoá sạch tab
//   tháng đó rồi ghi mới — vì tab chỉ chứa đúng 1 tháng nên đây là thao tác
//   NHẸ, không phụ thuộc đã có bao nhiêu năm dữ liệu khác.
// replace=false (các lô tiếp theo của CÙNG tháng này): chỉ nối thêm vào cuối,
//   không đọc/xoá gì cả — nhanh.
function actionReplaceMonth_(body) {
  const month = body.month;
  const sheetName = body.sheetName;
  const replace = body.replace;
  const rows = body.rows;
  if (!month || !sheetName || !Array.isArray(rows)) {
    return { error: 'Thiếu tham số (month/sheetName/rows).' };
  }
  const year = yearOfMonth_(month);
  if (!year) return { error: `Tham số month không hợp lệ: ${month}` };

  const ss = getYearSpreadsheet_(year);
  const tabName = monthTabName_(sheetName, month);
  const sheet = getOrCreateSheet_(ss, tabName);

  const lastRow = sheet.getLastRow();
  let header;

  if (replace) {
    // Xoá sạch toàn bộ tab (chỉ của đúng tháng này) rồi ghi lại từ đầu.
    if (lastRow > 0) {
      sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).clearContent();
    }
    header = rows.length > 0 ? Object.keys(rows[0]) : [];
    if (header.length > 0) {
      sheet.appendRow(header);
      sheet.getRange(1, 1, sheet.getMaxRows(), header.length).setNumberFormat('@');
    }
    if (rows.length > 0) {
      const rowArrays = rows.map((r) => header.map((h) => (r[h] !== undefined ? r[h] : '')));
      sheet.getRange(2, 1, rowArrays.length, header.length).setNumberFormat('@');
      sheet.getRange(2, 1, rowArrays.length, header.length).setValues(rowArrays);
    }
    return { ok: true, written: rows.length, mode: 'replace' };
  }

  // Nối thêm (không xoá gì) — dùng cho các lô sau của cùng 1 tháng.
  if (lastRow === 0) {
    header = rows.length > 0 ? Object.keys(rows[0]) : [];
    if (header.length > 0) {
      sheet.appendRow(header);
      sheet.getRange(1, 1, sheet.getMaxRows(), header.length).setNumberFormat('@');
    }
  } else {
    const lastCol = sheet.getLastColumn();
    header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  }
  if (rows.length > 0) {
    const rowArrays = rows.map((r) => header.map((h) => (r[h] !== undefined ? r[h] : '')));
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowArrays.length, header.length).setNumberFormat('@');
    sheet.getRange(startRow, 1, rowArrays.length, header.length).setValues(rowArrays);
  }
  return { ok: true, written: rows.length, mode: 'append-only' };
}

// body: { token, entries: [{student_id, name, branch, program, class_name, month}, ...] }
function actionUpsertStudentIndex_(body) {
  const entries = body.entries;
  if (!Array.isArray(entries)) return { error: 'Thiếu entries' };
  const sheet = getStudentIndexSheet_();
  const all = sheetToObjects_(sheet);
  const keyOf = function(sid, cls) { return `${sid}||${cls}`; };
  const rowIndexByKey = {};
  all.forEach((r, i) => { rowIndexByKey[keyOf(r.StudentID, r.ClassName)] = i; });

  const merged = {};
  entries.forEach((entry) => {
    const key = keyOf(entry.student_id, entry.class_name);
    if (!merged[key]) {
      merged[key] = Object.assign({}, entry, { months: new Set([String(entry.month)]) });
    } else {
      merged[key].months.add(String(entry.month));
    }
  });

  let updated = 0, inserted = 0;
  const toAppend = [];

  Object.keys(merged).forEach((key) => {
    const entry = merged[key];
    if (key in rowIndexByKey) {
      const rowIdx = rowIndexByKey[key];
      const existing = all[rowIdx];
      const months = new Set(String(existing.Months || '').split(',').map((m) => m.trim()).filter(Boolean));
      entry.months.forEach((m) => months.add(m));
      const newMonths = Array.from(months).sort((a, b) => {
        const pa = a.split('/'); const ma = Number(pa[0]); const ya = Number(pa[1]);
        const pb = b.split('/'); const mb = Number(pb[0]); const yb = Number(pb[1]);
        return ya !== yb ? ya - yb : ma - mb;
      }).join(',');
      sheet.getRange(rowIdx + 2, STUDENT_INDEX_HEADER.indexOf('Months') + 1).setValue(newMonths);
      updated++;
    } else {
      const monthsStr = Array.from(entry.months).join(',');
      toAppend.push([entry.student_id, entry.name, entry.branch, entry.program, entry.class_name, monthsStr]);
      inserted++;
    }
  });

  if (toAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, STUDENT_INDEX_HEADER.length).setValues(toAppend);
  }

  return { ok: true, updated, inserted };
}
