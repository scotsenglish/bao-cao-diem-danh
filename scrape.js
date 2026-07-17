/**
 * scrape.js — Tự động lấy dữ liệu Attendance từ i-Learning LMS (Playwright)
 * và xuất ra data/latest.xlsx (3 sheet), sau đó build.py sẽ build ra index.html.
 *
 * Cách hoạt động:
 *   1. Đăng nhập LMS bằng Playwright (dùng chung LMS_LOGIN_ID / LMS_LOGIN_PASSWORD
 *      với scraper tra cứu điểm i-Learning hiện có — copy nguyên cơ chế login
 *      từ scrape.js của repo ilearning-tra-cuu-diem).
 *   2. Trong context của trang đã đăng nhập, gọi lại đúng các API nội bộ
 *      (.asmx) mà file export tay của bạn đang dùng — không có vấn đề CORS
 *      vì fetch() chạy same-origin bên trong trang.
 *   3. Lặp qua từng chi nhánh, gom dữ liệu, ghi checkpoint sau mỗi chi nhánh
 *      để có thể resume nếu job bị gián đoạn giữa chừng.
 *   4. Xuất ra data/latest.xlsx với 2 sheet TỔNG HỢP (nhỏ gọn, build.py dùng
 *      để dựng Tổng quan/Theo Lớp/Theo Học viên):
 *        - Student Summary
 *        - Class Summary Monthly
 *   5. Đẩy dữ liệu CHI TIẾT (Number of Student, List of Student — nặng hơn
 *      nhiều) lên Google Apps Script/Google Sheets thay vì ghi vào xlsx, để
 *      không vượt giới hạn 100MB của GitHub dù dữ liệu tích luỹ nhiều năm.
 *      Tab "Chi tiết" và "Tra cứu Học viên" trên dashboard đọc dữ liệu này
 *      trực tiếp từ Apps Script (xem APPS_SCRIPT_URL/APPS_SCRIPT_TOKEN).
 *
 * Biến môi trường cần có (GitHub Actions secrets):
 *   LMS_LOGIN_ID, LMS_LOGIN_PASSWORD   — dùng chung với scraper i-Learning điểm số
 *
 * Biến có thể chỉnh:
 *   MONTHS_BACK   — "all" (mặc định) = lấy TOÀN BỘ lịch sử từ khi công ty bắt
 *                   đầu dùng LMS. Hoặc 1 số (vd "3") = chỉ lấy N tháng gần nhất
 *                   (tháng hiện tại + (N-1) tháng trước) — nhanh hơn, dùng khi
 *                   không cần cập nhật lại toàn bộ lịch sử mỗi lần chạy.
 *   STAFF_ID      — id nhân viên dùng để gọi API (mặc định 9072, lấy từ script
 *                   export tay của bạn). Đổi qua biến môi trường STAFF_ID nếu cần.
 */

const { chromium } = require('playwright');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ---------------------------------------------------------------------------
// CẤU HÌNH
// ---------------------------------------------------------------------------
const STAFF_ID = process.env.STAFF_ID ? Number(process.env.STAFF_ID) : 9072;
// MONTHS_BACK: số (vd "3") = chỉ lấy N tháng gần nhất. "all" (mặc định) = lấy
// TOÀN BỘ lịch sử từ khi công ty bắt đầu dùng LMS tới hôm nay.
const MONTHS_BACK_RAW = (process.env.MONTHS_BACK || 'all').trim().toLowerCase();
// Giới hạn RIÊNG cho 2 sheet raw "Number of Student" / "List of Student" (tab
// "Chi tiết" trên dashboard) — LUÔN chỉ lấy vài tháng gần nhất, KHÔNG theo
// MONTHS_BACK ở trên, dù MONTHS_BACK="all". Lý do: dữ liệu raw theo từng buổi/
// từng học viên rất nặng (hàng trăm nghìn dòng nếu để cả năm), embed hết vào
// index.html sẽ vượt giới hạn 100MB của GitHub. Class Summary Monthly / Student
// Summary (đã tổng hợp, nhỏ gọn) vẫn lấy đủ theo MONTHS_BACK như bình thường.
const DETAIL_MONTHS_BACK = process.env.DETAIL_MONTHS_BACK ? Number(process.env.DETAIL_MONTHS_BACK) : 2;
const CONCURRENCY = process.env.SCRAPE_CONCURRENCY ? Number(process.env.SCRAPE_CONCURRENCY) : 3;
const PAGE_TIMEOUT_MS = 45_000;

// Mốc ngày an toàn để lấy "toàn bộ": đặt sớm hơn nhiều so với ngày công ty
// thực tế bắt đầu dùng i-Learning LMS, để chắc chắn không bỏ sót dữ liệu nào.
// LMS sẽ tự trả về rỗng cho khoảng thời gian trước khi có dữ liệu thật, nên
// đặt sớm hơn thực tế không gây hại gì, chỉ là dư ra không ảnh hưởng kết quả.
const ALL_HISTORY_ANCHOR = '2025-01-01';

const REPO_ROOT = path.join(__dirname);
const OUTPUT_XLSX = path.join(REPO_ROOT, 'data', 'latest.xlsx');
const CHECKPOINT_FILE = path.join(REPO_ROOT, 'data', '.attendance_checkpoint.json');

const LMS_BASE = 'https://lms.scotsenglish.edu.vn';

// Backend Google Apps Script (tab "Chi tiết" + "Tra cứu Học viên" trên dashboard
// đọc dữ liệu từ đây, không nhúng trực tiếp vào index.html nữa — tránh vượt
// giới hạn 100MB của GitHub khi dữ liệu tích luỹ lâu dài).
const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL || '';
const APPS_SCRIPT_TOKEN = process.env.APPS_SCRIPT_TOKEN || '';

// ---------------------------------------------------------------------------
// TÍNH KHOẢNG THỜI GIAN
//   - MONTHS_BACK = "all": từ ALL_HISTORY_ANCHOR đến hôm nay (toàn bộ lịch sử)
//   - MONTHS_BACK = số N: rolling window N tháng gần nhất (tháng hiện tại + (N-1) tháng trước)
// ---------------------------------------------------------------------------
function computeDateRange(monthsBackRaw) {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const endY = now.getFullYear();
  const endM = now.getMonth() + 1;
  const lastDayOfEndMonth = new Date(endY, endM, 0).getDate();
  const dateTo = `${endY}-${pad(endM)}-${pad(lastDayOfEndMonth)}`;

  if (monthsBackRaw === 'all') {
    return { dateFrom: ALL_HISTORY_ANCHOR, dateTo };
  }

  const monthsBack = Number(monthsBackRaw);
  const startDate = new Date(endY, endM - monthsBack, 1); // lùi (monthsBack-1) tháng trước tháng hiện tại
  const startY = startDate.getFullYear();
  const startM = startDate.getMonth() + 1;
  const dateFrom = `${startY}-${pad(startM)}-01`;

  return { dateFrom, dateTo };
}

// ---------------------------------------------------------------------------
// ĐĂNG NHẬP LMS — copy nguyên từ scrape.js của repo ilearning-tra-cuu-diem
// (cùng 1 LMS, cùng tài khoản, cùng cơ chế login).
// ---------------------------------------------------------------------------
const LOGIN_URL = 'https://lms.scotsenglish.edu.vn/login.html';

async function loginToLMS(page) {
  const loginId = process.env.LMS_LOGIN_ID;
  const loginPassword = process.env.LMS_LOGIN_PASSWORD;
  if (!loginId || !loginPassword) {
    throw new Error('Thiếu biến môi trường LMS_LOGIN_ID / LMS_LOGIN_PASSWORD (GitHub Secrets).');
  }

  // LMS có thể hiện alert() báo sai tài khoản/mật khẩu — bắt lại để log ra,
  // tránh Playwright bị treo chờ vô thời hạn vì dialog chưa được xử lý.
  page.on('dialog', async (dialog) => {
    console.log(`[DIALOG từ trang] ${dialog.type()}: ${dialog.message()}`);
    await dialog.dismiss().catch(() => {});
  });

  console.log('== Đăng nhập LMS ==');
  await page.goto(LOGIN_URL, { waitUntil: 'networkidle' });
  await page.waitForSelector('#login_id', { state: 'visible' });
  await page.fill('#login_id', loginId);
  await page.fill('#login_password', loginPassword);
  await page.click('#btn_login');

  try {
    // Đợi rời khỏi trang login (Angular xử lý login() rồi điều hướng đi nơi khác)
    await page.waitForFunction(() => !location.href.includes('login.html'), { timeout: 60000 });
    console.log('Đăng nhập OK. URL hiện tại:', page.url());
  } catch (err) {
    // Không đăng nhập được -> chụp lại màn hình + HTML tại thời điểm lỗi để debug,
    // vì không thể xem trực tiếp máy chạy Actions.
    console.log('== ĐĂNG NHẬP THẤT BẠI — đang lưu ảnh chụp + HTML để debug ==');
    const debugDir = path.join(REPO_ROOT, 'debug');
    fs.mkdirSync(debugDir, { recursive: true });
    await page.screenshot({ path: path.join(debugDir, 'login-failed.png'), fullPage: true }).catch(() => {});
    fs.writeFileSync(
      path.join(debugDir, 'login-failed.html'),
      await page.content().catch(() => '(không lấy được HTML)')
    );
    console.log('Đã lưu debug/login-failed.png và debug/login-failed.html');
    console.log('URL tại thời điểm lỗi:', page.url());
    throw err;
  }
}

// ---------------------------------------------------------------------------
// CÁC HÀM GỌI API — chạy trong page.evaluate() (same-origin, có cookie login)
// ---------------------------------------------------------------------------

async function fetchBranches(page, staffId) {
  return page.evaluate(async (staffId) => {
    const safeParse = (res) => {
      try { return JSON.parse(res.d.result).Table || []; } catch { return []; }
    };
    const res = await fetch('/data/setup.asmx/CounStaffBranch', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json;charset=UTF-8' },
      body: JSON.stringify({ staff: { stf_id: staffId }, setup: { ug_type: 'i-GARTEN' } }),
    });
    const rawBranches = safeParse(await res.json());
    return rawBranches.filter((b) => !String(b.brch_name || '').trim().toLowerCase().includes('test'));
  }, staffId);
}

async function fetchBranchData(page, { staffId, branch, dateFrom, dateTo, detailDateFrom }) {
  return page.evaluate(async ({ staffId, branch, dateFrom, dateTo, detailDateFrom }) => {
    const safeParse = (res) => {
      try { return JSON.parse(res.d.result).Table || []; } catch { return []; }
    };
    const formatMonth = (dateText) => {
      if (!dateText) return '';
      const d = new Date(dateText);
      if (isNaN(d)) return '';
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      return `${mm}/${d.getFullYear()}`;
    };
    // Dùng Date object để so sánh (không dùng so sánh chuỗi) — an toàn hơn vì
    // không phụ thuộc API trả về đúng định dạng "YYYY-MM-DD" hay không (có thể
    // là "YYYY-MM-DDTHH:mm:ss" hoặc định dạng khác). Nếu không parse được ngày,
    // coi như "gần đây" (giữ lại) để tránh lỡ tay lọc mất dữ liệu.
    const detailCutoffTime = detailDateFrom ? new Date(`${detailDateFrom}T00:00:00`).getTime() : null;
    const isRecentEnough = (dateText) => {
      if (!detailCutoffTime) return true;
      const t = new Date(dateText).getTime();
      if (isNaN(t)) return true;
      return t >= detailCutoffTime;
    };

    const semRes = await fetch('/data/setup.asmx/CounSemester', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json;charset=UTF-8' },
      body: JSON.stringify({ staff: { stf_id: staffId }, setup: { hr_brch_id: branch.brch_id } }),
    });
    const semesters = safeParse(await semRes.json());
    if (!semesters.length) return { numberRows: [], sessionRows: [], listRawRows: [], _debugCounts: { numberDataTotal: 0, listDataTotal: 0 } };
    const bsemId = semesters[0].bsem_id;

    const callReport = async (workType) => {
      const res = await fetch('/data/setup.asmx/ReportStuAttendanceList', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json;charset=UTF-8' },
        body: JSON.stringify({
          ret: {
            rt_work_type: workType,
            rt_bsem_id: bsemId,
            rt_date_from: dateFrom,
            rt_date_to: dateTo,
            rt_cors_id: 0,
            rt_instructor: 0,
            rt_syl_id: 0,
            rt_attend_type: '',
          },
        }),
      });
      return safeParse(await res.json());
    };

    // Gọi API với khoảng ngày ĐẦY ĐỦ (dateFrom -> dateTo, có thể là toàn bộ
    // lịch sử) — cần vậy để Class Summary Monthly / Student Summary tính đúng
    // tổng số liệu qua các tháng. numberData/listData ở đây có thể rất nhiều dòng.
    const numberData = await callReport('NUMBER');
    const listData = await callReport('LIST');

    // numberRows / listRawRows: bản RAW cho tab "Chi tiết" trên dashboard —
    // CHỈ giữ lại các dòng gần đây (>= detailDateFrom) để tránh nhúng hàng
    // trăm nghìn dòng lịch sử vào index.html gây vượt giới hạn 100MB của GitHub.
    // sessionRows (dùng để tính tổng) thì KHÔNG lọc, vẫn dùng full listData.
    const numberRows = numberData
      .filter((r) => isRecentEnough(r.Date))
      .map((r) => ({
        ...r,
        Branch: branch.brch_name,
        Month: formatMonth(r.Date),
      }));

    // sessionRows: dùng RIÊNG cho phần tính tổng (buildAggregates) — giữ nguyên
    // các field đã xác nhận đúng tên (Program/Class/ID/Student/Grade/Type),
    // không đổi để không ảnh hưởng số liệu Class Summary Monthly / Student Summary.
    // Dùng FULL listData (không lọc theo detailDateFrom) để tổng số liệu đúng
    // cho toàn bộ khoảng thời gian đã chọn (MONTHS_BACK).
    const sessionRows = listData.map((r) => ({
      Branch: branch.brch_name,
      Program: r.Program,
      Class: r.Class,
      StudentID: r.ID,
      Student: r.Student,
      Grade: r.Grade,
      Date: r.Date,
      Month: formatMonth(r.Date),
      Type: r.Type,
    }));

    // listRawRows: bản RAW đầy đủ của "List of Student" (No, Assigned Staff, Day,
    // Start Time, School, Instructor, Room, Reason, ...) để phục vụ tab "Chi tiết"
    // trên dashboard — hiển thị y hệt bảng "List of Student" trên LMS.
    // LƯU Ý: tên cột chính xác phụ thuộc vào field thật API trả về, mình chưa
    // xác nhận được hết (không có quyền xem DevTools của bạn), nên giữ nguyên
    // spread toàn bộ để không đoán sai tên field. Cũng lọc theo detailDateFrom
    // như numberRows, vì đây là phần tốn dung lượng nhiều nhất.
    const listRawRows = listData
      .filter((r) => isRecentEnough(r.Date))
      .map((r) => ({
        ...r,
        Branch: branch.brch_name,
        Month: formatMonth(r.Date),
      }));

    return {
      numberRows, sessionRows, listRawRows,
      _debugCounts: { numberDataTotal: numberData.length, listDataTotal: listData.length },
    };
  }, { staffId, branch, dateFrom, dateTo, detailDateFrom });
}

// ---------------------------------------------------------------------------
// GOM DỮ LIỆU (aggregation) — giống hệt logic script export tay của bạn
// ---------------------------------------------------------------------------
function buildAggregates(numberRows, sessionRows) {
  const studentSummaryMap = new Map();
  const classMonthlyMap = new Map();

  sessionRows.forEach((r) => {
    const type = String(r.Type || '').trim().toLowerCase();

    const studentKey = [r.Branch, r.Program, r.Class, r.StudentID].join('||');
    if (!studentSummaryMap.has(studentKey)) {
      studentSummaryMap.set(studentKey, {
        Branch: r.Branch, Program: r.Program, Class: r.Class,
        Student: r.Student, StudentID: r.StudentID, Grade: r.Grade,
        'Attendance Records': 0, Attendance: 0, Absence: 0, Late: 0,
      });
    }
    const s = studentSummaryMap.get(studentKey);
    s['Attendance Records']++;
    if (type === 'attendance') s.Attendance++;
    else if (type === 'absence') s.Absence++;
    else if (type === 'late') s.Late++;

    const classKey = [r.Branch, r.Program, r.Class, r.Month].join('||');
    if (!classMonthlyMap.has(classKey)) {
      classMonthlyMap.set(classKey, {
        Branch: r.Branch, Program: r.Program, Class: r.Class, Month: r.Month,
        'Attendance Records': 0, Attendance: 0, Absence: 0, Late: 0,
      });
    }
    const c = classMonthlyMap.get(classKey);
    c['Attendance Records']++;
    if (type === 'attendance') c.Attendance++;
    else if (type === 'absence') c.Absence++;
    else if (type === 'late') c.Late++;
  });

  const withPct = (r) => ({
    ...r,
    'Attendance %': r['Attendance Records'] ? r.Attendance / r['Attendance Records'] : 0,
    'Absence %': r['Attendance Records'] ? r.Absence / r['Attendance Records'] : 0,
    'Late %': r['Attendance Records'] ? r.Late / r['Attendance Records'] : 0,
  });

  return {
    studentSummaryRows: Array.from(studentSummaryMap.values()).map(withPct),
    classSummaryRows: Array.from(classMonthlyMap.values()).map(withPct),
  };
}

// ---------------------------------------------------------------------------
// CHECKPOINT — cho phép resume nếu job bị gián đoạn giữa chừng
// ---------------------------------------------------------------------------
function loadCheckpoint() {
  try {
    return JSON.parse(fs.readFileSync(CHECKPOINT_FILE, 'utf-8'));
  } catch {
    return { doneBranchIds: [], numberRows: [], sessionRows: [], listRawRows: [] };
  }
}
function saveCheckpoint(state) {
  fs.mkdirSync(path.dirname(CHECKPOINT_FILE), { recursive: true });
  fs.writeFileSync(CHECKPOINT_FILE, JSON.stringify(state));
}
function clearCheckpoint() {
  try { fs.unlinkSync(CHECKPOINT_FILE); } catch { /* không có file thì thôi */ }
}

// ---------------------------------------------------------------------------
// ĐẨY DỮ LIỆU CHI TIẾT LÊN GOOGLE APPS SCRIPT (thay cho việc nhúng vào HTML)
// ---------------------------------------------------------------------------
function chunkArray(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function yearOfMonthStr_(monthStr) {
  const parts = String(monthStr || '').split('/');
  return parts.length === 2 ? parts[1] : null;
}

async function appsScriptPost(action, payload) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(Object.assign({ action, token: APPS_SCRIPT_TOKEN }, payload)),
    redirect: 'follow',
  });
  const text = await res.text();
  try {
    return JSON.parse(text);
  } catch {
    throw new Error(`Phản hồi từ Apps Script không phải JSON hợp lệ (có thể do sai URL/token, hoặc deploy chưa đúng): ${text.slice(0, 300)}`);
  }
}

// Gộp rows theo (year, month) rồi gửi lên Apps Script — chia nhỏ theo lô để
// tránh 1 request quá lớn (Apps Script có giới hạn kích thước request).
async function pushMonthlySheet(sheetName, rows) {
  if (!rows.length) return;
  const byYear = {};
  rows.forEach((r) => {
    const year = yearOfMonthStr_(r.Month);
    if (!year) return;
    (byYear[year] = byYear[year] || []).push(r);
  });

  for (const year of Object.keys(byYear)) {
    const yearRows = byYear[year];
    const monthsInYear = Array.from(new Set(yearRows.map((r) => r.Month)));
    const chunks = chunkArray(yearRows, 3000);
    console.log(`   📤 Đẩy ${sheetName} năm ${year}: ${yearRows.length} dòng, ${monthsInYear.length} tháng, chia ${chunks.length} lô...`);
    for (let i = 0; i < chunks.length; i++) {
      // Chỉ báo "months" (để Apps Script xoá dữ liệu cũ trước khi ghi) ở LÔ ĐẦU
      // TIÊN của mỗi năm — các lô sau chỉ nối thêm (không xoá lại, tránh xoá
      // mất dữ liệu vừa ghi ở lô trước đó trong CÙNG 1 lần chạy này).
      const months = i === 0 ? monthsInYear : [];
      const result = await appsScriptPost('upsertMonths', { year, sheetName, months, rows: chunks[i] });
      if (result && result.error) throw new Error(`Apps Script lỗi (${sheetName}, năm ${year}, lô ${i + 1}/${chunks.length}): ${result.error}`);
    }
  }
}

// Xây danh sách "chỉ mục học viên" (student_id + lớp -> những năm có dữ liệu)
// từ listRawRows, để tab Tra cứu Học viên biết cần mở năm nào khi tìm kiếm.
function buildStudentIndexEntries(listRawRows) {
  const seen = new Set();
  const entries = [];
  listRawRows.forEach((r) => {
    const year = yearOfMonthStr_(r.Month);
    if (!year || !r.ID || !r.Class) return;
    const key = `${r.ID}||${r.Class}||${year}`;
    if (seen.has(key)) return;
    seen.add(key);
    entries.push({
      student_id: r.ID,
      name: r.Student || '',
      branch: r.Branch,
      program: r.Program,
      class_name: r.Class,
      year,
    });
  });
  return entries;
}

async function pushStudentIndex(entries) {
  if (!entries.length) return;
  const chunks = chunkArray(entries, 3000);
  console.log(`   📤 Đẩy StudentIndex: ${entries.length} mục, chia ${chunks.length} lô...`);
  for (let i = 0; i < chunks.length; i++) {
    const result = await appsScriptPost('upsertStudentIndex', { entries: chunks[i] });
    if (result && result.error) throw new Error(`Apps Script lỗi (StudentIndex, lô ${i + 1}/${chunks.length}): ${result.error}`);
  }
}

async function pushDetailToAppsScript(state) {
  if (!APPS_SCRIPT_URL || !APPS_SCRIPT_TOKEN) {
    console.log('⚠️  Chưa cấu hình APPS_SCRIPT_URL / APPS_SCRIPT_TOKEN — bỏ qua bước đẩy dữ liệu tab "Chi tiết"/"Tra cứu Học viên" (các tab này sẽ không có dữ liệu mới).');
    return;
  }
  console.log('== Đang đẩy dữ liệu chi tiết lên Google Apps Script ==');
  await pushMonthlySheet('NumberOfStudent', state.numberRows);
  await pushMonthlySheet('ListOfStudent', state.listRawRows);
  await pushStudentIndex(buildStudentIndexEntries(state.listRawRows));
  console.log('== Đẩy dữ liệu chi tiết xong ==');
}

// ---------------------------------------------------------------------------
// MAIN
// ---------------------------------------------------------------------------
async function main() {
  const { dateFrom, dateTo } = computeDateRange(MONTHS_BACK_RAW);
  console.log(`📅 Khoảng thời gian lấy dữ liệu (tổng hợp): ${dateFrom} → ${dateTo} (MONTHS_BACK=${MONTHS_BACK_RAW})`);

  // Khoảng ngày riêng cho 2 sheet raw (tab "Chi tiết") — luôn chỉ vài tháng gần
  // nhất, không phụ thuộc MONTHS_BACK ở trên (xem giải thích ở khai báo hằng số).
  const { dateFrom: detailDateFrom } = computeDateRange(String(DETAIL_MONTHS_BACK));
  console.log(`📅 Khoảng thời gian cho tab "Chi tiết" (raw): từ ${detailDateFrom} trở đi (DETAIL_MONTHS_BACK=${DETAIL_MONTHS_BACK})`);

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext();
  const page = await context.newPage();

  console.log('🔐 Đang đăng nhập LMS...');
  await loginToLMS(page);
  console.log('✅ Đăng nhập thành công.');

  const branches = await fetchBranches(page, STAFF_ID);
  console.log(`🏢 ${branches.length} chi nhánh (đã loại chi nhánh test).`);

  const state = loadCheckpoint();
  const doneIds = new Set(state.doneBranchIds);
  if (doneIds.size) {
    console.log(`↻ Resume từ checkpoint: đã có ${doneIds.size}/${branches.length} chi nhánh.`);
  }

  const remaining = branches.filter((b) => !doneIds.has(b.brch_id));

  // Mở thêm page cho mỗi "worker" để tăng tốc (cùng context nên vẫn dùng
  // chung session đăng nhập, không cần login lại). QUAN TRỌNG: mỗi page mới
  // phải điều hướng vào đúng domain LMS trước, nếu không fetch() với đường
  // dẫn tương đối ('/data/setup.asmx/...') sẽ lỗi vì page đang ở about:blank
  // (không có origin hợp lệ để ghép thành URL đầy đủ).
  const pages = [page];
  for (let i = 1; i < CONCURRENCY; i++) {
    const p = await context.newPage();
    await p.goto(LMS_BASE, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS });
    pages.push(p);
  }

  let idx = 0;
  const workers = pages.map((p) => (async () => {
    while (idx < remaining.length) {
      const branch = remaining[idx++];
      try {
        console.log(`➡️  ${branch.brch_name}`);
        const { numberRows, sessionRows, listRawRows, _debugCounts } = await fetchBranchData(p, {
          staffId: STAFF_ID, branch, dateFrom, dateTo, detailDateFrom,
        });
        state.numberRows.push(...numberRows);
        state.sessionRows.push(...sessionRows);
        state.listRawRows.push(...listRawRows);
        state.doneBranchIds.push(branch.brch_id);
        saveCheckpoint(state); // checkpoint sau MỖI chi nhánh
        console.log(
          `   ✔️  ${branch.brch_name}: number raw ${numberRows.length}/${_debugCounts.numberDataTotal} ` +
          `(sau lọc/tổng), list raw ${listRawRows.length}/${_debugCounts.listDataTotal}, session (không lọc)=${sessionRows.length}`
        );
      } catch (err) {
        console.error(`   ❌ Lỗi ở chi nhánh ${branch.brch_name}:`, err.message);
        // Không throw — chi nhánh lỗi sẽ được thử lại ở lần chạy job kế tiếp
        // (chưa được thêm vào doneBranchIds nên vẫn nằm trong "remaining" lần sau).
      }
    }
  })());
  await Promise.all(workers);

  await browser.close();

  console.log(`🎉 Xong. Tổng: ${state.numberRows.length} dòng Number, ${state.sessionRows.length} lượt điểm danh, ${state.listRawRows.length} dòng List of Student (raw).`);

  const { studentSummaryRows, classSummaryRows } = buildAggregates(state.numberRows, state.sessionRows);

  const wb = XLSX.utils.book_new();
  // Chỉ ghi 2 sheet TỔNG HỢP vào xlsx (build.py chỉ cần 2 sheet này để dựng
  // Tổng quan/Theo Lớp/Theo Học viên). "Number of Student" và "List of Student"
  // (dữ liệu raw, rất nặng — có thể lên tới hàng trăm nghìn dòng) KHÔNG ghi vào
  // đây nữa, đã chuyển hẳn sang Google Sheets qua Apps Script (pushDetailToAppsScript
  // ở trên) để tránh vượt giới hạn 100MB của GitHub.
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(studentSummaryRows), 'Student Summary');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(classSummaryRows), 'Class Summary Monthly');

  fs.mkdirSync(path.dirname(OUTPUT_XLSX), { recursive: true });
  XLSX.writeFile(wb, OUTPUT_XLSX);
  const xlsxSizeMB = fs.statSync(OUTPUT_XLSX).size / (1024 * 1024);
  console.log(`💾 Đã ghi ${OUTPUT_XLSX} (${xlsxSizeMB.toFixed(1)} MB)`);
  if (xlsxSizeMB > 80) {
    console.log(
      `⚠️  CẢNH BÁO: data/latest.xlsx đang ${xlsxSizeMB.toFixed(1)} MB, gần/vượt giới hạn 100MB của GitHub. ` +
      `Nếu bước "git push" ở cuối bị lỗi "File ... exceeds GitHub's file size limit", hãy giảm DETAIL_MONTHS_BACK ` +
      `(hiện đang là ${DETAIL_MONTHS_BACK} tháng) khi chạy lại workflow.`
    );
  }

  try {
    await pushDetailToAppsScript(state);
  } catch (err) {
    console.log(`⚠️  Đẩy dữ liệu lên Apps Script thất bại: ${err.message}`);
    console.log('   (data/latest.xlsx vẫn được ghi/commit bình thường — tab "Chi tiết"/"Tra cứu Học viên" sẽ tạm chưa cập nhật cho tới lần chạy sau).');
  }

  clearCheckpoint(); // job chạy xong trọn vẹn -> xoá checkpoint để lần sau chạy từ đầu
}

main().catch((err) => {
  console.error('LỖI KHÔNG XỬ LÝ ĐƯỢC:', err);
  process.exit(1);
});
