/**
 * scrape.js — Tự động lấy dữ liệu Attendance từ i-Learning LMS (Playwright)
 * và xuất ra data/latest.xlsx (4 sheet), sau đó build.py sẽ build ra index.html.
 *
 * Cách hoạt động:
 *   1. Đăng nhập LMS bằng Playwright (dùng chung LMS_USERNAME / LMS_PASSWORD
 *      với scraper tra cứu điểm i-Learning hiện có).
 *   2. Trong context của trang đã đăng nhập, gọi lại đúng các API nội bộ
 *      (.asmx) mà file export tay của bạn đang dùng — không có vấn đề CORS
 *      vì fetch() chạy same-origin bên trong trang.
 *   3. Lặp qua từng chi nhánh, gom dữ liệu, ghi checkpoint sau mỗi chi nhánh
 *      để có thể resume nếu job bị gián đoạn giữa chừng.
 *   4. Xuất ra data/latest.xlsx với 3 sheet:
 *        - Number of Student
 *        - Student Summary
 *        - Class Summary Monthly
 *
 * Biến môi trường cần có (GitHub Actions secrets):
 *   LMS_USERNAME, LMS_PASSWORD   — dùng chung với scraper i-Learning điểm số
 *
 * Biến có thể chỉnh:
 *   MONTHS_BACK   — số tháng lùi về (mặc định 3: tháng hiện tại + 2 tháng trước).
 *                   Tăng số này nếu muốn dashboard tổng hợp được xa hơn về quá khứ,
 *                   nhưng lưu ý index.html sẽ phình to hơn tương ứng.
 *   STAFF_ID      — id nhân viên dùng để gọi API (mặc định 9072, lấy từ script
 *                   export tay của bạn). Đổi qua biến môi trường STAFF_ID nếu cần.
 *
 * *** PHẦN CẦN BẠN ĐIỀN: hàm loginToLMS() bên dưới ***
 * Mình chưa có URL/selector đăng nhập thật của lms.scotsenglish.edu.vn (không
 * truy cập được hệ thống nội bộ của bạn). Vì đây là CÙNG một LMS mà scraper
 * i-Learning tra cứu điểm của bạn đang đăng nhập thành công, cách nhanh nhất
 * là copy nguyên phần code đăng nhập từ scrape.js của repo i-Learning
 * (ilearning-tra-cuu-diem) dán vào hàm loginToLMS() bên dưới. Phần còn lại
 * của file này không cần đụng vào.
 */

const { chromium } = require('playwright');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ---------------------------------------------------------------------------
// CẤU HÌNH
// ---------------------------------------------------------------------------
const STAFF_ID = process.env.STAFF_ID ? Number(process.env.STAFF_ID) : 9072;
const MONTHS_BACK = process.env.MONTHS_BACK ? Number(process.env.MONTHS_BACK) : 3;
const CONCURRENCY = process.env.SCRAPE_CONCURRENCY ? Number(process.env.SCRAPE_CONCURRENCY) : 3;
const PAGE_TIMEOUT_MS = 45_000;

const REPO_ROOT = path.join(__dirname);
const OUTPUT_XLSX = path.join(REPO_ROOT, 'data', 'latest.xlsx');
const CHECKPOINT_FILE = path.join(REPO_ROOT, 'data', '.attendance_checkpoint.json');

const LMS_BASE = 'https://lms.scotsenglish.edu.vn';

// ---------------------------------------------------------------------------
// TÍNH KHOẢNG THỜI GIAN (rolling window MONTHS_BACK tháng gần nhất)
// ---------------------------------------------------------------------------
function computeDateRange(monthsBack) {
  const now = new Date();
  const endY = now.getFullYear();
  const endM = now.getMonth() + 1; // 1-12, tháng hiện tại
  const startDate = new Date(endY, endM - monthsBack, 1); // lùi (monthsBack-1) tháng trước tháng hiện tại
  const startY = startDate.getFullYear();
  const startM = startDate.getMonth() + 1;

  const pad = (n) => String(n).padStart(2, '0');
  const dateFrom = `${startY}-${pad(startM)}-01`;
  const lastDayOfEndMonth = new Date(endY, endM, 0).getDate();
  const dateTo = `${endY}-${pad(endM)}-${pad(lastDayOfEndMonth)}`;

  return { dateFrom, dateTo };
}

// ---------------------------------------------------------------------------
// ĐĂNG NHẬP LMS — *** CẦN BẠN THAY BẰNG CODE ĐĂNG NHẬP THẬT ***
// Copy y nguyên phần login từ scrape.js của repo ilearning-tra-cuu-diem,
// vì đây là cùng 1 hệ thống LMS, cùng tài khoản (LMS_USERNAME/LMS_PASSWORD).
// ---------------------------------------------------------------------------
async function loginToLMS(page) {
  const username = process.env.LMS_USERNAME;
  const password = process.env.LMS_PASSWORD;
  if (!username || !password) {
    throw new Error('Thiếu LMS_USERNAME / LMS_PASSWORD trong biến môi trường (GitHub Secrets).');
  }

  // TODO: thay các dòng dưới bằng đúng URL trang login + selector thật.
  // Đây chỉ là khung mẫu dựa theo pattern login LMS phổ biến.
  await page.goto(`${LMS_BASE}/login.aspx`, { waitUntil: 'domcontentloaded', timeout: PAGE_TIMEOUT_MS });
  await page.fill('#txtUsername', username);      // TODO: đổi selector đúng
  await page.fill('#txtPassword', password);      // TODO: đổi selector đúng
  await page.click('#btnLogin');                  // TODO: đổi selector đúng
  await page.waitForLoadState('networkidle', { timeout: PAGE_TIMEOUT_MS });

  // Kiểm tra login thành công (tuỳ chỉnh theo cách nhận biết đăng nhập OK
  // trên hệ thống thật, ví dụ chờ 1 phần tử chỉ xuất hiện sau khi login).
  const stillOnLogin = page.url().includes('login');
  if (stillOnLogin) {
    throw new Error('Đăng nhập LMS thất bại — kiểm tra lại LMS_USERNAME/LMS_PASSWORD hoặc selector trong loginToLMS().');
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

async function fetchBranchData(page, { staffId, branch, dateFrom, dateTo }) {
  return page.evaluate(async ({ staffId, branch, dateFrom, dateTo }) => {
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

    const semRes = await fetch('/data/setup.asmx/CounSemester', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json;charset=UTF-8' },
      body: JSON.stringify({ staff: { stf_id: staffId }, setup: { hr_brch_id: branch.brch_id } }),
    });
    const semesters = safeParse(await semRes.json());
    if (!semesters.length) return { numberRows: [], sessionRows: [] };
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

    const numberData = await callReport('NUMBER');
    const listData = await callReport('LIST');

    const numberRows = numberData.map((r) => ({
      Branch: branch.brch_name,
      Program: r.Program,
      Class: r.Class,
      Date: r.Date,
      Attendance: r.Attendance,
      Absence: r.Absence,
      Late: r.Late,
      Total: r.Total,
      Month: formatMonth(r.Date),
    }));

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

    return { numberRows, sessionRows };
  }, { staffId, branch, dateFrom, dateTo });
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
    return { doneBranchIds: [], numberRows: [], sessionRows: [] };
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
// MAIN
// ---------------------------------------------------------------------------
async function main() {
  const { dateFrom, dateTo } = computeDateRange(MONTHS_BACK);
  console.log(`📅 Khoảng thời gian lấy dữ liệu: ${dateFrom} → ${dateTo} (MONTHS_BACK=${MONTHS_BACK})`);

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
  // chung session đăng nhập, không cần login lại).
  const pages = [page];
  for (let i = 1; i < CONCURRENCY; i++) {
    pages.push(await context.newPage());
  }

  let idx = 0;
  const workers = pages.map((p) => (async () => {
    while (idx < remaining.length) {
      const branch = remaining[idx++];
      try {
        console.log(`➡️  ${branch.brch_name}`);
        const { numberRows, sessionRows } = await fetchBranchData(p, {
          staffId: STAFF_ID, branch, dateFrom, dateTo,
        });
        state.numberRows.push(...numberRows);
        state.sessionRows.push(...sessionRows);
        state.doneBranchIds.push(branch.brch_id);
        saveCheckpoint(state); // checkpoint sau MỖI chi nhánh
        console.log(`   ✔️  ${branch.brch_name}: N=${numberRows.length}, session=${sessionRows.length}`);
      } catch (err) {
        console.error(`   ❌ Lỗi ở chi nhánh ${branch.brch_name}:`, err.message);
        // Không throw — chi nhánh lỗi sẽ được thử lại ở lần chạy job kế tiếp
        // (chưa được thêm vào doneBranchIds nên vẫn nằm trong "remaining" lần sau).
      }
    }
  })());
  await Promise.all(workers);

  await browser.close();

  console.log(`🎉 Xong. Tổng: ${state.numberRows.length} dòng Number, ${state.sessionRows.length} lượt điểm danh (dùng để tính tổng, không xuất ra sheet riêng).`);

  const { studentSummaryRows, classSummaryRows } = buildAggregates(state.numberRows, state.sessionRows);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.numberRows), 'Number of Student');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(studentSummaryRows), 'Student Summary');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(classSummaryRows), 'Class Summary Monthly');

  fs.mkdirSync(path.dirname(OUTPUT_XLSX), { recursive: true });
  XLSX.writeFile(wb, OUTPUT_XLSX);
  console.log(`💾 Đã ghi ${OUTPUT_XLSX}`);

  clearCheckpoint(); // job chạy xong trọn vẹn -> xoá checkpoint để lần sau chạy từ đầu
}

main().catch((err) => {
  console.error('LỖI KHÔNG XỬ LÝ ĐƯỢC:', err);
  process.exit(1);
});
