/**
 * scripts/backup-drive.js — Gửi data/latest.xlsx lên Google Drive để backup
 * qua Apps Script Web App (xem google-apps-script/DriveBackup.gs).
 *
 * Đây CHỈ là bản backup — dashboard (index.html) vẫn dùng data/latest.xlsx
 * commit trong repo như bình thường, không phụ thuộc vào Google Drive.
 *
 * Biến môi trường cần có (GitHub Actions secrets):
 *   ILEARNING_APPS_SCRIPT_URL, ILEARNING_APPS_SCRIPT_TOKEN
 *   (xem hướng dẫn deploy trong google-apps-script/DriveBackup.gs)
 *
 * Lỗi ở bước này KHÔNG làm workflow thất bại — chỉ log cảnh báo, vì backup
 * là phụ, không ảnh hưởng tới việc dashboard chính có được cập nhật hay không.
 */

const fs = require('fs');
const path = require('path');

const REPO_ROOT = path.join(__dirname, '..');
const XLSX_PATH = path.join(REPO_ROOT, 'data', 'latest.xlsx');

const APPS_SCRIPT_URL = process.env.ILEARNING_APPS_SCRIPT_URL || '';
const APPS_SCRIPT_TOKEN = process.env.ILEARNING_APPS_SCRIPT_TOKEN || '';

function timestamp() {
  const tz = 'Asia/Ho_Chi_Minh';
  const parts = new Intl.DateTimeFormat('sv-SE', {
    timeZone: tz,
    year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit',
    hour12: false,
  }).formatToParts(new Date());
  const get = (type) => parts.find((p) => p.type === type).value;
  return `${get('year')}-${get('month')}-${get('day')}_${get('hour')}${get('minute')}${get('second')}`;
}

async function main() {
  if (!APPS_SCRIPT_URL || !APPS_SCRIPT_TOKEN) {
    console.log('⚠️  Chưa cấu hình ILEARNING_APPS_SCRIPT_URL / ILEARNING_APPS_SCRIPT_TOKEN — bỏ qua backup Google Drive.');
    return;
  }
  if (!fs.existsSync(XLSX_PATH)) {
    console.log(`⚠️  Không tìm thấy ${XLSX_PATH} — bỏ qua backup Google Drive.`);
    return;
  }

  const fileBuffer = fs.readFileSync(XLSX_PATH);
  const fileBase64 = fileBuffer.toString('base64');
  const fileName = `attendance_latest_${timestamp()}.xlsx`;

  console.log(`☁️  Đang backup ${XLSX_PATH} (${(fileBuffer.length / 1024 / 1024).toFixed(1)} MB) lên Google Drive dưới tên ${fileName}...`);

  try {
    const res = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ token: APPS_SCRIPT_TOKEN, fileBase64, fileName }),
      redirect: 'follow',
    });
    const text = await res.text();
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch {
      throw new Error(`Phản hồi không phải JSON hợp lệ: ${text.slice(0, 300)}`);
    }
    if (!parsed.ok) throw new Error(parsed.error || 'Lỗi không rõ từ Apps Script');
    console.log(`✅ Backup thành công. Xem tại: ${parsed.backupFolderUrl}`);
  } catch (err) {
    console.log(`⚠️  Backup lên Google Drive thất bại: ${err.message}`);
    console.log('   (data/latest.xlsx vẫn được commit bình thường vào repo — chỉ bản backup trên Drive tạm chưa có bản mới nhất).');
  }
}

main();
