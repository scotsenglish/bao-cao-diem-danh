/**
 * DriveBackup.gs — Backup data/latest.xlsx (dashboard Attendance i-Learning)
 * lên Google Drive
 * ---------------------------------------------------------------------------
 * Mỗi lần scripts/backup-drive.js chạy (sau khi scrape.js ghi xong
 * data/latest.xlsx), nó gửi NGUYÊN file .xlsx đó (mã hoá base64) lên đây.
 * Script chỉ có 1 việc: lưu file vào 1 thư mục Drive riêng "iLearning
 * Attendance Backups" (tự tạo nếu chưa có) — cùng cơ chế tạo folder riêng đã
 * làm cho thong-ke-diem-c-Learning / thong-ke-diem-i-Learning trước đó.
 *
 * File backup là bản sao Y HỆT data/latest.xlsx (2 sheet "Student Summary" +
 * "Class Summary Monthly") — nên có thể tải về từ Drive rồi dùng đúng nút
 * "📂 Tải file export mới (.xlsx)" sẵn có trên dashboard mà không cần chỉnh sửa gì.
 *
 * CÁCH THIẾT LẬP:
 * 1. Vào https://script.google.com -> New project
 * 2. Xoá code mặc định, paste toàn bộ nội dung file này vào
 * 3. Vào Project Settings (icon bánh răng bên trái) -> Script Properties ->
 *    Add script property:
 *      - Tên: WRITE_TOKEN
 *      - Giá trị: 1 chuỗi bí mật tự đặt (VD dùng lệnh `openssl rand -hex 16`)
 *        — token này dùng để xác thực scripts/backup-drive.js, không cho
 *        người lạ gọi lên ghi rác vào Drive của bạn.
 * 4. Bấm Deploy -> New deployment -> chọn loại "Web app"
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Bấm Deploy -> Copy URL (dạng https://script.google.com/macros/s/.../exec)
 * 5. URL đó chính là ILEARNING_APPS_SCRIPT_URL, token ở bước 3 chính là
 *    ILEARNING_APPS_SCRIPT_TOKEN -> thêm cả 2 vào GitHub Secrets của repo
 *    (Settings -> Secrets and variables -> Actions -> New repository secret).
 * 6. LẦN DEPLOY ĐẦU TIÊN, Google sẽ hiện màn hình xin thêm quyền truy cập
 *    Drive -> bấm "Advanced" -> "Go to (tên project) (unsafe)" -> Allow (đây
 *    là bước bảo mật bình thường của Google, không phải lỗi).
 * 7. LƯU Ý QUAN TRỌNG: mỗi lần sửa code trong editor, phải vào Deploy ->
 *    Manage deployments -> bấm nút sửa (icon bút chì) trên deployment "Web
 *    app" đang dùng -> Version: chọn "New" -> Deploy, thì Web App URL hiện tại
 *    (đang được scripts/backup-drive.js gọi tới) mới thực sự chạy code mới.
 *    Chỉ lưu (Ctrl+S) trong editor KHÔNG đủ.
 * 8. Verify nhanh không cần đợi workflow chạy: trong editor, chọn function
 *    "testBackupManually" ở dropdown "Select function" -> Run, rồi kiểm tra
 *    Drive có folder "Attendance Backups" chứa file .xlsx test.
 * ---------------------------------------------------------------------------
 */

const BACKUP_FOLDER_NAME = "Attendance Backups";

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const expectedToken = PropertiesService.getScriptProperties().getProperty("WRITE_TOKEN");

    if (!expectedToken) {
      return jsonOutput({ ok: false, error: "Chưa cấu hình WRITE_TOKEN trong Script Properties" });
    }
    if (body.token !== expectedToken) {
      return jsonOutput({ ok: false, error: "Token không đúng" });
    }

    const fileBase64 = body.fileBase64;
    const fileName = String(body.fileName || "").trim();
    if (!fileBase64 || !fileName) {
      return jsonOutput({ ok: false, error: "Thiếu fileBase64 hoặc fileName trong request" });
    }

    const folder = saveBackupFile(fileBase64, fileName);
    PropertiesService.getScriptProperties().setProperty("LAST_UPDATED", new Date().toISOString());

    return jsonOutput({ ok: true, fileName, backupFolderUrl: folder.getUrl() });
  } catch (err) {
    console.error("Backup thất bại: " + err);
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const lastUpdated = PropertiesService.getScriptProperties().getProperty("LAST_UPDATED") || null;
  return jsonOutput({ ok: true, lastUpdated });
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateBackupFolder() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty("BACKUP_FOLDER_ID");
  if (savedId) {
    try {
      const existing = DriveApp.getFolderById(savedId);
      existing.getName(); // kiểm tra folder vẫn còn tồn tại, chưa bị xoá thủ công
      return existing;
    } catch (e) {
      // ID cũ không còn hợp lệ (folder đã bị xoá) -> tạo lại bên dưới
    }
  }
  const folder = DriveApp.createFolder(BACKUP_FOLDER_NAME);
  props.setProperty("BACKUP_FOLDER_ID", folder.getId());
  return folder;
}

// fileBase64 là nội dung .xlsx gốc (data/latest.xlsx) đã mã hoá base64 —
// lưu thẳng thành file, KHÔNG dựng lại qua Spreadsheet trung gian, để đảm bảo
// bản backup y hệt file dashboard đang dùng (tương thích 100% với nút
// "Tải file export mới (.xlsx)" có sẵn trên dashboard).
function saveBackupFile(fileBase64, fileName) {
  const bytes = Utilities.base64Decode(fileBase64);
  const blob = Utilities.newBlob(bytes, MimeType.MICROSOFT_EXCEL, fileName);

  const folder = getOrCreateBackupFolder();
  // Tạo file rồi CHỦ ĐỘNG gán vào đúng folder + gỡ khỏi root — chắc chắn hơn
  // là chỉ gọi folder.createFile(blob).
  const newFile = DriveApp.createFile(blob);
  folder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile);
  return folder;
}

// Chạy tay function này trong editor Apps Script để test nhanh không cần đợi
// scripts/backup-drive.js gọi lên — tạo 1 Google Sheet nhỏ, export ra .xlsx
// rồi backup thử.
function testBackupManually() {
  const ss = SpreadsheetApp.create("test-backup-tmp");
  try {
    ss.getSheets()[0].getRange(1, 1, 1, 2).setValues([["Cột A", "Cột B"]]);
    SpreadsheetApp.flush();

    const fileId = ss.getId();
    const exportUrl = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    });
    const fileBase64 = Utilities.base64Encode(response.getContent());

    const folder = saveBackupFile(fileBase64, "test_backup.xlsx");
    Logger.log("Đã tạo file backup test trong folder: " + folder.getUrl());
  } finally {
    DriveApp.getFileById(ss.getId()).setTrashed(true);
  }
}
