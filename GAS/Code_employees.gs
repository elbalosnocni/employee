/**
 * EMPLOYEE SYNC API - PUBLIC EMPLOYEE DIRECTORY
 * Chỉ nhận/lưu trường nhân viên được phép chia sẻ.
 * KHÔNG nhận CCCD, ngày sinh, lương, phụ cấp hoặc dữ liệu nhạy cảm.
 *
 * Script Properties bắt buộc:
 *   EMPLOYEE_SPREADSHEET_ID = ID Google Sheet trung tâm
 *   BRANCH_TOKEN_<BRANCH_CODE> = SHA-256 hex của token chi nhánh
 *
 * Triển khai Web App: Execute as Me; Who has access: Anyone.
 * Endpoint được bảo vệ bằng token riêng từng chi nhánh.
 */

const EMP_CFG = {
  EMPLOYEE_SHEET: 'Employees',
  LOG_SHEET: 'SyncLog',
  HEADERS: ['BranchCode', 'EmployeeCode', 'FullName', 'TimeClockCode', 'Department', 'Section', 'Position', 'UpdatedAt'],
  MAX_ROWS: 10000,
  MAX_AGE_SECONDS: 600
};

function doGet() {
  return jsonOut_({status: 'ok', service: 'Employee Sync API', version: '2.0'});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOut_({status: 'error', message: 'Empty request'});
    }
    let raw = e.postData.contents || '';

  // Loại bỏ UTF-8 BOM nếu có
  raw = raw.replace(/^\uFEFF/, '').trim();

  let body;
  try {
    body = JSON.parse(raw);
  } catch (err) {
    throw new Error(
      'JSON không hợp lệ: ' + err.message +
      '. Dữ liệu nhận được: ' + raw.substring(0, 200)
    );
  }
    validatePayload_(body);

    const props = PropertiesService.getScriptProperties();
    const spreadsheetId = props.getProperty('EMPLOYEE_SPREADSHEET_ID');
    if (!spreadsheetId) throw new Error('Chưa cấu hình EMPLOYEE_SPREADSHEET_ID trong Script Properties');

    const branch = String(body.branch_code).trim().toUpperCase();
    const tokenKey = 'BRANCH_TOKEN_' + branch;
    const expectedHash = props.getProperty(tokenKey);
    const receivedHash = sha256Hex_(String(body.token || ''));
    if (!expectedHash || !constantTimeEqual_(expectedHash, receivedHash)) {
      throw new Error('Chi nhánh không được xác thực');
    }

    const now = Math.floor(Date.now() / 1000);
    const sentAt = Number(body.timestamp);
    if (!Number.isFinite(sentAt) || Math.abs(now - sentAt) > EMP_CFG.MAX_AGE_SECONDS) {
      throw new Error('Yêu cầu hết hạn; vui lòng đồng bộ lại');
    }

    const rows = body.employees;
    if (!Array.isArray(rows) || rows.length === 0) throw new Error('Không có nhân viên hợp lệ');
    if (rows.length > EMP_CFG.MAX_ROWS) throw new Error('Vượt quá giới hạn ' + EMP_CFG.MAX_ROWS + ' dòng/lần');

    lock.waitLock(30000);
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = getOrCreateSheet_(ss, EMP_CFG.EMPLOYEE_SHEET, EMP_CFG.HEADERS);
    const logSheet = getOrCreateSheet_(ss, EMP_CFG.LOG_SHEET,
      ['Timestamp', 'BranchCode', 'Received', 'Inserted', 'Updated', 'Skipped', 'Status', 'Note']);

    const last = sheet.getLastRow();
    const old = last > 1 ? sheet.getRange(2, 1, last - 1, EMP_CFG.HEADERS.length).getDisplayValues() : [];
    const index = Object.create(null);
    old.forEach((r, i) => {
      const key = makeKey_(r[0], r[1]);
      if (key) index[key] = i;
    });

    let inserted = 0, updated = 0, skipped = 0;
    rows.forEach((r) => {
      const clean = normalizeEmployee_(r, branch);
      if (!clean) { skipped++; return; }
      const key = makeKey_(clean[0], clean[1]);
      const row = clean.concat([new Date()]);
      if (Object.prototype.hasOwnProperty.call(index, key)) {
        old[index[key]] = row;
        updated++;
      } else {
        index[key] = old.length;
        old.push(row);
        inserted++;
      }
    });

    if (old.length) {
      // Chỉ cập nhật dữ liệu trong bảng Employees; không xóa các sheet riêng tư khác.
      sheet.getRange(2, 1, old.length, EMP_CFG.HEADERS.length).clearContent();
      sheet.getRange(2, 1, old.length, EMP_CFG.HEADERS.length).setValues(old);
      sheet.getRange(2, 8, old.length, 1).setNumberFormat('dd/mm/yyyy hh:mm:ss');
    }

    logSheet.appendRow([new Date(), branch, rows.length, inserted, updated, skipped, 'SUCCESS', '']);
    return jsonOut_({status: 'success', branch_code: branch, received: rows.length,
      inserted: inserted, updated: updated, skipped: skipped});
  } catch (err) {
    try {
      const props = PropertiesService.getScriptProperties();
      const id = props.getProperty('EMPLOYEE_SPREADSHEET_ID');
      if (id) {
        const ss = SpreadsheetApp.openById(id);
        const log = getOrCreateSheet_(ss, EMP_CFG.LOG_SHEET,
          ['Timestamp', 'BranchCode', 'Received', 'Inserted', 'Updated', 'Skipped', 'Status', 'Note']);
        log.appendRow([new Date(), '', 0, 0, 0, 0, 'ERROR', String(err.message || err).slice(0, 500)]);
      }
    } catch (_) {}
    return jsonOut_({status: 'error', message: String(err.message || err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function validatePayload_(b) {
  if (!b || b.version !== 2) throw new Error('Sai version API');
  if (!/^[A-Z0-9_-]{2,30}$/.test(String(b.branch_code || '').toUpperCase())) throw new Error('Mã chi nhánh không hợp lệ');
  if (!b.token || String(b.token).length < 24) throw new Error('Token không hợp lệ');
  if (!Array.isArray(b.employees)) throw new Error('employees phải là mảng');
}

function normalizeEmployee_(r, branch) {
  if (!r || typeof r !== 'object') return null;
  const code = safeText_(r.employee_code, 50);
  const name = safeText_(r.full_name, 150);
  if (!code || !name) return null;
  return [
    branch,
    code,
    name,
    safeText_(r.time_clock_code, 50),
    safeText_(r.department, 120),
    safeText_(r.section, 120),
    safeText_(r.position, 120)
  ];
}

function safeText_(v, max) {
  let s = String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim();
  if (s.length > max) s = s.substring(0, max);
  return s;
}
function makeKey_(branch, code) {
  const a = String(branch || '').trim().toUpperCase();
  const b = String(code || '').trim().toUpperCase();
  return a && b ? a + '|' + b : '';
}
function getOrCreateSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else {
    const existing = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    if (existing.join('|') !== headers.join('|')) {
      throw new Error('Cấu trúc sheet ' + name + ' không đúng. Không tự ý ghi đè tiêu đề hiện có.');
    }
  }
  return sh;
}
function sha256Hex_(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function constantTimeEqual_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}
function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Chạy thủ công trong Apps Script sau khi thay giá trị:
 * 1) setEmployeeSpreadsheetId('ID_SHEET_TRUNG_TAM')
 * 2) setBranchToken('BH', 'TOKEN_NGAU_NHIEN_DAI_TREN_24_KY_TU')
 * Không gửi token qua email/chat công khai.
 */
function setEmployeeSpreadsheetId(id) {
  if (!id || String(id).length < 20) throw new Error('Spreadsheet ID không hợp lệ');
  PropertiesService.getScriptProperties().setProperty('EMPLOYEE_SPREADSHEET_ID', String(id).trim());
}
function setBranchToken(branchCode, token) {
  const branch = String(branchCode || '').trim().toUpperCase();
  if (!/^[A-Z0-9_-]{2,30}$/.test(branch)) throw new Error('Mã chi nhánh không hợp lệ');
  if (!token || String(token).length < 24) throw new Error('Token phải có ít nhất 24 ký tự');
  PropertiesService.getScriptProperties().setProperty('BRANCH_TOKEN_' + branch, sha256Hex_(String(token)));
}
function setupEmployeeSheets() {
  const id = PropertiesService.getScriptProperties().getProperty('EMPLOYEE_SPREADSHEET_ID');
  if (!id) throw new Error('Hãy chạy setEmployeeSpreadsheetId trước');
  const ss = SpreadsheetApp.openById(id);
  getOrCreateSheet_(ss, EMP_CFG.EMPLOYEE_SHEET, EMP_CFG.HEADERS);
  getOrCreateSheet_(ss, EMP_CFG.LOG_SHEET,
    ['Timestamp', 'BranchCode', 'Received', 'Inserted', 'Updated', 'Skipped', 'Status', 'Note']);
}
function setupBranch() {
  setBranchToken(
    'VN',
    '123456789012345676'
  );
}
