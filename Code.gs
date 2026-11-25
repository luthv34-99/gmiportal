/**
 * Backend Web App — PT. Global Mega Indonesia (Portal, Laporan & Absensi Karyawan)
 * Spreadsheet terikat (SpreadsheetApp.getActiveSpreadsheet()).
 *
 * Panduan Deploy:
 *   1. Pastikan skrip terikat pada Google Spreadsheet perusahaan.
 *   2. Deploy -> Deployment baru -> Aplikasi web -> Jalankan sebagai: Saya -> Yang memiliki akses: Siapa saja.
 *   3. Salin URL .../exec ke file webapp.txt.
 */

const SHEET_NAME = 'GMI_Users';
const REPORTS_SHEET = 'GMI_Reports';
const CONFIG_SHEET = 'GMI_Config';
const ATTENDANCE_SHEET = 'GMI_Attendance';

const HEADERS = [
  'name', 'role', 'email', 'salt', 'passwordHash', 'verified', 'verifyCode',
  'createdAt', 'verifiedAt', 'sessionToken', 'sessionExpiry',
  'isAdmin', 'resetCode', 'resetExpiry'
];

const REPORT_HEADERS = [
  'id', 'noDoc', 'title', 'tanggal', 'model', 'customer', 'noPo', 'status',
  'createdBy', 'createdByName', 'createdAt', 'updatedAt', 'photoCount', 'payloadJson'
];

const ATTENDANCE_HEADERS = [
  'id', 'date', 'email', 'name', 'role',
  'clockIn', 'clockInLat', 'clockInLng', 'clockInAcc', 'clockInDist', 'clockInStatus', 'lateMinutes', 'clockInPhoto',
  'clockOut', 'clockOutLat', 'clockOutLng', 'clockOutAcc', 'clockOutDist', 'overtimeMinutes', 'clockOutPhoto',
  'deviceInfo'
];

const SESSION_DAYS = 7;
const RESET_TTL_MS = 30 * 60 * 1000;
const PHOTO_MAX_CHARS = 350000;
const REPORT_PAYLOAD_MAX_CHARS = 3500000;
const ADMIN_BOOTSTRAP_EMAIL = 'luthfinaufaloktafian@gmail.com';
const PUBLIC_URL_KEYS = ['logoUrl', 'qrUrl'];

function doGet(e) {
  return respond_(handle_((e && e.parameter) || {}));
}

function doPost(e) {
  let payload = {};
  try {
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    }
  } catch (err) {
    return respond_({ ok: false, message: 'JSON tidak valid: ' + err.message });
  }
  return respond_(handle_(payload));
}

function respond_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handle_(p) {
  try {
    switch (String(p.action || 'ping')) {
      case 'ping':                return { ok: true, time: new Date().toISOString() };
      case 'register':            return register_(p);
      case 'verify':              return verify_(p);
      case 'resend':              return resend_(p);
      case 'login':               return login_(p);
      case 'logout':              return logout_(p);
      case 'whoami':              return whoami_(p);
      case 'request_reset':       return requestReset_(p);
      case 'reset_password':      return resetPassword_(p);
      case 'list_users':          return listUsers_(p);
      case 'set_admin':           return setAdmin_(p);
      case 'get_public_config':   return getPublicConfig_();
      case 'get_config':          return getConfig_(p);
      case 'set_config':          return setConfig_(p);
      case 'save_report':         return saveReport_(p);
      case 'list_reports':        return listReports_(p);
      case 'get_report':          return getReport_(p);
      case 'delete_report':       return deleteReport_(p);
      case 'get_work_config':     return getWorkConfig_(p);
      case 'set_work_config':     return setWorkConfig_(p);
      case 'submit_attendance':   return submitAttendance_(p);
      case 'get_attendance_data': return getAttendanceData_(p);
      default:                    return { ok: false, message: 'Aksi tidak dikenal: ' + p.action };
    }
  } catch (err) {
    return { ok: false, message: 'Kesalahan server: ' + err.message };
  }
}

/* ======================= PENYIMPANAN & UTILITAS ======================= */

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  HEADERS.forEach((h, idx) => {
    if (existing[idx] !== h) sh.getRange(1, idx + 1).setValue(h);
  });
  sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), HEADERS.length).setNumberFormat('@');
  return sh;
}

function plainSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  headers.forEach((h, idx) => {
    if (existing[idx] !== h) sh.getRange(1, idx + 1).setValue(h);
  });
  return sh;
}

function rows_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const data = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
  return data.map((r, i) => {
    const o = { _row: i + 2 };
    HEADERS.forEach((k, c) => o[k] = r[c]);
    return o;
  });
}

function writeRow_(row, obj) {
  sheet_().getRange(row, 1, 1, HEADERS.length).setValues([HEADERS.map(k => obj[k] === undefined ? '' : obj[k])]);
}

function lower_(s) { return String(s || '').trim().toLowerCase(); }
function bool_(v) { return v === true || String(v).toUpperCase() === 'TRUE' || v === 1; }

function hex_(bytes) {
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function sha256_(text) {
  return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8));
}

function randomHex_(n) {
  return Utilities.getUuid().replace(/-/g, '').slice(0, n);
}

function sixDigitCode_() {
  const num = Math.floor(100000 + Math.random() * 900000);
  return String(num);
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function isBootstrapAdmin_(email) {
  return lower_(email) === ADMIN_BOOTSTRAP_EMAIL;
}

function sanitizeUser_(u) {
  return {
    name: String(u.name || ''),
    role: String(u.role || ''),
    email: lower_(u.email),
    verified: bool_(u.verified),
    isAdmin: bool_(u.isAdmin) || isBootstrapAdmin_(u.email),
    createdAt: String(u.createdAt || ''),
    verifiedAt: String(u.verifiedAt || '')
  };
}

function sessionUser_(token) {
  const tok = String(token || '').trim();
  if (!tok) return null;
  const u = rows_().find(r => String(r.sessionToken) === tok);
  if (!u || !bool_(u.verified)) return null;
  if (u.sessionExpiry && new Date(u.sessionExpiry).getTime() < Date.now()) return null;
  if (isBootstrapAdmin_(u.email) && !bool_(u.isAdmin)) {
    u.isAdmin = 'TRUE';
    writeRow_(u._row, u);
  }
  return u;
}

function requireAdmin_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return null;
  return (bool_(me.isAdmin) || isBootstrapAdmin_(me.email)) ? me : null;
}

/* ======================= KONFIGURASI UMUM & KOP ======================= */

function configSheet_() {
  const sh = plainSheet_(CONFIG_SHEET, ['key', 'value']);
  sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), 2).setNumberFormat('@');
  return sh;
}

function readConfig_() {
  const sh = configSheet_();
  const last = sh.getLastRow();
  const out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(r => {
    const k = String(r[0] || '').trim();
    if (k) out[k] = String(r[1] === undefined ? '' : r[1]);
  });
  return out;
}

function writeConfig_(kv) {
  const sh = configSheet_();
  const last = sh.getLastRow();
  const existing = last >= 2 ? sh.getRange(2, 1, last - 1, 2).getValues() : [];
  const index = {};
  existing.forEach((r, i) => {
    const k = String(r[0] || '').trim();
    if (k) index[k] = i + 2;
  });
  Object.keys(kv).forEach(k => {
    const val = kv[k] === undefined || kv[k] === null ? '' : String(kv[k]);
    if (index[k]) {
      sh.getRange(index[k], 1, 1, 2).setValues([[k, val]]);
    } else {
      sh.appendRow([k, val]);
    }
  });
}

function isValidConfigUrl_(u) {
  if (!u) return true;
  if (/^https?:\/\/[^\s]+$/i.test(u)) return true;
  if (/^data:image\/(png|jpeg|jpg|webp|svg\+xml);base64,[A-Za-z0-9+\/=]+$/i.test(u)) return true;
  return false;
}

function getPublicConfig_() {
  const all = readConfig_();
  const pub = {};
  PUBLIC_URL_KEYS.forEach(k => { pub[k] = all[k] || ''; });
  return { ok: true, config: pub };
}

function getConfig_(p) {
  if (!requireAdmin_(p)) return { ok: false, message: 'Akses ditolak: hanya admin' };
  const all = readConfig_();
  const pub = {};
  PUBLIC_URL_KEYS.forEach(k => { pub[k] = all[k] || ''; });
  return { ok: true, config: pub };
}

function setConfig_(p) {
  if (!requireAdmin_(p)) return { ok: false, message: 'Akses ditolak: hanya admin' };
  const upd = {};
  for (let i = 0; i < PUBLIC_URL_KEYS.length; i++) {
    const k = PUBLIC_URL_KEYS[i];
    if (p[k] !== undefined) {
      const v = String(p[k] || '').trim();
      if (!isValidConfigUrl_(v)) return { ok: false, message: 'Format URL/gambar tidak valid untuk ' + k };
      if (v.length > PHOTO_MAX_CHARS) return { ok: false, message: 'Ukuran gambar ' + k + ' terlalu besar' };
      upd[k] = v;
    }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try { writeConfig_(upd); } finally { lock.releaseLock(); }
  return getConfig_(p);
}

/* ======================= AUTENTIKASI & PENGGUNA ======================= */

function sendVerifyEmail_(email, name, code) {
  const subject = '[PT. Global Mega Indonesia] Kode Verifikasi Portal: ' + code;
  const body =
    'Halo ' + name + ',\n\n' +
    'Terima kasih telah mendaftar pada Portal PT. Global Mega Indonesia.\n' +
    'Kode verifikasi 6 digit Anda adalah:\n\n    ' + code + '\n\n' +
    'Masukkan kode tersebut pada halaman verifikasi untuk mengaktifkan akun Anda.\n\n' +
    'Hormat kami,\nPT. Global Mega Indonesia';
  MailApp.sendEmail({ to: email, subject: subject, body: body, name: 'PT. Global Mega Indonesia' });
}

function register_(p) {
  const name = String(p.name || '').trim();
  const role = String(p.role || '').trim();
  const email = lower_(p.email);
  const password = String(p.password || '');

  if (!name || !role || !email || !password) return { ok: false, message: 'Semua kolom wajib diisi' };
  if (!isValidEmail_(email)) return { ok: false, message: 'Format email tidak valid' };
  if (password.length < 6) return { ok: false, message: 'Password minimal 6 karakter' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const all = rows_();
    const existing = all.find(r => lower_(r.email) === email);
    const code = sixDigitCode_();
    const salt = randomHex_(16);
    const hash = sha256_(salt + ':' + password);
    const now = new Date().toISOString();
    const isAdmin = isBootstrapAdmin_(email) ? 'TRUE' : 'FALSE';

    if (existing) {
      if (bool_(existing.verified)) {
        return { ok: false, message: 'Email ini sudah terdaftar dan terverifikasi. Silakan langsung login.' };
      }
      writeRow_(existing._row, {
        ...existing,
        name, role, salt, passwordHash: hash,
        verified: 'FALSE', verifyCode: code,
        createdAt: now, isAdmin
      });
    } else {
      sheet_().appendRow([
        name, role, email, salt, hash, 'FALSE', code,
        now, '', '', '', isAdmin, '', ''
      ]);
    }
    sendVerifyEmail_(email, name, code);
    return { ok: true, message: 'Kode verifikasi telah dikirim ke ' + email };
  } finally {
    lock.releaseLock();
  }
}

function verify_(p) {
  const email = lower_(p.email);
  const code = String(p.code || '').trim();
  if (!email || !code) return { ok: false, message: 'Email dan kode wajib diisi' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const u = rows_().find(r => lower_(r.email) === email);
    if (!u) return { ok: false, message: 'Akun tidak ditemukan' };
    if (bool_(u.verified)) return { ok: true, message: 'Akun sudah terverifikasi' };
    if (String(u.verifyCode).trim() !== code) return { ok: false, message: 'Kode verifikasi salah' };

    u.verified = 'TRUE';
    u.verifyCode = '';
    u.verifiedAt = new Date().toISOString();
    if (isBootstrapAdmin_(u.email)) u.isAdmin = 'TRUE';
    writeRow_(u._row, u);
    return { ok: true, message: 'Verifikasi berhasil. Silakan login.' };
  } finally {
    lock.releaseLock();
  }
}

function resend_(p) {
  const email = lower_(p.email);
  const u = rows_().find(r => lower_(r.email) === email);
  if (!u) return { ok: false, message: 'Akun tidak ditemukan' };
  if (bool_(u.verified)) return { ok: false, message: 'Akun sudah terverifikasi' };
  const code = sixDigitCode_();
  u.verifyCode = code;
  writeRow_(u._row, u);
  sendVerifyEmail_(email, u.name, code);
  return { ok: true, message: 'Kode baru dikirim ke ' + email };
}

function login_(p) {
  const email = lower_(p.email);
  const password = String(p.password || '');
  const u = rows_().find(r => lower_(r.email) === email);
  if (!u) return { ok: false, message: 'Email atau password salah' };
  const hash = sha256_(u.salt + ':' + password);
  if (hash !== String(u.passwordHash)) return { ok: false, message: 'Email atau password salah' };
  if (!bool_(u.verified)) {
    return { ok: false, code: 'UNVERIFIED', message: 'Akun belum diverifikasi. Periksa kode di email Anda.' };
  }

  const token = randomHex_(32);
  const expiry = new Date(Date.now() + SESSION_DAYS * 86400 * 1000).toISOString();
  u.sessionToken = token;
  u.sessionExpiry = expiry;
  if (isBootstrapAdmin_(u.email)) u.isAdmin = 'TRUE';
  writeRow_(u._row, u);

  return { ok: true, token, user: sanitizeUser_(u) };
}

function logout_(p) {
  const u = sessionUser_(p.sessionToken);
  if (u) {
    u.sessionToken = '';
    u.sessionExpiry = '';
    writeRow_(u._row, u);
  }
  return { ok: true };
}

function whoami_(p) {
  const u = sessionUser_(p.sessionToken);
  if (!u) return { ok: false, message: 'Sesi tidak valid' };
  return { ok: true, user: sanitizeUser_(u) };
}

function requestReset_(p) {
  const email = lower_(p.email);
  const u = rows_().find(r => lower_(r.email) === email);
  if (!u || !bool_(u.verified)) {
    return { ok: true, message: 'Jika email terdaftar, kode pemulihan telah dikirim.' };
  }
  const code = sixDigitCode_();
  u.resetCode = code;
  u.resetExpiry = new Date(Date.now() + RESET_TTL_MS).toISOString();
  writeRow_(u._row, u);

  MailApp.sendEmail({
    to: email,
    subject: '[PT. Global Mega Indonesia] Kode Reset Password: ' + code,
    body: 'Halo ' + u.name + ',\n\nKode reset password Anda adalah: ' + code + ' (berlaku 30 menit).',
    name: 'PT. Global Mega Indonesia'
  });
  return { ok: true, message: 'Kode pemulihan telah dikirim ke ' + email };
}

function resetPassword_(p) {
  const email = lower_(p.email);
  const code = String(p.code || '').trim();
  const newPassword = String(p.newPassword || '');
  if (!email || !code || newPassword.length < 6) {
    return { ok: false, message: 'Data tidak lengkap atau password kurang dari 6 karakter' };
  }
  const u = rows_().find(r => lower_(r.email) === email);
  if (!u || String(u.resetCode).trim() !== code) {
    return { ok: false, message: 'Kode reset tidak valid' };
  }
  if (u.resetExpiry && new Date(u.resetExpiry).getTime() < Date.now()) {
    return { ok: false, message: 'Kode reset sudah kadaluarsa' };
  }
  const salt = randomHex_(16);
  u.salt = salt;
  u.passwordHash = sha256_(salt + ':' + newPassword);
  u.resetCode = '';
  u.resetExpiry = '';
  u.sessionToken = '';
  writeRow_(u._row, u);
  return { ok: true, message: 'Password berhasil diperbarui. Silakan login kembali.' };
}

function listUsers_(p) {
  if (!requireAdmin_(p)) return { ok: false, message: 'Akses ditolak: hanya admin' };
  return { ok: true, users: rows_().map(sanitizeUser_) };
}

function setAdmin_(p) {
  const admin = requireAdmin_(p);
  if (!admin) return { ok: false, message: 'Akses ditolak: hanya admin' };
  const targetEmail = lower_(p.email);
  if (isBootstrapAdmin_(targetEmail)) return { ok: false, message: 'Admin utama tidak dapat diubah' };
  const u = rows_().find(r => lower_(r.email) === targetEmail);
  if (!u) return { ok: false, message: 'Pengguna tidak ditemukan' };
  u.isAdmin = bool_(p.isAdmin) ? 'TRUE' : 'FALSE';
  writeRow_(u._row, u);
  return { ok: true, user: sanitizeUser_(u) };
}

/* ======================= LAPORAN INSPEKSI MESIN ======================= */

function reportRows_() {
  const sh = plainSheet_(REPORTS_SHEET, REPORT_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, REPORT_HEADERS.length).getValues().map((r, i) => {
    const o = { _row: i + 2 };
    REPORT_HEADERS.forEach((k, c) => o[k] = r[c]);
    return o;
  });
}

function saveReport_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid' };
  const rep = p.report || {};
  const payloadStr = JSON.stringify(rep);
  if (payloadStr.length > REPORT_PAYLOAD_MAX_CHARS) {
    return { ok: false, message: 'Ukuran laporan melebihi batas maksimal server' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = plainSheet_(REPORTS_SHEET, REPORT_HEADERS);
    const all = reportRows_();
    const id = String(rep.id || ('rep_' + Date.now()));
    const existing = all.find(r => String(r.id) === id);
    const now = new Date().toISOString();
    const photoCount = Array.isArray(rep.photos) ? rep.photos.length : 0;

    if (existing) {
      if (!bool_(me.isAdmin) && lower_(existing.createdBy) !== lower_(me.email)) {
        return { ok: false, message: 'Anda tidak memiliki izin mengubah laporan ini' };
      }
      const rowVals = [
        id, String(rep.noDoc || ''), String(rep.title || ''), String(rep.tanggal || ''),
        String(rep.model || ''), String(rep.customer || ''), String(rep.noPo || ''), String(rep.status || 'Draft'),
        existing.createdBy, existing.createdByName, existing.createdAt, now, photoCount, payloadStr
      ];
      sh.getRange(existing._row, 1, 1, REPORT_HEADERS.length).setValues([rowVals]);
    } else {
      sh.appendRow([
        id, String(rep.noDoc || ''), String(rep.title || ''), String(rep.tanggal || ''),
        String(rep.model || ''), String(rep.customer || ''), String(rep.noPo || ''), String(rep.status || 'Draft'),
        lower_(me.email), String(me.name), now, now, photoCount, payloadStr
      ]);
    }
    return { ok: true, id, updatedAt: now };
  } finally {
    lock.releaseLock();
  }
}

function listReports_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid' };
  const isAdmin = bool_(me.isAdmin);
  const mineOnly = bool_(p.mine);
  const list = reportRows_()
    .filter(r => (isAdmin && !mineOnly) || lower_(r.createdBy) === lower_(me.email))
    .map(r => ({
      id: r.id, noDoc: r.noDoc, title: r.title, tanggal: r.tanggal,
      model: r.model, customer: r.customer, noPo: r.noPo, status: r.status,
      createdBy: r.createdBy, createdByName: r.createdByName,
      createdAt: r.createdAt, updatedAt: r.updatedAt, photoCount: r.photoCount
    }))
    .sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
  return { ok: true, reports: list, isAdmin };
}

function getReport_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid' };
  const r = reportRows_().find(x => String(x.id) === String(p.id || ''));
  if (!r) return { ok: false, message: 'Laporan tidak ditemukan' };
  if (!bool_(me.isAdmin) && lower_(r.createdBy) !== lower_(me.email)) {
    return { ok: false, message: 'Akses ditolak' };
  }
  return { ok: true, report: JSON.parse(r.payloadJson || '{}') };
}

function deleteReport_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid' };
  const sh = plainSheet_(REPORTS_SHEET, REPORT_HEADERS);
  const r = reportRows_().find(x => String(x.id) === String(p.id || ''));
  if (!r) return { ok: false, message: 'Laporan tidak ditemukan' };
  if (!bool_(me.isAdmin) && lower_(r.createdBy) !== lower_(me.email)) {
    return { ok: false, message: 'Akses ditolak' };
  }
  sh.deleteRow(r._row);
  return { ok: true };
}

/* ======================= ABSENSI KARYAWAN ======================= */

function defaultWorkConfig_() {
  return {
    officeLat: -6.284500,        // Default Kawasan Industri Jababeka 1, Cikarang
    officeLng: 107.148500,
    officeRadius: 150,           // Radius geofence (meter)
    workStartTime: '08:00',      // Jam masuk standar
    workEndTime: '17:00',        // Jam pulang standar
    lateToleranceMin: 5,         // Toleransi keterlambatan (menit)
    overtimeMinThreshold: 30     // Minimal menit setelah jam pulang untuk dihitung lembur
  };
}

function readWorkConfig_() {
  const cfg = readConfig_();
  const def = defaultWorkConfig_();
  return {
    officeLat: cfg.officeLat !== undefined && cfg.officeLat !== '' ? Number(cfg.officeLat) : def.officeLat,
    officeLng: cfg.officeLng !== undefined && cfg.officeLng !== '' ? Number(cfg.officeLng) : def.officeLng,
    officeRadius: cfg.officeRadius !== undefined && cfg.officeRadius !== '' ? Number(cfg.officeRadius) : def.officeRadius,
    workStartTime: cfg.workStartTime || def.workStartTime,
    workEndTime: cfg.workEndTime || def.workEndTime,
    lateToleranceMin: cfg.lateToleranceMin !== undefined && cfg.lateToleranceMin !== '' ? Number(cfg.lateToleranceMin) : def.lateToleranceMin,
    overtimeMinThreshold: cfg.overtimeMinThreshold !== undefined && cfg.overtimeMinThreshold !== '' ? Number(cfg.overtimeMinThreshold) : def.overtimeMinThreshold
  };
}

function getWorkConfig_(p) {
  if (!sessionUser_(p.sessionToken)) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };
  return { ok: true, config: readWorkConfig_() };
}

function setWorkConfig_(p) {
  const admin = requireAdmin_(p);
  if (!admin) return { ok: false, message: 'Akses ditolak: hanya Admin yang dapat mengatur jam kerja & lokasi kantor' };

  const lat = Number(p.officeLat);
  const lng = Number(p.officeLng);
  const radius = Number(p.officeRadius);
  const start = String(p.workStartTime || '').trim();
  const end = String(p.workEndTime || '').trim();
  const tol = Number(p.lateToleranceMin);
  const otMin = Number(p.overtimeMinThreshold);

  if (isNaN(lat) || lat < -90 || lat > 90 || isNaN(lng) || lng < -180 || lng > 180) {
    return { ok: false, message: 'Koordinat Latitude / Longitude kantor tidak valid' };
  }
  if (isNaN(radius) || radius < 10 || radius > 5000) {
    return { ok: false, message: 'Radius geofence harus antara 10 hingga 5000 meter' };
  }
  if (!/^\d{2}:\d{2}$/.test(start) || !/^\d{2}:\d{2}$/.test(end)) {
    return { ok: false, message: 'Format jam kerja harus HH:MM (contoh: 08:00)' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    writeConfig_({
      officeLat: String(lat),
      officeLng: String(lng),
      officeRadius: String(Math.round(radius)),
      workStartTime: start,
      workEndTime: end,
      lateToleranceMin: String(Math.max(0, Math.round(isNaN(tol) ? 0 : tol))),
      overtimeMinThreshold: String(Math.max(0, Math.round(isNaN(otMin) ? 0 : otMin)))
    });
  } finally {
    lock.releaseLock();
  }
  return { ok: true, config: readWorkConfig_(), message: 'Pengaturan jam kerja & geofence berhasil disimpan' };
}

function haversineMeters_(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const toRad = function (deg) { return deg * Math.PI / 180; };
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
            Math.cos(toRad(lat1)) * Math.cos(toRad(lat2)) *
            Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return Math.round(R * c);
}

function parseMinutesOfDay_(hhmm) {
  const parts = String(hhmm || '00:00').split(':');
  return (parseInt(parts[0], 10) || 0) * 60 + (parseInt(parts[1], 10) || 0);
}

function attendanceRows_() {
  const sh = plainSheet_(ATTENDANCE_SHEET, ATTENDANCE_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ATTENDANCE_HEADERS.length).getValues().map(function (r, i) {
    const o = { _row: i + 2 };
    ATTENDANCE_HEADERS.forEach(function (k, c) { o[k] = r[c]; });
    return o;
  });
}

function submitAttendance_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };

  const type = String(p.type || '').toUpperCase();
  if (type !== 'IN' && type !== 'OUT') return { ok: false, message: 'Tipe absensi harus IN atau OUT' };

  const lat = Number(p.lat);
  const lng = Number(p.lng);
  const acc = Number(p.accuracy);
  const photo = String(p.photo || '');
  const deviceInfo = String(p.deviceInfo || '').slice(0, 180);

  if (isNaN(lat) || isNaN(lng) || (lat === 0 && lng === 0)) {
    return { ok: false, message: 'Koordinat GPS tidak valid. Pastikan GPS perangkat aktif.' };
  }
  if (!isNaN(acc) && acc > 200) {
    return { ok: false, message: 'Akurasi GPS terlalu rendah (' + Math.round(acc) + 'm). Pindah ke area terbuka dan coba lagi.' };
  }
  if (!/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(photo)) {
    return { ok: false, message: 'Foto selfie wajib diambil langsung dari kamera.' };
  }
  if (photo.length > PHOTO_MAX_CHARS) {
    return { ok: false, message: 'Ukuran foto selfie terlalu besar.' };
  }

  const cfg = readWorkConfig_();
  const dist = haversineMeters_(lat, lng, cfg.officeLat, cfg.officeLng);
  if (dist > cfg.officeRadius) {
    return {
      ok: false,
      code: 'OUT_OF_RANGE',
      distance: dist,
      maxRadius: cfg.officeRadius,
      message: 'Absen ditolak! Lokasi Anda berada ' + dist + ' meter dari perusahaan (maksimal radius ' + cfg.officeRadius + ' meter).'
    };
  }

  const nowDate = new Date();
  const todayStr = Utilities.formatDate(nowDate, 'Asia/Jakarta', 'yyyy-MM-dd');
  const timeStr = Utilities.formatDate(nowDate, 'Asia/Jakarta', 'HH:mm:ss');
  const nowMin = parseMinutesOfDay_(Utilities.formatDate(nowDate, 'Asia/Jakarta', 'HH:mm'));

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = plainSheet_(ATTENDANCE_SHEET, ATTENDANCE_HEADERS);
    const rows = attendanceRows_();
    const existing = rows.find(function (r) {
      return String(r.date) === todayStr && lower_(r.email) === lower_(me.email);
    });

    if (type === 'IN') {
      if (existing && existing.clockIn) {
        return { ok: false, message: 'Anda sudah melakukan Absen Masuk hari ini pada pukul ' + existing.clockIn };
      }
      const startMin = parseMinutesOfDay_(cfg.workStartTime);
      const limitMin = startMin + cfg.lateToleranceMin;
      const lateMinutes = nowMin > limitMin ? (nowMin - startMin) : 0;
      const status = lateMinutes > 0 ? 'TERLAMBAT' : 'TEPAT WAKTU';

      const newRow = [
        'att_' + Date.now(), todayStr, lower_(me.email), String(me.name), String(me.role),
        timeStr, String(lat), String(lng), String(Math.round(acc || 0)), String(dist), status, String(lateMinutes), photo,
        '', '', '', '', '', '0', '',
        deviceInfo
      ];
      sh.appendRow(newRow);

      return {
        ok: true,
        type: 'IN',
        status: status,
        time: timeStr,
        distance: dist,
        lateMinutes: lateMinutes,
        message: lateMinutes > 0
          ? 'Absen Masuk tercatat (' + timeStr + '). Anda terlambat ' + lateMinutes + ' menit.'
          : 'Absen Masuk tercatat (' + timeStr + '). Status: Tepat Waktu!'
      };
    } else {
      if (!existing || !existing.clockIn) {
        return { ok: false, message: 'Anda belum melakukan Absen Masuk hari ini.' };
      }
      if (existing.clockOut) {
        return { ok: false, message: 'Anda sudah melakukan Absen Pulang hari ini pada pukul ' + existing.clockOut };
      }
      const endMin = parseMinutesOfDay_(cfg.workEndTime);
      const rawOvertime = nowMin > endMin ? (nowMin - endMin) : 0;
      const overtimeMinutes = rawOvertime >= cfg.overtimeMinThreshold ? rawOvertime : 0;

      const updatedRow = [
        existing.id, existing.date, existing.email, existing.name, existing.role,
        existing.clockIn, existing.clockInLat, existing.clockInLng, existing.clockInAcc, existing.clockInDist, existing.clockInStatus, existing.lateMinutes, existing.clockInPhoto,
        timeStr, String(lat), String(lng), String(Math.round(acc || 0)), String(dist), String(overtimeMinutes), photo,
        existing.deviceInfo || deviceInfo
      ];
      sh.getRange(existing._row, 1, 1, ATTENDANCE_HEADERS.length).setValues([updatedRow]);

      return {
        ok: true,
        type: 'OUT',
        time: timeStr,
        distance: dist,
        overtimeMinutes: overtimeMinutes,
        message: overtimeMinutes > 0
          ? 'Absen Pulang tercatat (' + timeStr + '). Lembur terhitung: ' + overtimeMinutes + ' menit.'
          : 'Absen Pulang tercatat (' + timeStr + '). Terima kasih atas kerja keras Anda!'
      };
    }
  } finally {
    lock.releaseLock();
  }
}

function getAttendanceData_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };

  const cfg = readWorkConfig_();
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const all = attendanceRows_();
  const isAdmin = bool_(me.isAdmin);
  const viewAll = isAdmin && !bool_(p.mine);

  const todayRecord = all.find(function (r) {
    return String(r.date) === todayStr && lower_(r.email) === lower_(me.email);
  }) || null;

  const history = all
    .filter(function (r) { return viewAll || lower_(r.email) === lower_(me.email); })
    .sort(function (a, b) {
      const ka = String(a.date) + ' ' + String(a.clockIn);
      const kb = String(b.date) + ' ' + String(b.clockIn);
      return kb < ka ? -1 : 1;
    })
    .slice(0, 150)
    .map(function (r) {
      return {
        id: r.id, date: r.date, email: r.email, name: r.name, role: r.role,
        clockIn: r.clockIn, clockInLat: r.clockInLat, clockInLng: r.clockInLng,
        clockInAcc: r.clockInAcc, clockInDist: r.clockInDist,
        clockInStatus: r.clockInStatus, lateMinutes: Number(r.lateMinutes) || 0,
        clockInPhoto: r.clockInPhoto || '',
        clockOut: r.clockOut, clockOutLat: r.clockOutLat, clockOutLng: r.clockOutLng,
        clockOutAcc: r.clockOutAcc, clockOutDist: r.clockOutDist,
        overtimeMinutes: Number(r.overtimeMinutes) || 0,
        clockOutPhoto: r.clockOutPhoto || ''
      };
    });

  let absentToday = [];
  if (isAdmin) {
    const presentEmails = all
      .filter(function (r) { return String(r.date) === todayStr; })
      .map(function (r) { return lower_(r.email); });
    absentToday = rows_()
      .filter(function (u) { return bool_(u.verified) && presentEmails.indexOf(lower_(u.email)) === -1; })
      .map(function (u) { return { name: u.name, role: u.role, email: u.email }; });
  }

  return {
    ok: true,
    today: todayStr,
    config: cfg,
    todayRecord: todayRecord ? {
      clockIn: todayRecord.clockIn,
      clockInStatus: todayRecord.clockInStatus,
      lateMinutes: Number(todayRecord.lateMinutes) || 0,
      clockInDist: todayRecord.clockInDist,
      clockOut: todayRecord.clockOut,
      overtimeMinutes: Number(todayRecord.overtimeMinutes) || 0,
      clockOutDist: todayRecord.clockOutDist
    } : null,
    history: history,
    absentToday: absentToday,
    isAdmin: isAdmin
  };
}

function authorizeMail() {
  Logger.log('Kuota email tersisa: ' + MailApp.getRemainingDailyQuota());
  Logger.log('Spreadsheet terikat: ' + SpreadsheetApp.getActiveSpreadsheet().getName());
}
