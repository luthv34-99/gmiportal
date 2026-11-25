/**
 * ============================================================
 *  GMI Portal Backend (FINAL) — Google Apps Script
 *  PT. Global Mega Indonesia — Generator Laporan Inspeksi Mesin
 * ============================================================
 *  Cara pasang:
 *   1. Buat Google Sheet baru -> Ekstensi -> Apps Script -> tempel file ini.
 *   2. Jalankan fungsi  authorize()  sekali, setujui izin Sheets + Gmail.
 *   3. Deploy -> Deployment baru -> Jenis: Aplikasi web
 *        Jalankan sebagai : Saya
 *        Yang memiliki akses : Siapa saja (Anyone)
 *      (Setiap mengubah kode: Deploy -> Kelola deployment -> Versi baru)
 *   4. Daftar akun Anda lewat portal, lalu jalankan  makeAdmin()  sekali
 *      di editor (email diatur di ADMIN_EMAIL_TO_PROMOTE di bawah).
 * ============================================================
 */

const SHEET_NAME = 'GMI_Users';
const LEGACY_HEADERS = ['id', 'name', 'role', 'email', 'passwordHash', 'salt', 'isAdmin',
                        'verified', 'activationSent', 'activationToken', 'registeredAt', 'verifiedAt'];
// Status akun: pending (approved=FALSE) -> approved (Admin setuju, email aktivasi terkirim) -> aktif (verified=TRUE)
// photo = foto profil (data URL JPEG kecil) | passwordChangedAt = waktu ubah password terakhir (epoch ms)
const HEADERS = LEGACY_HEADERS.concat(['approved', 'approvedAt', 'photo', 'passwordChangedAt']);

const SENDER_NAME = 'Admin PT. Global Mega Indonesia';
const ADMIN_EMAIL_TO_PROMOTE = 'gmielectric1507@gmail.com';   // dipakai oleh makeAdmin()
const ALLOWED_PORTAL_PREFIX = 'https://luthv34-99.github.io/'; // tautan aktivasi hanya boleh ke portal ini

const SESSION_TTL = 21600;     // detik (maks 6 jam, batas CacheService)
const MAX_LOGIN_FAILS = 5;     // percobaan gagal per email
const LOCK_SECONDS = 900;      // 15 menit
const MAX_VERIFY_FAILS = 20;   // percobaan token salah (global) per 10 menit
const PASSWORD_CHANGE_INTERVAL_DAYS = 30;  // user hanya boleh ganti password 1x per 30 hari
const PHOTO_MAX_CHARS = 45000;             // batas sel Google Sheets = 50.000 karakter

// ---- Kop laporan (hanya Admin yang boleh mengubah) & database laporan ----
const CONFIG_SHEET = 'GMI_Config';          // key | value  (kop perusahaan resmi)
const REPORT_SHEET = 'GMI_Reports';         // 1 baris = 1 laporan (ringkasan)
const REPORT_DATA_SHEET = 'GMI_ReportData'; // isi laporan (JSON) dipecah per 45.000 karakter
const REPORT_HEADERS = ['id', 'docNo', 'customer', 'machineName', 'ownerEmail', 'ownerName',
                        'createdAt', 'updatedAt', 'chunks', 'sizeChars'];
const DATA_HEADERS = ['reportId', 'idx', 'chunk'];
const CONFIG_HEADERS = ['key', 'value'];
const CHUNK_SIZE = 45000;                   // batas sel Google Sheets = 50.000 karakter
const MAX_REPORT_CHARS = 4000000;           // batas ukuran satu laporan (~4 juta karakter)
const LOGO_MAX_CHARS = 45000;
const KOP_KEYS = ['compName', 'compAddress', 'compPhone', 'logo'];
const MAX_REPORT_LIST = 500;

/* ======================= ENTRY POINTS ======================= */

function doGet(e) { return handle_((e && e.parameter) || {}); }

function doPost(e) {
  let p = {};
  try {
    p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    p = (e && e.parameter) || {};
  }
  return handle_(p);
}

function handle_(p) {
  let out;
  try {
    switch (String(p.action || 'ping')) {
      case 'ping':              out = { ok: true, message: 'GMI portal backend aktif' }; break;
      case 'register':          out = register_(p); break;
      case 'login':             out = login_(p); break;
      case 'whoami':            out = whoami_(p); break;
      case 'logout':            out = logout_(p); break;
      case 'update_profile':    out = updateProfile_(p); break;       // user: foto profil
      case 'change_password':   out = changePassword_(p); break;      // user: maks 1x / 30 hari
      case 'verify_token':      out = verifyToken_(p); break;
      case 'list':              out = listUsers_(p); break;           // admin
      case 'activate_user':     out = activateUser_(p); break;        // admin
      case 'approve_user':      out = approveUser_(p); break;         // admin (approve + kirim email)
      case 'send_verification': out = sendVerification_(p); break;    // admin (kirim ulang email)
      case 'delete_user':       out = deleteUser_(p); break;          // admin
      case 'get_kop':           out = getKop_(p); break;              // semua user login
      case 'set_kop':           out = setKop_(p); break;              // HANYA admin
      case 'save_report':       out = saveReport_(p); break;          // user login (simpan/ubah milik sendiri)
      case 'list_reports':      out = listReports_(p); break;         // user: miliknya | admin: semua
      case 'get_report':        out = getReport_(p); break;           // pemilik / admin
      case 'delete_report':     out = deleteReport_(p); break;        // pemilik / admin
      case 'get_attendance_data':  out = attGetData_(p); break;    // user: data + riwayat (admin: semua)
      case 'submit_attendance':    out = attSubmit_(p); break;     // user: absen masuk/pulang (GPS + selfie)
      case 'get_attendance_photo': out = attPhoto_(p); break;      // pemilik / admin: foto bukti
      case 'set_work_config':      out = attConfigSet_(p); break;  // HANYA admin
      default: out = { ok: false, message: 'Action tidak dikenali: ' + p.action };
    }
  } catch (err) {
    out = { ok: false, message: 'Server error: ' + err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ======================= HELPERS ======================= */

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
  }
  if (sh.getMaxColumns() < HEADERS.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), HEADERS.length - sh.getMaxColumns());
  }
  const head = sh.getRange(1, 1, 1, HEADERS.length).getValues()[0].map(String);
  if (head.join('|') === HEADERS.join('|')) return sh;

  // Sheet versi lama (12 atau 14 kolom) -> tambahkan kolom baru secara otomatis
  let have = head.length;
  while (have > 0 && head[have - 1] === '') have--;
  const prefixOk = have >= LEGACY_HEADERS.length &&
                   head.slice(0, have).join('|') === HEADERS.slice(0, have).join('|');
  if (prefixOk) {
    migrateSchema_(sh, have);
    return sh;
  }
  throw new Error('Header sheet "' + SHEET_NAME + '" tidak sesuai. Hapus sheet tersebut agar dibuat ulang.');
}

function migrateSchema_(sh, have) {
  for (let c = have; c < HEADERS.length; c++) sh.getRange(1, c + 1).setValue(HEADERS[c]);
  if (have <= LEGACY_HEADERS.length) {   // kolom approved baru ditambahkan: isi untuk akun lama
    const iVerified = LEGACY_HEADERS.indexOf('verified');
    const iSent = LEGACY_HEADERS.indexOf('activationSent');
    const v = sh.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      const ok = bool_(v[i][iVerified]) || bool_(v[i][iSent]);
      sh.getRange(i + 1, LEGACY_HEADERS.length + 1).setValue(ok);
    }
  }
}

function rows_() {
  const v = sheet_().getDataRange().getValues();
  return v.slice(1).map(function (r, i) {
    const o = { _row: i + 2 };
    HEADERS.forEach(function (k, c) { o[k] = r[c]; });
    return o;
  });
}

function bool_(x) { return x === true || String(x).toUpperCase() === 'TRUE'; }
function lower_(x) { return String(x || '').trim().toLowerCase(); }
function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss') + ' WIB';
}

function hash_(salt, pw) {
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw, Utilities.Charset.UTF_8);
  return d.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function setCell_(row, header, value) {
  sheet_().getRange(row, HEADERS.indexOf(header) + 1).setValue(value);
}

function publicUser_(u, includeToken, self) {
  const o = {
    id: u.id, name: u.name, role: u.role, email: u.email,
    isAdmin: bool_(u.isAdmin), verified: bool_(u.verified),
    activationSent: bool_(u.activationSent), registeredAt: u.registeredAt,
    approved: bool_(u.approved) || bool_(u.verified), approvedAt: u.approvedAt
  };
  if (includeToken) o.activationToken = u.activationToken;
  if (self) {   // data milik user sendiri (tidak dikirim di daftar admin agar ringan)
    o.photo = String(u.photo || '');
    const next = nextPasswordChangeMs_(u);
    o.nextPasswordChangeAt = next && next > Date.now() ? new Date(next).toISOString() : '';
  }
  return o;
}

/* ---------- Sesi ---------- */

function newSession_(email) {
  const t = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('sess_' + t, lower_(email), SESSION_TTL);
  return t;
}

function sessionUser_(token) {
  if (!token) return null;
  const email = CacheService.getScriptCache().get('sess_' + token);
  if (!email) return null;
  const u = rows_().find(function (x) { return lower_(x.email) === email; });
  return (u && bool_(u.verified)) ? u : null;
}

function requireAdmin_(p) {
  const u = sessionUser_(p.sessionToken);
  return (u && bool_(u.isAdmin)) ? u : null;
}

/* ======================= ACTIONS ======================= */

function register_(p) {
  const name = String(p.name || '').trim().slice(0, 80);
  const role = String(p.role || '').trim().slice(0, 80);
  const email = lower_(p.email);
  const pw = String(p.password || '');

  if (!name || !role) return { ok: false, message: 'Nama dan jabatan wajib diisi' };
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { ok: false, message: 'Format email tidak valid' };
  if (pw.length < 6) return { ok: false, message: 'Password minimal 6 karakter' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (rows_().some(function (u) { return lower_(u.email) === email; })) {
      return { ok: false, message: 'Email sudah terdaftar' };
    }
    const salt = Utilities.getUuid();
    const token = 'ACT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
    sheet_().appendRow(['usr_' + Date.now(), name, role, email, hash_(salt, pw), salt,
                        false, false, false, token, now_(), '', false, '', '', '']);
    return { ok: true, message: 'Pendaftaran berhasil, menunggu persetujuan Admin' };
  } finally {
    lock.releaseLock();
  }
}

function login_(p) {
  const email = lower_(p.email);
  const pw = String(p.password || '');
  const cache = CacheService.getScriptCache();
  const fails = parseInt(cache.get('fail_' + email) || '0', 10);

  if (fails >= MAX_LOGIN_FAILS) {
    return { ok: false, message: 'Terlalu banyak percobaan gagal. Coba lagi dalam 15 menit.' };
  }

  const u = rows_().find(function (x) { return lower_(x.email) === email; });
  if (!u || hash_(u.salt, pw) !== u.passwordHash) {
    cache.put('fail_' + email, String(fails + 1), LOCK_SECONDS);
    return { ok: false, message: 'Email atau password salah' };
  }
  if (!bool_(u.verified)) {
    cache.remove('fail_' + email);
    if (!bool_(u.approved)) {
      return { ok: false, message: 'Akun Anda belum disetujui Admin. Mohon tunggu persetujuan.' };
    }
    return { ok: false, message: 'Akun sudah disetujui Admin. Buka email aktivasi lalu klik tautannya atau masukkan kodenya di tab Aktivasi.' };
  }
  cache.remove('fail_' + email);
  return { ok: true, user: publicUser_(u, false, true), token: newSession_(u.email) };
}

function whoami_(p) {
  const u = sessionUser_(p.sessionToken);
  if (!u) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };
  return { ok: true, user: publicUser_(u, false, true) };
}

function logout_(p) {
  if (p.sessionToken) CacheService.getScriptCache().remove('sess_' + p.sessionToken);
  return { ok: true };
}

/* ---------- Profil user ---------- */

function toMs_(x) {
  if (x === '' || x === null || x === undefined) return 0;
  if (Object.prototype.toString.call(x) === '[object Date]') return x.getTime();
  const n = Number(x);
  return isNaN(n) ? 0 : n;
}

/** Waktu (ms) paling awal user boleh mengganti password lagi; 0 = boleh sekarang. */
function nextPasswordChangeMs_(u) {
  const last = toMs_(u.passwordChangedAt);
  return last ? last + PASSWORD_CHANGE_INTERVAL_DAYS * 24 * 3600 * 1000 : 0;
}

/** User mengubah / menghapus foto profilnya sendiri. photo = data URL gambar kecil, atau '' untuk menghapus. */
function updateProfile_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };
  if (p.photo === undefined || p.photo === null) return { ok: false, message: 'Data foto tidak ada' };

  const photo = String(p.photo);
  if (photo === '') {
    setCell_(me._row, 'photo', '');
    return { ok: true, photo: '', message: 'Foto profil dihapus' };
  }
  if (!/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(photo)) {
    return { ok: false, message: 'Format foto tidak valid (gunakan JPG, PNG, atau WebP)' };
  }
  if (photo.length > PHOTO_MAX_CHARS) {
    return { ok: false, message: 'Ukuran foto terlalu besar setelah diproses. Coba foto lain.' };
  }
  setCell_(me._row, 'photo', photo);
  return { ok: true, photo: photo, message: 'Foto profil diperbarui' };
}

/** User mengganti password sendiri. Hanya boleh 1x setiap PASSWORD_CHANGE_INTERVAL_DAYS hari. */
function changePassword_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };

  const cur = String(p.currentPassword || '');
  const nw = String(p.newPassword || '');
  const email = lower_(me.email);
  const cache = CacheService.getScriptCache();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const u = findUser_(email);
    if (!u) return { ok: false, message: 'User tidak ditemukan' };

    const nextMs = nextPasswordChangeMs_(u);
    if (nextMs && Date.now() < nextMs) {
      return {
        ok: false,
        nextPasswordChangeAt: new Date(nextMs).toISOString(),
        message: 'Password hanya dapat diubah 1 kali setiap ' + PASSWORD_CHANGE_INTERVAL_DAYS +
                 ' hari. Perubahan berikutnya tersedia mulai ' +
                 Utilities.formatDate(new Date(nextMs), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm') + ' WIB.'
      };
    }

    const fails = parseInt(cache.get('fail_' + email) || '0', 10);
    if (fails >= MAX_LOGIN_FAILS) {
      return { ok: false, message: 'Terlalu banyak percobaan gagal. Coba lagi dalam 15 menit.' };
    }
    if (hash_(u.salt, cur) !== u.passwordHash) {
      cache.put('fail_' + email, String(fails + 1), LOCK_SECONDS);
      return { ok: false, message: 'Password saat ini salah' };
    }
    cache.remove('fail_' + email);

    if (nw.length < 6) return { ok: false, message: 'Password baru minimal 6 karakter' };
    if (nw.length > 100) return { ok: false, message: 'Password baru terlalu panjang' };
    if (nw === cur) return { ok: false, message: 'Password baru harus berbeda dari password saat ini' };

    const salt = Utilities.getUuid();
    const nowMs = Date.now();
    setCell_(u._row, 'passwordHash', hash_(salt, nw));
    setCell_(u._row, 'salt', salt);
    setCell_(u._row, 'passwordChangedAt', nowMs);
    return {
      ok: true,
      message: 'Password berhasil diubah',
      nextPasswordChangeAt: new Date(nowMs + PASSWORD_CHANGE_INTERVAL_DAYS * 24 * 3600 * 1000).toISOString()
    };
  } finally {
    lock.releaseLock();
  }
}

function verifyToken_(p) {
  const cache = CacheService.getScriptCache();
  const bad = parseInt(cache.get('verify_fails') || '0', 10);
  if (bad >= MAX_VERIFY_FAILS) {
    return { ok: false, message: 'Terlalu banyak kode salah. Coba lagi beberapa menit lagi.' };
  }

  const token = String(p.token || '').trim().toUpperCase();
  const email = lower_(p.email);
  const u = token ? rows_().find(function (x) {
    return String(x.activationToken).toUpperCase() === token && (!email || lower_(x.email) === email);
  }) : null;

  if (!u) {
    cache.put('verify_fails', String(bad + 1), 600);
    return { ok: false, message: 'Kode token tidak valid' };
  }
  if (!bool_(u.verified)) {
    if (!bool_(u.approved)) return { ok: false, message: 'Akun belum disetujui Admin' };
    setCell_(u._row, 'verified', true);
    setCell_(u._row, 'verifiedAt', now_());
  }
  return { ok: true, email: u.email };
}

function listUsers_(p) {
  if (!requireAdmin_(p)) return { ok: false, message: 'Akses ditolak: hanya Admin' };
  return { ok: true, users: rows_().map(function (u) { return publicUser_(u, true); }) };
}

function activateUser_(p) {
  if (!requireAdmin_(p)) return { ok: false, message: 'Akses ditolak: hanya Admin' };
  const email = lower_(p.email);
  const u = rows_().find(function (x) { return lower_(x.email) === email; });
  if (!u) return { ok: false, message: 'User tidak ditemukan' };
  setCell_(u._row, 'approved', true);
  if (!u.approvedAt) setCell_(u._row, 'approvedAt', now_());
  setCell_(u._row, 'verified', true);
  setCell_(u._row, 'verifiedAt', now_());
  return { ok: true, message: 'Akun ' + u.email + ' diaktifkan' };
}

function findUser_(email) {
  const e = lower_(email);
  return rows_().find(function (x) { return lower_(x.email) === e; }) || null;
}

function checkPortal_(portal) {
  return String(portal || '').indexOf(ALLOWED_PORTAL_PREFIX) === 0;
}

function sendActivationEmail_(u, portal) {
  const link = portal + '?activate_email=' + encodeURIComponent(u.email) +
               '&activate_token=' + encodeURIComponent(u.activationToken);
  const body =
    'Halo ' + u.name + ',\n\n' +
    'Pendaftaran akun Anda di Portal Generator Laporan Inspeksi Mesin (PT. Global Mega Indonesia) ' +
    'telah DISETUJUI oleh Admin.\n\n' +
    'Kode Aktivasi : ' + u.activationToken + '\n' +
    'Tautan verifikasi:\n' + link + '\n\n' +
    'Atau masukkan kode di tab "Aktivasi Kode Gmail" pada halaman login.\n\n' +
    'Hormat kami,\nAdmin PT. Global Mega Indonesia\nKawasan Industri Jababeka 1, Cikarang Utara - Bekasi';

  MailApp.sendEmail({
    to: u.email,
    subject: '[AKTIVASI AKUN GMI] Akun Anda Disetujui - Verifikasi Akses Laporan Inspeksi Mesin',
    body: body,
    name: SENDER_NAME
  });
}

/** Admin menyetujui akun -> status approved + email aktivasi dikirim otomatis. */
function approveUser_(p) {
  if (!requireAdmin_(p)) return { ok: false, message: 'Akses ditolak: hanya Admin' };
  const u = findUser_(p.email);
  if (!u) return { ok: false, message: 'User tidak ditemukan' };
  if (bool_(u.verified)) return { ok: false, message: 'Akun sudah aktif' };
  if (!checkPortal_(p.portalUrl)) return { ok: false, message: 'URL portal tidak diizinkan' };

  setCell_(u._row, 'approved', true);
  setCell_(u._row, 'approvedAt', now_());
  try {
    sendActivationEmail_(u, p.portalUrl);
    setCell_(u._row, 'activationSent', true);
    return { ok: true, emailSent: true, message: 'Akun ' + u.email + ' disetujui & email aktivasi terkirim' };
  } catch (err) {
    // Persetujuan tetap tersimpan; Admin bisa kirim ulang setelah masalah email diperbaiki.
    return { ok: true, emailSent: false, message: 'Akun disetujui, tetapi email gagal dikirim: ' + err.message };
  }
}

/** Kirim ulang email aktivasi untuk akun yang sudah disetujui. */
function sendVerification_(p) {
  if (!requireAdmin_(p)) return { ok: false, message: 'Akses ditolak: hanya Admin' };
  const u = findUser_(p.email);
  if (!u) return { ok: false, message: 'User tidak ditemukan' };
  if (bool_(u.verified)) return { ok: false, message: 'Akun sudah aktif' };
  if (!bool_(u.approved)) return { ok: false, message: 'Akun belum disetujui. Gunakan tombol Approve terlebih dahulu.' };
  if (!checkPortal_(p.portalUrl)) return { ok: false, message: 'URL portal tidak diizinkan' };

  sendActivationEmail_(u, p.portalUrl);
  setCell_(u._row, 'activationSent', true);
  return { ok: true, message: 'Email verifikasi terkirim ke ' + u.email };
}

/** Hapus akun user. Admin tidak bisa menghapus dirinya sendiri atau akun Admin lain. */
function deleteUser_(p) {
  const admin = requireAdmin_(p);
  if (!admin) return { ok: false, message: 'Akses ditolak: hanya Admin' };
  const email = lower_(p.email);
  if (email === lower_(admin.email)) return { ok: false, message: 'Anda tidak dapat menghapus akun Anda sendiri' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const u = findUser_(email);
    if (!u) return { ok: false, message: 'User tidak ditemukan' };
    if (bool_(u.isAdmin)) return { ok: false, message: 'Akun Admin tidak dapat dihapus dari panel' };
    sheet_().deleteRow(u._row);
    return { ok: true, message: 'Akun ' + u.email + ' dihapus' };
  } finally {
    lock.releaseLock();
  }
}

/* ======================= KOP LAPORAN (ADMIN) ======================= */

/** Sheet teks-biasa (format '@') agar nomor seperti 0001 tidak berubah jadi angka 1. */
function plainSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function readConfig_() {
  const sh = plainSheet_(CONFIG_SHEET, CONFIG_HEADERS);
  const last = sh.getLastRow();
  const cfg = {};
  if (last < 2) return cfg;
  sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) { cfg[String(r[0])] = String(r[1]); });
  return cfg;
}

function writeConfig_(obj) {
  const sh = plainSheet_(CONFIG_SHEET, CONFIG_HEADERS);
  const last = sh.getLastRow();
  const rowOf = {};
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { rowOf[String(r[0])] = i + 2; });
  }
  Object.keys(obj).forEach(function (k) {
    const v = String(obj[k]);
    if (rowOf[k]) {
      sh.getRange(rowOf[k], 2).setValue(v);
    } else {
      const nr = sh.getLastRow() + 1;
      if (sh.getMaxRows() < nr) sh.insertRowsAfter(sh.getMaxRows(), 1);
      sh.getRange(nr, 1, 1, 2).setValues([[k, v]]);
      rowOf[k] = nr;
    }
  });
}

/** Kop resmi untuk semua user yang sudah login. kop = null jika Admin belum pernah menyimpan. */
function getKop_(p) {
  if (!sessionUser_(p.sessionToken)) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };
  const cfg = readConfig_();
  if (!cfg.kopUpdatedAt) return { ok: true, kop: null };
  return {
    ok: true,
    kop: {
      compName: cfg.compName || '', compAddress: cfg.compAddress || '',
      compPhone: cfg.compPhone || '', logo: cfg.logo || '',
      updatedAt: cfg.kopUpdatedAt, updatedBy: cfg.kopUpdatedBy || ''
    }
  };
}

/** HANYA Admin: simpan kop resmi (nama, alamat, telepon, logo) untuk semua pengguna. */
function setKop_(p) {
  const admin = requireAdmin_(p);
  if (!admin) return { ok: false, message: 'Akses ditolak: hanya Admin yang dapat mengubah kop laporan' };

  const compName = String(p.compName || '').trim();
  const compAddress = String(p.compAddress || '').replace(/\r/g, '').trim();
  const compPhone = String(p.compPhone || '').trim();
  const logo = String(p.logo || '');

  if (!compName) return { ok: false, message: 'Nama perusahaan pada kop wajib diisi' };
  if (compName.length > 120) return { ok: false, message: 'Nama perusahaan terlalu panjang (maks 120 karakter)' };
  if (compAddress.length > 400) return { ok: false, message: 'Alamat terlalu panjang (maks 400 karakter)' };
  if (compPhone.length > 160) return { ok: false, message: 'Teks telepon terlalu panjang (maks 160 karakter)' };
  if (logo !== '') {
    if (!/^data:image\/(jpeg|png|webp);base64,[A-Za-z0-9+\/=]+$/.test(logo)) {
      return { ok: false, message: 'Format logo tidak valid (gunakan JPG, PNG, atau WebP)' };
    }
    if (logo.length > LOGO_MAX_CHARS) {
      return { ok: false, message: 'Ukuran logo terlalu besar setelah diproses. Gunakan gambar yang lebih sederhana.' };
    }
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    writeConfig_({
      compName: compName, compAddress: compAddress, compPhone: compPhone, logo: logo,
      kopUpdatedAt: now_(), kopUpdatedBy: lower_(admin.email)
    });
  } finally {
    lock.releaseLock();
  }
  return { ok: true, message: 'Kop laporan resmi disimpan untuk semua pengguna' };
}

/* ======================= DATABASE LAPORAN ======================= */

function reportRows_() {
  const sh = plainSheet_(REPORT_SHEET, REPORT_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, REPORT_HEADERS.length).getValues().map(function (r, i) {
    const o = { _row: i + 2 };
    REPORT_HEADERS.forEach(function (k, c) { o[k] = r[c]; });
    return o;
  });
}

function findReport_(id) {
  const sid = String(id || '');
  if (!sid) return null;
  return reportRows_().find(function (r) { return String(r.id) === sid; }) || null;
}

function canAccessReport_(user, rep) {
  return bool_(user.isAdmin) || lower_(rep.ownerEmail) === lower_(user.email);
}

/** Nomor baris (di sheet data) milik satu laporan. */
function chunkRowNumbers_(id) {
  const sh = plainSheet_(REPORT_DATA_SHEET, DATA_HEADERS);
  const last = sh.getLastRow();
  const rows = [];
  if (last < 2) return rows;
  sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) {
    if (String(r[0]) === String(id)) rows.push(i + 2);
  });
  return rows;
}

function deleteChunks_(id) {
  const sh = plainSheet_(REPORT_DATA_SHEET, DATA_HEADERS);
  const rows = chunkRowNumbers_(id);
  for (let i = rows.length - 1; i >= 0; i--) sh.deleteRow(rows[i]);   // dari bawah agar nomor baris tidak bergeser
}

function summaryOf_(r) {
  return {
    id: String(r.id), docNo: String(r.docNo), customer: String(r.customer),
    machineName: String(r.machineName), ownerEmail: String(r.ownerEmail), ownerName: String(r.ownerName),
    createdAt: String(r.createdAt), updatedAt: String(r.updatedAt), sizeChars: Number(r.sizeChars) || 0
  };
}

/**
 * Simpan laporan baru, atau perbarui laporan (reportId) milik sendiri.
 * p.dataJson = isi formulir (string JSON). Kop TIDAK ikut disimpan: selalu memakai kop resmi dari Admin.
 */
function saveReport_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };

  const dataJson = String(p.dataJson || '');
  if (!dataJson || dataJson.charAt(0) !== '{') return { ok: false, code: 'bad', message: 'Data laporan kosong atau tidak valid' };
  if (dataJson.length > MAX_REPORT_CHARS) {
    return { ok: false, code: 'toobig', message: 'Laporan terlalu besar (' + Math.round(dataJson.length / 1000) +
             ' ribu karakter, maks ' + Math.round(MAX_REPORT_CHARS / 1000) + ' ribu). Kurangi jumlah/ukuran foto.' };
  }
  const docNo = String(p.docNo || '').trim().slice(0, 40);
  const customer = String(p.customer || '').trim().slice(0, 120);
  const machineName = String(p.machineName || '').trim().slice(0, 120);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const repSheet = plainSheet_(REPORT_SHEET, REPORT_HEADERS);
    const dataSheet = plainSheet_(REPORT_DATA_SHEET, DATA_HEADERS);

    let existing = null;
    if (p.reportId) {
      existing = findReport_(p.reportId);
      if (!existing) return { ok: false, code: 'notfound', message: 'Laporan tidak ditemukan di database' };
      if (!canAccessReport_(me, existing)) return { ok: false, code: 'forbidden', message: 'Laporan ini milik pengguna lain' };
    }

    const id = existing ? String(existing.id) : 'rpt_' + Date.now() + '_' + Utilities.getUuid().slice(0, 6);
    if (existing) deleteChunks_(id);

    // isi laporan -> sel-sel berukuran aman; awalan "D" mencegah teks diperlakukan sebagai rumus
    const chunks = [];
    for (let i = 0; i < dataJson.length; i += CHUNK_SIZE) chunks.push(dataJson.slice(i, i + CHUNK_SIZE));
    const start = dataSheet.getLastRow() + 1;
    const need = start + chunks.length - 1;
    if (dataSheet.getMaxRows() < need) dataSheet.insertRowsAfter(dataSheet.getMaxRows(), need - dataSheet.getMaxRows());
    dataSheet.getRange(start, 1, chunks.length, 3)
      .setValues(chunks.map(function (c, i) { return [id, String(i), 'D' + c]; }));

    const ts = now_();
    if (existing) {
      repSheet.getRange(existing._row, 1, 1, REPORT_HEADERS.length).setValues([[
        id, docNo, customer, machineName, existing.ownerEmail, existing.ownerName,
        existing.createdAt, ts, String(chunks.length), String(dataJson.length)
      ]]);
    } else {
      const nr = repSheet.getLastRow() + 1;
      if (repSheet.getMaxRows() < nr) repSheet.insertRowsAfter(repSheet.getMaxRows(), 1);
      repSheet.getRange(nr, 1, 1, REPORT_HEADERS.length).setValues([[
        id, docNo, customer, machineName, lower_(me.email), String(me.name),
        ts, ts, String(chunks.length), String(dataJson.length)
      ]]);
    }
    return { ok: true, id: id, updated: Boolean(existing), updatedAt: ts,
             message: existing ? 'Laporan diperbarui di database' : 'Laporan disimpan ke database' };
  } finally {
    lock.releaseLock();
  }
}

/** Daftar ringkasan laporan: user biasa = miliknya sendiri, Admin = semua (atau hanya miliknya jika p.mine). */
function listReports_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };
  const all = bool_(me.isAdmin) && !bool_(p.mine);
  const list = reportRows_()
    .filter(function (r) { return all || lower_(r.ownerEmail) === lower_(me.email); })
    .sort(function (a, b) { return String(b.updatedAt) < String(a.updatedAt) ? -1 : 1; })
    .slice(0, MAX_REPORT_LIST)
    .map(summaryOf_);
  return { ok: true, reports: list, scope: all ? 'all' : 'mine' };
}

function getReport_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };
  const rep = findReport_(p.id);
  if (!rep) return { ok: false, code: 'notfound', message: 'Laporan tidak ditemukan' };
  if (!canAccessReport_(me, rep)) return { ok: false, code: 'forbidden', message: 'Akses ditolak: laporan milik pengguna lain' };

  const rows = chunkRowNumbers_(rep.id);
  if (!rows.length) return { ok: false, code: 'empty', message: 'Isi laporan tidak ditemukan di sheet ' + REPORT_DATA_SHEET };
  const sh = plainSheet_(REPORT_DATA_SHEET, DATA_HEADERS);
  const first = rows[0], lastRow = rows[rows.length - 1];
  const parts = [];
  sh.getRange(first, 1, lastRow - first + 1, 3).getValues().forEach(function (r) {
    if (String(r[0]) === String(rep.id)) parts.push({ idx: Number(r[1]), chunk: String(r[2]).slice(1) });
  });
  parts.sort(function (a, b) { return a.idx - b.idx; });
  const dataJson = parts.map(function (x) { return x.chunk; }).join('');
  if (dataJson.length !== Number(rep.sizeChars)) {
    return { ok: false, code: 'corrupt', message: 'Isi laporan tidak lengkap di database (ukuran tidak cocok).' };
  }
  return { ok: true, report: Object.assign(summaryOf_(rep), { dataJson: dataJson }) };
}

function deleteReport_(p) {
  const me = sessionUser_(p.sessionToken);
  if (!me) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const rep = findReport_(p.id);
    if (!rep) return { ok: false, code: 'notfound', message: 'Laporan tidak ditemukan' };
    if (!canAccessReport_(me, rep)) return { ok: false, code: 'forbidden', message: 'Akses ditolak: laporan milik pengguna lain' };
    deleteChunks_(rep.id);
    plainSheet_(REPORT_SHEET, REPORT_HEADERS).deleteRow(rep._row);
    return { ok: true, message: 'Laporan dihapus dari database' };
  } finally {
    lock.releaseLock();
  }
}

/* ======================= ABSENSI KARYAWAN ======================= */

const ATT_SHEET = 'GMI_Attendance';
const ATT_PHOTO_SHEET = 'GMI_AttPhotos';   // foto selfie dipisah agar tabel absensi tetap ringan
const ATT_HEADERS = ['date', 'email', 'name', 'role', 'clockIn', 'clockInStatus', 'lateMinutes', 'clockInLat', 'clockInLng',
  'clockInAcc', 'clockInDist', 'clockOut', 'overtimeMinutes', 'clockOutLat', 'clockOutLng', 'clockOutAcc', 'clockOutDist',
  'deviceIn', 'deviceOut', 'createdAt', 'updatedAt', 'flags'];
const ATT_PHOTO_HEADERS = ['key', 'photo'];
const ATT_DEFAULTS = { officeLat: '', officeLng: '', officeRadius: '100', workStartTime: '08:00', workEndTime: '17:00',
  lateToleranceMin: '10', overtimeMinThreshold: '30' };
const ATT_MAX_ACCURACY = 200;       // meter, sama dengan batas di halaman absensi
const ATT_PHOTO_MAX = 45000;        // batas sel Google Sheets = 50.000 karakter
const ATT_HISTORY_ADMIN = 300, ATT_HISTORY_USER = 120;

function attSheet_() { return plainSheet_(ATT_SHEET, ATT_HEADERS); }
function attPhotoSheet_() { return plainSheet_(ATT_PHOTO_SHEET, ATT_PHOTO_HEADERS); }

function attConfig_() {
  const cfg = readConfig_(), o = {};
  Object.keys(ATT_DEFAULTS).forEach(function (k) {
    const v = cfg['att_' + k];
    o[k] = (v === undefined || v === '') ? ATT_DEFAULTS[k] : v;
  });
  return {
    officeLat: o.officeLat === '' ? '' : parseFloat(o.officeLat), officeLng: o.officeLng === '' ? '' : parseFloat(o.officeLng),
    officeRadius: parseInt(o.officeRadius, 10) || 100, workStartTime: o.workStartTime, workEndTime: o.workEndTime,
    lateToleranceMin: parseInt(o.lateToleranceMin, 10) || 0, overtimeMinThreshold: parseInt(o.overtimeMinThreshold, 10) || 0
  };
}

function attNow_() {
  const d = new Date();
  return { date: Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM-dd'), time: Utilities.formatDate(d, 'Asia/Jakarta', 'HH:mm') };
}
function attValidTime_(s) { return /^([01]\d|2[0-3]):[0-5]\d$/.test(String(s)); }
function attMin_(hhmm) { const a = String(hhmm).split(':'); return parseInt(a[0], 10) * 60 + parseInt(a[1], 10); }
function attIsWeekday_(date) {   // Senin-Jumat
  const a = date.split('-').map(Number);
  const g = new Date(Date.UTC(a[0], a[1] - 1, a[2])).getUTCDay();
  return g >= 1 && g <= 5;
}
/** Penanda GPS mencurigakan (tidak memblokir, tetapi ditandai untuk Admin). */
function attFlags_(p, lat, lng, acc, rec) {
  const f = [];
  if (acc !== null && acc < 1) f.push('akurasi 0 m');
  const dec = function (x) { const m = String(x).match(/\.(\d+)/); return m ? m[1].length : 0; };
  if (dec(lat) <= 3 && dec(lng) <= 3) f.push('koordinat dibulatkan');
  const gt = attNum_(p.gpsTime);
  if (gt !== null && Math.abs(Date.now() - gt) > 120000) f.push('waktu GPS tidak sinkron');
  if (rec && rec.clockInLat !== '' && String(rec.clockInLat) === String(lat) && String(rec.clockInLng) === String(lng)) f.push('koordinat sama persis dengan absen masuk');
  return f;
}
function attNum_(x) { const n = parseFloat(x); return isFinite(n) ? n : null; }

function attDistance_(lat1, lng1, lat2, lng2) {   // haversine, meter
  const R = 6371000, rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad, dLng = (lng2 - lng1) * rad;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return Math.round(2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)));
}

function attRows_() {
  const sh = attSheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ATT_HEADERS.length).getValues().map(function (r, i) {
    const o = { _row: i + 2 };
    ATT_HEADERS.forEach(function (k, c) { o[k] = String(r[c] === undefined || r[c] === null ? '' : r[c]); });
    return o;
  });
}

function attPublic_(r) {
  if (!r) return null;
  const o = {};
  ATT_HEADERS.forEach(function (k) { if (k !== 'deviceIn' && k !== 'deviceOut') o[k] = r[k]; });
  o.lateMinutes = parseInt(r.lateMinutes, 10) || 0;
  o.overtimeMinutes = parseInt(r.overtimeMinutes, 10) || 0;
  return o;
}

function attWrite_(sh, rec, rowNo) {
  const arr = ATT_HEADERS.map(function (k) { return rec[k] === undefined || rec[k] === null ? '' : String(rec[k]); });
  if (!rowNo) {
    rowNo = sh.getLastRow() + 1;
    if (sh.getMaxRows() < rowNo) sh.insertRowsAfter(sh.getMaxRows(), 1);
  }
  sh.getRange(rowNo, 1, 1, ATT_HEADERS.length).setValues([arr]);
  return rowNo;
}

function attSavePhoto_(key, photo) {
  const sh = attPhotoSheet_(), last = sh.getLastRow();
  let rowNo = 0;
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) { if (String(r[0]) === key) rowNo = i + 2; });
  }
  if (!rowNo) {
    rowNo = last + 1;
    if (sh.getMaxRows() < rowNo) sh.insertRowsAfter(sh.getMaxRows(), 1);
  }
  sh.getRange(rowNo, 1, 1, 2).setValues([[key, photo]]);
}

function attPhotoOf_(key) {
  const sh = attPhotoSheet_(), last = sh.getLastRow();
  if (last < 2) return '';
  const v = sh.getRange(2, 1, last - 1, 2).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === key) return String(v[i][1]);
  return '';
}

/** Data untuk halaman absensi: konfigurasi, catatan hari ini, riwayat, dan (admin) daftar yang belum absen. */
function attGetData_(p) {
  const u = sessionUser_(p.sessionToken);
  if (!u) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };
  const isAdmin = bool_(u.isAdmin), cfg = attConfig_(), n = attNow_();
  const all = attRows_();
  const mine = lower_(u.email);
  const today = all.find(function (r) { return r.date === n.date && lower_(r.email) === mine; });
  const onlyMine = !isAdmin || p.mine === true || p.mine === 'true';
  const list = all.filter(function (r) { return onlyMine ? lower_(r.email) === mine : true; })
    .sort(function (a, b) { return (a.date + a.clockIn) < (b.date + b.clockIn) ? 1 : -1; })
    .slice(0, isAdmin ? ATT_HISTORY_ADMIN : ATT_HISTORY_USER).map(attPublic_);
  const out = { ok: true, isAdmin: isAdmin, config: cfg, todayRecord: attPublic_(today), history: list, serverDate: n.date, serverTime: n.time };
  if (isAdmin) {
    const done = {};
    all.forEach(function (r) { if (r.date === n.date) done[lower_(r.email)] = true; });
    const pastDeadline = attMin_(n.time) > attMin_(cfg.workStartTime) + cfg.lateToleranceMin;
    out.absentToday = (attIsWeekday_(n.date) && pastDeadline) ? rows_().filter(function (x) {
      return bool_(x.verified) && !bool_(x.isAdmin) && !done[lower_(x.email)];
    }).map(function (x) { return { name: x.name, role: x.role, email: lower_(x.email) }; }) : [];
  }
  return out;
}

function attSubmit_(p) {
  const u = sessionUser_(p.sessionToken);
  if (!u) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };
  const type = String(p.type || '').toUpperCase();
  if (type !== 'IN' && type !== 'OUT') return { ok: false, message: 'Jenis absen tidak valid' };
  const cfg = attConfig_();
  if (cfg.officeLat === '' || cfg.officeLng === '' || isNaN(cfg.officeLat) || isNaN(cfg.officeLng)) {
    return { ok: false, message: 'Admin belum mengatur lokasi kantor. Hubungi Admin.' };
  }
  const lat = attNum_(p.lat), lng = attNum_(p.lng), acc = attNum_(p.accuracy);
  if (lat === null || lng === null || Math.abs(lat) > 90 || Math.abs(lng) > 180) return { ok: false, message: 'Lokasi GPS tidak valid' };
  if (acc !== null && acc > ATT_MAX_ACCURACY) return { ok: false, message: 'Akurasi GPS terlalu rendah (±' + Math.round(acc) + ' m). Pindah ke area terbuka lalu refresh GPS.' };
  const dist = attDistance_(lat, lng, cfg.officeLat, cfg.officeLng);
  if (dist > cfg.officeRadius) return { ok: false, message: 'Anda berada ' + dist + ' m dari kantor. Maksimal ' + cfg.officeRadius + ' m.' };
  const photo = String(p.photo || '');
  if (!/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(photo)) return { ok: false, message: 'Foto selfie wajib diambil dari kamera' };
  if (photo.length > ATT_PHOTO_MAX) return { ok: false, message: 'Ukuran foto terlalu besar. Ambil ulang foto.' };
  const device = String(p.deviceInfo || '').slice(0, 200);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const n = attNow_(), stamp = n.date + ' ' + n.time, email = lower_(u.email);
    const sh = attSheet_();
    const rec = attRows_().find(function (r) { return r.date === n.date && lower_(r.email) === email; });
    if (type === 'IN') {
      if (rec && rec.clockIn) return { ok: false, message: 'Anda sudah absen masuk hari ini pukul ' + rec.clockIn + '.' };
      const diff = attMin_(n.time) - attMin_(cfg.workStartTime);
      const late = diff > cfg.lateToleranceMin;
      const row = { date: n.date, email: email, name: u.name, role: u.role, clockIn: n.time,
        clockInStatus: late ? 'TERLAMBAT' : 'TEPAT WAKTU', lateMinutes: late ? diff : 0,
        clockInLat: lat, clockInLng: lng, clockInAcc: acc === null ? '' : Math.round(acc), clockInDist: dist,
        overtimeMinutes: 0, deviceIn: device, createdAt: stamp, updatedAt: stamp,
        flags: attFlags_(p, lat, lng, acc, null).join('; ') };
      attSavePhoto_(n.date + '|' + email + '|IN', photo);
      attWrite_(sh, row, rec ? rec._row : 0);
      return { ok: true, message: late ? 'Absen masuk tercatat pukul ' + n.time + ' (terlambat ' + diff + ' menit)' : 'Absen masuk tercatat pukul ' + n.time + ' WIB. Tepat waktu!' };
    }
    if (!rec || !rec.clockIn) return { ok: false, message: 'Anda belum absen masuk hari ini.' };
    if (rec.clockOut) return { ok: false, message: 'Anda sudah absen pulang hari ini pukul ' + rec.clockOut + '.' };
    if (attMin_(n.time) < attMin_(rec.clockIn)) return { ok: false, message: 'Jam pulang tidak boleh sebelum jam masuk.' };
    const over = attMin_(n.time) - attMin_(cfg.workEndTime);
    rec.clockOut = n.time;
    rec.overtimeMinutes = over >= cfg.overtimeMinThreshold && over > 0 ? over : 0;
    rec.clockOutLat = lat; rec.clockOutLng = lng; rec.clockOutAcc = acc === null ? '' : Math.round(acc); rec.clockOutDist = dist;
    rec.deviceOut = device; rec.updatedAt = stamp;
    const fo = attFlags_(p, lat, lng, acc, rec);
    if (fo.length) rec.flags = (rec.flags ? rec.flags + '; ' : '') + 'pulang: ' + fo.join(', ');
    attSavePhoto_(n.date + '|' + email + '|OUT', photo);
    attWrite_(sh, rec, rec._row);
    return { ok: true, message: 'Absen pulang tercatat pukul ' + n.time + ' WIB' + (rec.overtimeMinutes ? ' (lembur ' + rec.overtimeMinutes + ' menit)' : '') };
  } finally {
    lock.releaseLock();
  }
}

/** Foto bukti selfie diambil terpisah (hanya pemilik catatan atau Admin). */
function attPhoto_(p) {
  const u = sessionUser_(p.sessionToken);
  if (!u) return { ok: false, message: 'Sesi tidak valid atau sudah berakhir' };
  const date = String(p.date || ''), email = lower_(p.email);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return { ok: false, message: 'Tanggal tidak valid' };
  if (!bool_(u.isAdmin) && lower_(u.email) !== email) return { ok: false, message: 'Tidak berhak melihat foto ini' };
  return { ok: true, in: attPhotoOf_(date + '|' + email + '|IN'), out: attPhotoOf_(date + '|' + email + '|OUT') };
}

function attConfigSet_(p) {
  if (!requireAdmin_(p)) return { ok: false, message: 'Hanya Admin yang boleh mengubah pengaturan' };
  const lat = attNum_(p.officeLat), lng = attNum_(p.officeLng), radius = parseInt(p.officeRadius, 10);
  const tol = parseInt(p.lateToleranceMin, 10), ot = parseInt(p.overtimeMinThreshold, 10);
  if (lat === null || lng === null || Math.abs(lat) > 90 || Math.abs(lng) > 180) return { ok: false, message: 'Koordinat kantor tidak valid' };
  if (!(radius >= 10 && radius <= 5000)) return { ok: false, message: 'Radius 10-5000 meter' };
  if (!attValidTime_(p.workStartTime) || !attValidTime_(p.workEndTime)) return { ok: false, message: 'Jam kerja tidak valid (HH:MM)' };
  if (attMin_(p.workEndTime) <= attMin_(p.workStartTime)) return { ok: false, message: 'Jam pulang harus setelah jam masuk' };
  if (!(tol >= 0 && tol <= 120)) return { ok: false, message: 'Toleransi telat 0-120 menit' };
  if (!(ot >= 0 && ot <= 240)) return { ok: false, message: 'Minimal lembur 0-240 menit' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    writeConfig_({ att_officeLat: lat, att_officeLng: lng, att_officeRadius: radius, att_workStartTime: p.workStartTime,
      att_workEndTime: p.workEndTime, att_lateToleranceMin: tol, att_overtimeMinThreshold: ot });
  } finally {
    lock.releaseLock();
  }
  return { ok: true, message: 'Pengaturan jam kerja & lokasi kantor disimpan' };
}

/* ======================= UTILITAS EDITOR ======================= */

/** Jalankan sekali: memberi izin Sheets + Gmail dan membuat sheet. */
function authorize() {
  MailApp.getRemainingDailyQuota();
  sheet_();
  plainSheet_(CONFIG_SHEET, CONFIG_HEADERS);
  plainSheet_(REPORT_SHEET, REPORT_HEADERS);
  plainSheet_(REPORT_DATA_SHEET, DATA_HEADERS);
  plainSheet_(ATT_SHEET, ATT_HEADERS);
  plainSheet_(ATT_PHOTO_SHEET, ATT_PHOTO_HEADERS);
  Logger.log('Izin OK. Sheet "' + SHEET_NAME + '" siap.');
}

/** Jalankan sekali SETELAH akun ADMIN_EMAIL_TO_PROMOTE mendaftar lewat portal. */
function makeAdmin() {
  const email = lower_(ADMIN_EMAIL_TO_PROMOTE);
  const u = rows_().find(function (x) { return lower_(x.email) === email; });
  if (!u) throw new Error('Email ' + email + ' belum mendaftar. Daftar dulu lewat portal.');
  setCell_(u._row, 'isAdmin', true);
  setCell_(u._row, 'approved', true);
  setCell_(u._row, 'verified', true);
  setCell_(u._row, 'verifiedAt', now_());
  Logger.log(email + ' sekarang Admin & terverifikasi.');
}

/**
 * Buat (atau reset password) akun Admin langsung dari editor.
 * 1) Isi PASSWORD di bawah dengan password baru, 2) jalankan fungsi ini sekali,
 * 3) kembalikan PASSWORD ke teks semula agar tidak tersimpan di kode.
 */
function createAdmin() {
  const EMAIL = ADMIN_EMAIL_TO_PROMOTE;
  const NAME = 'Admin PT. Global Mega Indonesia';
  const ROLE = 'System Administrator';
  const PASSWORD = 'GANTI_DENGAN_PASSWORD_BARU';   // minimal 6 karakter

  if (PASSWORD === 'GANTI_DENGAN_PASSWORD_BARU' || PASSWORD.length < 6) {
    throw new Error('Isi konstanta PASSWORD dengan password baru (min. 6 karakter) lalu jalankan lagi.');
  }
  const email = lower_(EMAIL);
  const salt = Utilities.getUuid();
  const existing = rows_().find(function (x) { return lower_(x.email) === email; });

  if (existing) {
    setCell_(existing._row, 'passwordHash', hash_(salt, PASSWORD));
    setCell_(existing._row, 'salt', salt);
    setCell_(existing._row, 'isAdmin', true);
    setCell_(existing._row, 'approved', true);
    setCell_(existing._row, 'verified', true);
    setCell_(existing._row, 'verifiedAt', now_());
  } else {
    sheet_().appendRow(['usr_admin_' + Date.now(), NAME, ROLE, email, hash_(salt, PASSWORD), salt,
                        true, true, true, 'ACT-ADMIN', now_(), now_(), true, now_(), '', '']);
  }
  Logger.log('Admin ' + email + ' siap. Kembalikan PASSWORD di kode ke teks semula.');
}

/**
 * Pindahkan akun dari tab lama "Users" (kolom: ID, Nama, Jabatan, Email, Password, IsAdmin,
 * Verified, ActivationSent, ActivationToken, RegisteredAt, VerifiedAt) ke tab GMI_Users.
 * Password teks biasa otomatis di-hash. Email yang sudah ada dilewati. Aman dijalankan berulang.
 */
function migrateOldUsers() {
  const OLD_SHEET = 'Users';
  const old = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(OLD_SHEET);
  if (!old) throw new Error('Tab "' + OLD_SHEET + '" tidak ditemukan.');

  const v = old.getDataRange().getValues();
  const h = v[0].map(function (x) { return String(x).trim().toLowerCase(); });
  const col = function (name) { return h.indexOf(name); };
  if (col('email') === -1 || col('password') === -1) {
    throw new Error('Tab "' + OLD_SHEET + '" harus punya kolom Email dan Password.');
  }

  const existing = rows_().map(function (x) { return lower_(x.email); });
  let added = 0;
  v.slice(1).forEach(function (r) {
    const email = lower_(r[col('email')]);
    const pw = String(r[col('password')] || '');
    if (!email || !pw || existing.indexOf(email) !== -1) return;

    const get = function (name, def) { const i = col(name); return (i === -1 || r[i] === '') ? def : r[i]; };
    const salt = Utilities.getUuid();
    sheet_().appendRow([
      get('id', 'usr_' + Date.now() + '_' + added),
      get('nama', ''), get('jabatan', ''), email,
      hash_(salt, pw), salt,
      bool_(get('isadmin', false)), bool_(get('verified', false)), bool_(get('activationsent', false)),
      get('activationtoken', 'ACT-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase()),
      get('registeredat', now_()), get('verifiedat', ''),
      bool_(get('verified', false)) || bool_(get('activationsent', false)), '', '', ''
    ]);
    existing.push(email);
    added++;
  });
  Logger.log('Migrasi selesai: ' + added + ' akun dipindahkan ke ' + SHEET_NAME + '.');
}
