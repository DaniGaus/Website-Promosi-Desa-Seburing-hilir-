/**
 * Database Website Promosi Desa Seburing
 * Google Apps Script backend: Nama, Email, Jenis (Pengaduan/Saran), Pesan.
 * Admin masuk lewat Username + Password (login), lalu server memberi token sesi.
 */

const SHEET_ID = '1E_bCv0o0IsthnqySqZ6cabVWOLMtL6VBGxritJ9XZH4';
const SHEET_NAME = 'Data';
const ADMIN_SHEET_NAME = 'Admin'; // sheet kedua: salinan username & password admin (agar bisa dilihat manual)

// Login untuk masuk website TETAP dicek lewat Script Properties (password
// di-hash, aman). Sheet "Admin" di bawah ini HANYA salinan agar kamu bisa
// melihat username/password yang kamu daftarkan; bukan sumber login.

// ====== GANTI kode ini sebelum deploy! Hanya kamu yang boleh tahu. ======
// Wajib dimasukkan saat mendaftar akun admin pertama kali, supaya orang lain
// yang kebetulan membuka admin.html DULUAN tidak bisa mencuri akun admin.
const SETUP_CODE = 'GANTI-KODE-RAHASIA-INI';
// ==========================================================================

const API_VERSION = 7;
const HEADERS = ['ID', 'Nama', 'Email', 'Pesan', 'Status', 'Waktu', 'Jenis'];
const JENIS_OK = ['Pengaduan', 'Saran', 'Lain-lain'];
const STATUS_OK = ['Baru', 'Diproses', 'Selesai'];
const SESSION_SECONDS = 60 * 60 * 2; // sesi login 2 jam
const MAX_FAIL = 5;                  // batas salah login
const LOCK_SECONDS = 60 * 5;         // dikunci 5 menit

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  } else if (sheet.getRange(1, 1).getValue() !== 'ID') {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  } else if (sheet.getRange(1, 7).getValue() !== 'Jenis') {
    sheet.getRange(1, 7).setValue('Jenis'); // migrasi sheet lama (6 kolom)
  }
  return sheet;
}

function getAdminSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(ADMIN_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(ADMIN_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Username', 'Password', 'Email Pemulihan', 'Dibuat']);
  }
  return sheet;
}

function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (data.action === 'status')   return status_();
    if (data.action === 'register') return register_(data);
    if (data.action === 'newAccount') return newAccount_(data);
    if (data.action === 'forgot')   return forgot_(data);
    if (data.action === 'reset')    return resetPassword_(data);
    if (data.action === 'login')    return login_(data);
    if (data.action === 'list')     return list_(data);
    if (data.action === 'update')   return updateStatus_(data);
    if (data.action === 'logout')   { CacheService.getScriptCache().remove('sess_' + data.token); return jsonResponse_({status:'success'}); }
    if (data.action) return jsonResponse_({status:'error', message:'Aksi tidak dikenal: ' + data.action});
    return submit_(data); // hanya kiriman formulir warga (tanpa action)
  } catch (err) {
    return jsonResponse_({status:'error', message: err.message});
  }
}

// Semua fitur (status, daftar, login, kirim formulir, lihat data, ubah status,
// logout) lewat SATU pintu masuk ini (doPost), dibedakan lewat "action".
// doGet tetap ada hanya sebagai cadangan kalau link dibuka langsung di browser.
function status_() {
  return jsonResponse_({status:'ok', version: API_VERSION, adminExists: hasAdmin_(), message:'API pengaduan & saran aktif.'});
}

function doGet() {
  return status_();
}

// ================= Akun admin (Script Properties, password di-hash) =================
// Mendukung BANYAK akun admin. Disimpan sebagai JSON di Script Properties
// 'admin_accounts': [{u: username, h: hash, e: email}, ...]
// Akun lama (1 akun, versi sebelumnya) otomatis dimigrasi saat pertama dibaca.
function getAccounts_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('admin_accounts');
  if (raw) {
    try { const arr = JSON.parse(raw); if (Array.isArray(arr)) return arr; } catch (err) {}
  }
  const u = props.getProperty('admin_user'), h = props.getProperty('admin_pass_hash');
  if (u && h) {
    const arr = [{u: u, h: h, e: props.getProperty('admin_email') || ''}];
    props.setProperty('admin_accounts', JSON.stringify(arr));
    return arr;
  }
  return [];
}

function saveAccounts_(arr) {
  PropertiesService.getScriptProperties().setProperty('admin_accounts', JSON.stringify(arr));
}

function findAccountIndex_(accounts, username) {
  const key = String(username || '').trim().toLowerCase();
  for (let i = 0; i < accounts.length; i++) {
    if (String(accounts[i].u).toLowerCase() === key) return i;
  }
  return -1;
}

function hasAdmin_() {
  return getAccounts_().length > 0;
}

function hashPass_(pass) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pass), Utilities.Charset.UTF_8);
  return raw.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

// Validasi isian akun; mengembalikan pesan error atau '' kalau valid.
function validateAccount_(username, password, email) {
  if (username.length < 3) return 'Username minimal 3 karakter.';
  if (password.length < 6) return 'Password minimal 6 karakter.';
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return 'Email pemulihan tidak valid.'; // email opsional
  return '';
}

// Daftar akun dari halaman login (publik): cukup username + password.
// Email pemulihan opsional (tanpa email, fitur lupa password tidak tersedia untuk akun itu).
function register_(data) {
  const username = String(data.username || '').trim();
  const password = String(data.password || '');
  const email = String(data.email || '').trim();
  const err = validateAccount_(username, password, email);
  if (err) return jsonResponse_({status:'error', message: err});
  if (!addAccount_(username, password, email)) return jsonResponse_({status:'error', message:'Username sudah dipakai. Pilih username lain atau langsung login.'});
  // Sengaja TIDAK auto-login: admin diarahkan ke Login untuk masuk sendiri.
  return jsonResponse_({status:'success', message:'Akun admin berhasil dibuat. Silakan login.'});
}

// Tambah akun baru (false kalau username sudah ada). Juga dicatat di sheet "Admin".
function addAccount_(username, password, email) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const accounts = getAccounts_();
    if (findAccountIndex_(accounts, username) >= 0) return false;
    accounts.push({u: username, h: hashPass_(password), e: email});
    saveAccounts_(accounts);
    const waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'dd-MM-yyyy HH:mm');
    getAdminSheet_().appendRow([username, password, email, waktu]);
    return true;
  } finally {
    lock.releaseLock();
  }
}

// ================= Tambah Akun Admin (dipakai saat SUDAH login) =================
// Dipanggil dari dalam halaman admin; butuh token sesi valid. Menambah akun,
// tidak menghapus akun yang sudah ada.
function newAccount_(data) {
  if (!validToken_(data.token)) return jsonResponse_({status:'error', message:'Sesi berakhir. Silakan login lagi.', expired:true});
  const username = String(data.username || '').trim();
  const password = String(data.password || '');
  const email = String(data.email || '').trim();
  const err = validateAccount_(username, password, email);
  if (err) return jsonResponse_({status:'error', message: err});
  if (!addAccount_(username, password, email)) return jsonResponse_({status:'error', message:'Username sudah dipakai.'});
  return jsonResponse_({status:'success', message:'Akun admin baru berhasil ditambahkan.'});
}

// ================= Lupa Password: kirim kode ke email pemulihan =================
function forgot_(data) {
  const accounts = getAccounts_();
  if (!accounts.length) return jsonResponse_({status:'error', message:'Belum ada akun admin. Silakan buat akun dulu.', needRegister:true});
  const inputEmail = String(data.email || '').trim().toLowerCase();
  const acc = accounts.filter(x => x.e && String(x.e).toLowerCase() === inputEmail)[0];
  if (!inputEmail || !acc) {
    return jsonResponse_({status:'error', message:'Email tidak cocok dengan email pemulihan yang terdaftar.'});
  }
  const cache = CacheService.getScriptCache();
  const sentRecently = cache.get('reset_sent');
  if (sentRecently) return jsonResponse_({status:'error', message:'Kode baru saja dikirim. Cek email kamu, atau tunggu 1 menit untuk kirim ulang.'});
  const code = String(Math.floor(100000 + Math.random() * 900000)); // 6 digit
  cache.put('reset_code', code, 600);   // berlaku 10 menit
  cache.put('reset_user', acc.u, 600);  // akun yang direset
  cache.put('reset_sent', '1', 60);     // jeda kirim ulang 1 menit
  try {
    MailApp.sendEmail({
      to: acc.e,
      subject: 'Kode Reset Password Admin - Website Desa Seburing',
      body: 'Username admin kamu: ' + acc.u + '\n\n' +
            'Kode reset password (berlaku 10 menit): ' + code + '\n\n' +
            'Kalau kamu tidak meminta ini, abaikan email ini.'
    });
  } catch (err) {
    return jsonResponse_({status:'error', message:'Gagal mengirim email: ' + err.message});
  }
  return jsonResponse_({status:'success', message:'Kode reset sudah dikirim ke email pemulihan.'});
}

function resetPassword_(data) {
  const cache = CacheService.getScriptCache();
  const savedCode = cache.get('reset_code');
  const resetUser = cache.get('reset_user');
  const code = String(data.code || '').trim();
  const newPassword = String(data.newPassword || '');
  if (!savedCode || !resetUser) return jsonResponse_({status:'error', message:'Kode sudah kedaluwarsa. Minta kode baru.'});
  if (code !== savedCode) return jsonResponse_({status:'error', message:'Kode reset salah.'});
  if (newPassword.length < 6) return jsonResponse_({status:'error', message:'Password baru minimal 6 karakter.'});

  const accounts = getAccounts_();
  const idx = findAccountIndex_(accounts, resetUser);
  if (idx < 0) return jsonResponse_({status:'error', message:'Akun tidak ditemukan.'});
  accounts[idx].h = hashPass_(newPassword);
  saveAccounts_(accounts);
  cache.remove('reset_code');
  cache.remove('reset_user');
  cache.remove('reset_sent');
  cache.remove('fails');

  // Perbarui juga kolom Password di sheet "Admin" (cari baris username-nya).
  const aSheet = getAdminSheet_();
  const rows = aSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toLowerCase() === String(resetUser).toLowerCase()) { aSheet.getRange(i + 1, 2).setValue(newPassword); break; }
  }

  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('sess_' + token, '1', SESSION_SECONDS);
  return jsonResponse_({status:'success', token: token, message:'Password berhasil diganti.'});
}

function submit_(data) {
  if (!validToken_(data.token)) return jsonResponse_({status:'error', message:'Silakan login terlebih dahulu.', expired:true});
  const nama = String(data.nama || '').trim();
  const email = String(data.email || '').trim();
  const pesan = String(data.pesan || '').trim();
  const jenis = JENIS_OK.indexOf(data.jenis) >= 0 ? data.jenis : 'Pengaduan';
  if (!nama || !email || !pesan) return jsonResponse_({status:'error', message:'Nama, email, dan pesan wajib diisi.'});
  if (nama.length > 100 || email.length > 150 || pesan.length > 2000) return jsonResponse_({status:'error', message:'Data terlalu panjang.'});
  const waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'dd-MM-yyyy HH:mm');
  getSheet_().appendRow([Utilities.getUuid(), nama, email, pesan, 'Baru', waktu, jenis]);
  return jsonResponse_({status:'success', message: jenis + ' berhasil disimpan.'});
}

function login_(data) {
  const accounts = getAccounts_();
  if (!accounts.length) return jsonResponse_({status:'error', message:'Belum ada akun admin. Silakan buat akun dulu.', needRegister:true});
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAIL) return jsonResponse_({status:'error', message:'Terlalu banyak percobaan. Coba lagi 5 menit lagi.'});
  const idx = findAccountIndex_(accounts, data.username);
  const ok = idx >= 0 && hashPass_(data.password || '') === accounts[idx].h;
  if (!ok) {
    cache.put('fails', String(fails + 1), LOCK_SECONDS);
    return jsonResponse_({status:'error', message:'Username atau password salah.'});
  }
  cache.remove('fails');
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('sess_' + token, '1', SESSION_SECONDS);
  return jsonResponse_({status:'success', token: token});
}

function validToken_(token) {
  return !!token && CacheService.getScriptCache().get('sess_' + token) === '1';
}

function list_(data) {
  if (!validToken_(data.token)) return jsonResponse_({status:'error', message:'Sesi berakhir. Silakan login lagi.', expired:true});
  const values = getSheet_().getDataRange().getValues();
  const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
  const fmt = v => (v instanceof Date) ? Utilities.formatDate(v, tz, 'dd-MM-yyyy HH:mm') : String(v || '');
  const rows = values.slice(1).filter(r => r[0]).map(r => ({
    id: String(r[0]), nama: String(r[1] || ''), email: String(r[2] || ''), pesan: String(r[3] || ''),
    status: r[4] || 'Baru', waktu: fmt(r[5]), jenis: r[6] || 'Pengaduan'
  })).reverse(); // terbaru di atas
  return jsonResponse_({
    status:'success',
    rows: rows,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit',
    adminSheetUrl: 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit#gid=' + getAdminSheet_().getSheetId()
  });
}

function updateStatus_(data) {
  if (!validToken_(data.token)) return jsonResponse_({status:'error', message:'Sesi berakhir. Silakan login lagi.', expired:true});
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const rowIndex = values.findIndex((r, i) => i > 0 && String(r[0]) === String(data.id));
  if (rowIndex < 1) return jsonResponse_({status:'error', message:'Data tidak ditemukan.'});
  const status = STATUS_OK.indexOf(data.status) >= 0 ? data.status : 'Baru';
  sheet.getRange(rowIndex + 1, 5).setValue(status);
  return jsonResponse_({status:'success', message:'Status diperbarui.'});
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
