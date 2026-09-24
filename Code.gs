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

const API_VERSION = 5;
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
function hasAdmin_() {
  const props = PropertiesService.getScriptProperties();
  return !!(props.getProperty('admin_user') && props.getProperty('admin_pass_hash'));
}

function hashPass_(pass) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pass), Utilities.Charset.UTF_8);
  return raw.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function register_(data) {
  if (hasAdmin_()) return jsonResponse_({status:'error', message:'Akun admin sudah terdaftar. Silakan login.'});
  const username = String(data.username || '').trim();
  const password = String(data.password || '');
  const email = String(data.email || '').trim();
  if (username.length < 3) return jsonResponse_({status:'error', message:'Username minimal 3 karakter.'});
  if (password.length < 6) return jsonResponse_({status:'error', message:'Password minimal 6 karakter.'});
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return jsonResponse_({status:'error', message:'Email pemulihan tidak valid.'});
  writeAccount_(username, password, email);
  // Setelah daftar, langsung login otomatis (kirim token sesi).
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('sess_' + token, '1', SESSION_SECONDS);
  return jsonResponse_({status:'success', token: token, message:'Akun admin berhasil dibuat.'});
}

// Simpan/timpa akun admin ke Script Properties + sheet "Admin" (baris ke-2,
// selalu hanya 1 baris). Dipakai oleh pendaftaran pertama dan "Buat Akun Baru".
function writeAccount_(username, password, email) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty('admin_user', username);
  props.setProperty('admin_pass_hash', hashPass_(password));
  props.setProperty('admin_email', email);
  const waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'dd-MM-yyyy HH:mm');
  const aSheet = getAdminSheet_();
  aSheet.getRange(2, 1, 1, 4).setValues([[username, password, email, waktu]]);
}

// ================= Buat Akun Admin Baru (dipakai saat SUDAH login) =================
// Berbeda dari register_: ini boleh dipakai walau akun admin sudah ada,
// asal token sesi valid (dipanggil dari dalam halaman admin, bukan publik).
// Akun lama otomatis diganti (ditimpa) dengan akun baru ini.
function newAccount_(data) {
  if (!validToken_(data.token)) return jsonResponse_({status:'error', message:'Sesi berakhir. Silakan login lagi.', expired:true});
  const username = String(data.username || '').trim();
  const password = String(data.password || '');
  const email = String(data.email || '').trim();
  if (username.length < 3) return jsonResponse_({status:'error', message:'Username minimal 3 karakter.'});
  if (password.length < 6) return jsonResponse_({status:'error', message:'Password minimal 6 karakter.'});
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return jsonResponse_({status:'error', message:'Email pemulihan tidak valid.'});
  writeAccount_(username, password, email);
  return jsonResponse_({status:'success', message:'Akun admin baru berhasil dibuat. Username/password lama sudah tidak berlaku.'});
}

// ================= Lupa Password: kirim kode ke email pemulihan =================
function forgot_(data) {
  if (!hasAdmin_()) return jsonResponse_({status:'error', message:'Belum ada akun admin. Silakan daftar dulu.', needRegister:true});
  const props = PropertiesService.getScriptProperties();
  const storedEmail = props.getProperty('admin_email');
  const inputEmail = String(data.email || '').trim().toLowerCase();
  if (!storedEmail) return jsonResponse_({status:'error', message:'Akun ini belum punya email pemulihan (dibuat sebelum fitur ini ada). Reset manual lewat Apps Script diperlukan.'});
  if (!inputEmail || inputEmail !== storedEmail.toLowerCase()) {
    return jsonResponse_({status:'error', message:'Email tidak cocok dengan email pemulihan yang terdaftar.'});
  }
  const cache = CacheService.getScriptCache();
  const sentRecently = cache.get('reset_sent');
  if (sentRecently) return jsonResponse_({status:'error', message:'Kode baru saja dikirim. Cek email kamu, atau tunggu 1 menit untuk kirim ulang.'});
  const code = String(Math.floor(100000 + Math.random() * 900000)); // 6 digit
  cache.put('reset_code', code, 600);   // berlaku 10 menit
  cache.put('reset_sent', '1', 60);     // jeda kirim ulang 1 menit
  try {
    MailApp.sendEmail({
      to: storedEmail,
      subject: 'Kode Reset Password Admin - Website Desa Seburing',
      body: 'Username admin kamu: ' + props.getProperty('admin_user') + '\n\n' +
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
  const code = String(data.code || '').trim();
  const newPassword = String(data.newPassword || '');
  if (!savedCode) return jsonResponse_({status:'error', message:'Kode sudah kedaluwarsa. Minta kode baru.'});
  if (code !== savedCode) return jsonResponse_({status:'error', message:'Kode reset salah.'});
  if (newPassword.length < 6) return jsonResponse_({status:'error', message:'Password baru minimal 6 karakter.'});

  const props = PropertiesService.getScriptProperties();
  props.setProperty('admin_pass_hash', hashPass_(newPassword));
  cache.remove('reset_code');
  cache.remove('reset_sent');
  cache.remove('fails');

  // Perbarui juga kolom Password di sheet "Admin".
  const aSheet = getAdminSheet_();
  if (aSheet.getLastRow() >= 2) aSheet.getRange(2, 2).setValue(newPassword);

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
  if (!hasAdmin_()) return jsonResponse_({status:'error', message:'Belum ada akun admin. Silakan daftar dulu.', needRegister:true});
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAIL) return jsonResponse_({status:'error', message:'Terlalu banyak percobaan. Coba lagi 5 menit lagi.'});
  const props = PropertiesService.getScriptProperties();
  const okUser = String(data.username || '') === props.getProperty('admin_user');
  const okPass = hashPass_(data.password || '') === props.getProperty('admin_pass_hash');
  if (!okUser || !okPass) {
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
