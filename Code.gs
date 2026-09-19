/**
 * Database Website Promosi Desa Seburing
 * Backend Google Apps Script — menyimpan & mengambil data Nama, Email, Pesan
 * dari/ke Google Sheet bernama "Database Website Promosi Desa Seburing".
 *
 * CARA PAKAI (bisa full dari HP):
 * 1. Buat Google Sheet baru, beri nama: Database Website Promosi Desa Seburing
 * 2. Buka https://script.google.com di browser HP, login akun Google yang sama.
 * 3. New Project -> hapus kode default -> tempel (paste) semua kode di file ini.
 * 4. Ganti SHEET_ID di bawah dengan ID Sheet kamu
 *    (ID = bagian di URL sheet antara /d/ dan /edit, contoh:
 *     https://docs.google.com/spreadsheets/d/INI_ID_NYA/edit)
 * 5. Klik Deploy -> New deployment -> pilih tipe "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Klik Deploy, copy URL Web App yang muncul (diakhiri /exec).
 * 7. Tempel URL itu ke variabel SCRIPT_URL di file database.html pada website.
 */

const SHEET_ID = '1E_bCv0o0IsthnqySqZ6cabVWOLMtL6VBGxritJ9XZH4';
const SHEET_NAME = 'Data';

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Nama', 'Email', 'Pesan', 'Waktu']);
  }
  return sheet;
}

// Menyimpan data baru (dipanggil oleh form di database.html lewat fetch POST)
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const nama = (data.nama || '').toString().trim();
    const email = (data.email || '').toString().trim();
    const pesan = (data.pesan || '').toString().trim();

    if (!nama || !email || !pesan) {
      return jsonResponse_({ status: 'error', message: 'Nama, Email, dan Pesan wajib diisi' });
    }

    const sheet = getSheet_();
    const waktu = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'dd-MM-yyyy HH:mm');
    sheet.appendRow([nama, email, pesan, waktu]);

    return jsonResponse_({ status: 'success', message: 'Data tersimpan' });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: err.message });
  }
}

// Mengambil semua data (dipanggil oleh database.html lewat fetch GET)
function doGet(e) {
  try {
    const sheet = getSheet_();
    const values = sheet.getDataRange().getValues();
    const rows = values.slice(1).map(function (r) {
      return { nama: r[0], email: r[1], pesan: r[2], waktu: r[3] };
    });
    return jsonResponse_(rows);
  } catch (err) {
    return jsonResponse_({ status: 'error', message: err.message });
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
