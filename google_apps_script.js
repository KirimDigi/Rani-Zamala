/**
 * ============================================================================
 * GOOGLE APPS SCRIPT - RSVP & BUKU UCAPAN (RANI ZAMALA)
 * ============================================================================
 * 
 * SPREADSHEET TARGET:
 * ID   : 1nEeclKaTKND14QF3-DFdhl0qSjDZrDlMx7n91mgNO2k
 * SHEET: Sheet1
 * 
 * KOLOM YANG DIGUNAKAN:
 * 1. timestamp
 * 2. nama tamu
 * 3. ucapan
 * 4. konfirmasi kehadiran
 * 5. jumlah tamu
 * 
 * CARA MEMASANG / DEPLOY:
 * 1. Buka spreadsheet Anda: https://docs.google.com/spreadsheets/d/1nEeclKaTKND14QF3-DFdhl0qSjDZrDlMx7n91mgNO2k/edit?gid=0#gid=0
 * 2. Klik menu "Ekstensi" (Extensions) -> "Apps Script".
 * 3. Hapus semua kode default di dalam editor (function myFunction() { ... }).
 * 4. Copy & Paste seluruh kode script di bawah ini.
 * 5. Klik icon Save (Disket / Ctrl+S).
 * 6. Klik tombol biru "Terapkan" (Deploy) di pojok kanan atas -> pilih "Penerapan baru" (New deployment).
 * 7. Pada icon Gerigi (Select type), pilih "Aplikasi Web" (Web app).
 * 8. Isi konfigurasi:
 *    - Deskripsi       : RSVP Rani Zamala
 *    - Jalankan sebagai: Saya (Email Anda)
 *    - Siapa yang memiliki akses (Who has access): SIAPA SAJA (Anyone)  <--- PENTING!
 * 9. Klik "Terapkan" (Deploy) -> Izinkan Akses (Authorize Access) jika diminta.
 * 10. Copy "URL Aplikasi Web" (akhiran /exec).
 * 11. Masukkan URL tersebut pada file index.html di variabel window.GOOGLE_APPS_SCRIPT_URL.
 * ============================================================================
 */

var SPREADSHEET_ID = "1nEeclKaTKND14QF3-DFdhl0qSjDZrDlMx7n91mgNO2k";
var SHEET_NAME = "Sheet1";

/**
 * Handle GET Request: Mengambil daftar ucapan untuk ditampilkan di website
 */
function doGet(e) {
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      sheet = doc.getSheets()[0];
    }
    
    var lastRow = sheet.getLastRow();
    
    // Jika belum ada data sama sekali selain header (atau sheet kosong)
    if (lastRow <= 1) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        total: 0,
        data: []
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // Ambil data dari baris ke-2 (melewati baris header)
    var values = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var list = [];
    
    // Urutkan dari data terbaru ke data terlama (reverse chronological)
    for (var i = values.length - 1; i >= 0; i--) {
      var row = values[i];
      var namaTamu = String(row[1] || '').trim();
      var ucapan = String(row[2] || '').trim();
      
      // Lewati jika baris benar-benar kosong
      if (!namaTamu && !ucapan) continue;
      
      var ts = row[0];
      if (ts instanceof Date) {
        ts = Utilities.formatDate(ts, "Asia/Jakarta", "dd/MM/yyyy HH:mm");
      } else {
        ts = String(ts || '');
      }
      
      list.push({
        timestamp: ts,
        nama: namaTamu,
        ucapan: ucapan,
        kehadiran: String(row[3] || '').trim(),
        jumlah: String(row[4] || '').trim()
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      total: list.length,
      data: list
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString(),
      data: []
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle POST Request: Menyimpan data RSVP & Ucapan baru dari formulir website
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  // Cegah bentrok pengiriman bersamaan dengan kunci selama 10 detik
  lock.tryLock(10000);
  
  try {
    var doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = doc.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      sheet = doc.getSheets()[0];
    }
    
    // Auto-generate header jika spreadsheet masih kosong
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "timestamp", 
        "nama tamu", 
        "ucapan", 
        "konfirmasi kehadiran", 
        "jumlah tamu"
      ]);
      // Format header agar tebal
      sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
    }
    
    // Parsing data dari body request
    var data = {};
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    }
    
    // Siapkan nilai tiap kolom
    var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
    var namaTamu = data.nama || data.nama_tamu || data['nama tamu'] || data.guestname || '';
    var ucapan = data.ucapan || data.messagestext || '';
    var konfirmasi = data.kehadiran || data.konfirmasi || data['konfirmasi kehadiran'] || data.confirmattend || '';
    var jumlahTamu = data.jumlah || data.jumlah_tamu || data['jumlah tamu'] || data.countpeople || '-';
    
    // Simpan baris baru ke Sheet1
    sheet.appendRow([
      timestamp, 
      namaTamu, 
      ucapan, 
      konfirmasi, 
      jumlahTamu
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Ucapan & RSVP berhasil disimpan!"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
