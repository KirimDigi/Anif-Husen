// ============================================
// Google Apps Script - Buku Tamu Anif & Husen
// Sheet: https://docs.google.com/spreadsheets/d/1SJ5poQuHYQXH9chy7XZzlqoHUugxxP4SR9CgGK8QE0o/edit
// Tab: Sheet1 | Kolom: timestamp | nama_tamu | ucapan | konfirmasi_kehadiran | jumlah_tamu
// Deploy: Publish > Deploy as web app > Execute as: Me > Who has access: Anyone
// Copy Web App URL dan paste ke index.html const APPS_SCRIPT_URL
// ============================================
const SPREADSHEET_ID = "1SJ5poQuHYQXH9chy7XZzlqoHUugxxP4SR9CgGK8QE0o";
const SHEET_NAME = "Sheet1";

function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // Header jika sheet kosong
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["timestamp", "nama_tamu", "ucapan", "konfirmasi_kehadiran", "jumlah_tamu"]);
    sheet.getRange(1,1,1,5).setFontWeight("bold").setBackground("#d16b8b").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// GET : ?action=read  atau langsung tanpa param => return semua ucapan
function doGet(e) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return jsonResponse({status: "success", data: []});
    }
    const values = sheet.getRange(2, 1, lastRow - 1, 5).getValues(); // skip header
    const data = values
      .filter(r => r[1] || r[2]) // minimal ada nama atau ucapan
      .map(r => ({
        timestamp: r[0] ? Utilities.formatDate(new Date(r[0]), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss") : "",
        nama_tamu: String(r[1] || "").trim(),
        ucapan: String(r[2] || "").trim(),
        konfirmasi_kehadiran: String(r[3] || "").trim(),
        konfirmasi: String(r[3] || "").trim(), // alias untuk kompatibilitas frontend lama
        jumlah_tamu: String(r[4] || "0").trim()
      }))
      .reverse(); // terbaru di atas

    return jsonResponse({status: "success", data: data});
  } catch (err) {
    return jsonResponse({status: "error", message: err.toString()});
  }
}

// POST : body JSON { nama_tamu, ucapan, konfirmasi, jumlah_tamu }
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.tryLock(10000);
    const sheet = getSheet();

    let payload = {};
    if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        // fallback form-urlencoded
        payload = e.parameter || {};
      }
    } else {
      payload = e.parameter || {};
    }

    // Support beberapa variasi key dari frontend
    const nama_tamu = (payload.nama_tamu || payload.nama || payload.author || "").toString().trim();
    const ucapan = (payload.ucapan || payload.comment || payload.pesan || "").toString().trim();
    const konfirmasi = (payload.konfirmasi || payload.konfirmasi_kehadiran || payload.confirm || "").toString().trim();
    let jumlah_tamu = (payload.jumlah_tamu || payload.jumlah || payload.guests || "0").toString().trim();

    // Validasi wajib
    if (!nama_tamu || !ucapan || !konfirmasi) {
      return jsonResponse({status: "error", message: "Nama, ucapan, dan konfirmasi wajib diisi."});
    }
    if (nama_tamu.length < 2 || ucapan.length < 2) {
      return jsonResponse({status: "error", message: "Nama/ucapan minimal 2 karakter."});
    }

    // Normalisasi jumlah tamu: jika Tidak hadir => 0, jika kosong => 0
    if (konfirmasi === "Tidak hadir" || konfirmasi === "Absen") {
      jumlah_tamu = "0";
    } else if (!jumlah_tamu || jumlah_tamu === "") {
      jumlah_tamu = "1";
    }

    const timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    sheet.appendRow([timestamp, nama_tamu, ucapan, konfirmasi, String(jumlah_tamu)]);

    return jsonResponse({status: "success", message: "Terima kasih! Ucapan Anda berhasil terkirim.", timestamp: timestamp});
  } catch (err) {
    return jsonResponse({status: "error", message: err.toString()});
  } finally {
    lock.releaseLock();
  }
}

// Handle CORS preflight (optional, untuk fetch)
function doOptions(e) {
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Test manual di Apps Script editor
function testGet() {
  Logger.log(doGet({}).getContent());
}
function testPost() {
  const mock = {
    postData: {
      contents: JSON.stringify({
        nama_tamu: "Tamu Tes",
        ucapan: "Selamat untuk Anif & Husen! SAMAWA!",
        konfirmasi: "Hadir",
        jumlah_tamu: "2"
      })
    }
  };
  Logger.log(doPost(mock).getContent());
}
