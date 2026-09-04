// GOOGLE APPS SCRIPT FOR RSVP & SPREADSHEET INTEGRATION
// Spreadsheet ID: 14h_BILBq0dZ6TWsGj9Csat0zz_lQ9xUtutenXE2TDtc
// Sheet Name: Sheet1

const SPREADSHEET_ID = "14h_BILBq0dZ6TWsGj9Csat0zz_lQ9xUtutenXE2TDtc";
const SHEET_NAME = "Sheet1";

function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const record = {};
      headers.forEach((header, index) => {
        let value = row[index];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        }
        record[header.toString().toLowerCase().replace(/\s+/g, '_')] = value;
      });
      rows.push(record);
    }
    
    // Return JSON response allowing CORS
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: rows }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    let params;
    if (e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else {
      params = e.parameter;
    }
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    
    // Check/create headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["timestamp", "nama tamu", "ucapan", "konfirmasi kehadiran", "jumlah tamu"]);
    }
    
    const timestamp = new Date();
    const namaTamu = params.nama_tamu || params.author || "";
    const ucapan = params.ucapan || params.comment || "";
    const konfirmasi = params.konfirmasi || "";
    const jumlahTamu = params.jumlah_tamu || "0";
    
    sheet.appendRow([timestamp, namaTamu, ucapan, konfirmasi, jumlahTamu]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
