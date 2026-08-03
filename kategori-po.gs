// ==========================================
// API KATEGORI PO - Google Apps Script
// ==========================================
// 1. Buat project Apps Script baru di spreadsheet Kategori PO.
// 2. Paste kode ini ke Code.gs.
// 3. Ganti SPREADSHEET_ID dengan ID spreadsheet Kategori PO kamu.
//    (Bisa dilihat di URL spreadsheet: https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID_KAMU>/edit)
// 4. Ganti SHEET_NAME jika nama sheetnya bukan "Sheet1".
// 5. Deploy as Web App -> Access: Anyone.

const SPREADSHEET_ID = "GANTI_DENGAN_SPREADSHEET_ID_KAMU"; 
const SHEET_NAME = "Sheet1"; // Ganti jika nama sheetnya beda

function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("Sheet tidak ditemukan!");

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return output.setContent(JSON.stringify({ success: true, data: [] }));
    }

    const headers = data[0].map(h => String(h).trim());
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const obj = { _rowIndex: i + 1 }; // Simpan nomor baris asli untuk update
      headers.forEach((h, j) => {
        obj[h] = formatCell(data[i][j]);
      });
      // Hanya ambil baris yang ada Kategori PO-nya (tidak kosong)
      if (obj["KATEGORI PO"]) {
        rows.push(obj);
      }
    }

    output.setContent(JSON.stringify({ success: true, data: rows }));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, error: err.message }));
  }
  return output;
}

function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    const payload = JSON.parse(e.postData.contents);
    
    // Harus ada rowIndex (nomor baris) dan fields (kolom yang mau diubah)
    if (!payload.rowIndex || !payload.fields) {
      throw new Error("Missing rowIndex or fields");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    // Ambil header untuk mengetahui index kolom
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    
    // Update setiap field yang dikirim
    for (const key in payload.fields) {
      const colIndex = headers.indexOf(key);
      if (colIndex !== -1) {
        // Update sel (baris ke-rowIndex, kolom ke-(colIndex+1))
        sheet.getRange(payload.rowIndex, colIndex + 1).setValue(payload.fields[key]);
      }
    }
    
    output.setContent(JSON.stringify({ success: true, message: "Berhasil diupdate" }));
  } catch (err) {
    output.setContent(JSON.stringify({ success: false, error: err.message }));
  }
  return output;
}

// Untuk menangani CORS Preflight request dari browser
function doOptions(e) {
  const output = ContentService.createTextOutput("");
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function formatCell(val) {
  if (val instanceof Date && !isNaN(val.getTime())) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, "0");
    const d = String(val.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return String(val ?? "").trim();
}
