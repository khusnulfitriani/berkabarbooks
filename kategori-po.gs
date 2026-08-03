// ==========================================
// API KATEGORI PO - Google Apps Script
// ==========================================
// 1. Buat project Apps Script baru di spreadsheet Kategori PO.
// 2. Paste kode ini ke Code.gs.
// 3. Ganti SPREADSHEET_ID dengan ID spreadsheet Kategori PO kamu.
//    (Bisa dilihat di URL spreadsheet: https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID_KAMU>/edit)
// 4. Ganti SHEET_NAME jika nama sheetnya bukan "Sheet1".
// 5. Deploy as Web App -> Access: Anyone.

const SPREADSHEET_ID = "1wIk-qkQsuyPslcyTolo3gikRfuqs-W8Liyhr_cbKs8Y"; 
const SHEET_NAME = "ACTIVE"; // Ganti jika nama sheetnya beda

function doGet(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    // 1. Cek apakah ada data di Cache
    const cache = CacheService.getScriptCache();
    const cachedData = cache.get("kategori_po_data");
    
    if (cachedData) {
      // Jika ada, langsung kembalikan data dari Cache (SUPER CEPAT)
      return output.setContent(cachedData);
    }

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

    const jsonString = JSON.stringify({ success: true, data: rows });
    
    // 2. Simpan hasil ke Cache selama 1 Jam (3600 detik)
    // Jika data terlalu besar (batas cache GAS adalah 100KB), try-catch akan mengabaikannya
    try {
      cache.put("kategori_po_data", jsonString, 3600);
    } catch(e) {
      // Abaikan jika data terlalu besar untuk di-cache
    }

    output.setContent(jsonString);
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
    
    if (!payload.rowIndex || !payload.fields) {
      throw new Error("Missing rowIndex or fields");
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    
    for (const key in payload.fields) {
      const colIndex = headers.indexOf(key);
      if (colIndex !== -1) {
        sheet.getRange(payload.rowIndex, colIndex + 1).setValue(payload.fields[key]);
      }
    }
    
    // 3. Hapus Cache karena data baru saja diubah, agar halaman menampilkan versi terbaru!
    const cache = CacheService.getScriptCache();
    cache.remove("kategori_po_data");

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
