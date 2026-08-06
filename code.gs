// =============================================
// TRACKING ORDER - Google Apps Script Backend
// =============================================
// Cara deploy:
// 1. Buka spreadsheet → Extensions → Apps Script
// 2. Hapus semua kode yang ada, paste kode ini
// 3. Klik Save (ikon disket)
// 4. Klik Deploy → Manage deployments → Edit → New version → Deploy
// 5. URL tidak berubah, tidak perlu update index.html

const SHEET_ORDERS   = "Sheet1";      // Sheet pesanan aktif
const SHEET_SENT     = "Terkirim";    // Sheet pesanan terkirim
const SHEET_CUSTOMER = "Sheet2";      // Sheet data customer
const COL_ID         = "Customer ID"; // Nama kolom customer id
const CACHE_SECONDS  = 600;           // Cache 10 menit (600 detik)

function doGet(e) {
  const id = (e.parameter.id || "").trim().toLowerCase();

  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  if (!id) {
    output.setContent(JSON.stringify({ error: "id required" }));
    return output;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Ambil data sheet dari cache atau baca langsung
    const dataOrders   = getSheetData(ss, SHEET_ORDERS);
    const dataSent     = getSheetData(ss, SHEET_SENT);
    const dataCustomer = getSheetData(ss, SHEET_CUSTOMER);

    const orders   = filterById(dataOrders,   id);
    const sent     = filterById(dataSent,     id);
    const customer = findCustomer(dataCustomer, id);

    // Hitung summary gabungan
    const allRows = [...orders, ...sent];
    const totalBelanja = allRows.reduce((s, r) => s + toNum(r["Harga Total"]), 0);
    const totalDeposit = allRows.reduce((s, r) => s + toNum(r["DP"]), 0);
    const totalQty     = allRows.reduce((s, r) => s + toNum(r["Qty"]), 0);

    const result = {
      id,
      nama: customer ? customer["Nama"] : (orders[0]?.Nama || sent[0]?.Nama || ""),
      totalBelanja,
      totalDeposit,
      sisaTanggungan: Math.max(totalBelanja - totalDeposit, 0),
      totalQty,
      orders,
      sent,
    };

    output.setContent(JSON.stringify(result));
  } catch (err) {
    output.setContent(JSON.stringify({ error: err.message }));
  }

  return output;
}

// ─── CACHE LAYER ──────────────────────────────────────────────
// Membaca sheet dari cache; kalau belum ada / expired → baca spreadsheet & simpan ke cache
function getSheetData(ss, sheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "sheet_" + sheetName;

  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached); // Data sudah ada di cache → langsung pakai
  }

  // Belum ada di cache → baca dari spreadsheet
  const parsed = readSheet(ss, sheetName);

  // Simpan ke cache (max 100KB per entry; cukup untuk ratusan baris)
  try {
    cache.put(cacheKey, JSON.stringify(parsed), CACHE_SECONDS);
  } catch (e) {
    // Kalau data terlalu besar untuk cache, lanjut tanpa cache
  }

  return parsed;
}

// Baca semua baris dari sheet, hasilkan array of {header: value}
function readSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    headers.forEach((h, j) => { obj[h] = formatCell(data[i][j]); });
    rows.push(obj);
  }
  return rows;
}

// ─── FILTER ───────────────────────────────────────────────────
function filterById(rows, id) {
  return rows.filter(r => String(r[COL_ID] || "").trim().toLowerCase() === id);
}

function findCustomer(rows, id) {
  return rows.find(r => String(r[COL_ID] || "").trim().toLowerCase() === id) || null;
}

// ─── FORMAT ───────────────────────────────────────────────────
// Format nilai sel: Date → "YYYY-MM-DD", lainnya → string
function formatCell(val) {
  if (val instanceof Date && !isNaN(val.getTime())) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, "0");
    const d = String(val.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return String(val ?? "").trim();
}

// Konversi nilai uang / angka ke number
function toNum(v) {
  const c = String(v || "").replace(/Rp/gi, "").replace(/[^\d,-]/g, "").replace(/\./g, "").replace(",", ".");
  const n = parseFloat(c);
  return isFinite(n) ? n : 0;
}

// =============================================
// SYNC KE CLOUDFLARE D1
//
// ADA 2 FUNGSI:
//   syncOrdersAndCustomers() → pasang trigger SETIAP 30 MENIT
//     (orders ~5860 + customers ~2778 = ~9 chunks → ~30-60 detik)
//
//   syncSentFull() → jalankan MANUAL atau trigger MINGGUAN
//     (sent ~47949 = ~48 chunks → ~3-5 menit, aman untuk 6 menit limit)
// =============================================
const WORKER_SYNC_URL = "https://berkabarbooks-api.khusnulfitrianinund.workers.dev/sync";
const WORKER_SECRET   = "enye0201132705";
const CHUNK_SIZE      = 1000; // 1000 baris per request (jauh lebih sedikit HTTP calls)

// ─── Sync Orders + Customers + Resi (jalankan tiap 30 menit) ──
function syncOrdersAndCustomers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("📖 Membaca orders & customers...");
  const orders    = readSheet(ss, SHEET_ORDERS);
  const customers = readSheet(ss, SHEET_CUSTOMER);
  Logger.log("✔ orders: " + orders.length + ", customers: " + customers.length);

  const payload  = {
    secret:         WORKER_SECRET,
    clearOrders:    true,
    clearCustomers: true,
    clearSent:      false,
    orders:         orders,
    sent:           [],
    customers:      customers,
  };

  Logger.log("📤 Mengirim " + orders.length + " orders dan " + customers.length + " customers...");
  if (!_doSyncRequest(payload, "orders/cust")) return;

  Logger.log("✅ Sync orders & customers selesai!");

  // Sekalian sync resi (data kecil, cepat)
  syncResi();

  // Sekalian sync baris-baris BARU di sheet Terkirim (incremental, cepat)
  syncSentIncremental();
}

// ─── Sync Resi dari "5. LIST PENGIRIMAN" ───────────────────────
const SHEET_RESI = "5. LIST PENGIRIMAN";

function syncResi() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_RESI);
  if (!sheet) {
    Logger.log("⚠️ Sheet resi tidak ditemukan: " + SHEET_RESI);
    return;
  }
  const rows = readSheet(ss, SHEET_RESI);

  if (!rows.length) {
    Logger.log("⚠️ Sheet resi kosong: " + SHEET_RESI);
    return;
  }

  // Ambil nama kolom pertama (kolom A)
  const dataRange = sheet.getDataRange().getValues();
  const headers = dataRange[0].map(h => String(h).trim());
  const colAHeader = headers[0];

  // Tambahkan property "no_antrian" ke setiap row
  rows.forEach(r => {
    r["no_antrian"] = r[colAHeader] || "";
  });

  Logger.log("📖 Membaca resi: " + rows.length + " baris → " + Math.ceil(rows.length / CHUNK_SIZE) + " batch");

  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    const batchNum = Math.floor(i / CHUNK_SIZE) + 1;
    const payload  = {
      secret:    WORKER_SECRET,
      clearResi: (i === 0),   // hapus resi hanya di batch pertama
      orders:    [],
      sent:      [],
      customers: [],
      resi:      rows.slice(i, i + CHUNK_SIZE),
    };
    if (!_doSyncRequest(payload, "resi #" + batchNum)) return;
  }

  Logger.log("✅ Sync resi selesai!");
}

// ─── Sync Sent FULL (jalankan manual saja untuk rebuild baseline) ─
// Setelah selesai, counter incremental di-reset ke total baris saat ini
// sehingga syncSentIncremental() berikutnya hanya ambil baris BARU.
function syncSentFull() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("📖 Membaca sent (ini besar, harap tunggu)...");
  const sent = readSheet(ss, SHEET_SENT);
  Logger.log("✔ sent: " + sent.length + " baris → " + Math.ceil(sent.length / CHUNK_SIZE) + " batch");

  let success = true;
  for (let i = 0; i < sent.length; i += CHUNK_SIZE) {
    const batchNum = Math.floor(i / CHUNK_SIZE) + 1;
    const payload  = {
      secret:    WORKER_SECRET,
      clearSent: (i === 0),    // hapus sent hanya di batch pertama
      orders:    [],
      sent:      sent.slice(i, i + CHUNK_SIZE),
      customers: [],
    };
    if (!_doSyncRequest(payload, "sent #" + batchNum)) { success = false; break; }
  }

  if (success) {
    // Simpan total baris sebagai baseline untuk incremental sync berikutnya
    PropertiesService.getScriptProperties().setProperty("SENT_LAST_SYNCED_COUNT", String(sent.length));
    Logger.log("✅ Sync sent full selesai! Baseline incremental diset ke " + sent.length + " baris.");
  } else {
    Logger.log("⚠️ Sync sent full GAGAL di salah satu batch. Baseline TIDAK diupdate.");
  }
}

// ─── Sync Sent INCREMENTAL (otomatis tiap 30 menit bersama syncOrdersAndCustomers) ─
// Hanya mengirim baris-baris BARU yang belum pernah di-sync ke D1.
// Aman karena sheet Terkirim hanya append (baris lama tidak berubah).
function syncSentIncremental() {
  const props = PropertiesService.getScriptProperties();
  const lastCount = parseInt(props.getProperty("SENT_LAST_SYNCED_COUNT") || "0");

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SENT);
  if (!sheet) {
    Logger.log("⚠️ Sheet Terkirim tidak ditemukan: " + SHEET_SENT);
    return;
  }

  const allData   = sheet.getDataRange().getValues();
  const totalRows = allData.length - 1; // dikurangi 1 baris header

  if (totalRows <= lastCount) {
    Logger.log("📭 Incremental sent: tidak ada baris baru (total=" + totalRows + ", lastSynced=" + lastCount + ")");
    return;
  }

  const newRowCount = totalRows - lastCount;
  Logger.log("📦 Incremental sent: ditemukan " + newRowCount + " baris baru (baris " + (lastCount + 1) + "–" + totalRows + ")");

  // Susun header dari baris pertama
  const headers = allData[0].map(function(h) { return String(h).trim(); });

  // Ambil hanya baris baru (index mulai dari lastCount+1 karena baris 0 adalah header)
  const newRows = [];
  for (var i = lastCount + 1; i < allData.length; i++) {
    var obj = {};
    headers.forEach(function(h, j) { obj[h] = formatCell(allData[i][j]); });
    newRows.push(obj);
  }

  // Kirim ke Worker dalam chunk, TANPA clearSent (tidak hapus data lama di D1)
  var success = true;
  for (var c = 0; c < newRows.length; c += CHUNK_SIZE) {
    var batchNum = Math.floor(c / CHUNK_SIZE) + 1;
    var payload  = {
      secret:    WORKER_SECRET,
      clearSent: false,   // PENTING: jangan hapus data lama!
      orders:    [],
      sent:      newRows.slice(c, c + CHUNK_SIZE),
      customers: [],
    };
    if (!_doSyncRequest(payload, "sent-incremental #" + batchNum)) { success = false; break; }
  }

  if (success) {
    // Update counter hanya kalau semua batch berhasil
    props.setProperty("SENT_LAST_SYNCED_COUNT", String(totalRows));
    Logger.log("✅ Incremental sent selesai! Counter diupdate ke " + totalRows + " baris.");
  } else {
    Logger.log("⚠️ Incremental sent GAGAL. Counter TIDAK diupdate (akan dicoba ulang 30 menit lagi).");
  }
}

// ─── Full sync (manual saja, tidak cocok untuk trigger otomatis) ─
function syncToD1() {
  Logger.log("=== FULL SYNC DIMULAI ===");
  syncOrdersAndCustomers();
  syncSentFull();
  Logger.log("=== FULL SYNC SELESAI ===");
}

// ─── Helper: kirim satu chunk, return true jika OK ────────────
function _doSyncRequest(payload, label) {
  try {
    const res  = UrlFetchApp.fetch(WORKER_SYNC_URL, {
      method:             "post",
      contentType:        "application/json",
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    const text = res.getContentText();

    if (code !== 200) {
      Logger.log("❌ [" + label + "] HTTP " + code + ": " + text.substring(0, 200));
      return false;
    }
    const json = JSON.parse(text);
    if (!json.ok) {
      Logger.log("❌ [" + label + "] error: " + JSON.stringify(json));
      return false;
    }
    Logger.log("✔ [" + label + "] orders:" + json.orders + " sent:" + json.sent + " cust:" + json.customers);
    return true;

  } catch (e) {
    Logger.log("❌ [" + label + "] exception: " + e.message);
    return false;
  }
}
