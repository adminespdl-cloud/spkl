// ╔══════════════════════════════════════════════════════════╗
// ║  SPKL Padalarang - API Petugas & Rekap Lembur (Lengkap)  ║
// ║  CARA DEPLOY ULANG:                                      ║
// ║  1. Buka https://script.google.com                       ║
// ║  2. Buka project yang sebelumnya dibuat                  ║
// ║  3. Hapus semua kode, paste kode baru ini                ║
// ║  4. Klik "Deploy" -> "Manage deployments"                ║
// ║  5. Edit (icon pensil) -> Version: "New version"         ║
// ║  6. Klik Deploy -> Berikan otorisasi Drive (jika diminta)║
// ╚══════════════════════════════════════════════════════════╝

const SPREADSHEET_ID  = '1YiXREQJ4Xl39L1N6-FjPY40YoqZZcyCNLnEPi5drIJ4';
const SHEET_PETUGAS   = 'Petugas';
const SHEET_DATA_SPKL = 'Rekap Lembur';

// ── GET: Ambil Data Petugas (Login) ──────────────────────────────────
function doGet(e) {
  try {
    const action = e.parameter && e.parameter.action;
    
    // Jika tidak ada action=getUsers / getPetugas, coba kembalikan format lama 
    // agar login fallback tetap aman
    if (action !== 'getUsers' && action !== 'getPetugas') {
      return respond({ error: 'Action tidak dikenali atau kosong' }, e);
    }

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_PETUGAS);
    
    if (!sheet) {
      return respond({ status: 'error', message: 'Sheet "' + SHEET_PETUGAS + '" tidak ditemukan.' }, e);
    }

    const rows  = sheet.getDataRange().getValues();
    const users = [];

    // Mulai dari baris ke-2 (index 1) — skip header
    for (let i = 1; i < rows.length; i++) {
      const r   = rows[i];
      const nip = String(r[0] || '').trim();
      const pwd = String(r[1] || '').trim();
      if (!nip || !pwd) continue;
      users.push({
        username:    nip,
        password:    pwd,
        nama:        String(r[2] || '').trim(),
        jabatan:     String(r[3] || '').trim(),
        penempatan:  String(r[4] || '').trim(),
        role:        String(r[5] || 'operator').trim().toLowerCase()
      });
    }

    const sheetData = ss.getSheetByName(SHEET_DATA_SPKL);
    const dataRowCount = sheetData ? sheetData.getLastRow() : 0;

    return respond({ status: 'success', data: users, dataRowCount: dataRowCount, spreadsheetId: SPREADSHEET_ID }, e);

  } catch (err) {
    return respond({ status: 'error', message: err.message }, e);
  }
}


// ── POST: Simpan Data SPKL ke Rekap Lembur ───────────────────────────
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_DATA_SPKL);
    if (!sheet) throw new Error("Sheet '" + SHEET_DATA_SPKL + "' tidak ditemukan!");

    const data = JSON.parse(e.postData.contents);
    
    // Fungsi pembantu untuk memproses foto base64
    function processPhoto(b64Data, mimeType, fileName) {
      if (!b64Data) return "";
      try {
        let cleanBase64 = b64Data;
        // Jika mengandung "base64,", ambil string sesudahnya
        if (cleanBase64.indexOf('base64,') !== -1) {
          cleanBase64 = cleanBase64.split('base64,')[1];
        } else if (cleanBase64.indexOf('data:') === 0) {
          // Jika masih mengandung header data URI utuh
          const commaIdx = cleanBase64.indexOf(',');
          if (commaIdx !== -1) cleanBase64 = cleanBase64.substring(commaIdx + 1);
        }
        
        // Buat file jika cleanBase64 memiliki panjang yang masuk akal
        if (cleanBase64 && cleanBase64.length > 100) {
          const bytes = Utilities.base64Decode(cleanBase64);
          const blob = Utilities.newBlob(bytes, mimeType || "image/jpeg", fileName || "Eviden.jpg");
          const file = DriveApp.createFile(blob);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          return file.getUrl();
        }
      } catch (err) {
        return "Data foto error: " + err.message;
      }
      return "Data foto tidak valid";
    }

    const fotoUrl1 = processPhoto(data.fotoBase64, data.fotoMimeType, data.fotoFileName || "Eviden1.jpg");
    const fotoUrl2 = processPhoto(data.fotoBase64_2, data.fotoMimeType_2, data.fotoFileName_2 || "Eviden2.jpg");

    // Susunan kolom sesuai permintaan:
    // NO INDUK | NAMA | JABATAN | PENEMPATAN | KATEGORI | KETERANGAN LEMBUR | UPAH | TGL LEMBUR | SHIFT | JAM KERJA | WAKTU LEMBUR/KJK | Jml JAM | KETERANGAN | EVIDEN FOTO 1 | EVIDEN FOTO 2 | MINGGU KE | TIMESTAMP
    const rowData = [
      data.noInduk || "",
      data.nama || "",
      data.jabatan || "",
      data.penempatan || "",
      data.kategori || "",
      data.keteranganLembur || "",
      "",                             // UPAH (Dikosongkan / belum ada perhitungan di form)
      data.tglLembur || "",
      data.shift || "",
      data.jamKerja || "",            // JAM KERJA (dari dropdown: 08.00-16.00, dll)
      data.waktuLembur || "",         // WAKTU LEMBUR / KJK (contoh: 08:00 - 13:00)
      data.jmlJam || "",              // Jml JAM (contoh: 5)
      data.keteranganTambahan || data.keterangan || "", 
      fotoUrl1,                       // EVIDEN FOTO 1 (Link Google Drive)
      fotoUrl2,                       // EVIDEN FOTO 2 (Link Google Drive)
      data.mingguKe || "",
      Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss") // TIMESTAMP
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush(); // Memaksa script menulis data ke sheet detik ini juga

    const insertedRow = sheet.getLastRow();

    return respond({ 
      status: "success", 
      message: "Data berhasil disimpan",
      insertedAtRow: insertedRow,
      sheetName: sheet.getName()
    });
  } catch (error) {
    return respond({ status: "error", message: error.message });
  }
}

// Helper: kirim response JSON 
function respond(data, e) {
  const json = JSON.stringify(data);
  const cb   = e && e.parameter && e.parameter.callback;
  if (cb) {
    return ContentService
      .createTextOutput(cb + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
