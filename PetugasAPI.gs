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

// ── GET: Ambil Data Petugas & Riwayat Lembur ────────────────────────
function doGet(e) {
  try {
    const action = e.parameter && e.parameter.action;
    
    // Jika tidak ada action, kembalikan error
    if (!action) {
      return respond({ error: 'Action tidak dikenali atau kosong' }, e);
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // ACTION 1: getPetugas / getUsers
    if (action === 'getUsers' || action === 'getPetugas') {
      const sheet = ss.getSheetByName(SHEET_PETUGAS);
      if (!sheet) return respond({ status: 'error', message: 'Sheet "' + SHEET_PETUGAS + '" tidak ditemukan.' }, e);

      const rows  = sheet.getDataRange().getValues();
      const users = [];

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
    }
    
    // ACTION 2: getHistory
    else if (action === 'getHistory') {
      const sheetData = ss.getSheetByName(SHEET_DATA_SPKL);
      if (!sheetData) return respond({ status: 'error', message: 'Sheet "' + SHEET_DATA_SPKL + '" tidak ditemukan.' }, e);
      
      const rows = sheetData.getDataRange().getValues();
      const history = [];
      
      // Susunan kolom: NO INDUK(0) | NAMA(1) | JABATAN(2) | PENEMPATAN(3) | KATEGORI(4) | KETERANGAN LEMBUR(5) | UPAH(6) | TGL LEMBUR(7) | SHIFT(8) | JAM KERJA(9) | WAKTU LEMBUR(10) | Jml JAM(11) | KETERANGAN(12) | EVIDEN 1(13) | EVIDEN 2(14) | MINGGU KE(15) | TIMESTAMP(16)
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[1]) continue; // Skip jika NAMA kosong
        
        let tglStr = String(r[7] || '');
        if (r[7] instanceof Date) {
          tglStr = Utilities.formatDate(r[7], "Asia/Jakarta", "yyyy-MM-dd");
        }
        
        history.unshift({ // unshift supaya yang terbaru (bawah) ada di atas array
          noInduk: String(r[0] || ''),
          nama: String(r[1] || ''),
          jabatan: String(r[2] || ''),
          penempatan: String(r[3] || ''),
          kategori: String(r[4] || ''),
          keteranganLembur: String(r[5] || ''),
          tglLembur: tglStr,
          shift: String(r[8] || ''),
          jamKerja: String(r[9] || ''),
          waktuLembur: String(r[10] || ''),
          jmlJam: String(r[11] || '0'),
          keteranganTambahan: String(r[12] || ''),
          fotoUrl1: String(r[13] || ''),
          fotoUrl2: String(r[14] || ''),
          mingguKe: String(r[15] || ''),
          timestamp: String(r[16] || ''),
          status: 'DISETUJUI' // Karena ada di database, berarti status disetujui
        });
      }
      
      // Batasi 200 data terakhir agar memori HP/cache tidak terlalu berat
      if (history.length > 200) history.length = 200;
      
      return respond({ status: 'success', data: history }, e);
    }
    
    // Action 3: getPending
    else if (action === 'getPending') {
      const sheetAntrian = ss.getSheetByName('Antrian SPKL');
      if (!sheetAntrian) return respond({ status: 'success', data: [] }, e);
      
      const rows = sheetAntrian.getDataRange().getValues();
      const pending = [];
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r[0]) continue; // Skip header or empty
        try {
          const payload = JSON.parse(r[6] || '{}');
          payload.status = r[2] || payload.status;
          payload.rejectionNote = r[5] || payload.rejectionNote;
          pending.unshift(payload);
        } catch(err) { }
      }
      return respond({ status: 'success', data: pending }, e);
    }
    
    else {
      return respond({ error: 'Action tidak dikenali' }, e);
    }

  } catch (err) {
    return respond({ status: 'error', message: err.message }, e);
  }
}


// ── POST: Simpan Data SPKL ke Rekap Lembur / Antrian ───────────────────────────
function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const data = JSON.parse(e.postData.contents);
    const action = data.action || 'directSubmit';

    // Fungsi pembantu untuk memproses foto base64
    function processPhoto(b64Data, mimeType, fileName) {
      if (!b64Data) return "";
      try {
        let cleanBase64 = b64Data;
        if (cleanBase64.indexOf('base64,') !== -1) {
          cleanBase64 = cleanBase64.split('base64,')[1];
        } else if (cleanBase64.indexOf('data:') === 0) {
          const commaIdx = cleanBase64.indexOf(',');
          if (commaIdx !== -1) cleanBase64 = cleanBase64.substring(commaIdx + 1);
        }
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

    if (action === 'submitPending') {
      let sheetAntrian = ss.getSheetByName('Antrian SPKL');
      if (!sheetAntrian) {
        sheetAntrian = ss.insertSheet('Antrian SPKL');
        sheetAntrian.appendRow(['ID', 'TIMESTAMP', 'STATUS', 'SUBMITTED_ROLE', 'APPROVER_ROLE', 'REJECT_NOTE', 'PAYLOAD']);
      }
      
      const fotoUrl1 = processPhoto(data.fotoBase64, data.fotoMimeType, data.fotoFileName || "Eviden1.jpg");
      const fotoUrl2 = processPhoto(data.fotoBase64_2, data.fotoMimeType_2, data.fotoFileName_2 || "Eviden2.jpg");
      
      data.fotoUrl1 = fotoUrl1;
      data.fotoUrl2 = fotoUrl2;
      delete data.fotoBase64;
      delete data.fotoBase64_2;
      delete data.fotoMimeType;
      delete data.fotoMimeType_2;

      sheetAntrian.appendRow([
        data.id,
        new Date().toISOString(),
        'MENUNGGU',
        data.submittedByRole,
        data.approverRole,
        '',
        JSON.stringify(data)
      ]);
      SpreadsheetApp.flush();
      return respond({ status: "success", message: "Disimpan di antrian" });
    }
    else if (action === 'approvePending') {
      const sheetAntrian = ss.getSheetByName('Antrian SPKL');
      const sheetData = ss.getSheetByName(SHEET_DATA_SPKL);
      if (!sheetAntrian || !sheetData) throw new Error("Sheet tidak ditemukan");
      
      const rows = sheetAntrian.getDataRange().getValues();
      let foundRowIndex = -1;
      let payloadData = null;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
           foundRowIndex = i + 1;
           try { payloadData = JSON.parse(rows[i][6]); } catch(e) {}
           break;
        }
      }
      
      if (foundRowIndex > 0 && payloadData) {
        const rowData = [
          payloadData.noInduk || "",
          payloadData.nama || "",
          payloadData.jabatan || "",
          payloadData.penempatan || "",
          payloadData.kategori || "",
          payloadData.keteranganLembur || "",
          "",
          payloadData.tglLembur || "",
          payloadData.shift || "",
          payloadData.jamKerja || "",
          payloadData.waktuLembur || "",
          payloadData.jmlJam || "",
          payloadData.keteranganTambahan || payloadData.keterangan || "", 
          payloadData.fotoUrl1 || "",
          payloadData.fotoUrl2 || "",
          payloadData.mingguKe || "",
          Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss")
        ];
        sheetData.appendRow(rowData);
        sheetAntrian.deleteRow(foundRowIndex);
        SpreadsheetApp.flush();
        return respond({ status: "success", message: "SPKL Disetujui" });
      } else {
        throw new Error("Data antrian tidak ditemukan");
      }
    }
    else if (action === 'rejectPending') {
      const sheetAntrian = ss.getSheetByName('Antrian SPKL');
      if (!sheetAntrian) throw new Error("Sheet tidak ditemukan");
      const rows = sheetAntrian.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
           sheetAntrian.getRange(i + 1, 3).setValue('DITOLAK');
           sheetAntrian.getRange(i + 1, 6).setValue(data.rejectionNote || '');
           SpreadsheetApp.flush();
           return respond({ status: "success", message: "SPKL Ditolak" });
        }
      }
      throw new Error("Data antrian tidak ditemukan");
    }
    else if (action === 'resubmitPending') {
      const sheetAntrian = ss.getSheetByName('Antrian SPKL');
      if (!sheetAntrian) throw new Error("Sheet tidak ditemukan");
      const rows = sheetAntrian.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
           sheetAntrian.getRange(i + 1, 3).setValue('MENUNGGU');
           sheetAntrian.getRange(i + 1, 6).setValue('');
           
           const fotoUrl1 = processPhoto(data.fotoBase64, data.fotoMimeType, data.fotoFileName || "Eviden1.jpg");
           const fotoUrl2 = processPhoto(data.fotoBase64_2, data.fotoMimeType_2, data.fotoFileName_2 || "Eviden2.jpg");
           if (fotoUrl1) data.fotoUrl1 = fotoUrl1;
           if (fotoUrl2) data.fotoUrl2 = fotoUrl2;
           delete data.fotoBase64; delete data.fotoBase64_2; delete data.fotoMimeType; delete data.fotoMimeType_2;
           
           let oldPayload = {};
           try { oldPayload = JSON.parse(rows[i][6]); } catch(e) {}
           const newPayload = { ...oldPayload, ...data, id: oldPayload.id };
           sheetAntrian.getRange(i + 1, 7).setValue(JSON.stringify(newPayload));
           
           SpreadsheetApp.flush();
           return respond({ status: "success", message: "Diajukan ulang" });
        }
      }
      throw new Error("Data antrian tidak ditemukan");
    }
    else if (action === 'dismissPending') {
      const sheetAntrian = ss.getSheetByName('Antrian SPKL');
      if (!sheetAntrian) throw new Error("Sheet tidak ditemukan");
      const rows = sheetAntrian.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][0] === data.id) {
           sheetAntrian.deleteRow(i + 1);
           SpreadsheetApp.flush();
           return respond({ status: "success", message: "Dihapus" });
        }
      }
      throw new Error("Data antrian tidak ditemukan");
    }
    else {
      // Direct Submit (legacy code)
      const sheet = ss.getSheetByName(SHEET_DATA_SPKL);
      if (!sheet) throw new Error("Sheet '" + SHEET_DATA_SPKL + "' tidak ditemukan!");
      
      const fotoUrl1 = processPhoto(data.fotoBase64, data.fotoMimeType, data.fotoFileName || "Eviden1.jpg");
      const fotoUrl2 = processPhoto(data.fotoBase64_2, data.fotoMimeType_2, data.fotoFileName_2 || "Eviden2.jpg");

      const rowData = [
        data.noInduk || "",
        data.nama || "",
        data.jabatan || "",
        data.penempatan || "",
        data.kategori || "",
        data.keteranganLembur || "",
        "",
        data.tglLembur || "",
        data.shift || "",
        data.jamKerja || "",
        data.waktuLembur || "",
        data.jmlJam || "",
        data.keteranganTambahan || data.keterangan || "", 
        fotoUrl1 || "",
        fotoUrl2 || "",
        data.mingguKe || "",
        Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss")
      ];

      sheet.appendRow(rowData);
      SpreadsheetApp.flush();

      return respond({ 
        status: "success", 
        message: "Data berhasil disimpan",
        insertedAtRow: sheet.getLastRow(),
        sheetName: sheet.getName()
      });
    }
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
