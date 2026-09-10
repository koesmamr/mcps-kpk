const SPREADSHEET_ID = '1kucE7kAQ3wJ_XAQ9JefRaDaRE4W_uBe2ICOXuEwt1kA';
const FOLDER_NAME = 'Uploads_MCPS_KPK';

/**
 * Handle HTTP GET request (untuk test status API)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Backend API MCPS KPK siap digunakan.'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle HTTP POST request dari formulir frontend
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    // 1. Ambil atau Buat Folder Penyimpanan di Google Drive
    const targetFolder = getOrCreateFolder(FOLDER_NAME);
    
    // 2. Upload file foto jika ada
    const fotoR2Url = data.fotoR2 ? saveBase64File(data.fotoR2, 'R2_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : '';
    const fotoR4Url = data.fotoR4 ? saveBase64File(data.fotoR4, 'R4_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : '';
    const fotoRumahUrl = data.fotoRumah ? saveBase64File(data.fotoRumah, 'Rumah_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : '';
    const fotoAlatUrl = data.fotoAlat ? saveBase64File(data.fotoAlat, 'Alat_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : '';

    // 3. Buka Spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Target sheet: cari sheet bernama 'Sheet1'
    let sheet = ss.getSheetByName('Sheet1');
    if (!sheet) {
      sheet = ss.getSheets()[0];
    }

    // 4. Tentukan baris tujuan untuk mengisi data
    // Baris data dimulai dari baris 9
    // Format kolom:
    // Kolom 1 (A): NO
    // Kolom 2 (B): NAMA PEGAWAI
    // Kolom 3 (C): JABATAN / STATUS PEGAWAI
    // Kolom 4 (D): R2 (Merk, Type dan Plat Nomor)
    // Kolom 5 (E): FOTO R2 (IMAGE formula)
    // Kolom 6 (F): R4 (Merk, Type dan Plat Nomor)
    // Kolom 7 (G): FOTO R4 (IMAGE formula)
    // Kolom 8 (H): Rumah Dinas (Nama / Alamat)
    // Kolom 9 (I): FOTO RUMAH (IMAGE formula)
    // Kolom 10 (J): Peralatan Kantor (Merk, Type, Tahun)
    // Kolom 11 (K): FOTO PERALATAN (IMAGE formula)
    // Kolom 12 (L): TANDA TANGAN (No. urut ttd)

    let targetRow = 9;
    const maxRows = sheet.getMaxRows();
    
    // Cari baris pertama yang Kolom B (Nama Pegawai) kosong atau berisi dummy XXXX pada baris 9-10
    while (targetRow <= maxRows) {
      const valB = sheet.getRange(targetRow, 2).getValue().toString().trim();
      if (!valB || valB === 'XXXX') {
        break;
      }
      targetRow++;
    }

    // Jika melebihi baris yang ada, tambah baris
    if (targetRow > maxRows) {
      sheet.insertRowAfter(maxRows);
    }

    // Hitung nomor urut
    const nomorUrut = targetRow - 8;

    // Gabungkan Jabatan dan Status Pegawai jika perlu
    const jabatanStatus = (data.jabatan || '') + (data.statusPegawai ? '\n' + data.statusPegawai : '');

    // Siapkan formula IMAGE untuk menampilkan foto di cell
    const formulaR2 = fotoR2Url ? '=IMAGE(\"' + getDirectImageUrl(fotoR2Url) + '\", 1)' : '';
    const formulaR4 = fotoR4Url ? '=IMAGE(\"' + getDirectImageUrl(fotoR4Url) + '\", 1)' : '';
    const formulaRumah = fotoRumahUrl ? '=IMAGE(\"' + getDirectImageUrl(fotoRumahUrl) + '\", 1)' : '';
    const formulaAlat = fotoAlatUrl ? '=IMAGE(\"' + getDirectImageUrl(fotoAlatUrl) + '\", 1)' : '';

    // Isi sel data
    sheet.getRange(targetRow, 1).setValue(nomorUrut);
    sheet.getRange(targetRow, 2).setValue(data.namaPegawai || '');
    sheet.getRange(targetRow, 3).setValue(jabatanStatus);
    sheet.getRange(targetRow, 4).setValue(data.kendaraanR2 || '-');
    if (formulaR2) {
      sheet.getRange(targetRow, 5).setFormula(formulaR2);
    } else {
      sheet.getRange(targetRow, 5).setValue('-');
    }

    sheet.getRange(targetRow, 6).setValue(data.kendaraanR4 || '-');
    if (formulaR4) {
      sheet.getRange(targetRow, 7).setFormula(formulaR4);
    } else {
      sheet.getRange(targetRow, 7).setValue('-');
    }

    sheet.getRange(targetRow, 8).setValue(data.rumahDinas || '-');
    if (formulaRumah) {
      sheet.getRange(targetRow, 9).setFormula(formulaRumah);
    } else {
      sheet.getRange(targetRow, 9).setValue('-');
    }

    sheet.getRange(targetRow, 10).setValue(data.peralatanKantor || '-');
    if (formulaAlat) {
      sheet.getRange(targetRow, 11).setFormula(formulaAlat);
    } else {
      sheet.getRange(targetRow, 11).setValue('-');
    }

    // Kolom tanda tangan
    sheet.getRange(targetRow, 12).setValue(nomorUrut + '...............');

    // Atur tinggi baris agar foto terlihat proporsional dan jelas (tinggi 95-100px)
    sheet.setRowHeight(targetRow, 100);

    // Atur perataan vertikal ke tengah
    sheet.getRange(targetRow, 1, 1, 12).setVerticalAlignment('middle');

    // Catat juga ke tab 'Jawaban Formulir 1' jika ada sebagai arsip raw
    const responseSheet = ss.getSheetByName('Jawaban Formulir 1');
    if (responseSheet) {
      const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
      responseSheet.appendRow([
        timestamp,
        data.namaPegawai || '',
        data.nip || '',
        data.jabatan || '',
        data.statusPegawai || '',
        data.kendaraanR2 || '',
        data.kendaraanR4 || '',
        data.rumahDinas || '',
        data.peralatanKantor || '',
        data.pernyataan || 'Setuju / Benar',
        fotoR2Url,
        fotoR4Url,
        fotoRumahUrl,
        fotoAlatUrl
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Data dan foto berhasil disimpan langsung ke dalam tabel Sheet1 baris ' + targetRow + '!',
      row: targetRow
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Ubah Google Drive sharing URL ke direct content URL agar fungsi =IMAGE(...) dapat menampilkan foto
 */
function getDirectImageUrl(driveUrl) {
  if (!driveUrl) return '';
  const match = driveUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return 'https://lh3.googleusercontent.com/d/' + match[1];
  }
  return driveUrl;
}

/**
 * Simpan file Base64 ke folder Google Drive dan jadikan link dapat diakses
 */
function saveBase64File(base64DataUrl, fileNamePrefix, folder) {
  try {
    const parts = base64DataUrl.split(',');
    const meta = parts[0];
    const base64Str = parts[1];
    
    let mimeType = 'image/jpeg';
    let ext = '.jpg';
    if (meta.indexOf('image/png') !== -1) {
      mimeType = 'image/png';
      ext = '.png';
    } else if (meta.indexOf('image/webp') !== -1) {
      mimeType = 'image/webp';
      ext = '.webp';
    }

    const decoded = Utilities.base64Decode(base64Str);
    const blob = Utilities.newBlob(decoded, mimeType, fileNamePrefix + ext);
    const file = folder.createFile(blob);
    
    // Set permission agar foto dapat dirender oleh rumus =IMAGE()
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (e) {
    return '';
  }
}

/**
 * Dapatkan folder Google Drive, atau buat jika belum ada
 */
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function sanitize(str) {
  if (!str) return 'anonim';
  return str.replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 30);
}
