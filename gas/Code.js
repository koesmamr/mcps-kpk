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
    
    // 2. Upload file foto jika ada dan dapatkan objek File
    const fileR2 = data.fotoR2 ? saveBase64File(data.fotoR2, 'R2_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : null;
    const fileR4 = data.fotoR4 ? saveBase64File(data.fotoR4, 'R4_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : null;
    const fileRumah = data.fotoRumah ? saveBase64File(data.fotoRumah, 'Rumah_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : null;
    const fileAlat = data.fotoAlat ? saveBase64File(data.fotoAlat, 'Alat_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : null;

    // 3. Buka Spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Target sheet: cari sheet bernama 'Sheet1'
    let sheet = ss.getSheetByName('Sheet1');
    if (!sheet) {
      sheet = ss.getSheets()[0];
    }

    // 4. Tentukan baris tujuan untuk mengisi data
    // Baris data dimulai dari baris 9
    let targetRow = 9;
    const maxRows = sheet.getMaxRows();
    
    // Cari baris pertama yang Kolom B (Nama Pegawai) kosong atau berisi dummy awal
    while (targetRow <= maxRows) {
      const valB = sheet.getRange(targetRow, 2).getValue().toString().trim();
      if (!valB || valB === 'XXXX' || valB === 'TEST VERIFIKASI SISTEM' || valB === 'Contoh Pegawai Uji Coba') {
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

    // Gabungkan Jabatan dan Status Pegawai
    const jabatanStatus = (data.jabatan || '') + (data.statusPegawai ? '\n' + data.statusPegawai : '');

    // Isi sel data teks
    sheet.getRange(targetRow, 1).setValue(nomorUrut);
    sheet.getRange(targetRow, 2).setValue(data.namaPegawai || '');
    sheet.getRange(targetRow, 3).setValue(jabatanStatus);
    sheet.getRange(targetRow, 4).setValue(data.kendaraanR2 || '-');
    sheet.getRange(targetRow, 6).setValue(data.kendaraanR4 || '-');
    sheet.getRange(targetRow, 8).setValue(data.rumahDinas || '-');
    sheet.getRange(targetRow, 10).setValue(data.peralatanKantor || '-');
    sheet.getRange(targetRow, 12).setValue(nomorUrut + '...............');

    // Masukkan foto langsung ke dalam sel (In-Cell Image)
    insertImageIntoCell(sheet, targetRow, 5, fileR2);
    insertImageIntoCell(sheet, targetRow, 7, fileR4);
    insertImageIntoCell(sheet, targetRow, 9, fileRumah);
    insertImageIntoCell(sheet, targetRow, 11, fileAlat);

    // Atur tinggi baris agar foto terlihat proporsional dan jelas
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
        fileR2 ? fileR2.getUrl() : '',
        fileR4 ? fileR4.getUrl() : '',
        fileRumah ? fileRumah.getUrl() : '',
        fileAlat ? fileAlat.getUrl() : ''
      ]);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Data dan foto berhasil masuk ke dalam cell Sheet1 baris ' + targetRow + '!',
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
 * Menyisipkan foto ke dalam cell menggunakan CellImageBuilder
 */
function insertImageIntoCell(sheet, row, col, file) {
  const cell = sheet.getRange(row, col);
  if (!file) {
    cell.setValue('-');
    return;
  }

  try {
    const directUrl = 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
    const cellImage = SpreadsheetApp.newCellImage()
      .setSourceUrl(directUrl)
      .setAltTextTitle('Foto Bukti Fisik')
      .build();
    cell.setValue(cellImage);
  } catch (err) {
    // Fallback: rumus HYPERLINK ke foto di Google Drive
    cell.setFormula('=HYPERLINK(\"' + file.getUrl() + '\", \"Lihat Foto\")');
  }
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
    
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file;
  } catch (e) {
    return null;
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
