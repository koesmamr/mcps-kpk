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

    // 3. Format isian dengan menyertakan link foto jika ada
    const r2Detail = formatFieldWithLink(data.kendaraanR2, fotoR2Url);
    const r4Detail = formatFieldWithLink(data.kendaraanR4, fotoR4Url);
    const rumahDetail = formatFieldWithLink(data.rumahDinas, fotoRumahUrl);
    const alatDetail = formatFieldWithLink(data.peralatanKantor, fotoAlatUrl);

    // 4. Buka Spreadsheet dan tulis baris baru
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets()[0]; // Ambil sheet pertama

    // Format Timestamp
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    const row = [
      timestamp,
      data.namaPegawai || '',
      data.nip || '',
      data.jabatan || '',
      data.statusPegawai || '',
      r2Detail,
      r4Detail,
      rumahDetail,
      alatDetail,
      data.pernyataan || 'Setuju / Benar'
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Data dan dokumen berhasil disimpan ke Google Sheets & Google Drive!',
      timestamp: timestamp
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Gabungkan keterangan aset dengan link foto jika tersedia
 */
function formatFieldWithLink(text, fileUrl) {
  const keterangan = (text || '').trim();
  if (!keterangan && !fileUrl) return '-';
  if (keterangan && fileUrl) {
    return keterangan + '\nFoto: ' + fileUrl;
  }
  return keterangan || ('Foto: ' + fileUrl);
}

/**
 * Simpan file Base64 ke folder Google Drive dan jadikan link dapat diakses
 */
function saveBase64File(base64DataUrl, fileNamePrefix, folder) {
  try {
    // base64DataUrl format: "data:image/jpeg;base64,/9j/4AAQSkZ..."
    const parts = base64DataUrl.split(',');
    const meta = parts[0];
    const base64Str = parts[1];
    
    // Deteksi mime type dan ekstensi
    let mimeType = 'image/jpeg';
    let ext = '.jpg';
    if (meta.indexOf('image/png') !== -1) {
      mimeType = 'image/png';
      ext = '.png';
    } else if (meta.indexOf('image/webp') !== -1) {
      mimeType = 'image/webp';
      ext = '.webp';
    } else if (meta.indexOf('application/pdf') !== -1) {
      mimeType = 'application/pdf';
      ext = '.pdf';
    }

    const decoded = Utilities.base64Decode(base64Str);
    const blob = Utilities.newBlob(decoded, mimeType, fileNamePrefix + ext);
    const file = folder.createFile(blob);
    
    // Set permission agar file bisa dilihat oleh siapapun yang memiliki tautan
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (e) {
    return 'Gagal upload file: ' + e.message;
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
