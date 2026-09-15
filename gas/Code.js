const SPREADSHEET_ID = '1pbXNUIck_2nlKmVPltYkyfwHP-99i8iztjHw57ehP0k';
const FOLDER_ID = '19gMa6kFzEjO3UfMlvl8I44Nl8u-qPNtt';

/**
 * Handle HTTP GET request (untuk test status API)
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Backend API MCPS KPK (koesmamr@gmail.com) siap digunakan.'
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Inisialisasi struktur sheet dan tabel
 */
function initSheetStructure() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName('Sheet1');
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName('Sheet1');
  }

  // Set judul & header jika masih kosong
  if (sheet.getRange(1, 1).getValue() === '') {
    sheet.getRange('A1:L1').merge().setValue('LAMPIRAN BERITA ACARA PAKTA INTEGRITAS').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('A2:L2').merge().setValue('PENGGUNAAN BARANG BARANG MILIK DAERAH').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('A3:L3').merge().setValue('BADAN/DINAS/KANTOR/KECAMATAN/UPT.................... KAB TUBAN').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('A4:L4').merge().setValue('PEMENUHAN MCSP KPK-RI TAHUN 2026').setFontWeight('bold').setHorizontalAlignment('center');

    // Header Tabel Baris 6 - 8
    sheet.getRange('A6:A8').merge().setValue('NO').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange('B6:B8').merge().setValue('NAMA PEGAWAI').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange('C6:C7').merge().setValue('JABATAN/').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
    sheet.getRange('C8').setValue('STATUS PEGAWAI').setFontWeight('bold').setHorizontalAlignment('center');

    sheet.getRange('D6:K6').merge().setValue('PENGGUNAAN BARANG MILIK DAERAH').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('D7:E7').merge().setValue('Kendaraan Dinas R2').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('D8').setValue('Merk, Type dan Plat Nomor').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('E8').setValue('Foto').setFontWeight('bold').setHorizontalAlignment('center');

    sheet.getRange('F7:G7').merge().setValue('Kendaraan Dinas R4').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('F8').setValue('Merk, Type dan Plat Nomor').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('G8').setValue('Foto').setFontWeight('bold').setHorizontalAlignment('center');

    sheet.getRange('H7:I7').merge().setValue('Rumah Dinas').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('H8').setValue('Rumah Jabatan / Alamat').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('I8').setValue('Foto').setFontWeight('bold').setHorizontalAlignment('center');

    sheet.getRange('J7:K7').merge().setValue('Peralatan Kantor').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('J8').setValue('Merk, Type dan Tahun').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('K8').setValue('Foto').setFontWeight('bold').setHorizontalAlignment('center');

    sheet.getRange('L6:L8').merge().setValue('TANDA TANGAN').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');

    sheet.setColumnWidth(1, 45);   // No
    sheet.setColumnWidth(2, 200);  // Nama
    sheet.setColumnWidth(3, 180);  // Jabatan / Status
    sheet.setColumnWidth(4, 180);  // R2
    sheet.setColumnWidth(5, 140);  // Foto R2
    sheet.setColumnWidth(6, 180);  // R4
    sheet.setColumnWidth(7, 140);  // Foto R4
    sheet.setColumnWidth(8, 200);  // Rumah Dinas
    sheet.setColumnWidth(9, 140);  // Foto Rumah
    sheet.setColumnWidth(10, 200); // Alat Kantor
    sheet.setColumnWidth(11, 140); // Foto Alat
    sheet.setColumnWidth(12, 130); // TTD
  }

  // Pastikan ada sheet Response/Raw
  let responseSheet = ss.getSheetByName('Jawaban Formulir 1');
  if (!responseSheet) {
    responseSheet = ss.insertSheet('Jawaban Formulir 1');
    responseSheet.appendRow([
      'Cap waktu',
      'Nama Pegawai',
      'NIP / NIPPPK',
      'Jabatan',
      'Status Pegawai',
      'Kendaraan R2: Merk, Type dan Plat Nomor',
      'Kendaraan R4: Merk, Type dan Plat Nomor',
      'Rumah Dinas: Nama / Jabatan / Alamat',
      'Peralatan Kantor: Merk, Type dan Tahun',
      'Pernyataan Kebenaran Data',
      'Link Foto R2',
      'Link Foto R4',
      'Link Foto Rumah Dinas',
      'Link Foto Alat Kantor'
    ]);
  }
}

/**
 * Handle HTTP POST request dari formulir frontend
 */
function doPost(e) {
  try {
    initSheetStructure();

    const data = JSON.parse(e.postData.contents);
    const targetFolder = DriveApp.getFolderById(FOLDER_ID);

    // 1. Upload file foto jika ada
    const fileR2 = data.fotoR2 ? saveBase64File(data.fotoR2, 'R2_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : null;
    const fileR4 = data.fotoR4 ? saveBase64File(data.fotoR4, 'R4_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : null;
    const fileRumah = data.fotoRumah ? saveBase64File(data.fotoRumah, 'Rumah_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : null;
    const fileAlat = data.fotoAlat ? saveBase64File(data.fotoAlat, 'Alat_' + sanitize(data.namaPegawai) + '_' + Date.now(), targetFolder) : null;

    // 2. Buka Spreadsheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Sheet1') || ss.getSheets()[0];

    // 3. Tentukan baris data (mulai dari baris 9)
    let targetRow = 9;
    const maxRows = sheet.getMaxRows();
    
    while (targetRow <= maxRows) {
      const valB = sheet.getRange(targetRow, 2).getValue().toString().trim();
      if (!valB || valB === 'XXXX') {
        break;
      }
      targetRow++;
    }

    if (targetRow > maxRows) {
      sheet.insertRowAfter(maxRows);
    }

    const nomorUrut = targetRow - 8;
    const jabatanStatus = (data.jabatan || '') + (data.statusPegawai ? '\n' + data.statusPegawai : '');

    // 4. Masukkan data teks
    sheet.getRange(targetRow, 1).setValue(nomorUrut);
    sheet.getRange(targetRow, 2).setValue(data.namaPegawai || '');
    sheet.getRange(targetRow, 3).setValue(jabatanStatus);
    sheet.getRange(targetRow, 4).setValue(data.kendaraanR2 || '-');
    sheet.getRange(targetRow, 6).setValue(data.kendaraanR4 || '-');
    sheet.getRange(targetRow, 8).setValue(data.rumahDinas || '-');
    sheet.getRange(targetRow, 10).setValue(data.peralatanKantor || '-');
    sheet.getRange(targetRow, 12).setValue(nomorUrut + '...............');

    // 5. Masukkan foto langsung ke dalam sel (In-Cell Image)
    insertImageIntoCell(sheet, targetRow, 5, fileR2);
    insertImageIntoCell(sheet, targetRow, 7, fileR4);
    insertImageIntoCell(sheet, targetRow, 9, fileRumah);
    insertImageIntoCell(sheet, targetRow, 11, fileAlat);

    // Styling baris data
    sheet.setRowHeight(targetRow, 100);
    sheet.getRange(targetRow, 1, 1, 12).setVerticalAlignment('middle');

    // 6. Simpan arsip raw ke tab 'Jawaban Formulir 1'
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
      message: 'Data dan foto berhasil disimpan di spreadsheet akun koesmamr@gmail.com (baris ' + targetRow + ')!',
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
    cell.setFormula('=HYPERLINK(\"' + file.getUrl() + '\", \"Lihat Foto\")');
  }
}

/**
 * Simpan file Base64 ke folder Google Drive
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

function sanitize(str) {
  if (!str) return 'anonim';
  return str.replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 30);
}
