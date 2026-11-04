/**
 * APLIKASI PENDAFTARAN PEMILIH & GENERATE UNDANGAN PDF
 * =============================================================
 * Kolom sheet Pemilih:
 * A No | B NIK | C Nama | D Token | E SudahDicetak | F SudahMemilih |
 * G Timestamp | H NoWhatsApp | I Blok | J NoRumah | K NoBlok
 */
const SHEET_PEMILIH = 'Pemilih';

// ==== KONFIGURASI FOLDER & TEMPLATE GOOGLE DRIVE ====
const FOLDER_UNDANGAN_ID = '1EmJ4lBKym19anF79mfV97jABf4Qoj3DK';
const TEMPLATE_DOC_ID = '18m2mDlKFFh-DZ5s8Kdzo2C5NVzy5wNFjzUW9JCrCbYE';

// ==== PIN MANUAL ====
const PIN_INPUT = '123456'; // PIN halaman INPUT (daftar NIK)
const PIN_CETAK = '654321'; // PIN halaman CETAK undangan & DAFTAR pemilih

/**
 * Jalankan SEKALI dari menu Google Sheet untuk membuat struktur sheet.
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_PEMILIH);
  if (!sh) sh = ss.insertSheet(SHEET_PEMILIH);
  sh.clear();
  sh.appendRow(['No', 'NIK', 'Nama', 'Token', 'SudahDicetak', 'SudahMemilih',
    'Timestamp', 'NoWhatsApp', 'Blok', 'NoRumah', 'NoBlok']);
  sh.setFrozenRows(1);
  sh.getRange('B:B').setNumberFormat('@'); // NIK
  sh.getRange('H:H').setNumberFormat('@'); // NoWhatsApp
  sh.getRange('K:K').setNumberFormat('@'); // NoBlok
  sh.getRange('A:A').setNumberFormat('0');
  sh.autoResizeColumns(1, 11);
  const pesan = 'Setup selesai! Sheet Pemilih sudah siap digunakan.';
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
}

/**
 * MIGRASI: Tambahkan kolom WhatsApp & Alamat jika belum ada
 */
function migrasiTambahKolomKontak() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEMILIH);
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  if (header.indexOf('NoWhatsApp') !== -1) {
    const pesan = 'Kolom NoWhatsApp/Blok/NoRumah sudah ada.';
    try { SpreadsheetApp.getUi().alert(pesan); } catch (e) { Logger.log(pesan); }
    return;
  }
  const kolomBaruMulai = sh.getLastColumn() + 1;
  sh.getRange(1, kolomBaruMulai, 1, 3).setValues([['NoWhatsApp', 'Blok', 'NoRumah']]);
  sh.getRange(1, kolomBaruMulai, sh.getMaxRows(), 1).setNumberFormat('@');
  sh.autoResizeColumns(kolomBaruMulai, 3);
  const pesan = 'Migrasi selesai!';
  try { SpreadsheetApp.getUi().alert(pesan); } catch (e) { Logger.log(pesan); }
}

/**
 * MIGRASI: Tambahkan kolom NoBlok (kolom K) jika belum ada
 */
function migrasiTambahNoBlok() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEMILIH);
  const header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 11)).getValues()[0];
  let pesan;
  if (header[10] === 'NoBlok') {
    pesan = 'Kolom NoBlok sudah ada (kolom K).';
  } else {
    sh.getRange(1, 11).setValue('NoBlok');
    sh.getRange('K:K').setNumberFormat('@');
    sh.autoResizeColumns(11, 1);
    pesan = 'Kolom NoBlok berhasil ditambahkan (kolom K).';
  }
  try { SpreadsheetApp.getUi().alert(pesan); } catch (e) { Logger.log(pesan); }
}

/**
 * Menu di Google Sheet
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🗳 Pendaftaran RW')
    .addItem('1. Setup Awal (jalankan sekali)', 'setupSheet')
    .addItem('🔧 Migrasi: Tambah Kolom WhatsApp/Alamat', 'migrasiTambahKolomKontak')
    .addItem('🔧 Migrasi: Tambah Kolom NoBlok', 'migrasiTambahNoBlok')
    .addItem('📊 Lihat Jumlah Terdaftar', 'lihatJumlahTerdaftar')
    .addToUi();
}

function lihatJumlahTerdaftar() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEMILIH);
  const total = Math.max(sh.getLastRow() - 1, 0);
  const belumDicetak = sh.getLastRow() > 1
    ? sh.getRange(2, 5, sh.getLastRow() - 1, 1).getValues().filter(r => r[0] !== true && r[0] !== 'TRUE').length
    : 0;
  const pesan = 'Total terdaftar: ' + total + ' orang\nBelum dikirim undangannya: ' + belumDicetak + ' orang';
  try {
    SpreadsheetApp.getUi().alert(pesan);
  } catch (e) {
    Logger.log(pesan);
  }
}

/** Entry point Web App */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || 'input';
  const namaFile = page === 'cetak' ? 'Cetak' : (page === 'daftar' ? 'Daftar' : 'Input');
  const judul = page === 'daftar' ? 'Daftar Pemilih RW' : 'Pendaftaran Pemilih RW';
  const template = HtmlService.createTemplateFromFile(namaFile);
  return template.evaluate()
    .setTitle(judul)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * API untuk halaman di GitHub Pages (PWA).
 * Dipanggil lewat fetch POST dengan isi: {"action":"namaFungsi","args":[...]}.
 * Hanya fungsi di daftar putih di bawah yang boleh dipanggil dari luar.
 * PIN tetap diperiksa di dalam masing-masing fungsi.
 */
const AKSI_API = {
  verifikasiPin: verifikasiPin,
  daftarkanPemilih: daftarkanPemilih,
  ubahPemilih: ubahPemilih,
  getSemuaPemilih: getSemuaPemilih,
  getDaftarBelumDicetak: getDaftarBelumDicetak,
  getDaftarSudahDicetak: getDaftarSudahDicetak,
  cariPemilih: cariPemilih,
  tandaiSudahDicetak: tandaiSudahDicetak,
  kirimUndanganPdfPerWarga: kirimUndanganPdfPerWarga
};

function doPost(e) {
  let hasil;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const fungsi = Object.prototype.hasOwnProperty.call(AKSI_API, req.action) ? AKSI_API[req.action] : null;
    if (!fungsi) throw new Error('Aksi tidak dikenal.');
    const args = Array.isArray(req.args) ? req.args : [];
    hasil = { ok: true, data: fungsi.apply(null, args) };
  } catch (err) {
    hasil = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(hasil))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Cek PIN */
function verifikasiPin(pinInput, jenis) {
  const pinBenar = jenis === 'cetak' ? PIN_CETAK : PIN_INPUT;
  if (!pinBenar) {
    return { valid: false, message: 'PIN belum diisi di kode Apps Script.' };
  }
  if (String(pinInput).trim() !== String(pinBenar)) {
    return { valid: false, message: 'PIN salah.' };
  }
  return { valid: true };
}

/** Daftarkan satu pemilih baru */
function daftarkanPemilih(pinInput, nik, nama, noWhatsapp, blok, noBlok, noRumah) {
  const cekPin = verifikasiPin(pinInput, 'input');
  if (!cekPin.valid) return { success: false, message: cekPin.message };

  nik = String(nik).trim();
  nama = String(nama).trim();
  noWhatsapp = String(noWhatsapp).trim().replace(/[\s-]/g, '');
  blok = String(blok).trim().toUpperCase();
  noBlok = String(noBlok).trim();
  noRumah = String(noRumah).trim();

  if (!/^\d{16}$/.test(nik)) {
    return { success: false, message: 'NIK harus berupa 16 digit angka.' };
  }
  if (!nama) {
    return { success: false, message: 'Nama tidak boleh kosong.' };
  }
  if (!/^\+?\d{9,15}$/.test(noWhatsapp)) {
    return { success: false, message: 'No WhatsApp tidak valid.' };
  }
  if (!/^[A-Z]+$/.test(blok)) {
    return { success: false, message: 'Blok hanya boleh huruf kapital (A-Z).' };
  }
  if (!/^\d+$/.test(noBlok)) {
    return { success: false, message: 'No Blok hanya boleh angka.' };
  }
  if (!noRumah) {
    return { success: false, message: 'No Rumah tidak boleh kosong.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEMILIH);
    const lastRow = sh.getLastRow();
    const nikList = lastRow > 1 ? sh.getRange(2, 2, lastRow - 1, 1).getValues().flat().map(String) : [];
    if (nikList.indexOf(nik) !== -1) {
      return { success: false, message: 'NIK sudah terdaftar.' };
    }
    const tokenList = lastRow > 1 ? sh.getRange(2, 4, lastRow - 1, 1).getValues().flat().map(String) : [];
    let token;
    do {
      token = Utilities.getUuid().split('-')[0].toUpperCase();
    } while (tokenList.indexOf(token) !== -1);
    const nomorBaru = lastRow;
    sh.appendRow([nomorBaru, nik, nama, token, false, false, new Date(),
      noWhatsapp, blok, noRumah, noBlok]);
    // pastikan NoBlok tersimpan sebagai teks
    sh.getRange(sh.getLastRow(), 11).setNumberFormat('@').setValue(noBlok);
    return { success: true, message: 'Berhasil didaftarkan: ' + nama, token: token };
  } finally {
    lock.releaseLock();
  }
}

/** Ubah data satu pemilih (untuk perbaikan typo). Kunci baris = kolom No. */
function ubahPemilih(pinInput, no, nik, nama, noWhatsapp, blok, noBlok, noRumah) {
  const cekPin = verifikasiPin(pinInput, 'input');
  if (!cekPin.valid) return { success: false, message: cekPin.message };

  nik = String(nik).trim();
  nama = String(nama).trim();
  noWhatsapp = String(noWhatsapp).trim().replace(/[\s-]/g, '');
  blok = String(blok).trim().toUpperCase();
  noBlok = String(noBlok).trim();
  noRumah = String(noRumah).trim();

  if (!/^\d{16}$/.test(nik)) {
    return { success: false, message: 'NIK harus berupa 16 digit angka.' };
  }
  if (!nama) {
    return { success: false, message: 'Nama tidak boleh kosong.' };
  }
  if (!/^\+?\d{9,15}$/.test(noWhatsapp)) {
    return { success: false, message: 'No WhatsApp tidak valid.' };
  }
  if (!/^[A-Z]+$/.test(blok)) {
    return { success: false, message: 'Blok hanya boleh huruf kapital (A-Z).' };
  }
  if (!/^\d+$/.test(noBlok)) {
    return { success: false, message: 'No Blok hanya boleh angka.' };
  }
  if (!noRumah) {
    return { success: false, message: 'No Rumah tidak boleh kosong.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEMILIH);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return { success: false, message: 'Data pemilih kosong.' };

    const data = sh.getRange(2, 1, lastRow - 1, 2).getValues(); // kolom A (No) & B (NIK)
    let barisSheet = -1;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]) === String(no)) { barisSheet = i + 2; break; }
    }
    if (barisSheet === -1) {
      return { success: false, message: 'Data dengan No ' + no + ' tidak ditemukan.' };
    }

    // NIK tidak boleh sama dengan pemilih LAIN
    for (let i = 0; i < data.length; i++) {
      if (i + 2 !== barisSheet && String(data[i][1]) === nik) {
        return { success: false, message: 'NIK sudah dipakai pemilih lain (No ' + data[i][0] + ').' };
      }
    }

    sh.getRange(barisSheet, 2).setNumberFormat('@').setValue(nik);          // NIK
    sh.getRange(barisSheet, 3).setValue(nama);                              // Nama
    sh.getRange(barisSheet, 8).setNumberFormat('@').setValue(noWhatsapp);   // NoWhatsApp
    sh.getRange(barisSheet, 9).setValue(blok);                              // Blok
    sh.getRange(barisSheet, 10).setValue(noRumah);                          // NoRumah
    sh.getRange(barisSheet, 11).setNumberFormat('@').setValue(noBlok);      // NoBlok

    return { success: true, message: 'Data berhasil diperbarui: ' + nama };
  } finally {
    lock.releaseLock();
  }
}

/** Ambil SEMUA pemilih untuk halaman Daftar */
function getSemuaPemilih(pinInput, jenisPin) {
  const cekPin = verifikasiPin(pinInput, jenisPin === 'input' ? 'input' : 'cetak');
  if (!cekPin.valid) throw new Error(cekPin.message);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEMILIH);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, 11).getValues();
  return data.map(row => ({
    no: row[0],
    nik: String(row[1]),
    nama: row[2],
    sudahDicetak: row[4] === true || row[4] === 'TRUE',
    noWhatsapp: String(row[7] || ''),
    blok: String(row[8] || ''),
    noBlok: String(row[10] || ''),
    noRumah: String(row[9] || '')
  }));
}

/** Ambil daftar pemilih yang BELUM dikirim undangannya */
function getDaftarBelumDicetak(pinInput) {
  const cekPin = verifikasiPin(pinInput, 'cetak');
  if (!cekPin.valid) throw new Error(cekPin.message);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEMILIH);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, 11).getValues();
  return data
    .filter(row => row[4] !== true && row[4] !== 'TRUE')
    .map(row => ({
      no: row[0],
      nik: String(row[1]),
      nama: row[2],
      token: row[3],
      noWhatsapp: String(row[7] || ''),
      blok: row[8] || '',
      noBlok: String(row[10] || ''),
      noRumah: row[9] || ''
    }));
}

/** Ambil daftar pemilih yang SUDAH dikirim undangannya */
function getDaftarSudahDicetak(pinInput) {
  const cekPin = verifikasiPin(pinInput, 'cetak');
  if (!cekPin.valid) throw new Error(cekPin.message);

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEMILIH);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, 11).getValues();
  return data
    .filter(row => row[4] === true || row[4] === 'TRUE')
    .map(row => ({
      no: row[0],
      nik: String(row[1]),
      nama: row[2],
      token: row[3],
      sudahDicetak: true,
      noWhatsapp: String(row[7] || ''),
      blok: row[8] || '',
      noBlok: String(row[10] || ''),
      noRumah: row[9] || ''
    }));
}

/** Cari pemilih berdasarkan NIK atau nama */
function cariPemilih(pinInput, keyword) {
  const cekPin = verifikasiPin(pinInput, 'cetak');
  if (!cekPin.valid) throw new Error(cekPin.message);

  keyword = String(keyword).trim().toLowerCase();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEMILIH);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const data = sh.getRange(2, 1, lastRow - 1, 11).getValues();
  return data
    .filter(row => String(row[1]).includes(keyword) ||
      String(row[2]).toLowerCase().includes(keyword))
    .map(row => ({
      no: row[0],
      nik: String(row[1]),
      nama: row[2],
      token: row[3],
      sudahDicetak: row[4] === true || row[4] === 'TRUE',
      noWhatsapp: String(row[7] || ''),
      blok: row[8] || '',
      noBlok: String(row[10] || ''),
      noRumah: row[9] || ''
    }));
}

/** Tandai status "SudahDicetak" menjadi true */
function tandaiSudahDicetak(pinInput, nikList) {
  const cekPin = verifikasiPin(pinInput, 'cetak');
  if (!cekPin.valid) return { success: false, message: cekPin.message };

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEMILIH);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: true };

  const data = sh.getRange(2, 2, lastRow - 1, 1).getValues();
  const nikSet = new Set(nikList.map(String));
  const updates = [];

  data.forEach((row, idx) => {
    if (nikSet.has(String(row[0]))) {
      updates.push(idx + 2);
    }
  });

  updates.forEach(rowNum => sh.getRange(rowNum, 5).setValue(true));
  return { success: true, jumlahDitandai: updates.length };
}

// =================================================================
// FUNGSI BUAT PDF & SIMPAN KE DRIVE UNTUK WA
// =================================================================

/**
 * Dipanggil saat tombol "Kirim WA" diklik
 */
function kirimUndanganPdfPerWarga(pinInput, nik) {
  const cekPin = verifikasiPin(pinInput, 'cetak');
  if (!cekPin.valid) return { success: false, message: cekPin.message };

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PEMILIH);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: false, message: 'Data pemilih kosong.' };

  const data = sh.getRange(2, 1, lastRow - 1, 11).getValues();
  const barisWarga = data.find(row => String(row[1]) === String(nik));

  if (!barisWarga) {
    return { success: false, message: 'Data warga tidak ditemukan.' };
  }

  const warga = {
    noDpt: barisWarga[0],
    nik: String(barisWarga[1]),
    nama: barisWarga[2],
    token: barisWarga[3],
    noWhatsapp: String(barisWarga[7] || ''),
    blok: barisWarga[8] || '',
    noBlok: String(barisWarga[10] || ''),
    noRumah: barisWarga[9] || ''
  };

  try {
    // 1. Buat PDF & Simpan ke Drive
    const urlPdf = buatFilePdfSatuWarga_(warga);

    // 2. (Penandaan "sudah dikirim" dilakukan saat tombol Buka WhatsApp diklik,
    //     lewat fungsi tandaiSudahDicetak dari halaman Cetak)

    // 3. Kirim balik hasilnya ke Web App
    return {
      success: true,
      message: 'PDF berhasil disimpan ke Drive!',
      pdfUrl: urlPdf,
      noWa: warga.noWhatsapp,
      nama: warga.nama
    };
  } catch (e) {
    return { success: false, message: 'Gagal membuat PDF: ' + e.toString() };
  }
}

/**
 * Generasi PDF dari Google Docs Template dengan QR Code Token
 */
function buatFilePdfSatuWarga_(warga) {
  const folder = DriveApp.getFolderById(FOLDER_UNDANGAN_ID);
  const template = DriveApp.getFileById(TEMPLATE_DOC_ID);

  const namaFile = `Undangan_${warga.nama}_DPT${warga.noDpt}`;
  const salinanDoc = template.makeCopy(namaFile, folder);
  const doc = DocumentApp.openById(salinanDoc.getId());
  const body = doc.getBody();

  // Contoh hasil: Blok C, No Blok 5, No Rumah 12 -> "C5/12"
  const alamat = warga.blok + warga.noBlok + '/' + warga.noRumah;

  // 1. Mengisi text placeholder biasa
  body.replaceText('{{NAMA}}', warga.nama);
  body.replaceText('{{NO_DPT}}', String(warga.noDpt));
  body.replaceText('{{BLOK}}', alamat);

  // 2. Generate gambar QR Code dari Nilai TOKEN Warga
  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=250x250&margin=1&data=${encodeURIComponent(warga.token)}`;
  const qrBlob = UrlFetchApp.fetch(qrUrl).getBlob();

  // 3. Cari placeholder {{BARCODE}} atau {{QR_CODE}} di Google Docs dan ganti dengan Gambar QR
  const placeholder = body.findText('{{BARCODE}}') || body.findText('{{QR_CODE}}');
  if (placeholder) {
    const el = placeholder.getElement();
    const parent = el.getParent();

    // Sisipkan gambar QR Code tepat di posisi placeholder
    if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const img = parent.asParagraph().appendInlineImage(qrBlob);
      img.setWidth(120);
      img.setHeight(120);
      el.deleteText(0, el.getText().length - 1); // Hapus teks placeholder
    }
  }

  doc.saveAndClose();

  // 4. Ubah Dokumen ke PDF
  const pdfBlob = salinanDoc.getAs('application/pdf');
  pdfBlob.setName(`${namaFile}.pdf`);
  const filePdf = folder.createFile(pdfBlob);

  // 5. Bersihkan file sementara & atur akses link PDF
  salinanDoc.setTrashed(true);
  filePdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return filePdf.getUrl();
}

