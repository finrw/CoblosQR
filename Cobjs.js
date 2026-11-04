/**
 * =====================================================================
 *  PEMILIHAN DIGITAL RW - BACKEND (code.gs)
 *  Hosting tampilan (HTML/CSS/JS) ada di GitHub Pages.
 *  Script ini hanya menjadi API JSON (Web App).
 * =====================================================================
 *
 *  SPREADSHEET KHUSUS PENCOBLOSAN (terpisah dari PILIH_RW).
 *  Script ini TERIKAT ke spreadsheet pencoblosan tersebut.
 *
 *  SHEET YANG DIPAKAI:
 *   - Pemilih  : daftar token dari PILIH_RW. Terisi otomatis saat token
 *                di-scan (sinkron langsung ke PILIH_RW), atau lewat menu
 *                "Pemilihan > Impor / Sinkron Token".
 *                Kolom: Token | SudahMemilih
 *   - Kandidat : dibuat otomatis oleh setup() -> NoKandidat | Nama | Foto
 *   - Suara    : dibuat otomatis oleh setup() -> Kode | NoKandidat |
 *                StatusRekap | WaktuRekap
 *
 *  KERAHASIAAN:
 *   - Sheet Pemilih hanya diberi tanda SudahMemilih = TRUE (tanpa waktu).
 *   - Sheet Suara TIDAK menyimpan identitas/token/waktu coblos.
 *   - Baris suara disisipkan di posisi acak supaya urutan baris tidak
 *     bisa dicocokkan dengan urutan orang yang memilih.
 *
 *  ACTION API (POST, body JSON, Content-Type: text/plain):
 *   cekToken    {token}               -> data pemilih + daftar kandidat
 *   coblos      {token, noKandidat}   -> {kode} (isi barcode resi)
 *   rekapScan   {kode, pin}           -> tandai 1 resi sudah dihitung
 *   getRekap    {pin}                 -> hasil per kandidat
 * =====================================================================
 */

// ======================= KONFIGURASI =================================
var CFG = {
  SHEET_PEMILIH: 'Pemilih',
  SHEET_KANDIDAT: 'Kandidat',
  SHEET_SUARA: 'Suara',
  PANJANG_KODE: 10,
  ALFABET_KODE: 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789' // tanpa I, O, 0, 1 agar mudah dibaca
};

// ======================= ENTRY POINT =================================
function doGet(e) {
  return json_({ ok: true, pesan: 'API Pemilihan RW aktif' });
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents || '{}');
    switch (req.action) {
      case 'cekToken':  return json_(cekToken_(req.token));
      case 'coblos':    return json_(coblos_(req.token, req.noKandidat));
      case 'rekapScan': return json_(rekapScan_(req.kode, req.pin));
      case 'getRekap':  return json_(getRekap_(req.pin));
      case 'ambilSuara': return json_(ambilSuara_(req.pin));
      default:          return json_({ ok: false, error: 'AKSI_TIDAK_DIKENAL' });
    }
  } catch (err) {
    return json_({ ok: false, error: 'SERVER_ERROR', detail: String(err) });
  }
}

// ======================= SETUP (jalankan 1x) =========================
/**
 * Jalankan manual dari editor: membuat sheet Kandidat & Suara jika belum ada.
 */
function setup() {
  var ss = getSS_();

  var kan = ss.getSheetByName(CFG.SHEET_KANDIDAT);
  if (!kan) {
    kan = ss.insertSheet(CFG.SHEET_KANDIDAT);
    kan.appendRow(['NoKandidat', 'Nama', 'Foto']);
    kan.appendRow([1, 'Nama Kandidat 1', '']);
    kan.appendRow([2, 'Nama Kandidat 2', '']);
    kan.setFrozenRows(1);
  }

  var sua = ss.getSheetByName(CFG.SHEET_SUARA);
  if (!sua) {
    sua = ss.insertSheet(CFG.SHEET_SUARA);
    sua.appendRow(['Kode', 'NoKandidat', 'StatusRekap', 'WaktuRekap']);
    sua.setFrozenRows(1);
  }

  var pem = ss.getSheetByName(CFG.SHEET_PEMILIH);
  if (!pem) {
    pem = ss.insertSheet(CFG.SHEET_PEMILIH);
    pem.appendRow(['Token', 'SudahMemilih']);
    pem.setFrozenRows(1);
  }

  Logger.log('Setup selesai. Edit sheet Kandidat, isi SOURCE_SS_ID, lalu impor pemilih.');
}

/** Menu di spreadsheet agar panitia tidak perlu membuka editor kode. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Pemilihan')
    .addItem('1. Setup sheet', 'setup')
    .addItem('2. Impor / Sinkron Token', 'importPilihan')
    .addToUi();
}

/**
 * Isi ID spreadsheet PILIH_RW (bagian URL antara /d/ dan /edit), jalankan 1x,
 * lalu boleh dikembalikan ke teks semula.
 */
function setSumberPemilihan() {
  var ID_SUMBER = '1PHVsHaghFCFdtZiE-RuyHsC7O9BKaUgnJAFj2MnbZh4';
  if (ID_SUMBER === 'GANTI_ID_SPREADSHEET_PILIH_RW') throw new Error('Ganti dulu ID_SUMBER');
  PropertiesService.getScriptProperties().setProperty('SOURCE_SS_ID', ID_SUMBER);
  Logger.log('ID spreadsheet sumber tersimpan.');
}

/**
 * Salin SEMUA token dari PILIH_RW (sheet "Pemilih", kolom Token) ke spreadsheet ini.
 * Aman dijalankan berulang (mis. ada pendaftar baru):
 *  - token baru ditambahkan
 *  - token yang sudah ada dilewati
 *  - kolom SudahMemilih TIDAK pernah ditimpa
 * Jangan jalankan ini saat pencoblosan sedang ramai-ramainya.
 */
function importPilihan() {
  var props = PropertiesService.getScriptProperties();
  var idSumber = props.getProperty('SOURCE_SS_ID');
  if (!idSumber) throw new Error('Jalankan setSumberPemilihan() dulu.');

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Server sibuk, coba lagi.');

  try {
    var sumber = SpreadsheetApp.openById(idSumber).getSheetByName('Pemilih');
    if (!sumber) throw new Error('Sheet "Pemilih" tidak ada di spreadsheet sumber.');

    var sv = sumber.getDataRange().getValues();
    var sh = sv[0].map(function (x) { return String(x).trim(); });
    var sToken = sh.indexOf('Token');
    if (sToken < 0) throw new Error('Kolom Token tidak ada di sumber.');

    var tujuan = getSS_().getSheetByName(CFG.SHEET_PEMILIH);
    if (!tujuan) throw new Error('Jalankan setup() dulu.');
    var tv = tujuan.getDataRange().getValues();

    // peta token -> nomor baris di sheet tujuan
    var peta = {};
    for (var i = 1; i < tv.length; i++) {
      var t = String(tv[i][0]).trim().toUpperCase();
      if (t) peta[t] = i + 1;
    }

    var baru = [], dilewati = 0;
    for (var r = 1; r < sv.length; r++) {
      var token = String(sv[r][sToken]).trim().toUpperCase();
      if (!token) continue;
      if (peta[token]) {
        dilewati++;
      } else {
        peta[token] = true;            // cegah token ganda di sumber
        baru.push([token, false]);
      }
    }
    if (baru.length) {
      tujuan.getRange(tujuan.getLastRow() + 1, 1, baru.length, 2).setValues(baru);
    }
    SpreadsheetApp.flush();
    Logger.log('Impor selesai. Token baru: ' + baru.length + ', sudah ada: ' + dilewati);
    try { SpreadsheetApp.getActive().toast('Baru: ' + baru.length + ', sudah ada: ' + dilewati, 'Impor selesai'); } catch (e) {}
  } finally {
    lock.releaseLock();
  }
}

/**
 * Isi PIN di bawah, jalankan 1x, lalu HAPUS angkanya dari kode.
 * PIN dipakai untuk halaman rekap/penghitungan.
 */
function setPinPanitia() {
  var PIN_BARU = 'GANTI_PIN_ANDA'; // <-- ganti, jalankan, lalu kembalikan ke teks ini
  if (PIN_BARU === 'GANTI_PIN_ANDA') throw new Error('Ganti dulu PIN_BARU');
  PropertiesService.getScriptProperties().setProperty('PIN_PANITIA', PIN_BARU);
  Logger.log('PIN panitia tersimpan.');
}

// ======================= 1. CEK TOKEN ================================
function cekToken_(token) {
  // sinkron=true: kalau token belum ada di sini, cek langsung ke PILIH_RW
  var cari = cariPemilih_(token, true);
  if (!cari.ok) return cari;

  return {
    ok: true,
    kandidat: daftarKandidat_()
  };
}

// ======================= 2. COBLOS ===================================
function coblos_(token, noKandidat) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'SERVER_SIBUK' };

  try {
    // validasi ulang di dalam lock (TANPA sinkron, agar tidak kunci ganda)
    var cari = cariPemilih_(token);
    if (!cari.ok) return cari;

    var kandidat = daftarKandidat_();
    var valid = kandidat.some(function (k) { return String(k.no) === String(noKandidat); });
    if (!valid) return { ok: false, error: 'KANDIDAT_TIDAK_VALID' };

    var ss = getSS_();
    var pem = ss.getSheetByName(CFG.SHEET_PEMILIH);
    var sua = ss.getSheetByName(CFG.SHEET_SUARA);

    // 1) kunci token dulu
    pem.getRange(cari.baris, cari.kolMemilih).setValue(true);
    SpreadsheetApp.flush();

    // 2) simpan suara; jika gagal, buka kunci kembali
    try {
      var kode = buatKodeUnik_(sua);
      var baris = [kode, Number(noKandidat), 'BELUM', ''];
      sisipAcak_(sua, baris);
      SpreadsheetApp.flush();
      return { ok: true, kode: kode };
    } catch (err) {
      pem.getRange(cari.baris, cari.kolMemilih).setValue(false);
      SpreadsheetApp.flush();
      return { ok: false, error: 'GAGAL_SIMPAN', detail: String(err) };
    }
  } finally {
    lock.releaseLock();
  }
}

// ======================= 3. REKAP: SCAN RESI =========================
function rekapScan_(kode, pin) {
  if (!cekPin_(pin)) return { ok: false, error: 'PIN_SALAH' };
  kode = String(kode || '').trim().toUpperCase();
  if (!kode) return { ok: false, error: 'KODE_KOSONG' };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'SERVER_SIBUK' };

  try {
    var sua = getSS_().getSheetByName(CFG.SHEET_SUARA);
    var data = sua.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === kode) {
        if (String(data[i][2]).trim().toUpperCase() === 'SUDAH') {
          return { ok: false, error: 'SUDAH_DIHITUNG' };
        }
        sua.getRange(i + 1, 3, 1, 2).setValues([['SUDAH', new Date()]]);
        SpreadsheetApp.flush();
        // sengaja TIDAK mengembalikan nomor kandidat ke layar scanner
        return { ok: true, pesan: 'Resi sah dan sudah dihitung', total: hitungDirekap_() };
      }
    }
    return { ok: false, error: 'KODE_TIDAK_VALID' };
  } finally {
    lock.releaseLock();
  }
}

// ======================= 4. HASIL REKAP ==============================
function getRekap_(pin) {
  if (!cekPin_(pin)) return { ok: false, error: 'PIN_SALAH' };

  var ss = getSS_();
  var kandidat = daftarKandidat_();
  var hitung = {};
  kandidat.forEach(function (k) { hitung[k.no] = 0; });

  // hitung suara dari resi yang sudah di-scan (StatusRekap = SUDAH)
  var sua = ss.getSheetByName(CFG.SHEET_SUARA).getDataRange().getValues();
  var totalSuara = 0, sudahDirekap = 0;
  for (var i = 1; i < sua.length; i++) {
    if (!String(sua[i][0]).trim()) continue;
    totalSuara++;
    if (String(sua[i][2]).trim().toUpperCase() === 'SUDAH') {
      sudahDirekap++;
      var no = sua[i][1];
      if (hitung[no] !== undefined) hitung[no]++;
    }
  }

  // statistik pemilih
  var pem = ss.getSheetByName(CFG.SHEET_PEMILIH);
  var pv = pem.getDataRange().getValues();
  var h = pv[0].map(function (x) { return String(x).trim(); });
  var iMemilih = h.indexOf('SudahMemilih');
  var totalPemilih = 0, sudahMemilih = 0;
  for (var r = 1; r < pv.length; r++) {
    if (!pv[r][h.indexOf('Token')]) continue;
    totalPemilih++;
    if (iMemilih >= 0 && pv[r][iMemilih] === true) sudahMemilih++;
  }

  return {
    ok: true,
    kandidat: kandidat.map(function (k) {
      return { no: k.no, nama: k.nama, foto: k.foto, suara: hitung[k.no] };
    }),
    totalPemilih: totalPemilih,
    sudahMemilih: sudahMemilih,
    resiTercetak: totalSuara,
    resiDihitung: sudahDirekap,
    resiBelumDihitung: totalSuara - sudahDirekap
  };
}

// ======================= HELPER ======================================
function getSS_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw new Error('Script tidak terikat spreadsheet. Isi Script Property SS_ID.');
  return SpreadsheetApp.openById(id);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Cari token (sinkron=true: jika belum ada, cek PILIH_RW lalu tambahkan).
 * Tidak menyimpan/mengembalikan identitas apa pun.
 */
function cariPemilih_(token, sinkron) {
  token = String(token || '').trim().toUpperCase();
  if (!token) return { ok: false, error: 'TOKEN_KOSONG' };

  var hasil = cariDiSheet_(token);
  if (hasil.error === 'TOKEN_TIDAK_DITEMUKAN' && sinkron) {
    if (!tambahDariSumber_(token)) return hasil;
    hasil = cariDiSheet_(token);
  }
  return hasil;
}

/** Cari token di sheet Pemilih spreadsheet ini + pastikan belum dipakai. */
function cariDiSheet_(token) {
  token = String(token || '').trim().toUpperCase();
  if (!token) return { ok: false, error: 'TOKEN_KOSONG' };

  var pem = getSS_().getSheetByName(CFG.SHEET_PEMILIH);
  var data = pem.getDataRange().getValues();
  var h = data[0].map(function (x) { return String(x).trim(); });

  var iToken = h.indexOf('Token');
  var iMemilih = h.indexOf('SudahMemilih');
  if (iToken < 0 || iMemilih < 0) return { ok: false, error: 'KOLOM_TIDAK_LENGKAP' };

  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iToken]).trim().toUpperCase() !== token) continue;

    if (data[r][iMemilih] === true || String(data[r][iMemilih]).toUpperCase() === 'TRUE') {
      return { ok: false, error: 'SUDAH_MEMILIH' };
    }
    return {
      ok: true,
      baris: r + 1,
      kolMemilih: iMemilih + 1
    };
  }
  return { ok: false, error: 'TOKEN_TIDAK_DITEMUKAN' };
}

/** Cek token di PILIH_RW; kalau ada, tambahkan ke sheet Pemilih ini. */
function tambahDariSumber_(token) {
  var idSumber = PropertiesService.getScriptProperties().getProperty('SOURCE_SS_ID');
  if (!idSumber) return false;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return false;
  try {
    // mungkin sudah ditambahkan oleh scan lain barusan
    if (cariDiSheet_(token).error !== 'TOKEN_TIDAK_DITEMUKAN') return true;

    var sumber = SpreadsheetApp.openById(idSumber).getSheetByName('Pemilih');
    if (!sumber) return false;
    var sv = sumber.getDataRange().getValues();
    var iToken = sv[0].map(function (x) { return String(x).trim(); }).indexOf('Token');
    if (iToken < 0) return false;

    for (var r = 1; r < sv.length; r++) {
      if (String(sv[r][iToken]).trim().toUpperCase() === token) {
        getSS_().getSheetByName(CFG.SHEET_PEMILIH).appendRow([token, false]);
        return true;
      }
    }
    return false;
  } finally {
    lock.releaseLock();
  }
}

function daftarKandidat_() {
  var sh = getSS_().getSheetByName(CFG.SHEET_KANDIDAT);
  if (!sh) throw new Error('Sheet Kandidat belum ada. Jalankan setup().');
  var data = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === '' || data[i][0] === null) continue;
    out.push({ no: data[i][0], nama: data[i][1], foto: data[i][2] || '' });
  }
  return out;
}

function buatKodeUnik_(sheetSuara) {
  var ada = {};
  var data = sheetSuara.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) ada[String(data[i][0]).toUpperCase()] = true;

  for (var coba = 0; coba < 50; coba++) {
    var kode = '';
    for (var j = 0; j < CFG.PANJANG_KODE; j++) {
      kode += CFG.ALFABET_KODE.charAt(Math.floor(Math.random() * CFG.ALFABET_KODE.length));
    }
    if (!ada[kode]) return kode;
  }
  throw new Error('Gagal membuat kode unik');
}

/** Sisipkan baris di posisi acak (bukan selalu paling bawah). */
function sisipAcak_(sheet, rowValues) {
  var last = sheet.getLastRow();            // termasuk header
  if (last <= 1) {
    sheet.appendRow(rowValues);
    return;
  }
  var posisi = 2 + Math.floor(Math.random() * last); // 2..last+1
  if (posisi > last) {
    sheet.appendRow(rowValues);
  } else {
    sheet.insertRowBefore(posisi);
    sheet.getRange(posisi, 1, 1, rowValues.length).setValues([rowValues]);
  }
}

function hitungDirekap_() {
  var data = getSS_().getSheetByName(CFG.SHEET_SUARA).getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim().toUpperCase() === 'SUDAH') n++;
  }
  return n;
}

function cekPin_(pin) {
  var simpan = PropertiesService.getScriptProperties().getProperty('PIN_PANITIA');
  return !!simpan && String(pin || '') === simpan;
}

/** Pastikan kolom-kolom ada di sheet; jika belum, tambahkan di kolom paling kanan. */
function kolom_(sheet, namaList) {
  var lastCol = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).trim(); });
  namaList.forEach(function (nama) {
    if (header.indexOf(nama) < 0) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(nama);
      header.push(nama);
    }
  });
}

// ======================= 5. AMBIL SUARA UNTUK HITUNG =================
// Hanya kode acak + nomor kandidat. Tanpa token, tanpa waktu, tanpa identitas.
function ambilSuara_(pin) {
  if (!cekPin_(pin)) return { ok: false, error: 'PIN_SALAH' };

  var ss = getSS_();
  var sua = ss.getSheetByName(CFG.SHEET_SUARA).getDataRange().getValues();
  var suara = [];
  for (var i = 1; i < sua.length; i++) {
    var kode = String(sua[i][0]).trim();
    if (!kode) continue;
    suara.push({ kode: kode, no: sua[i][1] });
  }

  var pv = ss.getSheetByName(CFG.SHEET_PEMILIH).getDataRange().getValues();
  var h = pv[0].map(function (x) { return String(x).trim(); });
  var iToken = h.indexOf('Token'), iMemilih = h.indexOf('SudahMemilih');
  var sudahMemilih = 0;
  for (var r = 1; r < pv.length; r++) {
    if (!pv[r][iToken]) continue;
    if (iMemilih >= 0 && pv[r][iMemilih] === true) sudahMemilih++;
  }

  return { ok: true, kandidat: daftarKandidat_(), suara: suara, sudahMemilih: sudahMemilih };
}
