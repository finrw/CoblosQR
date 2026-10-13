/**
 * KONFIGURASI FRONTEND PENCOBLOSAN
 * Edit file ini saja, tidak perlu menyentuh index.html / rekap.html.
 */
window.APP_CONFIG = {
  // URL Web App hasil Deploy di Google Apps Script (berakhiran /exec)
  API_URL: "https://script.google.com/macros/s/AKfycbyc6SuSBtNsUjEZUQ09-MBJPQ75f6ONzWBGVuhiCmEsrq_u_nUCzbFXXgoOA4KwHJ4Ghg/exec",

  // Judul di bagian atas halaman dan di resi
  JUDUL: "Pemilihan RW 016",

  // Teks kecil di bawah QR pada resi
  KETERANGAN_RESI: "Masukkan resi ke kotak suara. Jangan difoto atau dibagikan.",

  // Setelah resi tampil, otomatis kembali ke scanner (detik). 0 = tidak otomatis.
  AUTO_KEMBALI_DETIK: 45,

  // Kolom ketik token manual di halaman scan.
  // true  = tampil (untuk masa uji / kamera bermasalah)
  // false = disembunyikan
  TAMPIL_INPUT_MANUAL: true
};

