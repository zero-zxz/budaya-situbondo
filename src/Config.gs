/**
 * ============================================================
 * Web Kebudayaan Situbondo - Google Apps Script Backend
 * Berkas: Config.gs
 * ============================================================
 */

var CONFIG = {
  // Masukkan Spreadsheet ID Google Sheets Anda di sini.
  // Jika script dijalankan dari dalam spreadsheet (Extensions > Apps Script),
  // kosongkan string ini dan sistem otomatis menggunakan active spreadsheet.
  SPREADSHEET_ID: '',

  // Nama tab di Google Sheets
  SHEET_KATALOG: 'Katalog_Budaya',
  SHEET_SEJARAH: 'Sejarah_Situbondo',
  SHEET_KATEGORI: 'Kategori',

  // Identitas & Kebanggaan Budaya Situbondo
  APP_TITLE: 'Warisan Budaya Situbondo - Pusaka Cagar Budaya, Tradisi, & Sejarah',
  APP_SUBTITLE: 'Inventarisasi Digital 17 Kecamatan, Seni Pandalungan, & Jejak Karesidenan Besuki',
  
  // Cache durasi detik (opsional)
  CACHE_DURATION_SEC: 300
};
