# Web Warisan Budaya Situbondo (Google Apps Script)

Aplikasi web modern berbasis **Google Apps Script (HTML Service)** dengan antarmuka bertema warisan sejarah nusantara (heritage aesthetic), linimasa sejarah Situbondo berbasis literatur valid, serta filter berbasis 17 kecamatan se-Kabupaten Situbondo.

---

## Struktur Berkas Proyek

```
budaya-gas-web/
├── preview.html                  <- Pratinjau mandiri instan di browser lokal (klik ganda)
├── PRD.md                        <- Dokumen spesifikasi kebutuhan produk
├── README.md                     <- Panduan instalasi dan deployment ini
├── appsscript.json               <- Manifest runtime V8 & konfigurasi Google Apps Script
├── .clasp.json                   <- Konfigurasi deployment clasp (opsional)
├── src/                          <- Berkas sumber Google Apps Script
│   ├── appsscript.json           <- Salinan manifest untuk direktori root clasp
│   ├── Config.gs                 <- Konfigurasi ID spreadsheet & metadata Situbondo
│   ├── Database.gs               <- Data Access Layer, linimasa sejarah, & 17 kecamatan
│   ├── Api.gs                    <- Endpoint API terbungkus aman untuk google.script.run
│   ├── Code.gs                   <- Entry point web app (doGet) & injeksi server data
│   ├── Index.html                <- Template utama HTML (Katalog, Linimasa, Pandalungan)
│   ├── Styles.html               <- CSS modular bertema heritage (Anti-Slop, zero emoji)
│   └── Scripts.html              <- Logika interaktif klien, pencarian, filter kecamatan, modal
└── data/                         <- Data referensi & skema database
    ├── SKEMA_GOOGLE_SHEETS.md    <- Panduan struktur kolom & inisialisasi sheet otomatis
    ├── katalog_budaya.json       <- 43 entri warisan budaya format JSON (lengkap foto 1:1 tiap objek)
    ├── katalog_budaya.csv        <- Berkas CSV siap impor langsung ke Google Sheets (43 objek)
    ├── sejarah_situbondo.json    <- 12 tonggak sejarah Situbondo dengan literatur valid
    └── foto/                     <- 43 berkas foto asli cagar budaya, kesenian, & tradisi (zero duplikasi)
```

---

## 1. Fitur Khusus untuk Masyarakat Situbondo

1. **Linimasa Kronik Sejarah (Megalitikum hingga Modern)**:
   - Disusun berdasarkan literasi kuat: *Kitab Nagarakretagama* (Mpu Prapanca 1365 M), catatan *Suma Oriental* (Tomé Pires 1512 M), *Babad Besuki* (Kyai Pate Alos), arsip *Residentie Besoeki*, *De Grote Postweg* Daendels Anyer-Panarukan (1808 M), industri gula lokomotif uap PG Olean (1846 M), *Harjakasi 15 Agustus*, perjuangan Pahlawan Nasional *KHR. As'ad Syamsul Arifin* (Sukorejo Banyuputih), hingga *PP No. 28 Tahun 1972* yang melegitimasi nama Kabupaten Situbondo.
2. **Filter 17 Kecamatan se-Kabupaten Situbondo**:
   - Warga dari Banyuglugur, Jatibanteng, Besuki, Suboh, Mlandingan, Sumbermalang, Bungatan, Kendit, Panarukan, Situbondo Kota, Mangaran, Panji, Kapongan, Arjasa, Jangkar, Asembagus, hingga Banyuputih dapat menyaring warisan di tanah kelahiran masing-masing dalam satu klik.
3. **Harmoni Budaya Pandalungan & Kota Santri Pancasila**:
   - Memadukan narasi asimilasi kultural Madura-Jawa pesisir dengan sejarah Muktamar NU Sukorejo 1984 yang menegaskan Situbondo sebagai tonggak Kota Santri Pancasila.
4. **Bebas Emoji 100% & Full Vektor SVG**:
   - Menggunakan 290+ ikon SVG vektor presisi tinggi tanpa satupun karakter emoji generik.

---

## 2. Pratinjau Cepat di Komputer Lokal

Anda dapat menguji seluruh fungsi web secara offline di komputer Anda:
- Klik ganda berkas **`preview.html`**, atau
- Buka melalui browser: `file:///D:/Coding/Apps/budaya-gas-web/preview.html`

Seluruh fitur navigasi tab (Katalog Warisan, Linimasa Sejarah, 17 Kecamatan), pencarian kata kunci, filter kecamatan, dan modal rincian aktif sepenuhnya.

---

## 3. Cara Deploy ke Google Apps Script (Metode Web Editor)

### Langkah 1: Buat Proyek Apps Script
1. Buka [Google Apps Script Home](https://script.google.com/) atau buat langsung dari Google Spreadsheet baru via menu **Extensions > Apps Script**.
2. Beri nama proyek: `Web Kebudayaan Situbondo`.

### Langkah 2: Salin Berkas Sumber (`src/`)
Pada panel editor Apps Script, buat berkas sesuai struktur berikut:
1. Buat berkas Script (`.gs`):
   - `Config.gs`: Salin seluruh isi dari `src/Config.gs`.
   - `Database.gs`: Salin seluruh isi dari `src/Database.gs`.
   - `Api.gs`: Salin seluruh isi dari `src/Api.gs`.
   - `Code.gs`: Salin seluruh isi dari `src/Code.gs`.
2. Buat berkas HTML (`.html`):
   - `Index.html`: Salin isi dari `src/Index.html`.
   - `Styles.html`: Salin isi dari `src/Styles.html`.
   - `Scripts.html`: Salin isi dari `src/Scripts.html`.

### Langkah 3: Inisialisasi Otomatis Google Sheets
1. Buat Google Spreadsheet kosong baru di Google Drive.
2. Salin Spreadsheet ID dari tautan browser.
3. Tempelkan ID tersebut ke `src/Config.gs` pada variabel `SPREADSHEET_ID`.
4. Pada dropdown fungsi toolbar editor Apps Script, pilih **`initSheets`**, lalu klik **Run (Jalankan)**.
5. Berikan izin saat jendela otorisasi muncul.
6. Dua tab utama (`Katalog_Budaya` dan `Sejarah_Situbondo`) akan otomatis terbentuk dan terisi data lengkap!

### Langkah 4: Publikasikan Sebagai Web App
1. Klik tombol **Deploy > New deployment**.
2. Pilih tipe **Web app**.
3. Atur konfigurasi:
   - **Execute as**: `Me (email-anda@gmail.com)`
   - **Who has access**: `Anyone` (Siapa saja agar dapat diakses oleh seluruh warga masyarakat Situbondo tanpa harus login)
4. Klik **Deploy** dan salin tautan Web App URL untuk disebarluaskan.
