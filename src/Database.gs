/**
 * ============================================================
 * Web Kebudayaan Situbondo - Google Apps Script Backend
 * Berkas: Database.gs
 * Data Access Layer, Linimasa Sejarah, & Inventaris 17 Kecamatan
 * ============================================================
 */

function getSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID && CONFIG.SPREADSHEET_ID.trim() !== '') {
    try {
      return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID.trim());
    } catch (err) {
      Logger.log('[Database] Gagal openById: ' + err.message + '. Beralih ke active spreadsheet.');
    }
  }
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (err) {
    Logger.log('[Database] Tidak ada active spreadsheet: ' + err.message);
    return null;
  }
}

function getKecamatanList() {
  return [
    'Banyuglugur',
    'Jatibanteng',
    'Besuki',
    'Suboh',
    'Mlandingan',
    'Sumbermalang',
    'Bungatan',
    'Kendit',
    'Panarukan',
    'Situbondo',
    'Mangaran',
    'Panji',
    'Kapongan',
    'Arjasa',
    'Jangkar',
    'Asembagus',
    'Banyuputih'
  ];
}

function getAllCultureItems() {
  var ss = getSpreadsheet();
  if (!ss) return getSeedData();

  var sheet = ss.getSheetByName(CONFIG.SHEET_KATALOG);
  if (!sheet) return getSeedData();

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return getSeedData();

  var headers = data[0].map(function(h) {
    return String(h).trim().toLowerCase();
  });

  var items = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1]) continue;

    var item = {};
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j];
      var val = row[j];
      if (key === 'unggulan') {
        item[key] = (val === true || String(val).toUpperCase() === 'TRUE');
      } else {
        item[key] = (val !== null && val !== undefined) ? String(val).trim() : '';
      }
    }
    items.push(item);
  }

  return items.length > 0 ? items : getSeedData();
}

function getTimelineMilestones() {
  var ss = getSpreadsheet();
  if (!ss) return getSeedTimeline();

  var sheet = ss.getSheetByName(CONFIG.SHEET_SEJARAH || 'Sejarah_Situbondo');
  if (!sheet) return getSeedTimeline();

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return getSeedTimeline();

  var headers = data[0].map(function(h) {
    return String(h).trim().toLowerCase();
  });

  var timeline = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] && !row[1]) continue;

    var item = {};
    for (var j = 0; j < headers.length; j++) {
      var key = headers[j];
      var val = row[j];
      item[key] = (val !== null && val !== undefined) ? String(val).trim() : '';
    }
    timeline.push(item);
  }

  return timeline.length > 0 ? timeline : getSeedTimeline();
}

function getCultureItemById(id) {
  var items = getAllCultureItems();
  for (var i = 0; i < items.length; i++) {
    if (items[i].id === id) return items[i];
  }
  return null;
}

function filterCultureItems(query, category, kecamatan) {
  var items = getAllCultureItems();
  var q = query ? String(query).toLowerCase().trim() : '';
  var cat = category ? String(category).trim() : '';
  var kec = kecamatan ? String(kecamatan).trim() : '';

  return items.filter(function(item) {
    var matchCat = !cat || cat === 'Semua' || (item.kategori && item.kategori.toLowerCase() === cat.toLowerCase());
    var matchKec = !kec || kec === 'Semua' || (item.kecamatan && item.kecamatan.toLowerCase() === kec.toLowerCase());
    var matchQuery = !q || (
      (item.nama && item.nama.toLowerCase().indexOf(q) !== -1) ||
      (item.lokasi && item.lokasi.toLowerCase().indexOf(q) !== -1) ||
      (item.kecamatan && item.kecamatan.toLowerCase().indexOf(q) !== -1) ||
      (item.jenis && item.jenis.toLowerCase().indexOf(q) !== -1) ||
      (item.deskripsi && item.deskripsi.toLowerCase().indexOf(q) !== -1)
    );
    return matchCat && matchKec && matchQuery;
  });
}

function initSheets() {
  var ss = getSpreadsheet();
  if (!ss) {
    throw new Error('Spreadsheet belum terhubung. Harap isi CONFIG.SPREADSHEET_ID terlebih dahulu.');
  }

  // 1. Tab Katalog_Budaya
  var sheetKatalog = ss.getSheetByName(CONFIG.SHEET_KATALOG);
  if (!sheetKatalog) sheetKatalog = ss.insertSheet(CONFIG.SHEET_KATALOG);

  var headersKatalog = [
    'id', 'nama', 'kategori', 'jenis', 'lokasi', 'kecamatan', 'tahun',
    'unggulan', 'deskripsi', 'status', 'sumber', 'detail_tambahan', 'foto', 'foto_url'
  ];

  sheetKatalog.clear();
  sheetKatalog.appendRow(headersKatalog);
  var headerRange1 = sheetKatalog.getRange(1, 1, 1, headersKatalog.length);
  headerRange1.setFontWeight('bold');
  headerRange1.setBackground('#e7e4df');

  var seedsKatalog = getSeedData();
  var rowsKatalog = seedsKatalog.map(function(item) {
    return [
      item.id,
      item.nama,
      item.kategori,
      item.jenis,
      item.lokasi,
      item.kecamatan || '',
      item.tahun,
      item.unggulan ? 'TRUE' : 'FALSE',
      item.deskripsi,
      item.status,
      item.sumber,
      item.detail_tambahan,
      item.foto || '',
      item.foto_url || ''
    ];
  });
  if (rowsKatalog.length > 0) {
    sheetKatalog.getRange(2, 1, rowsKatalog.length, headersKatalog.length).setValues(rowsKatalog);
  }
  sheetKatalog.autoResizeColumns(1, headersKatalog.length);

  // 2. Tab Sejarah_Situbondo
  var sheetSejarah = ss.getSheetByName(CONFIG.SHEET_SEJARAH || 'Sejarah_Situbondo');
  if (!sheetSejarah) sheetSejarah = ss.insertSheet(CONFIG.SHEET_SEJARAH || 'Sejarah_Situbondo');

  var headersSejarah = ['id', 'era', 'tahun', 'judul', 'ringkasan', 'uraian', 'sumber', 'lokasi', 'foto'];
  sheetSejarah.clear();
  sheetSejarah.appendRow(headersSejarah);
  var headerRange2 = sheetSejarah.getRange(1, 1, 1, headersSejarah.length);
  headerRange2.setFontWeight('bold');
  headerRange2.setBackground('#e7e4df');

  var seedsSejarah = getSeedTimeline();
  var rowsSejarah = seedsSejarah.map(function(t) {
    return [t.id, t.era, t.tahun, t.judul, t.ringkasan, t.uraian, t.sumber, t.lokasi, t.foto || ''];
  });
  if (rowsSejarah.length > 0) {
    sheetSejarah.getRange(2, 1, rowsSejarah.length, headersSejarah.length).setValues(rowsSejarah);
  }
  sheetSejarah.autoResizeColumns(1, headersSejarah.length);

  Logger.log('[Database] Inisialisasi ganda berhasil: ' + rowsKatalog.length + ' katalog dan ' + rowsSejarah.length + ' babak sejarah.');
  return 'Inisialisasi berhasil: ' + rowsKatalog.length + ' entri warisan dan ' + rowsSejarah.length + ' tonggak sejarah ditambahkan.';
}

/**
 * Dataset Warisan Budaya Situbondo: Lengkap 17 Kecamatan & Bebas Foto Duplikat
 */
function getSeedData() {
  return [
    {
      id: "batu-silindris-banyuglugur",
      nama: "Batu Silindris Banyuglugur",
      kategori: "Cagar Budaya",
      jenis: "Benda Megalitik",
      lokasi: "Desa Banyuglugur",
      kecamatan: "Banyuglugur",
      tahun: "Zaman Prasejarah",
      unggulan: false,
      deskripsi: "Tinggalan batu silindris purbakala berbahan andesit alami yang ditemukan di wilayah pesisir barat Situbondo, menjadi bukti aktivitas manusia megalitikum di pintu masuk barat kabupaten.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 100.3.3.2/185/431.013/2024)",
      sumber: "Naskah Kajian TACB Kabupaten Situbondo",
      detail_tambahan: "SK Bupati 2024 | Bahan: Andesit Purbakala",
      foto: "batu-silindris-banyuglugur.jpg",
      foto_url: "data/foto/batu-silindris-banyuglugur.jpg"
    },
    {
      id: "mangrove-dubibir-banyuglugur",
      nama: "Kawasan Mangrove Pantai Dubibir",
      kategori: "Tradisi",
      jenis: "Tradisi & Lingkungan Pesisir",
      lokasi: "Pantai Dubibir, Desa Banyuglugur",
      kecamatan: "Banyuglugur",
      tahun: "Turun-temurun",
      unggulan: true,
      deskripsi: "Kawasan konservasi mangrove dan tradisi penjagaan sempadan pantai oleh warga pesisir Banyuglugur untuk menahan abrasi gelombang Selat Madura dan melestarikan habitat biota laut.",
      status: "Kawasan Konservasi & Tradisi Pesisir",
      sumber: "Dokumentasi Komunitas Peduli Pesisir Situbondo",
      detail_tambahan: "Fungsi: Penahan Abrasi & Perlindungan Habitat Bahari",
      foto: "pita-mangrove.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/1/14/Mangrove_Pantai_Dubibir_Situbondo.jpg/800px-Mangrove_Pantai_Dubibir_Situbondo.jpg"
    },
    {
      id: "pantai-utama-raya-banyuglugur",
      nama: "Pantai Utama Raya Banyuglugur",
      kategori: "Sejarah",
      jenis: "Geografi Bersejarah Pesisir",
      lokasi: "Desa Banyuglugur",
      kecamatan: "Banyuglugur",
      tahun: "Kontemporer",
      unggulan: false,
      deskripsi: "Pesisir pantai utara di gerbang barat Situbondo yang berbatasan langsung dengan Probolinggo, menjadi simpul persinggahan jalur Daendels dan rest area bahari terpadu.",
      status: "Kawasan Bahari Gerbang Barat",
      sumber: "Dinas Pariwisata Situbondo",
      detail_tambahan: "Lokasi: Jalur Pantura Gerbang Barat",
      foto: "pita-utama-raya.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/1/1a/Pantai_Utama_Raya_Situbondo.jpg/800px-Pantai_Utama_Raya_Situbondo.jpg"
    },
    {
      id: "gelang-renteng-jatibanteng",
      nama: "Gelang Renteng Jatibanteng",
      kategori: "Cagar Budaya",
      jenis: "Benda Logam Megalitik",
      lokasi: "Kawasan Perbukitan Jatibanteng",
      kecamatan: "Jatibanteng",
      tahun: "Era Perunggu - Besi",
      unggulan: false,
      deskripsi: "Rangkaian perhiasan gelang logam perunggu purbakala yang ditemukan dalam bekal kubur batu di wilayah Jatibanteng, membuktikan tingginya keahlian metalurgi masyarakat kuno.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 188/317/P/001.3/2023)",
      sumber: "Naskah Kajian TACB Kabupaten Situbondo",
      detail_tambahan: "Bahan: Perunggu Purbakala | SK Bupati 2023",
      foto: "gelang-renteng-jatibanteng.jpg",
      foto_url: "data/foto/gelang-renteng-jatibanteng.jpg"
    },
    {
      id: "selamatan-bumi-baritan",
      nama: "Ritus Baritan & Selamatan Desa",
      kategori: "Tradisi",
      jenis: "Ritus Tolak Bala Agraris",
      lokasi: "Desa-desa di Jatibanteng",
      kecamatan: "Jatibanteng",
      tahun: "Turun-temurun",
      unggulan: false,
      deskripsi: "Tradisi selamatan tolak bala di perempatan jalan desa atau sumber air dengan membawa takir daun pisang berisi nasi tumpeng mini untuk dinikmati bersama oleh warga tani.",
      status: "Tercatat Tradisi Agraris Kerakyatan",
      sumber: "Himpunan Kebudayaan Desa Situbondo",
      detail_tambahan: "Makna: Kerukunan dan Syukur Musim Tanam",
      foto: "baritan-jatibanteng.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/1/18/Hasil_bumi_untuk_di_sedekah_kan.jpg/800px-Hasil_bumi_untuk_di_sedekah_kan.jpg"
    },
    {
      id: "rumah-residen-besuki",
      nama: "Rumah Dinas Residen Besuki",
      kategori: "Cagar Budaya",
      jenis: "Bangunan Kolonial",
      lokasi: "Jl. Madura, Desa Besuki",
      kecamatan: "Besuki",
      tahun: "Abad ke-19",
      unggulan: true,
      deskripsi: "Pusat kediaman dan kantor administrasi pejabat tertinggi Residen Besuki semasa Hindia Belanda. Arsitektur bergaya Indische Empire dengan pilar-pilar kokoh yang menjadi saksi pusat pemerintahan kawasan timur Jawa Timur.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 188/392/P/004.2/2020)",
      sumber: "Naskah Kajian TACB Situbondo & Register Dapobud",
      detail_tambahan: "Gaya Arsitektur: Indische Empire Style",
      foto: "pendopo-besuki.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/1/14/Pendopo_Pate_Alos_Besuki_Situbondo.jpg/800px-Pendopo_Pate_Alos_Besuki_Situbondo.jpg"
    },
    {
      id: "kantor-kawedanan-besuki",
      nama: "Kantor Kawedanan Besuki",
      kategori: "Cagar Budaya",
      jenis: "Bangunan Pemerintahan",
      lokasi: "Desa Besuki",
      kecamatan: "Besuki",
      tahun: "Abad ke-19",
      unggulan: false,
      deskripsi: "Gedung administrasi pembantu bupati wilayah Besuki pada masa lampau, memiliki beranda depan luas dan ventilasi lengkung khas bangunan tropis kolonial.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 188/392/P/004.2/2020)",
      sumber: "Naskah Kajian TACB Situbondo",
      detail_tambahan: "SK Bupati 2020 | Bangunan Bersejarah",
      foto: "kawedanan-besuki.jpg",
      foto_url: "data/foto/kawedanan-besuki.jpg"
    },
    {
      id: "klenteng-poo-tong-bio",
      nama: "Klenteng Poo Tong Bio Besuki",
      kategori: "Cagar Budaya",
      jenis: "Tempat Peribadatan Kuno",
      lokasi: "Pecinan Besuki",
      kecamatan: "Besuki",
      tahun: "Abad ke-18",
      unggulan: true,
      deskripsi: "Klenteng tertua di wilayah Besuki yang didirikan komunitas pedagang Tionghoa pesisir. Memiliki ornamen ukiran naga, altar klasik, serta papan pembakar dupa kuno lambang kerukunan antarwarga.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 188/443/P/004.2/2021)",
      sumber: "Naskah Kajian TACB Situbondo",
      detail_tambahan: "SK Bupati 2021 | Simbol Multikultural Besuki",
      foto: "klenteng-poo-tong-bio.jpg",
      foto_url: "data/foto/klenteng-poo-tong-bio.jpg"
    },
    {
      id: "ngalos-pate-alos",
      nama: "Tradisi Ziarah Pasarean Kyai Pate Alos",
      kategori: "Tradisi",
      jenis: "Ritus Ziarah Leluhur",
      lokasi: "Makam Pate Alos, Besuki",
      kecamatan: "Besuki",
      tahun: "Ritus Tahunan",
      unggulan: false,
      deskripsi: "Tradisi berziarah ke makam pendiri Besuki dan bupati terdahulu (Raden Bagus Kasim / Kyai Pate Alos) setiap menjelang Hari Jadi Kabupaten dan bulan Ruwah.",
      status: "Tradisi Religi Masyarakat Besuki",
      sumber: "Dinas Pendidikan dan Kebudayaan Situbondo",
      detail_tambahan: "Tokoh: Raden Bagus Kasim (Kyai Pate Alos)",
      foto: "pasarean-pate-alos.jpg",
      foto_url: "data/foto/pasarean-pate-alos.jpg"
    },
    {
      id: "situs-cemara",
      nama: "Situs Sarkofagus Cemara",
      kategori: "Cagar Budaya",
      jenis: "Situs Megalitik",
      lokasi: "Desa Cemara",
      kecamatan: "Suboh",
      tahun: "± 1500 SM",
      unggulan: true,
      deskripsi: "Sarkofagus batu monolitik berukuran panjang 220 cm, lebar 120 cm, dan tebal 50 cm. Berfungsi sebagai peti kubur batu penghormatan pemimpin suku prasejarah di perbukitan Suboh.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 100.3.3.2/185/431.013/2024)",
      sumber: "TACB Situbondo & BPK Wilayah XI",
      detail_tambahan: "Dimensi: 220 x 120 x 50 cm | Batu Andesit",
      foto: "situs-cemara-suboh.jpg",
      foto_url: "data/foto/situs-cemara-suboh.jpg"
    },
    {
      id: "tari-glipang-suboh",
      nama: "Seni Tari Glipang Suboh",
      kategori: "Kesenian",
      jenis: "Seni Tari Rakyat",
      lokasi: "Kecamatan Suboh",
      kecamatan: "Suboh",
      tahun: "Tradisi Lisan",
      unggulan: false,
      deskripsi: "Kesenian tari keprajuritan rakyat berakar tradisi Pandalungan dengan gerak rancak gagah, hentakan kaki berirama kendang, dan syair pujian moral.",
      status: "Tercatat Kesenian Tradisional Suboh",
      sumber: "Paguyuban Seni Pandalungan Situbondo",
      detail_tambahan: "Karakter: Tari Keprajuritan Rakyat",
      foto: "tari-glipang-suboh.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/f/f1/Topeng_Dalang_-_Sumenep.jpg/800px-Topeng_Dalang_-_Sumenep.jpg"
    },
    {
      id: "menara-pandang-mlandingan",
      nama: "Fondasi Menara Pandang Mlandingan",
      kategori: "Cagar Budaya",
      jenis: "Struktur Pengawasan Laut",
      lokasi: "Pesisir Mlandingan",
      kecamatan: "Mlandingan",
      tahun: "Abad ke-19",
      unggulan: false,
      deskripsi: "Struktur batu bata dan andesit fondasi menara pantau mercusuar masa Hindia Belanda untuk memandu perahu nelayan dan kapal niaga di Selat Madura.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 188/443/P/004.2/2021)",
      sumber: "Naskah Kajian TACB Situbondo",
      detail_tambahan: "SK Bupati 2021 | Lokasi Pesisir Mlandingan",
      foto: "menara-lonceng-kuno.jpg",
      foto_url: "data/foto/menara-lonceng-kuno.jpg"
    },
    {
      id: "petik-laut-mlandingan",
      nama: "Tradisi Selamatan Laut Mlandingan",
      kategori: "Tradisi",
      jenis: "Ritus Bahari Nelayan",
      lokasi: "Desa Selomukti, Mlandingan",
      kecamatan: "Mlandingan",
      tahun: "Turun-temurun",
      unggulan: false,
      deskripsi: "Upacara tahunan doa bersama dan larung tumpeng hasil bumi oleh warga nelayan Mlandingan sebagai ungkapan syukur atas keselamatan melaut.",
      status: "Tradisi Bahari Pesisir",
      sumber: "Rukun Nelayan Mlandingan",
      detail_tambahan: "Peserta: Warga Nelayan Selat Madura",
      foto: "petik-laut-mlandingan.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d9/Petik_Laut_Pantai_Darmo_Probolinggo.jpg/730px-Petik_Laut_Pantai_Darmo_Probolinggo.jpg"
    },
    {
      id: "struktur-mellek-sumbermalang",
      nama: "Struktur Batur & Pagar Mellek",
      kategori: "Cagar Budaya",
      jenis: "Struktur Punden Purbakala",
      lokasi: "Desa Sumbermalang, Lereng Argopuro",
      kecamatan: "Sumbermalang",
      tahun: "Zaman Klasik Purbakala",
      unggulan: false,
      deskripsi: "Susunan batu berundak purbakala di ketinggian lereng Gunung Argopuro yang dahulu digunakan sebagai tempat pemujaan dan pertapaan resi masa lampau.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 100.3.3.2/185/431.013/2024)",
      sumber: "TACB Situbondo",
      detail_tambahan: "SK Bupati 2024 | Kawasan Pegunungan Argopuro",
      foto: "struktur-mellek.jpg",
      foto_url: "data/foto/struktur-mellek.jpg"
    },
    {
      id: "kethek-ogleng-situbondo",
      nama: "Kesenian Kethek Ogleng Sumbermalang",
      kategori: "Kesenian",
      jenis: "Seni Tari Rakyat",
      lokasi: "Kecamatan Sumbermalang",
      kecamatan: "Sumbermalang",
      tahun: "Tradisi Lisan",
      unggulan: false,
      deskripsi: "Seni pertunjukan tari rakyat dengan penari bertopeng monyet putih lincah yang menirukan gerakan kera hutan lereng Argopuro, diiringi gamelan rancak dan humor kerakyatan.",
      status: "Tercatat Kesenian Rakyat Tradisional",
      sumber: "Paguyuban Seniman Situbondo",
      detail_tambahan: "Unsur Seni: Akrobatik, Tari Topeng, Teater Rakyat",
      foto: "kethek-ogleng.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d5/Kethek_Ogleng_Kesenian_berasal_dari_Sampung_Ponorogo_a.jpg/676px-Kethek_Ogleng_Kesenian_berasal_dari_Sampung_Ponorogo_a.jpg"
    },
    {
      id: "sarkofagus-bungatan",
      nama: "Kompleks 26 Sarkofagus Selowogo",
      kategori: "Cagar Budaya",
      jenis: "Benda Purbakala",
      lokasi: "Desa Selowogo",
      kecamatan: "Bungatan",
      tahun: "Zaman Megalithikum",
      unggulan: true,
      deskripsi: "Klaster makam batu megalitik terbanyak di Jawa Timur yang tersebar di perkebunan warga Bungatan. Terdiri atas lebih dari 26 unit sarkofagus dengan pahatan purba.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 188/443/P/004.2/2021)",
      sumber: "Daftar 26 Objek Sarkofagus SK Bupati 2021",
      detail_tambahan: "Terdiri dari Sarkofagus 1 s.d. 26",
      foto: "sarkofagus-selowogo.jpg",
      foto_url: "data/foto/sarkofagus-selowogo.jpg"
    },
    {
      id: "batik-situbondo",
      nama: "Batik Khas Pesisir Motif Kerang",
      kategori: "Sejarah",
      jenis: "Kriya Tekstil Tradisi",
      lokasi: "Sentra Selowogo",
      kecamatan: "Bungatan",
      tahun: "1994",
      unggulan: true,
      deskripsi: "Kain batik khas Situbondo yang terinspirasi oleh kekayaan laut utara Jawa. Memiliki ragam hias khas biota bahari seperti kerang mahkota, karang rengganis, dan daun tembakau.",
      status: "Warisan Kriya Khas Kabupaten Situbondo",
      sumber: "Asosiasi Pengrajin Batik Situbondo",
      detail_tambahan: "Ciri Khas: Motif Kerang & Biota Bahari",
      foto: "batik-situbondo-kerang.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e1/Batik_Madura_16.jpg/800px-Batik_Madura_16.jpg"
    },
    {
      id: "pantai-pasir-putih-bungatan",
      nama: "Pantai Pasir Putih Situbondo",
      kategori: "Tradisi",
      jenis: "Kawasan Tradisi Bahari",
      lokasi: "Desa Pasir Putih",
      kecamatan: "Bungatan",
      tahun: "Turun-temurun",
      unggulan: true,
      deskripsi: "Ikon bahari kebanggaan masyarakat Situbondo berbentuk teluk melengkung yang tenang. Menjadi lokasi tradisi lomba perahu layar tradisional dan pesta petik laut tahunan.",
      status: "Ikon Bahari Kebanggaan Situbondo",
      sumber: "Dinas Pariwisata Situbondo",
      detail_tambahan: "Pusat Parade Perahu Layar Tradisional",
      foto: "hero-pasir-putih.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/5/5f/Perahu_Layar_di_Pantai_Pasir_Putih_Situbondo.jpg/800px-Perahu_Layar_di_Pantai_Pasir_Putih_Situbondo.jpg"
    },
    {
      id: "kesenian-ojhung",
      nama: "Atraksi Adat Ojhung Bugeman",
      kategori: "Kesenian",
      jenis: "Seni Bela Diri Tradisi",
      lokasi: "Desa Bugeman",
      kecamatan: "Kendit",
      tahun: "Turun-temurun",
      unggulan: true,
      deskripsi: "Tradisi adu ketangkasan rotan antara dua pendekar bertelanjang dada dengan pelindung kepala dari sabut kelapa. Digelar saat selamatan desa memohon hujan dan tolak bala.",
      status: "Warisan Budaya Takbenda (WBTb) Daerah",
      sumber: "Dinas Pariwisata dan Kebudayaan Situbondo",
      detail_tambahan: "Lokasi Sakral: Desa Bugeman, Kendit",
      foto: "kesenian-ojhung.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/6/67/Kesenian_Ojung.jpg/800px-Kesenian_Ojung.jpg"
    },
    {
      id: "kampung-kerapu-kendit",
      nama: "Kawasan Pesisir Kampung Kerapu",
      kategori: "Tradisi",
      jenis: "Tradisi Nelayan Budidaya",
      lokasi: "Dusun Gundil, Desa Klatakan",
      kecamatan: "Kendit",
      tahun: "Turun-temurun",
      unggulan: false,
      deskripsi: "Sentra budidaya ikan kerapu cantang dan kerapu tikus unggulan di pesisir Kendit, dilengkapi dermaga kayu apung yang membentang ke perairan Selat Madura.",
      status: "Kawasan Bahari dan Budidaya Khas Situbondo",
      sumber: "Himpunan Nelayan Budidaya Situbondo",
      detail_tambahan: "Sentra Budidaya Kerapu Nasional",
      foto: "pita-kerapu.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/c/cf/Kampung_Kerapu_Situbondo.jpg/800px-Kampung_Kerapu_Situbondo.jpg"
    },
    {
      id: "titik-nol-daendels",
      nama: "Titik Nol Jalan Pos Daendels Panarukan",
      kategori: "Sejarah",
      jenis: "Monumen Sejarah Nasional",
      lokasi: "Kawasan Pelabuhan Panarukan",
      kecamatan: "Panarukan",
      tahun: "1808",
      unggulan: true,
      deskripsi: "Ujung akhir pembangunan monumental Jalan Raya Pos (De Grote Postweg) sepanjang 1.000 km dari Anyer ke Panarukan oleh Daendels. Menjadikan Panarukan dermaga ekspor rempah, gula, dan kopi terkemuka.",
      status: "Monumen Sejarah Nasional",
      sumber: "Arsip Sejarah Daendels & Pemkab Situbondo",
      detail_tambahan: "Panjang Lintasan: 1.000 Kilometer dari Anyer",
      foto: "titik-nol-panarukan.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/a/a9/Dermaga_Panarukan.jpg/800px-Dermaga_Panarukan.jpg"
    },
    {
      id: "benteng-panarukan",
      nama: "Benteng & Pos Jaga VOC Panarukan",
      kategori: "Cagar Budaya",
      jenis: "Benteng Pertahanan",
      lokasi: "Pelabuhan Panarukan",
      kecamatan: "Panarukan",
      tahun: "Abad ke-18",
      unggulan: true,
      deskripsi: "Tinggalan pos pertahanan militer dan pengawasan pelabuhan Belanda yang mengontrol arus ekspor komoditas hasil bumi di Selat Madura dan Laut Jawa.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 188/443/P/004.2/2021)",
      sumber: "Naskah Kajian TACB Situbondo",
      detail_tambahan: "Fungsi Lama: Pengawas Pelabuhan Daendels",
      foto: "benteng-panarukan.jpg",
      foto_url: "data/foto/benteng-panarukan.jpg"
    },
    {
      id: "gereja-maria-bintang-samudera",
      nama: "Menara Gereja Maria Bintang Samudera",
      kategori: "Cagar Budaya",
      jenis: "Bangunan Religi Bersejarah",
      lokasi: "Jl. Pelabuhan Panarukan",
      kecamatan: "Panarukan",
      tahun: "Awal Abad ke-20",
      unggulan: false,
      deskripsi: "Bangunan menara dan struktur lonceng peribadatan kuno yang melayani komunitas pelaut dan warga pesisir Panarukan sejak masa kolonial.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 100.3.3.2/170/431.013/2025)",
      sumber: "SK Bupati Situbondo Tahun 2025",
      detail_tambahan: "SK Penetapan Terbaru 2025",
      foto: "menara-gereja-maria-panarukan.jpg",
      foto_url: "data/foto/menara-gereja-maria-panarukan.jpg"
    },
    {
      id: "pg-olean",
      nama: "Kompleks Pabrik Gula Olean",
      kategori: "Cagar Budaya",
      jenis: "Situs Industri Sejarah",
      lokasi: "Desa Olean",
      kecamatan: "Situbondo",
      tahun: "1846",
      unggulan: true,
      deskripsi: "Pabrik gula tertua yang didirikan tahun 1846 dan menjadi living museum industri gula dunia. Hingga kini masih mengoperasikan lokomotif uap kuno berbahan bakar ampas tebu di jalur lori perkebunan.",
      status: "Terdata Warisan Industri Bersejarah",
      sumber: "Dinas Pendidikan dan Kebudayaan Situbondo",
      detail_tambahan: "Tahun Berdiri: 1846 | Lokomotif Uap Aktif",
      foto: "pg-olean.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/e/e6/COLLECTIE_TROPENMUSEUM_Huize_%27Olean%27_in_de_omgeving_van_Pasirpoetih_op_Oost-Java_TMnr_60009804.jpg"
    },
    {
      id: "ancak-agung",
      nama: "Tradisi Ancak Agung Maulid Nabi",
      kategori: "Tradisi",
      jenis: "Tradisi Keagamaan Kolosal",
      lokasi: "Alun-alun Kabupaten Situbondo",
      kecamatan: "Situbondo",
      tahun: "Tradisi Tahunan",
      unggulan: true,
      deskripsi: "Peringatan Maulid Nabi Muhammad SAW dengan mengarak ribuan ancak bambu bertingkat berisi tumpeng nasi dan hasil bumi yang didoakan bersama oleh santri dan masyarakat luas.",
      status: "Agenda Tahunan Resmi Budaya Daerah",
      sumber: "Pemerintah Kabupaten Situbondo",
      detail_tambahan: "Waktu: Bulan Rabiul Awal | Karakter: Gotong Royong Warga",
      foto: "ancak-agung.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/5/52/Ancak_Agung_Situbondo.jpg/800px-Ancak_Agung_Situbondo.jpg"
    },
    {
      id: "tari-landhung",
      nama: "Tari Landhung (Lajher Pandhelungan)",
      kategori: "Kesenian",
      jenis: "Seni Tari Khas",
      lokasi: "Kabupaten Situbondo",
      kecamatan: "Situbondo",
      tahun: "2019",
      unggulan: true,
      deskripsi: "Tari kreasi khas daerah yang merepresentasikan garis pantai Situbondo terpanjang di Jawa Timur (150 km). Diciptakan oleh Hosnatun (Cak Tutun), tarian ini memadukan kelenturan gerak pesisir dan hentakan dinamis Pandalungan.",
      status: "Kesenian Resmi Khas Kabupaten Situbondo",
      sumber: "Dinas Pendidikan dan Kebudayaan Situbondo",
      detail_tambahan: "Koreografer: Hosnatun (Cak Tutun) | Tahun: 2019",
      foto: "tari-landhung.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/f/f7/Dokumentasi_tari_remo_munali_fatah.jpg/800px-Dokumentasi_tari_remo_munali_fatah.jpg"
    },
    {
      id: "petik-laut-kalbut",
      nama: "Tradisi Petik Laut Pelabuhan Kalbut",
      kategori: "Tradisi",
      jenis: "Ritus Bahari Pesisir",
      lokasi: "Pelabuhan Kalbut, Desa Semiring",
      kecamatan: "Mangaran",
      tahun: "Turun-temurun",
      unggulan: true,
      deskripsi: "Upacara syukuran nelayan Kalbut atas melimpahnya tangkapan laut dan doa keselamatan berlayar. Ditandai pelarungan replika perahu sesaji (gitek) ke tengah laut diiringi ratusan perahu hias berlayar.",
      status: "Warisan Budaya Takbenda Kemaritiman",
      sumber: "Rukun Nelayan Kalbut Mangaran",
      detail_tambahan: "Pusat Tradisi Petik Laut Utama Situbondo",
      foto: "hero-petik-laut.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/Petik_laut.jpg/1600px-Petik_laut.jpg"
    },
    {
      id: "pelabuhan-kalbut",
      nama: "Pelabuhan Tradisional Kalbut",
      kategori: "Sejarah",
      jenis: "Dermaga Bersejarah Pesisir",
      lokasi: "Desa Semiring",
      kecamatan: "Mangaran",
      tahun: "Abad ke-19",
      unggulan: false,
      deskripsi: "Pelabuhan rakyat bersejarah yang menghubungkan pengangkutan barang dagangan dan ternak sapi antara Situbondo dan pulau-pulau di Madura seperti Pulau Raas dan Sapudi.",
      status: "Situs Sejarah Maritim Lokal",
      sumber: "Dinas Perhubungan Situbondo",
      detail_tambahan: "Jalur Tradisional Situbondo - Kepulauan Madura",
      foto: "pelabuhan-kalbut.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e3/Layar_Terkembang.jpg/800px-Layar_Terkembang.jpg"
    },
    {
      id: "pg-panji",
      nama: "Kompleks Pabrik Gula Panji",
      kategori: "Cagar Budaya",
      jenis: "Situs Industri Gula",
      lokasi: "Kecamatan Panji",
      kecamatan: "Panji",
      tahun: "1884",
      unggulan: false,
      deskripsi: "Kompleks pabrik gula peninggalan kolonial Belanda yang dibangun tahun 1884. Menjadi pusat penggilingan tebu rakyat dan penggerak ekonomi agraria di lembah Panji.",
      status: "Situs Sejarah Industri Gula",
      sumber: "Arsip PG Panji & Pemkab Situbondo",
      detail_tambahan: "Tahun Berdiri: 1884 | Arsitektur Pabrik Abad 19",
      foto: "pg-panji.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/f/f9/KITLV_185566_-_Halte_Pandji_bij_suikeronderneming_Pandji_bij_Kapongan_bij_Sitoebondo.jpg/600px-KITLV_185566_-_Halte_Pandji_bij_suikeronderneming_Pandji_bij_Kapongan_bij_Sitoebondo.jpg"
    },
    {
      id: "tari-remo-trisnawati",
      nama: "Tari Remo Trisnawati Panji",
      kategori: "Kesenian",
      jenis: "Seni Tari Klasik",
      lokasi: "Kecamatan Panji",
      kecamatan: "Panji",
      tahun: "1981",
      unggulan: false,
      deskripsi: "Varian tari Remo gaya Situbondo yang dipelopori penari perempuan Trisnawati sejak 1981. Mengekspresikan karakter ksatria wanita dengan sampur berayun tegas dan lonceng kaki bergemerincing.",
      status: "Tercatat Kesenian Tari Daerah",
      sumber: "Publikasi Kesenian Daerah Situbondo",
      detail_tambahan: "Gaya Khas: Remo Putri Karakter Tegas",
      foto: "tari-remo-panji.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/a/a0/Tari_Remo_Remaja.jpg/800px-Tari_Remo_Remaja.jpg"
    },
    {
      id: "gelang-renteng-seletreng",
      nama: "Gelang Renteng Purbakala Seletreng",
      kategori: "Cagar Budaya",
      jenis: "Benda Logam Purbakala",
      lokasi: "Desa Seletreng",
      kecamatan: "Kapongan",
      tahun: "Era Prasejarah Logam",
      unggulan: false,
      deskripsi: "Artefak gelang logam perunggu kuno yang tersusun bertautan, ditemukan pada penggalian situs hunian megalitik di kawasan Seletreng Kapongan.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 188/317/P/001.3/2023)",
      sumber: "Naskah Kajian TACB Situbondo",
      detail_tambahan: "SK Bupati 2023 | Bahan: Perunggu Logam Purba",
      foto: "gelang-renteng-seletreng.jpg",
      foto_url: "data/foto/gelang-renteng-seletreng.jpg"
    },
    {
      id: "cottoan-peleyan-kapongan",
      nama: "Kain Batik Kuno Cotto'an Peleyan",
      kategori: "Sejarah",
      jenis: "Kriya Tekstil Leluhur",
      lokasi: "Desa Peleyan",
      kecamatan: "Kapongan",
      tahun: "Abad ke-19",
      unggulan: false,
      deskripsi: "Peninggalan kain batik tertua di tanah Situbondo yang dinamai cotto'an. Teknik pewarnaan alami indigo dan soga yang diwariskan turun-temurun oleh para pembatik sepuh Peleyan.",
      status: "Warisan Kriya Leluhur Situbondo",
      sumber: "Dokumentasi Sentra Batik Peleyan",
      detail_tambahan: "Asal Nama: Cotto'an (Pewarnaan Tradisional)",
      foto: "batik-cottoan-peleyan.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/c/cf/Batik_Madura_dengan_Motif_Batik.jpg/800px-Batik_Madura_dengan_Motif_Batik.jpg"
    },
    {
      id: "macapat-madura",
      nama: "Tembang Macapat Mamaca Madura",
      kategori: "Kesenian",
      jenis: "Sastra Tutur & Tembang",
      lokasi: "Kecamatan Arjasa",
      kecamatan: "Arjasa",
      tahun: "Abad ke-18",
      unggulan: false,
      deskripsi: "Seni melantunkan tembang naskah kuno (mamaca) beraksara pegon atau Carakan Madura yang berisi wejangan budi pekerti luhur, kisah para nabi, tasawuf, dan riwayat leluhur.",
      status: "Tercatat Warisan Sastra Lisan",
      sumber: "Komunitas Pembaca Naskah Kuno Situbondo",
      detail_tambahan: "Media: Kitab Naskah Lontar & Daluwang",
      foto: "macapat-madura.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/4/45/De_gamelan_oor_De_Wali_Negara_van_Madoera%2C_P.A.A._Tjakraningrat%2C_geschonken_aan%2C_Bestanddeelnr_15962.jpg/800px-De_gamelan_oor_De_Wali_Negara_van_Madoera%2C_P.A.A._Tjakraningrat%2C_geschonken_aan%2C_Bestanddeelnr_15962.jpg"
    },
    {
      id: "batu-karang-arjasa",
      nama: "Tinggalan Megalitik Arjasa",
      kategori: "Cagar Budaya",
      jenis: "Situs Purbakala",
      lokasi: "Perbukitan Arjasa",
      kecamatan: "Arjasa",
      tahun: "Prasejarah",
      unggulan: false,
      deskripsi: "Struktur formasi batu alam megalitik yang dahulu dimanfaatkan sebagai tempat pertemuan para tetua adat prasejarah di perbukitan timur Situbondo.",
      status: "Tercatat dalam Inventaris ODCB Daerah",
      sumber: "Dinas Pendidikan dan Kebudayaan Situbondo",
      detail_tambahan: "Lokasi: Kawasan Perbukitan Arjasa",
      foto: "lumpang-batu-arjasa.jpg",
      foto_url: "data/foto/lumpang-batu-arjasa.jpg"
    },
    {
      id: "karapan-sapi-situbondo",
      nama: "Kesenian & Olahraga Karapan Sapi",
      kategori: "Kesenian",
      jenis: "Olahraga & Atraksi Tradisi",
      lokasi: "Kecamatan Jangkar",
      kecamatan: "Jangkar",
      tahun: "Sejak Era Kolonial 1910",
      unggulan: true,
      deskripsi: "Atraksi pacuan sepasang sapi jantan berhias kaleles kayu ukir khas budaya Pandalungan. Berakar dari tradisi membajak sawah yang berkembang menjadi ajang ketangkasan bergengsi dan silaturahmi warga.",
      status: "Warisan Budaya Tradisi Pandalungan",
      sumber: "Koleksi Dokumentasi Tropenmuseum 1910 & Paguyuban Karapan",
      detail_tambahan: "Sapi Pilihan dengan Hiasan Ukir Kaleles Khas",
      foto: "kerapan-sitoebondo.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/6/6b/COLLECTIE_TROPENMUSEUM_Stierenwedloop_%28Kerappan%29_te_Sitoebondo_op_Oost-Java_tijdens_het_bezoek_van_de_Soesoehoenan_van_Solo_TMnr_60009914.jpg/800px-COLLECTIE_TROPENMUSEUM_Stierenwedloop_%28Kerappan%29_te_Sitoebondo_op_Oost-Java_tijdens_het_bezoek_van_de_Soesoehoenan_van_Solo_TMnr_60009914.jpg"
    },
    {
      id: "pelabuhan-jangkar",
      nama: "Pelabuhan Penyeberangan Jangkar",
      kategori: "Sejarah",
      jenis: "Dermaga Bersejarah Antarpulau",
      lokasi: "Desa Jangkar",
      kecamatan: "Jangkar",
      tahun: "Abad ke-20",
      unggulan: false,
      deskripsi: "Dermaga penyeberangan laut strategis di timur Situbondo yang menghubungkan Pulau Jawa dengan Pulau Madura (Kalianget), Pulau Raas, Sapudi, hingga Kepulauan Nusa Tenggara.",
      status: "Dermaga Maritim Strategis Daerah",
      sumber: "Dinas Perhubungan Situbondo",
      detail_tambahan: "Jalur Laut Utama ke Kepulauan Madura & NTB",
      foto: "pelabuhan-jangkar.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/a/a3/Ketapang_Ferry_Port%2C_East_Java%2C_20220821_1054_9794.jpg/800px-Ketapang_Ferry_Port%2C_East_Java%2C_20220821_1054_9794.jpg"
    },
    {
      id: "dolmen-mojosari",
      nama: "Dolmen Meja Batu Mojosari",
      kategori: "Cagar Budaya",
      jenis: "Struktur Megalitik",
      lokasi: "Desa Mojosari",
      kecamatan: "Asembagus",
      tahun: "Zaman Megalithikum",
      unggulan: true,
      deskripsi: "Meja batu purbakala berkaki batu penyangga yang dahulu difungsikan sebagai altar persembahan sesaji serta tempat musyawarah pemuka adat prasejarah.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 100.3.3.2/185/431.013/2024)",
      sumber: "Naskah Kajian TACB Kabupaten Situbondo",
      detail_tambahan: "Bentuk: Meja Batu Altar Sesaji Purba",
      foto: "dolmen-mojosari.jpg",
      foto_url: "data/foto/dolmen-mojosari.jpg"
    },
    {
      id: "komantan-korong",
      nama: "Tradisi Pengantin Komantan Korong",
      kategori: "Tradisi",
      jenis: "Tradisi Adat Pernikahan",
      lokasi: "Desa Asembagus",
      kecamatan: "Asembagus",
      tahun: "Turun-temurun",
      unggulan: true,
      deskripsi: "Ritus pernikahan khas masyarakat Pandalungan di mana pengantin pria dimasukkan ke dalam sangkar hias (korong) yang diselubungi kain batik dan bunga melati, lalu dipikul beramai-ramai menuju rumah mempelai wanita.",
      status: "Tercatat Warisan Adat Tradisi",
      sumber: "Lembaga Adat Desa Asembagus",
      detail_tambahan: "Makna: Tanggung Jawab & Perlindungan Suami",
      foto: "komantan-korong.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/3/32/Baju_Adat_Pengantin_Madura_Lilin_dari_Sumenep_01.jpg/800px-Baju_Adat_Pengantin_Madura_Lilin_dari_Sumenep_01.jpg"
    },
    {
      id: "trolingkung",
      nama: "Musik Patrol Bambu Trolingkung",
      kategori: "Kesenian",
      jenis: "Seni Musik Tradisi",
      lokasi: "Kecamatan Asembagus",
      kecamatan: "Asembagus",
      tahun: "Abad ke-20",
      unggulan: false,
      deskripsi: "Ensambel musik kentongan bambu berlaras nada khas yang dimainkan secara ritmis keliling kampung saat ronda malam dan perayaan hari besar kemerdekaan.",
      status: "Tercatat Tradisi Kesenian Rakyat",
      sumber: "Himpunan Seniman Musik Bambu Asembagus",
      detail_tambahan: "Instrumen: Kentongan Bambu Multi-Nada",
      foto: "trolingkung-asembagus.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/0/02/Angklung_Paglak.jpg/800px-Angklung_Paglak.jpg"
    },
    {
      id: "menhir-rajekwesi",
      nama: "Menhir Batu Tegak Rajekwesi Baluran",
      kategori: "Cagar Budaya",
      jenis: "Benda Megalitik",
      lokasi: "Kawasan Savana Baluran",
      kecamatan: "Banyuputih",
      tahun: "Zaman Batu Besar",
      unggulan: true,
      deskripsi: "Tegakan batu andesit tunggal yang didirikan tegak di kawasan timur Baluran sebagai monumen pemujaan arwah leluhur serta tanda batas kawasan sakral prasejarah.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 100.3.3.2/185/431.013/2024)",
      sumber: "Naskah Kajian TACB Kabupaten Situbondo",
      detail_tambahan: "SK Bupati 2024 | Kawasan Taman Nasional Baluran",
      foto: "sarkofagus-rajekwesi.jpg",
      foto_url: "data/foto/sarkofagus-rajekwesi.jpg"
    },
    {
      id: "batu-lesung-sumberejo",
      nama: "Situs Batu Lesung Sumberejo",
      kategori: "Cagar Budaya",
      jenis: "Situs Megalitik",
      lokasi: "Desa Sumberejo",
      kecamatan: "Banyuputih",
      tahun: "Prasejarah",
      unggulan: false,
      deskripsi: "Gugusan 5 lumpang batu berlubang purba yang berjejer rapi di kawasan timur Situbondo, menjadi bukti awal tradisi penumbukan bulir pangan agraris nusantara.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 188/392/P/004.2/2020)",
      sumber: "TACB Situbondo",
      detail_tambahan: "Jumlah: 5 Formasi Lumpang Batu Kuno",
      foto: "batu-lesung-sumberejo.jpg",
      foto_url: "data/foto/batu-lesung-sumberejo.jpg"
    },
    {
      id: "batu-togu",
      nama: "Situs Menhir Batu Togu Krajan",
      kategori: "Cagar Budaya",
      jenis: "Benda Megalitik",
      lokasi: "Desa Krajan",
      kecamatan: "Banyuputih",
      tahun: "Zaman Megalithikum",
      unggulan: false,
      deskripsi: "Menhir batu tegak purba berukuran tinggi 85 cm, lebar 40 cm, dan tebal 17 cm yang dikeramatkan sebagai penanda wilayah kepemimpinan masa lampau.",
      status: "Cagar Budaya Peringkat Kabupaten (SK Bupati 188/443/P/004.2/2021)",
      sumber: "Naskah Kajian TACB Situbondo",
      detail_tambahan: "Dimensi: 85 x 40 x 17 cm | SK Bupati 2021",
      foto: "batu-togu-krajan.jpg",
      foto_url: "data/foto/batu-togu-krajan.jpg"
    },
    {
      id: "perjuangan-khr-asad",
      nama: "Markas Perjuangan KHR. As'ad Syamsul Arifin",
      kategori: "Sejarah",
      jenis: "Peristiwa Kemerdekaan RI",
      lokasi: "Pondok Pesantren Sukorejo",
      kecamatan: "Banyuputih",
      tahun: "1945 - 1949",
      unggulan: true,
      deskripsi: "Jejak markas perjuangan Laskar Sabilillah dan Hizbullah di Sukorejo yang dipimpin ulama kharismatik KHR. As'ad Syamsul Arifin dalam membendung agresi militer penjajah di tapal kuda timur Jawa Timur. Beliau dianugerahi gelar Pahlawan Nasional RI pada 2016.",
      status: "Situs Tokoh Pahlawan Nasional RI",
      sumber: "Kementerian Sosial RI & Biografi KHR. As'ad Syamsul Arifin",
      detail_tambahan: "Gelar: Pahlawan Nasional RI (Keppres 115/TK/2016)",
      foto: "khr-asad-syamsul-arifin.jpg",
      foto_url: "https://upload.wikimedia.org/wikipedia/commons/thumb/e/ee/KHR._As%27ad_Syamsul_Arifin.jpg/600px-KHR._As%27ad_Syamsul_Arifin.jpg"
    }
  ];
}

function getSeedTimeline() {
  // Linimasa 12 babak sejarah resmi Situbondo
  return [
    {
      id: "era-megalitikum",
      era: "Zaman Prasejarah",
      tahun: "± 2500 - 500 SM",
      judul: "Peradaban Megalitikum Lereng Argopuro & Baluran",
      ringkasan: "Sebaran tinggalan sarkofagus, menhir, dan batu dolmen membuktikan Situbondo adalah salah satu sentra peradaban megalitik tertua dan terpadat di Jawa Timur.",
      uraian: "Jejak hunian purbakala di wilayah Situbondo terbentang dari lereng Gunung Argopuro di sisi barat hingga kawasan savana Baluran di timur. Ditemukan lebih dari puluhan sarkofagus batu monolitik di Bungatan, Suboh, dan Cemara, menhir tegak di Rajekwesi Banyuputih, serta dolmen meja batu di Mojosari Asembagus. Tinggalan ini mencerminkan tatanan masyarakat agraris purba yang telah mengenal sistem kepercayaan penghormatan leluhur.",
      sumber: "Balai Pelestarian Kebudayaan (BPK) Wilayah XI & Naskah Kajian TACB Situbondo",
      lokasi: "Bungatan, Suboh, Asembagus, Banyuputih",
      foto: "pita-baluran.jpg"
    },
    {
      id: "era-majapahit",
      era: "Kerajaan Klasik Nusantara",
      tahun: "1365 M",
      judul: "Panarukan Bandar Utama Majapahit (Kitab Nagarakretagama)",
      ringkasan: "Panarukan tercatat dalam pupuh 28 Kitab Nagarakretagama karya Mpu Prapanca sebagai pelabuhan bandar laut strategis Majapahit.",
      uraian: "Dalam Kitab Desawarnana (Nagarakretagama) pupuh 28 karya Mpu Prapanca yang ditulis pada tahun 1365 Masehi semasa pemerintahan Prabu Hayam Wuruk, wilayah Patukangan (Pecaron) dan Panarukan dicatat sebagai pelabuhan transit resmi kekaisaran Majapahit menuju wilayah timur nusantara. Panarukan menjadi gerbang niaga maritim yang mempertemukan saudagar rempah-rempah Maluku, beras Jawa, dan pedagang antarpulau.",
      sumber: "Kitab Nagarakretagama (Mpu Prapanca, 1365) & Sejarah Kebudayaan Jawa Timur",
      lokasi: "Kecamatan Panarukan",
      foto: "hero-pasir-putih.jpg"
    },
    {
      id: "era-blambangan-portugis",
      era: "Penjelajahan Samudera",
      tahun: "1512 - 1515 M",
      judul: "Kesaksian Tomé Pires (Suma Oriental) tentang Panarukan",
      ringkasan: "Penjelajah Portugis Tomé Pires mencatat Panarukan sebagai kota pelabuhan besar di bawah kedaulatan Kerajaan Blambangan.",
      uraian: "Dalam naskah penjelajahan legendaris Suma Oriental karangan musafir Portugis Tomé Pires (1512-1515), Panarukan digambarkan sebagai pelabuhan internasional yang ramai dan benteng pertahanan maritim Kerajaan Blambangan. Kapal-kapal dagang dari berbagai penjuru singgah mengisi perbekalan air tawar dan bertransaksi komoditas pertanian serta hasil hutan sebelum menyeberang ke kepulauan Nusa Tenggara.",
      sumber: "The Suma Oriental of Tomé Pires (Armando Cortesão, The Hakluyt Society, 1944)",
      lokasi: "Pesisir Teluk Panarukan",
      foto: "pita-laut-jawa.jpg"
    },
    {
      id: "era-pate-alos",
      era: "Pendirian Kadipaten",
      tahun: "1743 - 1764 M",
      judul: "Kepemimpinan Kyai Pate Alos & Berdirinya Kadipaten Besuki",
      ringkasan: "Raden Bagus Kasim (Kyai Pate Alos / Pangeran Aryo Soerjoadiningrat I) membuka wilayah dan mendirikan pusat pemerintahan pertama di Besuki.",
      uraian: "Sosok sentral perintis pemerintahan Situbondo adalah Raden Bagus Kasim yang bergelar Tumenggung Joyolelono atau Kyai Pate Alos. Beliau memimpin pembukaan hutan belantara dan penataan permukiman di kawasan Besuki, menjadikannya kadipaten yang makmur dan aman. Atas kepemimpinannya yang arif dan berwibawa (alos), beliau sangat dihormati oleh seluruh masyarakat Madura dan Jawa hingga hari ini, dengan makam pusaka yang terus diziarahi di Besuki.",
      sumber: "Babad Besuki, Arsip Kearsipan Jawa Timur, & Dokumen Balai Pelestarian Nilai Budaya",
      lokasi: "Kecamatan Besuki",
      foto: "pendopo-besuki.jpg"
    },
    {
      id: "era-karesidenan-besuki",
      era: "Karesidenan Hindia Belanda",
      tahun: "1800-an",
      judul: "Besuki Menjadi Ibu Kota Karesidenan Besuki (Residentie Besoeki)",
      ringkasan: "Pemerintah kolonial Hindia Belanda menetapkan Besuki sebagai ibu kota karesidenan yang membawahi lima afdeeling di ujung timur Jawa.",
      uraian: "Pada awal abad ke-19, posisi Besuki semakin strategis hingga diangkat menjadi ibu kota Karesidenan Besuki (Residentie Besoeki). Wilayah administratif ini membawahi Afdeeling Besuki, Panarukan, Bondowoso, Jember, dan Banyuwangi. Di masa ini dibangun Gedung Rumah Dinas Residen Besuki, Kantor Kawedanan, Rumah Bupati Besuki I, serta Klenteng Poo Tong Bio sebagai simpul peribadatan saudagar Tionghoa yang kini seluruhnya telah berstatus cagar budaya resmi.",
      sumber: "Regeeringsalmanak voor Nederlandsch-Indie & SK Penetapan Cagar Budaya Bupati Situbondo 2020",
      lokasi: "Kota Tua Besuki",
      foto: ""
    },
    {
      id: "era-jalan-daendels",
      era: "Jalan Raya Pos",
      tahun: "1808 M",
      judul: "Ujung Timur Jalan Raya Pos Anyer - Panarukan (De Grote Postweg)",
      ringkasan: "Gubernur Jenderal Daendels menetapkan Panarukan sebagai titik akhir lintasan jalan raya pos sepanjang 1.000 kilometer.",
      uraian: "Gubernur Jenderal Herman Willem Daendels memerintahkan pembuatan jalan raya pos dari Anyer (Banten) hingga Panarukan. Panarukan dipilih sebagai ujung akhir karena memiliki teluk pelabuhan yang tenang dan terlindung dari gelombang samudera, menjadikannya pintu keluar utama komoditas perkebunan kopi, gula, dan tembakau dari Karesidenan Besuki menuju pasar dunia.",
      sumber: "Arsip Nasional Republik Indonesia (ANRI) & De Grote Postweg (Pramoedya Ananta Toer)",
      lokasi: "Panarukan (Titik Nol Km)",
      foto: ""
    },
    {
      id: "era-industri-gula",
      era: "Industrialisasi Pesisir",
      tahun: "1846 - 1897 M",
      judul: "Kejayaan Industri Gula & Jalur Kereta Api Staatsspoorwegen",
      ringkasan: "Pendirian pabrik-pabrik gula bersejarah (PG Olean, PG Panji, PG Asembagus) dan jalur kereta api Kalisat - Panarukan.",
      uraian: "Situbondo berkembang pesat sebagai salah satu sentra industri gula terbesar di Hindia Belanda. Didirikan PG Olean (1846) yang hingga kini masih mempertahankan lokomotif uap kuno, PG Panji (1884), PG Demas, PG Wringinanom, dan PG Asembagus (1891). Untuk mengalirkan berton-ton gula ke kapal ekspor, Staatsspoorwegen (SS) membangun jalur rel kereta api Kalisat-Panarukan yang diresmikan pada tahun 1897.",
      sumber: "Arsip Suikerfabriek Oost-Java & Dokumentasi Sejarah Perkeretaapian Indonesia",
      lokasi: "Panji, Olean, Asembagus",
      foto: "pita-utama-raya.jpg"
    },
    {
      id: "era-harjakasi",
      era: "Hari Jadi Resmi",
      tahun: "15 Agustus",
      judul: "Penetapan Hari Jadi Kabupaten Situbondo (Harjakasi)",
      ringkasan: "Peringatan resmi berdirinya pemerintahan Kabupaten Situbondo berpijak pada penataan afdeeling dan pemindahan ibu kota ke Situbondo.",
      uraian: "Hari Jadi Kabupaten Situbondo (Harjakasi) ditetapkan setiap tanggal 15 Agustus melalui Peraturan Daerah. Momentum ini memperingati babak bersejarah pemindahan dan penataan pusat pemerintahan dari Besuki ke wilayah Situbondo (pusat kota saat ini), yang kemudian diperingati setiap tahun melalui kirab budaya agung, pawai pusaka, ziarah leluhur Kyai Pate Alos, dan pergelaran kesenian rakyat.",
      sumber: "Peraturan Daerah Kabupaten Situbondo tentang Hari Jadi Kabupaten Situbondo",
      lokasi: "Kabupaten Situbondo",
      foto: "ancak-agung.jpg"
    },
    {
      id: "era-pahlawan-asad",
      era: "Perang Kemerdekaan",
      tahun: "1945 - 1949 M",
      judul: "Perjuangan KHR. As'ad Syamsul Arifin & Laskar Sabilillah",
      ringkasan: "Pesantren Sukorejo Banyuputih menjadi basis perlawanan bersenjata dan penggemblengan pejuang kemerdekaan melawan penjajah.",
      uraian: "Ulama kharismatik KHR. As'ad Syamsul Arifin dari Pondok Pesantren Salafiyah Syafi'iyah Sukorejo memimpin Laskar Sabilillah dan Hizbullah di wilayah karesidenan Besuki. Beliau mengobarkan jihad membela tanah air, mengoordinasikan pertempuran gerilya di daerah pelabuhan, serta menyokong logistik pejuang kemerdekaan. Atas jasa kepahlawanannya, Presiden RI menganugerahkan gelar Pahlawan Nasional kepada KHR. As'ad pada tahun 2016.",
      sumber: "Keputusan Presiden RI No. 115/TK/Tahun 2016 & Biografi KHR. As'ad Syamsul Arifin",
      lokasi: "Sukorejo, Kecamatan Banyuputih",
      foto: "pita-kerapu.jpg"
    },
    {
      id: "era-pp-28-1972",
      era: "Legitimasi Modern",
      tahun: "1972 M",
      judul: "Perubahan Nama Kabupaten Panarukan Menjadi Kabupaten Situbondo",
      ringkasan: "Pemerintah menerbitkan PP No. 28 Tahun 1972 yang secara resmi mengubah nama Kabupaten Panarukan menjadi Kabupaten Situbondo.",
      uraian: "Berdasarkan Peraturan Pemerintah Republik Indonesia No. 28 Tahun 1972 tanggal 21 September 1972, nama 'Kabupaten Panarukan' secara yuridis resmi diubah menjadi 'Kabupaten Situbondo'. Penyesuaian ini menyelaraskan nama daerah otonom tingkat II dengan nama ibu kota kabupaten tempat berkedudukannya kantor bupati dan pusat kegiatan aparatur pemerintahan.",
      sumber: "Lembaran Negara Republik Indonesia No. 39 Tahun 1972 (Peraturan Pemerintah No. 28/1972)",
      lokasi: "Kabupaten Situbondo",
      foto: ""
    },
    {
      id: "era-muktamar-pancasila",
      era: "Kota Santri Pancasila",
      tahun: "1984 M",
      judul: "Muktamar NU Sukorejo & Deklarasi Asas Tunggal Pancasila",
      ringkasan: "Muktamar Nahdlatul Ulama ke-27 di Sukorejo Situbondo melahirkan keputusan monumental penerimaan Pancasila sebagai dasar negara.",
      uraian: "Pada bulan Desember 1984 di Pondok Pesantren Sukorejo, Situbondo mencatat tinta emas sejarah bangsa saat Muktamar NU ke-27 menetapkan Deklarasi tentang Hubungan Islam dan Pancasila. Diprakarsai oleh KHR. As'ad Syamsul Arifin bersama para kiai sepuh, NU menyatakan bahwa penerimaan Pancasila tidak bertentangan dengan syariat Islam dan NKRI adalah bentuk final. Dari peristiwa ini lahir identitas kuat Situbondo sebagai Kota Santri Pancasila.",
      sumber: "Hasil Muktamar NU ke-27 di Situbondo (1984) & Risalah Kebangsaan PBNU",
      lokasi: "Sukorejo, Kecamatan Banyuputih",
      foto: "pita-mangrove.jpg"
    },
    {
      id: "era-kontemporer",
      era: "Era Pelestarian Digital",
      tahun: "2020 - Sekarang",
      judul: "Konservasi Cagar Budaya & Kebangkitan Seni Tradisi Situbondo",
      ringkasan: "Penetapan puluhan cagar budaya resmi SK Bupati, pengukuhan Tari Landhung, dan digitalisasi warisan kebudayaan daerah.",
      uraian: "Kabupaten Situbondo giat melakukan registrasi dan penyelamatan warisan leluhur. Sebanyak 77 objek cagar budaya (bangunan, struktur, situs, dan benda purbakala) telah resmi ditetapkan melalui SK Bupati Situbondo rentang 2020 hingga 2025. Di bidang seni, Tari Landhung resmi menjadi ikon tarian daerah, festival ancak agung dan petik laut digelar rutin, serta inventarisasi budaya kini dapat diakses secara terbuka oleh seluruh masyarakat Situbondo.",
      sumber: "Dinas Pendidikan dan Kebudayaan Kabupaten Situbondo & Registrasi Cagar Budaya Nasional",
      lokasi: "17 Kecamatan se-Kabupaten Situbondo",
      foto: "hero-petik-laut.jpg"
    }
  ];
}
