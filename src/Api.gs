/**
 * ============================================================
 * Web Kebudayaan Situbondo - Google Apps Script Backend
 * Berkas: Api.gs
 * API Endpoints untuk pemanggilan google.script.run
 * ============================================================
 */

/**
 * Mengambil daftar katalog dengan filter pencarian, kategori, dan kecamatan Situbondo.
 * @param {Object} params { query: string, category: string, kecamatan: string }
 * @return {Object} { success: boolean, data: Array, total: number, error: string|null }
 */
function apiGetKatalog(params) {
  try {
    var p = params || {};
    var results = filterCultureItems(p.query, p.category, p.kecamatan);
    return {
      success: true,
      data: results,
      total: results.length,
      error: null
    };
  } catch (err) {
    Logger.log('[Api] apiGetKatalog error: ' + err.toString());
    return {
      success: false,
      data: [],
      total: 0,
      error: err.message || 'Gagal memuat katalog warisan budaya.'
    };
  }
}

/**
 * Mengambil seluruh linimasa sejarah Situbondo dari Karesidenan Besuki hingga kini.
 * @return {Object} { success: boolean, data: Array, error: string|null }
 */
function apiGetSejarah() {
  try {
    var timeline = getTimelineMilestones();
    return {
      success: true,
      data: timeline,
      total: timeline.length,
      error: null
    };
  } catch (err) {
    Logger.log('[Api] apiGetSejarah error: ' + err.toString());
    return {
      success: false,
      data: [],
      error: err.message || 'Gagal memuat linimasa sejarah Situbondo.'
    };
  }
}

/**
 * Mengambil daftar 17 Kecamatan resmi di Kabupaten Situbondo.
 */
function apiGetKecamatan() {
  return {
    success: true,
    data: getKecamatanList()
  };
}

/**
 * Mengambil detail satu entri berdasarkan ID.
 */
function apiGetDetail(id) {
  try {
    if (!id) throw new Error('Parameter ID tidak boleh kosong.');
    var item = getCultureItemById(String(id).trim());
    if (!item) {
      return {
        success: false,
        data: null,
        error: 'Data dengan ID tersebut tidak ditemukan.'
      };
    }
    return {
      success: true,
      data: item,
      error: null
    };
  } catch (err) {
    return {
      success: false,
      data: null,
      error: err.message
    };
  }
}

/**
 * Ringkasan statistik inventarisasi kebudayaan Situbondo.
 */
function apiGetRingkasan() {
  try {
    var items = getAllCultureItems();
    var stats = {
      total: items.length,
      unggulan: 0,
      totalKecamatan: getKecamatanList().length,
      perKategori: {
        'Cagar Budaya': 0,
        'Kesenian': 0,
        'Tradisi': 0,
        'Sejarah': 0
      }
    };

    items.forEach(function(item) {
      if (item.unggulan) stats.unggulan++;
      var kat = item.kategori || 'Lainnya';
      if (stats.perKategori[kat] !== undefined) {
        stats.perKategori[kat]++;
      } else {
        stats.perKategori[kat] = 1;
      }
    });

    return {
      success: true,
      data: stats,
      error: null
    };
  } catch (err) {
    return {
      success: false,
      data: null,
      error: err.message
    };
  }
}
