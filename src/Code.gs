/**
 * ============================================================
 * Web Kebudayaan Situbondo - Google Apps Script Backend
 * Berkas: Code.gs
 * Entry Point Web App (doGet) & Injeksi Server-Side Data
 * ============================================================
 */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('Index');
  
  template.appTitle = CONFIG.APP_TITLE || 'Warisan Budaya Situbondo - Pusaka Cagar Budaya, Tradisi, & Sejarah';
  template.appSubtitle = CONFIG.APP_SUBTITLE || 'Inventarisasi Digital 17 Kecamatan, Seni Pandalungan, & Jejak Karesidenan Besuki';
  
  // Injeksi server data awal untuk rendering instan bebas flicker
  var initialItems = getAllCultureItems();
  var initialTimeline = getTimelineMilestones();
  var initialKecamatan = getKecamatanList();

  template.initialItemsJson = JSON.stringify(initialItems);
  template.initialTimelineJson = JSON.stringify(initialTimeline);
  template.initialKecamatanJson = JSON.stringify(initialKecamatan);

  return template
    .evaluate()
    .setTitle(template.appTitle)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
