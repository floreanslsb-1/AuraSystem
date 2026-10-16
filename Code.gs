/** =====================================================================
 *  AURA COSMETICS — Web App Toko (Cetak Label + Barang Masuk)
 *  Project Apps Script berdiri sendiri (bukan extension di sheet).
 *  Karyawan cukup buka URL web app — tidak perlu akses ke Master Data.
 *  ===================================================================== */

const CFG = {
  MASTER_ID     : '1h9OLC-qFtw6mZ-yduVUiOMuzljby5W7Ur6oAPpLzy_s',   // ← ID dari URL Google Sheets master
  SHEET_PRODUK  : 'Produk',
  PDF_FOLDER    : 'Label Aura Cosmetics',
  LOGO_FILE_ID  : '11kp5gVM4kAP4_yJs7cckRFKHfAxdhSUH',
  STORE_NAME    : 'Aura Cosmetics',
  LABEL_W_MM    : 50,
  LABEL_H_MM    : 35,
  GAP_MM        : 2,
  PAGE_MARGIN_MM: 5,
  FIT_COLS      : 4,
  PAGE_SIZE     : 'A4 portrait',
  THEME         : 't-cream',
  SHOW_SKU      : true          // label tunggal menampilkan SKU kecil di bawah; label grup tidak
};

// Judul kolom di tab Produk (harus persis sama dengan header master)
const H = {
  sku    : 'SKU',
  nama   : 'Nama Produk',
  brand  : 'Brand',
  barcode: 'Barcode',
  het    : 'HET (Rp)',
  jual   : 'Harga Jual / Aura (Rp)',
  label  : 'Nama di Label',
  labHet : 'HET Label Tercetak',
  labJual: 'Harga Label Tercetak',
  labTgl : 'Tgl Label Dicetak',
  labStat: 'Status Label'
};

/** ================= WEB APP ================= **/
function doGet() {
  const t = HtmlService.createTemplateFromFile('App');
  t.cfg  = CFG;
  t.logo = getLogoDataUri_();
  return t.evaluate()
    .setTitle('Aura — Label & Barang Masuk')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** ================= DATA LABEL ================= **/
// Satu entri per label. Produk dengan "Nama di Label" sama = satu label,
// kecuali harganya tidak seragam → dipecah jadi label per produk (pakai Nama Produk).
function getLabels() {
  const sh = produk_();
  const c  = cols_(sh);
  const n  = sh.getLastRow() - 1;
  if (n < 1) return [];
  const vals = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();

  const groups = {}, order = [];
  vals.forEach(function (r) {
    const sku = String(r[c.sku] || '').trim();
    if (!sku) return;
    const nama  = String(r[c.nama] || '').trim();
    const teks  = String(r[c.label] || '').trim() || nama;
    const key   = teks.toUpperCase();
    if (!groups[key]) { groups[key] = { teks: teks, rows: [] }; order.push(key); }
    groups[key].rows.push({
      sku: sku, nama: nama, brand: String(r[c.brand] || ''),
      barcode: String(r[c.barcode] || ''),
      het: num_(r[c.het]), jual: num_(r[c.jual]),
      status: String(r[c.labStat] || '')
    });
  });

  const out = [];
  order.forEach(function (key) {
    const g = groups[key];
    const prices = {};
    g.rows.forEach(function (x) { prices[x.jual + '|' + x.het] = 1; });
    const seragam = Object.keys(prices).length === 1;
    if (seragam || g.rows.length === 1) {
      out.push(item_(g.teks, g.rows, false));
    } else {
      g.rows.forEach(function (x) { out.push(item_(x.nama, [x], true)); });
    }
  });
  return out;
}

function item_(teks, rows, dipecah) {
  const a = rows[0];
  const promo = a.het > a.jual && a.jual > 0;
  let status = 'OK';
  rows.forEach(function (x) { status = worst_(status, x.status); });
  return {
    id      : rows.map(function (x) { return x.sku; }).join(','),
    nama    : teks,
    brand   : a.brand,
    skus    : rows.map(function (x) { return x.sku; }),
    jumlah  : rows.length,
    het     : a.het,
    jual    : a.jual,
    promo   : promo,
    diskon  : promo ? Math.floor((1 - a.jual / a.het) * 100 + 1e-9) : 0,
    status  : status,
    dipecah : dipecah,
    search  : (teks + ' ' + rows.map(function (x) { return x.nama + ' ' + x.sku + ' ' + x.barcode; }).join(' ')).toLowerCase(),
    barcodes: rows.map(function (x) { return x.barcode; }).filter(String)
  };
}

function worst_(a, b) {
  const rank = { 'OK': 0, 'Belum dicetak': 1, 'Perlu cetak ulang': 2 };
  return (rank[b] || 0) > (rank[a] || 0) ? b : a;
}

/** Dipanggil setelah label dicetak: catat harga yang tercetak ke master. */
function markPrinted(skus) {
  if (!skus || !skus.length) return 0;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = produk_();
    const c  = cols_(sh);
    const n  = sh.getLastRow() - 1;
    const vals = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();
    const set = {};
    skus.forEach(function (s) { set[String(s).trim()] = 1; });

    const outHet = [], outJual = [], outTgl = [];
    const now = new Date();
    let count = 0;
    vals.forEach(function (r) {
      const hit = set[String(r[c.sku] || '').trim()];
      outHet.push([hit ? r[c.het] : r[c.labHet]]);
      outJual.push([hit ? r[c.jual] : r[c.labJual]]);
      outTgl.push([hit ? now : r[c.labTgl]]);
      if (hit) count++;
    });
    sh.getRange(2, c.labHet + 1, n, 1).setValues(outHet);
    sh.getRange(2, c.labJual + 1, n, 1).setValues(outJual);
    sh.getRange(2, c.labTgl + 1, n, 1).setValues(outTgl);
    return count;
  } finally {
    lock.releaseLock();
  }
}

/** ================= PDF ================= **/
function savePdf(html, count) {
  const name = 'Label Aura ' +
    Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH.mm') + ' (' + count + ').pdf';
  const blob = Utilities.newBlob(html, 'text/html', 'label.html').getAs('application/pdf').setName(name);
  const it = DriveApp.getFoldersByName(CFG.PDF_FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(CFG.PDF_FOLDER);
  return folder.createFile(blob).getUrl();
}

/** ================= HELPERS ================= **/
function produk_() {
  const sh = SpreadsheetApp.openById(CFG.MASTER_ID).getSheetByName(CFG.SHEET_PRODUK);
  if (!sh) throw new Error('Tab "' + CFG.SHEET_PRODUK + '" tidak ditemukan di Master Data.');
  return sh;
}
function cols_(sh) {
  const hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  const m = {};
  Object.keys(H).forEach(function (k) {
    const i = hdr.indexOf(H[k]);
    if (i < 0) throw new Error('Kolom "' + H[k] + '" tidak ditemukan di tab ' + CFG.SHEET_PRODUK + '.');
    m[k] = i;
  });
  return m;
}
function num_(v) {
  if (typeof v === 'number') return v;
  const s = String(v || '').replace(/[^\d,.-]/g, '').replace(/\./g, '').replace(',', '.');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}
function getLogoDataUri_() {
  if (!CFG.LOGO_FILE_ID) return '';
  const cache = CacheService.getScriptCache();
  const hit = cache.get('auralogo');
  if (hit) return hit;
  try {
    const blob = DriveApp.getFileById(CFG.LOGO_FILE_ID).getBlob();
    const uri = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    if (uri.length < 95000) cache.put('auralogo', uri, 21600);
    return uri;
  } catch (e) {
    return '';
  }
}
