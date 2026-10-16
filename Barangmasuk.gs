/** =====================================================================
 *  AURA — BARANG MASUK (upload nota → cocokkan → setujui → file import)
 *  File tambahan di project web app yang sama (berdampingan dengan Code.gs).
 *  Butuh Script Properties:  GEMINI_API_KEY  dan  OWNER_PIN
 *  ===================================================================== */

const BM = {
  // dicoba berurutan; model yang sudah tidak tersedia / tanpa kuota gratis otomatis dilewati
  GEMINI_MODELS: ['gemini-3.8-flash', 'gemini-3.5-flash', 'gemini-3.1-flash-lite'],
  CLAUDE_MODEL : 'claude-haiku-4-5-20251001',   // dipakai bila CLAUDE_API_KEY diisi
  // penyedia gratis lain (OpenAI-compatible). Ganti nama model bila penyedia mengubahnya.
  MISTRAL_MODEL: 'mistral-small-latest',
  GROQ_MODEL   : 'meta-llama/llama-4-scout-17b-16e-instruct',
  FOLDER_NOTA  : 'Aura - Foto Nota',
  FOLDER_IMPORT: 'Aura - File Import Odoo',
  T_NOTA       : 'Nota',
  T_PRICELIST  : 'Price List Vendor',
  T_KODE       : 'Kode Vendor',
  T_RIWAYAT    : 'Riwayat Beli',
  T_SET        : 'Pengaturan'
};
const OWNER_EMAIL = 'floreanslsb@gmail.com';   // hardcoded: tidak bisa diubah dari app maupun sheet

const P = { // header tab Produk
  sku:'SKU', nama:'Nama Produk', brand:'Brand', barcode:'Barcode', het:'HET (Rp)',
  beli:'Modal / Harga Beli Terakhir (Rp)', jual:'Harga Jual / Aura (Rp)', mm:'Margin Manual (%)',
  tgl:'Tgl Beli Terakhir', dist:'Distributor Terakhir', label:'Nama di Label', id:'ID Odoo',
  rule:'ID Aturan Pricelist', track:'Pelacakan', pos:'Tersedia di POS', note:'Catatan'
};
const FORMULA_COLS = ['Margin (Rp)','Margin (%)','Harga Jual Saran (Rp)','Diskon dari HET','Status Harga','Cek Grup','Status Label'];

/** ================= DATA AWAL UNTUK HALAMAN ================= **/
function bm_init() {
  const ss = master_();
  ensureTabs_(ss);
  const sh = ss.getSheetByName(CFG.SHEET_PRODUK), c = hmap_(sh);
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const products = [], prefix = {}, iLbl = c['Status Label']; let labelPerlu = 0;
  vals.forEach(function (r) {
    const sku = String(r[c[P.sku]] || '').trim(); if (!sku) return;
    if (iLbl !== undefined) { const sl = String(r[iLbl] || '').trim(); if (sl && sl !== 'OK') labelPerlu++; }
    const brand = String(r[c[P.brand]] || '');
    products.push({ sku: sku, nama: String(r[c[P.nama]]), brand: brand, barcode: String(r[c[P.barcode]] || ''), label: String(r[c[P.label]] || ''),
      het: num_(r[c[P.het]]), jual: num_(r[c[P.jual]]), beli: num_(r[c[P.beli]]),
      mm: r[c[P.mm]] === '' ? null : num_(r[c[P.mm]]) });
    if (brand && !prefix[brand]) prefix[brand] = sku.split('-')[0];
  });
  return { products: products, prefix: prefix, vendors: vendors_(ss), settings: settings_(ss), notas: listNota_(ss), labelPerlu: labelPerlu };
}

/** ================= BACA NOTA DENGAN GEMINI ================= **/
function bm_readNota(files) {
  const props = PropertiesService.getScriptProperties();
  const claudeKey = props.getProperty('CLAUDE_API_KEY'), key = props.getProperty('GEMINI_API_KEY');
  const mistralKey = props.getProperty('MISTRAL_API_KEY'), groqKey = props.getProperty('GROQ_API_KEY');
  if (!claudeKey && !key && !mistralKey && !groqKey)
    throw new Error('Isi salah satu: MISTRAL_API_KEY, GROQ_API_KEY, GEMINI_API_KEY atau CLAUDE_API_KEY di Script Properties.');
  if (!files || !files.length) throw new Error('Belum ada foto nota.');

  const parts = [{ text: PROMPT_ }];
  files.forEach(function (f) { parts.push({ inline_data: { mime_type: f.mime, data: f.data } }); });
  // coba penyedia berurutan; kalau satu gagal (kuota habis / error), lanjut ke berikutnya
  const urut = [];
  if (claudeKey)  urut.push(function () { return callClaude_(claudeKey, files); });
  if (mistralKey) urut.push(function () { return callOpenAI_('https://api.mistral.ai/v1/chat/completions', mistralKey, BM.MISTRAL_MODEL, files, 'Mistral'); });
  if (groqKey)    urut.push(function () { return callOpenAI_('https://api.groq.com/openai/v1/chat/completions', groqKey, BM.GROQ_MODEL, files, 'Groq'); });
  if (key)        urut.push(function () { return callGemini_(key, parts); });
  let txt = '', gagal = [];
  for (let i = 0; i < urut.length && !txt; i++) {
    try { txt = urut[i](); } catch (e) { gagal.push(e.message); }
  }
  if (!txt) throw new Error(gagal.join(' | '));
  const ai = JSON.parse(txt.replace(/```json|```/g, '').trim());

  // foto baru disimpan ke Drive setelah nota berhasil dibaca
  const folder = folder_(BM.FOLDER_NOTA), urls = [];
  files.forEach(function (f, i) {
    const blob = Utilities.newBlob(Utilities.base64Decode(f.data), f.mime, f.name || ('nota-' + i));
    urls.push(folder.createFile(blob).getUrl());
  });

  const ss = master_();
  const draft = buildDraft_(ss, ai);
  draft.foto = urls;
  const id = 'N' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd-HHmmss');
  saveNota_(ss, id, draft, 'Draft');
  return { id: id, draft: draft };
}

function callOpenAI_(url, key, model, files, nama) {
  const content = [{ type: 'text', text: PROMPT_ + '\nJawab hanya dengan JSON, tanpa teks lain.' }];
  files.forEach(function (f) {
    if (f.mime === 'application/pdf') throw new Error(nama + ' tidak menerima PDF — upload sebagai foto (JPG).');
    content.push({ type: 'image_url', image_url: { url: 'data:' + f.mime + ';base64,' + f.data } });
  });
  const res = UrlFetchApp.fetch(url, {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify({ model: model, temperature: 0, max_tokens: 8000,
      response_format: { type: 'json_object' }, messages: [{ role: 'user', content: content }] })
  });
  const code = res.getResponseCode(), body = res.getContentText();
  if (code !== 200) throw new Error(nama + ' error ' + code + ': ' + body.slice(0, 250));
  const txt = JSON.parse(body).choices[0].message.content || '';
  const a = txt.indexOf('{'), z = txt.lastIndexOf('}');
  return a > -1 ? txt.slice(a, z + 1) : txt;
}

function callClaude_(key, files) {
  const content = files.map(function (f) {
    return f.mime === 'application/pdf'
      ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: f.data } }
      : { type: 'image', source: { type: 'base64', media_type: f.mime, data: f.data } };
  });
  content.push({ type: 'text', text: PROMPT_ + '\nJawab hanya dengan JSON, tanpa teks lain.' });
  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: BM.CLAUDE_MODEL, max_tokens: 8000, temperature: 0,
      messages: [{ role: 'user', content: content }] })
  });
  const code = res.getResponseCode(), body = res.getContentText();
  if (code !== 200) throw new Error('Claude API error ' + code + ': ' + body.slice(0, 300));
  const txt = JSON.parse(body).content.filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; }).join('');
  const a = txt.indexOf('{'), z = txt.lastIndexOf('}');
  return a > -1 ? txt.slice(a, z + 1) : txt;
}

function callGemini_(key, parts) {
  const errs = [];
  for (let i = 0; i < BM.GEMINI_MODELS.length; i++) {
    const model = BM.GEMINI_MODELS[i];
    const res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + model +
      ':generateContent?key=' + key, {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ contents: [{ parts: parts }],
          generationConfig: { responseMimeType: 'application/json', temperature: 0 } })
      });
    const code = res.getResponseCode(), body = res.getContentText();
    if (code === 200) {
      try { return JSON.parse(body).candidates[0].content.parts.map(function (p) { return p.text || ''; }).join(''); }
      catch (e) { throw new Error('Jawaban Gemini (' + model + ') tidak terbaca.'); }
    }
    const noFree = code === 429 && /limit:\s*0|free_tier/i.test(body);
    if (code === 404 || code === 400 && /not found|not supported/i.test(body) || noFree) { errs.push(model + ' (' + code + ')'); continue; }
    if (code === 429) throw new Error('Kuota gratis Gemini sedang habis (429). Coba lagi 1 menit lagi.');
    throw new Error('Gemini error ' + code + ' (' + model + '): ' + body.slice(0, 300));
  }
  throw new Error('Tidak ada model Gemini yang bisa dipakai: ' + errs.join(', ') +
    '. Cek nama model gratis di aistudio.google.com lalu ubah BM.GEMINI_MODELS.');
}

const PROMPT_ = [
  'Kamu membaca foto FAKTUR/NOTA PEMBELIAN barang kosmetik dari distributor di Indonesia.',
  'Kembalikan HANYA JSON dengan bentuk:',
  '{"vendor":"nama perusahaan penjual","no_faktur":"","tanggal":"YYYY-MM-DD",',
  ' "ppn_termasuk": true/false, "ppn_persen": 11, "total_nota": angka,',
  ' "baris":[{"kode":"kode barang/barcode di nota atau \\"\\"","nama":"nama barang persis seperti di nota",',
  '   "qty": angka dibayar, "bonus": angka qty gratis (0 bila tidak ada), "harga": harga satuan sebelum diskon,',
  '   "diskon":[persen diskon berurutan, mis. [2,1] untuk 2%+1%], "jumlah": nilai baris di nota}]}',
  'Aturan: angka tanpa pemisah ribuan (1.234,56 → 1234.56). Baris bonus yang harganya 0 masukkan ke "bonus" pada barang yang sama bila jelas sama barangnya.',
  'ppn_termasuk = true bila harga di nota sudah termasuk PPN (mis. total = DPP + PPN tanpa tambahan, atau tertulis termasuk/include PPN).',
  'ppn_termasuk = false bila PPN ditambahkan di bawah subtotal sehingga total > subtotal.',
  'Jangan mengarang; isi "" atau 0 bila tidak terbaca. Gabungkan semua halaman jadi satu nota.'
].join('\n');

/** Susun draft: hitung harga beli, cocokkan SKU, isi HET dari price list. */
function buildDraft_(ss, ai) {
  const st = settings_(ss);
  const kode = kodeVendor_(ss), prod = produkIndex_(ss), pl = priceList_(ss);
  // vendor: yang pernah dipakai → nama di Price List Vendor (diisi pemilik, persis Odoo) → nama di nota apa adanya
  const plVend = pl.map(function (r) { return r.vendor; }).filter(function (v, i, a) { return a.indexOf(v) === i; });
  const vend = matchVendor_(vendors_(ss), ai.vendor) || matchVendor_(plVend, ai.vendor) || String(ai.vendor || '').trim();
  const ppnIn = ai.ppn_termasuk !== false;
  const ppn = ppnIn ? 0 : (num_(ai.ppn_persen) || st.ppn * 100) / 100;
  const vpre = vend ? vendorPrefixes_(ss, vend) : [];
  const brandOf = {}; prod.list.forEach(function (p) { if (!brandOf[p.pre]) brandOf[p.pre] = p.brand; });
  const vBrand = vpre.length === 1 ? brandOf[vpre[0]] || '' : '';

  const lines = (ai.baris || []).map(function (b, i) {
    const harga = num_(b.harga), disk = (b.diskon || []).map(num_);
    let net = harga; disk.forEach(function (d) { net *= (1 - d / 100); });
    const beli = Math.round(net * (1 + ppn) * 100) / 100;
    const m = matchLine_(vend, b, kode, prod, vpre);
    const L = { no: i + 1, kode: String(b.kode || ''), namaNota: String(b.nama || ''), qty: num_(b.qty), bonus: num_(b.bonus),
      harga: harga, diskon: disk, beli: beli, jumlahNota: num_(b.jumlah),
      status: m.status, sku: m.sku || '', kandidat: m.kandidat || [],
      baru: { nama: vBrand ? (norm_(b.nama).indexOf(norm_(vBrand)) === 0 ? titleCase_(b.nama) : vBrand + ' ' + titleCase_(b.nama)) : titleCase_(b.nama),
              brand: vBrand, prefix: vBrand ? vpre[0] : '', barcode: String(b.barcode || '').trim(), het: '', hetSumber: '' },
      ed: edAwal_(b), tanpaEd: !!(b.eds && b.eds.tanpa), hargaJualBaru: '' };
    if (m.status !== 'Restock') {
      const h = hetFromPriceList_(pl, vend, harga, b.nama);
      if (h) { L.baru.het = h.het; L.baru.hetSumber = h.nama; }
    }
    return L;
  });
  // ---- alur "brand baru": semua baris langsung Baru dengan brand + prefix dari atas ----
  let bb = null;
  if (ai.brandBaru && ai.brandBaru.brand) {
    const brand = String(ai.brandBaru.brand).trim(), pre = String(ai.brandBaru.prefix || '').trim().toUpperCase();
    if (!/^[A-Z0-9]{2,6}$/.test(pre)) throw new Error('Prefix SKU harus 2–6 huruf/angka (mis. SKF, WRD).');
    const ada = prod.list.filter(function (p) { return p.brand.toLowerCase() === brand.toLowerCase(); })[0];
    if (ada) throw new Error('Brand "' + brand + '" sudah ada di Master Data (prefix ' + ada.pre + '). Pakai alur "Brand sudah ada" / "Distributor baru".');
    const pakai = prod.list.filter(function (p) { return p.pre === pre; })[0];
    if (pakai) throw new Error('Prefix ' + pre + ' sudah dipakai brand "' + pakai.brand + '". Pilih prefix lain.');
    bb = { brand: brand, prefix: pre };
    lines.forEach(function (L) {
      L.status = 'Baru'; L.sku = ''; L.kandidat = [];
      L.baru.brand = brand; L.baru.prefix = pre;
      L.baru.nama = norm_(L.namaNota).indexOf(norm_(brand)) === 0 ? titleCase_(L.namaNota) : brand + ' ' + titleCase_(L.namaNota);
      if (!L.baru.het) { const h = hetFromPriceList_(pl, vend, L.harga, L.namaNota); if (h) { L.baru.het = h.het; L.baru.hetSumber = h.nama; } }
    });
  }
  return { vendor: vend, vendorNota: ai.vendor || '', noFaktur: ai.no_faktur || '', tanggal: ai.tanggal || '',
           ppnTermasuk: ppnIn, ppnPersen: ppnIn ? 0 : ppn * 100, totalNota: num_(ai.total_nota), lines: lines, brandBaru: bb };
}

/** ED dari tabel tempel: satu/lebih tanggal (+qty). Qty kosong → dibagi otomatis ke ED terakhir. */
function edAwal_(b) {
  const tot = num_(b.qty) + num_(b.bonus);
  const list = (b.eds && b.eds.list && b.eds.list.length) ? b.eds.list : (b.ed ? [{ tgl: String(b.ed), qty: null }] : []);
  if (!list.length) return [{ tgl: '', qty: tot }];
  if (list.length === 1) return [{ tgl: list[0].tgl, qty: tot }];
  let isi = 0; list.forEach(function (e) { if (e.qty) isi += num_(e.qty); });
  const kosong = list.filter(function (e) { return !e.qty; });
  return list.map(function (e) {
    return { tgl: e.tgl, qty: e.qty ? num_(e.qty) : (kosong.length === 1 ? Math.max(0, tot - isi) : 0) };
  });
}

/** ================= PRICE LIST VENDOR (tempel dari AI) ================= **/
function bm_savePriceList(rows, pin) {
  cekPin_(pin);
  const ss = master_(); ensureTabs_(ss);
  const sh = ss.getSheetByName(BM.T_PRICELIST);
  const n = Math.max(0, sh.getLastRow() - 1);
  const v = n ? sh.getRange(2, 1, n, 5).getValues() : [];
  const pos = {}; v.forEach(function (r, i) { pos[(String(r[0]).trim() + '|' + norm_(r[1])).toLowerCase()] = i; });
  let baru = 0, upd = 0; const add = [], vend = {}, tambah = {};
  (rows || []).forEach(function (r) {
    const row = [String(r[0]).trim(), String(r[1]).trim(), num_(r[2]), num_(r[3]) || '', String(r[4] || '').trim()];
    if (!row[0] || !row[1] || !row[2]) return;
    vend[row[0]] = 1;
    const k = (row[0] + '|' + norm_(row[1])).toLowerCase();
    if (k in pos) { v[pos[k]] = row; upd++; }
    else if (k in tambah) { add[tambah[k]] = row; }
    else { tambah[k] = add.length; add.push(row); baru++; }
  });
  if (n) sh.getRange(2, 1, n, 5).setValues(v);
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 5).setValues(add);
  return { baru: baru, update: upd, vendor: Object.keys(vend) };
}

/** ================= SIMPAN / KIRIM / SETUJUI ================= **/
function bm_getNota(id) { const r = findNota_(master_(), id); return { id: id, status: r.status, draft: r.draft }; }

function bm_save(id, draft, kirim) {
  const ss = master_();
  guardNota_(ss, id, draft);
  if (!kirim) { saveNota_(ss, id, draft, 'Draft'); return { status: 'Draft' }; }
  validate_(draft);
  cekBarcodeDobel_(ss, draft);
  const alasan = alasanApproval_(ss, draft);
  if (alasan.length) {
    let lama = ''; try { lama = findNota_(ss, id).status; } catch (e) {}
    saveNota_(ss, id, draft, 'Menunggu Persetujuan');
    if (lama !== 'Menunggu Persetujuan') emailPemilik_(id, draft, alasan);   // tidak dobel bila staf klik Kirim lagi
    return { status: 'Menunggu Persetujuan' };
  }
  return generate_(ss, id, draft);
}

function bm_approve(id, draft, pin) {
  const ok = PropertiesService.getScriptProperties().getProperty('OWNER_PIN');
  if (!ok) throw new Error('OWNER_PIN belum diisi di Script Properties.');
  if (String(pin).trim() !== String(ok).trim()) throw new Error('PIN salah.');
  guardNota_(master_(), id, draft);
  validate_(draft);
  cekBarcodeDobel_(master_(), draft);
  draft.lines.forEach(function (l) {
    if (l.status === 'Baru' && !num_(l.baru.het)) throw new Error('HET produk baru "' + (l.baru.nama || l.namaNota) + '" belum diisi.');
  });
  return generate_(master_(), id, draft);
}

/** Nota dari tabel yang ditempel (hasil Claude/Gemini chat atau isi manual) — tanpa API AI. */
function bm_fromTable(t) {
  if (!t || !t.baris || !t.baris.length) throw new Error('Tabel kosong — tidak ada baris barang.');
  const ss = master_();
  ensureTabs_(ss);
  const draft = buildDraft_(ss, t);
  draft.foto = [];
  const id = 'N' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd-HHmmss');
  saveNota_(ss, id, draft, 'Draft');
  const dup = fakturDipakai_(ss, id, draft);
  return { id: id, draft: draft, dup: dup };
}

/** Pengaman: nota yang sudah dibuatkan file tidak boleh diproses lagi, dan satu no faktur hanya sekali. */
function guardNota_(ss, id, d) {
  try {
    const cur = findNota_(ss, id);
    if (cur.status === 'File Dibuat' || cur.status === 'Selesai')
      throw new Error('Nota ini sudah berstatus "' + cur.status + '" — file import sudah pernah dibuat. Pakai 🔄 Cek vs Master Data untuk revisi, jangan disetujui ulang.');
  } catch (e) { if (/sudah berstatus/.test(e.message)) throw e; }
  const dup = fakturDipakai_(ss, id, d);
  if (dup) throw new Error('No faktur ' + d.noFaktur + ' dari ' + d.vendor + ' sudah diproses di nota ' + dup + '. Hapus nota ini atau pakai nota ' + dup + '.');
}
function fakturDipakai_(ss, id, d) {
  const sh = ss.getSheetByName(BM.T_NOTA); if (sh.getLastRow() < 2 || !d.noFaktur) return '';
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  for (let i = 0; i < v.length; i++) {
    if (v[i][0] !== id && String(v[i][4]).trim() === String(d.noFaktur).trim() && String(v[i][3]) === String(d.vendor)
        && (v[i][2] === 'File Dibuat' || v[i][2] === 'Selesai')) return v[i][0];
  }
  return '';
}

/** Isi HET produk baru yang masih kosong dari tab Price List Vendor (cocok lewat harga nota = HNT). */
function bm_fillHet(d) {
  const pl = priceList_(master_()); let n = 0;
  d.lines.forEach(function (l) {
    if (l.status === 'Restock' || num_(l.baru.het)) return;
    const h = hetFromPriceList_(pl, d.vendor, num_(l.harga), l.namaNota);
    if (h) { l.baru.het = h.het; l.baru.hetSumber = h.nama; n++; }
  });
  return { draft: d, n: n };
}

function bm_deleteNota(id) {
  const ss = master_(), sh = ss.getSheetByName(BM.T_NOTA);
  const v = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 10).getValues();
  for (let i = 0; i < v.length; i++) {
    if (v[i][0] !== id) continue;
    if (['Draft', 'Menunggu Persetujuan'].indexOf(v[i][2]) < 0)
      throw new Error('Nota berstatus "' + v[i][2] + '" tidak bisa dihapus karena Master Data sudah ter-update.');
    try {
      (JSON.parse(v[i][9]).foto || []).forEach(function (u) {
        const m = String(u).match(/[-\w]{25,}/); if (m) DriveApp.getFileById(m[0]).setTrashed(true);
      });
    } catch (e) {}
    sh.deleteRow(i + 2);
    return 'Terhapus';
  }
  throw new Error('Nota ' + id + ' tidak ditemukan.');
}

/** ================= SINKRON NOTA ↔ MASTER DATA ================= **/
const SYNC_FIELDS = { nama: 'Nama produk', barcode: 'Barcode', het: 'HET', jual: 'Harga jual' };

function syncRows_(ss, d) {
  const st = settings_(ss);
  const sh = ss.getSheetByName(CFG.SHEET_PRODUK), c = hmap_(sh);
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const rowOf = {}; vals.forEach(function (r, i) { rowOf[String(r[c[P.sku]])] = i; });
  const out = [];
  d.lines.forEach(function (l, li) {
    const baru = l.status === 'Baru', sku = baru ? l.skuBaru : l.sku;
    if (!sku || rowOf[sku] === undefined) return;               // produk baru yang belum masuk master dilewati
    const r = vals[rowOf[sku]];
    const m = { nama: String(r[c[P.nama]]), barcode: String(r[c[P.barcode]] || ''), het: num_(r[c[P.het]]), jual: num_(r[c[P.jual]]) };
    let n = {};
    if (baru) n = { nama: l.baru.nama, barcode: String(l.baru.barcode || ''), het: num_(l.baru.het),
                    jual: num_(l.jualFinal) || num_(l.hargaJualBaru) || saran_(l.beli, num_(l.baru.het), null, st) };
    else {
      if (num_(l.hetBaru)) n.het = num_(l.hetBaru);
      if (num_(l.hargaJualBaru)) n.jual = num_(l.hargaJualBaru);
    }
    Object.keys(n).forEach(function (f) {
      if (String(n[f]).trim() !== String(m[f]).trim())
        out.push({ i: li, sku: sku, produk: m.nama, field: f, label: SYNC_FIELDS[f], nota: n[f], master: m[f], row: rowOf[sku] });
    });
  });
  return { diffs: out, sh: sh, c: c };
}

function bm_compare(id) {
  const ss = master_(), r = findNota_(ss, id);
  return syncRows_(ss, r.draft).diffs.map(function (x) { delete x.row; return x; });
}

/** choices: [{i, field, pilih:'master'|'nota'}] */
function bm_applySync(id, choices) {
  const ss = master_(), nota = findNota_(ss, id), d = nota.draft;
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const s = syncRows_(ss, d), sh = s.sh, c = s.c;
    const pick = {}; (choices || []).forEach(function (x) { pick[x.i + '|' + x.field] = x.pilih; });
    let keNota = 0, keMaster = 0;
    s.diffs.forEach(function (x) {
      const p = pick[x.i + '|' + x.field]; if (!p) return;
      const l = d.lines[x.i], baru = l.status === 'Baru';
      if (p === 'master') {
        if (x.field === 'nama') l.baru.nama = x.master;
        if (x.field === 'barcode') l.baru.barcode = x.master;
        if (x.field === 'het') { if (baru) l.baru.het = x.master; else l.hetBaru = x.master; }
        if (x.field === 'jual') { l.hargaJualBaru = x.master; if (baru) l.jualFinal = x.master; }
        keNota++;
      } else {
        const col = { nama: P.nama, barcode: P.barcode, het: P.het, jual: P.jual }[x.field];
        sh.getRange(x.row + 2, c[col] + 1).setValue(x.nota);
        if (x.field === 'nama') {                              // label ikut bila masih sama dengan nama lama
          const lc = sh.getRange(x.row + 2, c[P.label] + 1);
          if (String(lc.getValue()) === String(x.master)) lc.setValue(x.nota);
        }
        keMaster++;
      }
    });
    let url = d.fileImport || '';
    if (nota.status === 'File Dibuat') url = d.fileImport = rebuildFiles_(ss, d);
    saveNota_(ss, id, d, nota.status);
    return { keNota: keNota, keMaster: keMaster, url: url, status: nota.status };
  } finally { lock.releaseLock(); }
}

/** Buat ulang file import dari isi nota (tanpa membuat SKU baru / tanpa mengubah master). */
function rebuildFiles_(ss, d) {
  const st = settings_(ss);
  const sh = ss.getSheetByName(CFG.SHEET_PRODUK), c = hmap_(sh);
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const rowOf = {}; vals.forEach(function (r, i) { rowOf[String(r[c[P.sku]])] = i; });
  const impProduk = [['id','name','default_code','barcode','categ_id','list_price','type','is_storable','tracking','use_expiration_date','available_in_pos','taxes_id','supplier_taxes_id']];
  const updHet = [['id','list_price']];
  const impPL = [['id','item_ids/id','item_ids/applied_on','item_ids/product_tmpl_id/id','item_ids/compute_price','item_ids/fixed_price']];
  const manual = [['SKU','Nama Produk','Harga Jual Lama','Harga Jual Baru','Yang perlu dilakukan']];
  const impPO = [['partner_id','partner_ref','order_line/product_id','order_line/product_qty','order_line/price_unit']];
  const cek = [['SKU','Nama Produk','Qty Diterima','Tanggal Kedaluwarsa','Nomor Lot','Qty per Lot']];
  const cetakBc = [['Nama Produk','Barcode','Jumlah Stiker','SKU']];
  d.lines.forEach(function (l) {
    let sku, nama;
    if (l.status === 'Baru') {
      sku = l.skuBaru; nama = l.baru.nama;
      const xmlId = 'aura_prod_' + sku.toLowerCase().replace(/-/g, '_'), het = num_(l.baru.het);
      const jual = num_(l.jualFinal) || num_(l.hargaJualBaru) || saran_(l.beli, het, null, st);
      impProduk.push([xmlId, nama, sku, String(l.baru.barcode), l.baru.brand, het, 'consu', 1,
                      l.tanpaEd ? 'none' : 'lot', l.tanpaEd ? 0 : 1, 1, '', '']);
      impPL.push(['', '', '1_product', '__import__.' + xmlId, 'fixed', jual]);
      if (l.baru.barcodeAuto) cetakBc.push([nama, String(l.baru.barcode), num_(l.qty) + num_(l.bonus), sku]);
    } else {
      sku = l.sku; const r = vals[rowOf[sku]] || []; nama = r[c[P.nama]] || sku;
      if (num_(l.hetBaru) && num_(l.hetBaru) !== num_(l.hetLama) && r[c[P.id]]) updHet.push([String(r[c[P.id]]), num_(l.hetBaru)]);
      const baruJ = num_(l.hargaJualBaru);
      if (baruJ && baruJ !== num_(l.jualLama))
        ubahHarga_(st, r[c[P.brand]], sku, nama, l.jualLama, baruJ, String(r[c[P.rule]] || ''), String(r[c[P.id]]), impPL, manual);
    }
    if (num_(l.qty)) impPO.push([impPO.length === 1 ? d.vendor : '', impPO.length === 1 ? d.noFaktur : '', sku, num_(l.qty), l.beli]);
    if (num_(l.bonus)) impPO.push([impPO.length === 1 ? d.vendor : '', impPO.length === 1 ? d.noFaktur : '', sku, num_(l.bonus), 0]);
    const tot = num_(l.qty) + num_(l.bonus);
    if (l.tanpaEd) cek.push([sku, nama, tot, '-', '-', tot]);
    else l.ed.forEach(function (e) { cek.push([sku, nama, tot, e.tgl, lotName_(e.tgl), num_(e.qty)]); });
  });
  if (impPL.length > 1 && !impPL[1][0]) impPL[1][0] = st.pricelistId;
  const lots = lotRows_(d, st);
  const lamaInfo = d.fileInfo || {};
  d.fileInfo = infoFile_({ produk: impProduk, het: updHet, pl: impPL, manual: manual, po: impPO, lot: lots });
  d.fileInfo.vendorBaru = lamaInfo.vendorBaru || ''; d.fileInfo.brandBaru = lamaInfo.brandBaru || [];
  d.fileInfo.bc = cetakBc.length - 1;
  return buildXlsx_('Import ' + d.vendor + ' ' + d.noFaktur.replace(/[\/\\]/g, '-') + ' (revisi ' +
    Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd-MM HH.mm') + ')', [
    ['Import Produk', impProduk], ['Update HET Produk', updHet], ['Import Pricelist', impPL], ['Hapus Promo Lama', manual],
    ['Import PO', impPO], ['Import Lot', lots], ['Ceklis Penerimaan', cek]].concat(cetakBc.length > 1 ? [['Cetak Barcode', cetakBc]] : []));
}

/** Centang langkah panduan di Odoo (disimpan di nota). */
function bm_saveSteps(id, langkah) {
  const ss = master_(), r = findNota_(ss, id);
  r.draft.langkah = langkah || {};
  saveNota_(ss, id, r.draft, r.status);
  return true;
}

/** Isi kolom "ID Aturan Pricelist" di master dari export aturan pricelist Harga Aura.
 *  pairs: [[ID aturan, ID eksternal produk (template)], ...] */
function bm_syncRuleIds(pairs, pin) {
  cekPin_(pin);
  const ss = master_(), sh = ss.getSheetByName(CFG.SHEET_PRODUK), c = hmap_(sh);
  const n = sh.getLastRow() - 1, v = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();
  const map = {}; (pairs || []).forEach(function (p) { if (p[0] && p[1]) map[String(p[1]).trim()] = String(p[0]).trim(); });
  let isi = 0, beda = 0;
  const out = v.map(function (r) {
    const rid = map[String(r[c[P.id]]).trim()];
    if (rid) { if (String(r[c[P.rule]]) !== rid) beda++; isi++; return [rid]; }
    return [r[c[P.rule]]];
  });
  sh.getRange(2, c[P.rule] + 1, n, 1).setValues(out);
  return { cocok: isi, berubah: beda, tidakCocok: Object.keys(map).length - isi };
}
function cekPin_(pin) {
  const ok = PropertiesService.getScriptProperties().getProperty('OWNER_PIN');
  if (!ok) throw new Error('OWNER_PIN belum diisi di Script Properties.');
  if (String(pin).trim() !== String(ok).trim()) throw new Error('PIN salah.');
}

function bm_doneOdoo(id) {
  const ss = master_(), r = findNota_(ss, id);
  const skus = r.draft.lines.filter(function (l) { return l.status === 'Baru'; }).map(function (l) { return l.skuBaru; });
  if (skus.length) {
    const sh = ss.getSheetByName(CFG.SHEET_PRODUK), c = hmap_(sh);
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    v.forEach(function (row, i) {
      if (skus.indexOf(String(row[c[P.sku]])) > -1) sh.getRange(i + 2, c[P.pos] + 1).setValue('Ya');
    });
  }
  saveNota_(ss, id, r.draft, 'Selesai');
  return 'Selesai';
}

/** Produk Baru yang barcodenya sudah ada di Master Data / dobel di nota ini → tolak dengan pesan jelas. */
function cekBarcodeDobel_(ss, d) {
  const prod = produkIndex_(ss), seen = {}, err = [];
  d.lines.forEach(function (l, i) {
    if (l.status !== 'Baru') return;
    const bc = String(l.baru.barcode || '').trim(); if (!bc) return;
    const nm = 'Baris ' + (i + 1) + ' "' + (l.baru.nama || l.namaNota) + '"';
    const sku = prod.byBarcode[bc];
    if (sku) err.push(nm + ': barcode ' + bc + ' sudah dipakai ' + sku + ' — ' + prod.bySku[sku].nama + '. Ubah Jenis ke Restock dan pilih SKU ' + sku + '.');
    else if (seen[bc]) err.push(nm + ': barcode ' + bc + ' sama dengan baris ' + seen[bc] + '.');
    else seen[bc] = i + 1;
  });
  if (err.length) throw new Error('Produk sudah ada di Master Data / barcode dobel:\n• ' + err.join('\n• '));
}

function validate_(d) {
  if (!d.vendor) throw new Error('Vendor belum dipilih.');
  if (!d.noFaktur) throw new Error('No faktur belum diisi.');
  d.lines.forEach(function (l) {
    const nm = l.namaNota;
    if (l.status === 'Ragu') throw new Error('Baris "' + nm + '" masih Ragu — pilih SKU atau tandai Baru.');
    if (l.status === 'Restock' && !l.sku) throw new Error('Baris "' + nm + '" belum ada SKU.');
    if (l.status === 'Baru') {
      if (!l.baru.nama) throw new Error('Nama produk baru untuk "' + nm + '" belum diisi.');
      if (!l.baru.brand) throw new Error('Brand produk baru untuk "' + nm + '" belum diisi.');
      if (!l.baru.barcode && !l.baru.barcodeAuto) throw new Error('Barcode produk baru "' + l.baru.nama + '" belum diisi (atau pilih "buat otomatis").');
      if (!/^[A-Z0-9]{2,6}$/.test(String(l.baru.prefix || '').toUpperCase())) throw new Error('Prefix SKU brand "' + l.baru.brand + '" harus 2–6 huruf/angka (mis. VIV, AVO).');
    }
    if (!l.tanpaEd) {
      let tot = 0;
      l.ed.forEach(function (e) { if (!e.tgl) throw new Error('ED "' + nm + '" belum diisi.'); tot += num_(e.qty); });
      if (tot !== num_(l.qty) + num_(l.bonus)) throw new Error('Jumlah qty per ED "' + nm + '" (' + tot + ') ≠ qty diterima (' + (num_(l.qty) + num_(l.bonus)) + ').');
    }
  });
}

/** Daftar alasan nota butuh persetujuan pemilik (kosong = tidak perlu). */
function alasanApproval_(ss, d) {
  const prod = produkIndex_(ss), out = [];
  d.lines.forEach(function (l) {
    const p = prod.bySku[l.sku];
    if (l.status === 'Baru') { out.push('Produk baru: ' + (l.baru.nama || l.namaNota)); return; }
    const nm = p ? p.nama : l.namaNota;
    if (num_(l.hargaJualBaru)) out.push('Harga jual diubah: ' + nm);
    if (num_(l.hetBaru)) out.push('HET diubah: ' + nm);
    if (p && l.beli > p.beli + 1) out.push('Modal naik: ' + nm + ' (' + rpTxt_(p.beli) + ' → ' + rpTxt_(l.beli) + ')');
  });
  return out;
}
function needsApproval_(ss, d) { return alasanApproval_(ss, d).length > 0; }

function emailPemilik_(id, d, alasan) {
  try {
    const url = ScriptApp.getService().getUrl();
    const li = alasan.slice(0, 30).map(function (a) { return '<li>' + escMail_(a) + '</li>'; }).join('')
      + (alasan.length > 30 ? '<li>… dan ' + (alasan.length - 30) + ' lainnya</li>' : '');
    MailApp.sendEmail({
      to: OWNER_EMAIL,
      subject: '[Aura] Perlu persetujuan: ' + (d.vendor || d.vendorNota || 'Nota') + ' ' + (d.noFaktur || ''),
      htmlBody: '<p>Ada nota barang masuk yang menunggu persetujuan pemilik.</p>'
        + '<p><b>Vendor:</b> ' + escMail_(d.vendor || d.vendorNota) + '<br><b>No faktur:</b> ' + escMail_(d.noFaktur)
        + '<br><b>Tanggal faktur:</b> ' + escMail_(d.tanggal) + '<br><b>Jumlah baris:</b> ' + d.lines.length
        + '<br><b>Total nota:</b> ' + rpTxt_(d.totalNota) + '</p>'
        + '<p><b>Perlu dicek:</b></p><ul>' + li + '</ul>'
        + (url ? '<p><a href="' + url + '">Buka Aura Cosmetics System</a> → 📦 Barang Masuk → buka nota → 🔑 Mode pemilik → Setujui.</p>' : '')
    });
  } catch (e) { console.error('Email pemilik gagal: ' + e.message); }   // email gagal tidak boleh menggagalkan Kirim
}
function escMail_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;'); }
function rpTxt_(n) { return 'Rp ' + Math.round(num_(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }

/** Jalankan SEKALI dari editor untuk memberi izin kirim email. */
function tesEmailPemilik() {
  MailApp.sendEmail(OWNER_EMAIL, '[Aura] Tes email', 'Notifikasi pemilik dari Aura Cosmetics System sudah aktif.');
}

/** ================= GENERATE FILE + UPDATE MASTER ================= **/
function generate_(ss, id, d) {
  const lock = LockService.getScriptLock(); lock.waitLock(30000);
  try {
    const st = settings_(ss);
    const sh = ss.getSheetByName(CFG.SHEET_PRODUK), c = hmap_(sh), H = headers_(sh);
    const n = sh.getLastRow() - 1;
    const vals = sh.getRange(2, 1, n, sh.getLastColumn()).getValues();
    // ---- vendor / brand baru ----
    const preOf = {}, brandOfPre = {};
    vals.forEach(function (r) { const b = String(r[c[P.brand]] || '').trim(), p = String(r[c[P.sku]] || '').split('-')[0];
      if (b && p) { preOf[b.toLowerCase()] = preOf[b.toLowerCase()] || p; brandOfPre[p] = brandOfPre[p] || b; } });
    const vendorBaru = vendors_(ss).indexOf(d.vendor) < 0, brandBaru = {};
    d.lines.forEach(function (l) {
      if (l.status !== 'Baru') return;
      const b = String(l.baru.brand).trim(), p = String(l.baru.prefix || '').toUpperCase();
      const ada = preOf[b.toLowerCase()];
      if (ada && ada !== p) throw new Error('Brand "' + b + '" sudah memakai prefix ' + ada + ', bukan ' + p + '.');
      if (!ada) {
        if (brandOfPre[p]) throw new Error('Prefix ' + p + ' sudah dipakai brand "' + brandOfPre[p] + '". Pilih prefix lain untuk "' + b + '".');
        brandBaru[b] = p;
      }
    });
    const rowOf = {}; vals.forEach(function (r, i) { rowOf[String(r[c[P.sku]])] = i; });
    const used = {}; vals.forEach(function (r) { used[String(r[c[P.sku]])] = 1; });
    const tgl = tglLokal_(d.tanggal);
    const bcAda = {}; vals.forEach(function (r) { const b = String(r[c[P.barcode]] || '').trim(); if (b) bcAda[b] = String(r[c[P.sku]]); });
    d.lines.forEach(function (l) {
      const b = l.status === 'Baru' ? String(l.baru.barcode || '').trim() : '';
      if (b && bcAda[b]) throw new Error('"' + (l.baru.nama || l.namaNota) + '" ditandai Baru, padahal barcode ' + b + ' sudah ada sebagai ' + bcAda[b] + '. Ubah jenisnya ke Restock (SKU ' + bcAda[b] + ').');
    });
    // barcode internal (produk tanpa barcode pabrik): EAN-13 berawalan 20 — dibuat berurutan, unik di master
    const cetakBc = [['Nama Produk','Barcode','Jumlah Stiker','SKU']];
    d.lines.forEach(function (l) {
      if (l.status !== 'Baru' || !l.baru.barcodeAuto || l.baru.barcode) return;
      l.baru.barcode = barcodeInternal_(bcAda); bcAda[l.baru.barcode] = '(nota ini)';
    });

    const impProduk = [['id','name','default_code','barcode','categ_id','list_price','type','is_storable','tracking','use_expiration_date','available_in_pos','taxes_id','supplier_taxes_id']];
    const impPL = [['id','item_ids/id','item_ids/applied_on','item_ids/product_tmpl_id/id','item_ids/compute_price','item_ids/fixed_price']];
    const manual = [['SKU','Nama Produk','Harga Jual Lama','Harga Jual Baru','Yang perlu dilakukan']];
    const updHet = [['id','list_price']];
    const impPO = [['partner_id','partner_ref','order_line/product_id','order_line/product_qty','order_line/price_unit']];
    const cek = [['SKU','Nama Produk','Qty Diterima','Tanggal Kedaluwarsa','Nomor Lot','Qty per Lot']];
    const riwayat = [], newRows = [], upd = [];

    d.lines.forEach(function (l) {
      let sku, nama, het, jual, xmlId;
      if (l.status === 'Baru') {
        const pre = (l.baru.prefix || '').toUpperCase();
        if (!pre) throw new Error('Prefix SKU untuk brand "' + l.baru.brand + '" belum diisi.');
        sku = nextSku_(pre, used); used[sku] = 1; l.skuBaru = sku;
        nama = rapikanNama_(l.baru.nama); l.baru.nama = nama; het = num_(l.baru.het);
        jual = num_(l.hargaJualBaru) || saran_(l.beli, het, null, st); l.jualFinal = jual;
        xmlId = 'aura_prod_' + sku.toLowerCase().replace(/-/g, '_');
        impProduk.push([xmlId, nama, sku, String(l.baru.barcode), l.baru.brand, het, 'consu', 1,
                        l.tanpaEd ? 'none' : 'lot', l.tanpaEd ? 0 : 1, 1, '', '']);
        impPL.push(['', '', '1_product', '__import__.' + xmlId, 'fixed', jual]);
        if (l.baru.barcodeAuto) cetakBc.push([nama, String(l.baru.barcode), num_(l.qty) + num_(l.bonus), sku]);
        const row = H.map(function () { return ''; });
        row[c[P.sku]] = sku; row[c[P.nama]] = nama; row[c[P.brand]] = l.baru.brand; row[c[P.barcode]] = String(l.baru.barcode);
        row[c[P.het]] = het; row[c[P.beli]] = l.beli; row[c[P.jual]] = jual; row[c[P.tgl]] = tgl; row[c[P.dist]] = d.vendor;
        row[c[P.label]] = String(l.baru.label || '').trim() || nama; row[c[P.id]] = '__import__.' + xmlId; row[c[P.track]] = l.tanpaEd ? 'Tidak Dilacak' : 'Menurut Lot';
        row[c[P.pos]] = 'Draft (belum di Odoo)'; row[c[P.note]] = 'Dari nota ' + d.noFaktur;
        newRows.push(row);
      } else {
        sku = l.sku; const i = rowOf[sku];
        if (i === undefined) throw new Error('SKU ' + sku + ' tidak ada di tab Produk.');
        const r = vals[i]; nama = r[c[P.nama]]; het = num_(r[c[P.het]]);
        const lama = num_(r[c[P.jual]]), baruJ = num_(l.hargaJualBaru), hetB = num_(l.hetBaru);
        l.jualLama = lama; l.hetLama = het;
        if (hetB && hetB !== het) {                       // pemilik mengubah HET produk lama
          upd.push([i, c[P.het], hetB]);
          if (r[c[P.id]]) updHet.push([String(r[c[P.id]]), hetB]);
          het = hetB;
        }
        upd.push([i, c[P.beli], l.beli], [i, c[P.tgl], tgl], [i, c[P.dist], d.vendor]);
        if (baruJ && baruJ !== lama) {
          upd.push([i, c[P.jual], baruJ]);
          ubahHarga_(st, r[c[P.brand]], sku, nama, lama, baruJ, String(r[c[P.rule]] || ''), String(r[c[P.id]]), impPL, manual);
        }
      }
      // PO + ceklis + riwayat
      if (num_(l.qty)) impPO.push([impPO.length === 1 ? d.vendor : '', impPO.length === 1 ? d.noFaktur : '', sku, num_(l.qty), l.beli]);
      if (num_(l.bonus)) impPO.push([impPO.length === 1 ? d.vendor : '', impPO.length === 1 ? d.noFaktur : '', sku, num_(l.bonus), 0]);
      const tot = num_(l.qty) + num_(l.bonus);
      if (l.tanpaEd) cek.push([sku, nama, tot, '-', '-', tot]);
      else l.ed.forEach(function (e) { cek.push([sku, nama, tot, e.tgl, lotName_(e.tgl), num_(e.qty)]); });
      riwayat.push([tgl, d.vendor, d.noFaktur, sku, l.namaNota, tot, num_(l.bonus) || '', het, l.beli, 'Nota ' + id]);
    });
    if (impPL.length > 1) impPL[1][0] = st.pricelistId;

    // ---- tulis ke master (hanya sel yang berubah) ----
    upd.forEach(function (u) { sh.getRange(u[0] + 2, u[1] + 1).setValue(u[2]); });
    if (newRows.length) appendProduk_(sh, H, newRows);
    const rw = ss.getSheetByName(BM.T_RIWAYAT);
    rw.getRange(rw.getLastRow() + 1, 1, riwayat.length, riwayat[0].length).setValues(riwayat);
    learnKode_(ss, d);

    // ---- file import ----
    const lots = lotRows_(d, st);
    const url = buildXlsx_('Import ' + d.vendor + ' ' + d.noFaktur.replace(/[\/\\]/g, '-'), [
      ['Import Produk', impProduk], ['Update HET Produk', updHet], ['Import Pricelist', impPL], ['Hapus Promo Lama', manual],
      ['Import PO', impPO], ['Import Lot', lots], ['Ceklis Penerimaan', cek]].concat(cetakBc.length > 1 ? [['Cetak Barcode', cetakBc]] : []));
    d.fileImport = url;
    d.fileInfo = infoFile_({ produk: impProduk, het: updHet, pl: impPL, manual: manual, po: impPO, lot: lots });
    d.fileInfo.vendorBaru = vendorBaru ? d.vendor : '';
    d.fileInfo.brandBaru = Object.keys(brandBaru);
    d.fileInfo.bc = cetakBc.length - 1;
    d.langkah = {};
    // brand baru belum punya promo lama → langsung ikut Harga Aura
    if (d.fileInfo.brandBaru.length) tambahBrandPL_(ss, d.fileInfo.brandBaru);
    saveNota_(ss, id, d, 'File Dibuat');
    return { status: 'File Dibuat', url: url, baru: newRows.length, manual: manual.length - 1 };
  } finally { lock.releaseLock(); }
}

function appendProduk_(sh, H, rows) {
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, H.length).setValues(rows);
  H.forEach(function (h, j) {
    if (FORMULA_COLS.indexOf(h) < 0) return;
    const f = sh.getRange(2, j + 1).getFormulaR1C1();
    if (f) sh.getRange(start, j + 1, rows.length, 1).setFormulaR1C1(f);
  });
  const bc = H.indexOf(P.barcode); if (bc > -1) sh.getRange(start, bc + 1, rows.length, 1).setNumberFormat('@');
}
function learnKode_(ss, d) {
  const sh = ss.getSheetByName(BM.T_KODE), have = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .forEach(function (r) { have[(r[0] + '|' + r[1] + '|' + r[2]).toLowerCase()] = 1; });
  const add = [];
  d.lines.forEach(function (l) {
    const sku = l.status === 'Baru' ? l.skuBaru : l.sku, k = (d.vendor + '|' + l.kode + '|' + l.namaNota).toLowerCase();
    if (sku && (l.kode || l.namaNota) && !have[k]) { add.push([d.vendor, l.kode, l.namaNota, sku, 'Nota ' + d.noFaktur]); have[k] = 1; }
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 5).setValues(add);
}
function buildXlsx_(name, sheets) {
  const tmp = SpreadsheetApp.create(name);
  sheets.forEach(function (s, i) {
    const sh = i === 0 ? tmp.getSheets()[0].setName(s[0]) : tmp.insertSheet(s[0]);
    const data = s[1], w = Math.max.apply(null, data.map(function (r) { return r.length; }));
    const norm = data.map(function (r) { while (r.length < w) r.push(''); return r; });
    for (let j = 0; j < w; j++) {           // kolom teks diformat teks (barcode, SKU, ID); kolom angka dibiarkan angka
      const teks = norm.slice(1).some(function (r) { return typeof r[j] === 'string' && r[j] !== ''; });
      if (teks) sh.getRange(1, j + 1, norm.length, 1).setNumberFormat('@');
    }
    sh.getRange(1, 1, norm.length, w).setValues(norm.map(function (r) {
      return r.map(function (v) { return v instanceof Date ? Utilities.formatDate(v, 'Asia/Jakarta', 'yyyy-MM-dd') : v; });
    }));
    sh.setFrozenRows(1);
  });
  SpreadsheetApp.flush();
  const blob = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName(name + '.xlsx');
  const file = folder_(BM.FOLDER_IMPORT).createFile(blob);
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return file.getUrl();
}

/** ================= PENCOCOKAN ================= **/
function matchLine_(vendor, b, kode, prod, vpre) {
  const kd = String(b.kode || '').trim(), nm = norm_(b.nama);
  const k1 = kode[(vendor + '|k|' + kd).toLowerCase()];
  if (kd && k1) return { status: 'Restock', sku: k1 };
  const k2 = kode[(vendor + '|n|' + nm).toLowerCase()];
  if (nm && k2) return { status: 'Restock', sku: k2 };
  if (/^\d{8,14}$/.test(kd) && prod.byBarcode[kd]) return { status: 'Restock', sku: prod.byBarcode[kd] };
  const bc = String(b.barcode || '').trim();
  if (bc && prod.byBarcode[bc]) return { status: 'Restock', sku: prod.byBarcode[bc] };
  // kemiripan nama → hanya usulan; dibatasi ke brand yang pernah dibeli dari vendor ini
  const t = tokens_(b.nama), scored = [];
  if (t.length) prod.list.forEach(function (p) {
    if (vpre.length && vpre.indexOf(p.pre) < 0) return;
    const s = sim_(t, p.tok); if (s >= 0.5) scored.push({ sku: p.sku, nama: p.nama, skor: Math.round(s * 100) });
  });
  scored.sort(function (a, b2) { return b2.skor - a.skor; });
  return scored.length ? { status: 'Ragu', kandidat: scored.slice(0, 3) } : { status: 'Baru' };
}
// prefix SKU yang pernah dibeli dari vendor ini (dari Kode Vendor + Riwayat Beli)
function vendorPrefixes_(ss, vendor) {
  const set = {};
  [[BM.T_KODE, 1, 4], [BM.T_RIWAYAT, 2, 4]].forEach(function (x) {
    const sh = ss.getSheetByName(x[0]); if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
      if (String(r[x[1] - 1]) === vendor && r[x[2] - 1]) set[String(r[x[2] - 1]).split('-')[0]] = 1;
    });
  });
  return Object.keys(set);
}
function hetFromPriceList_(pl, vendor, harga, nama) {
  const vk = function (v) { return norm_(v).replace(/^(pt|cv|ud|tb|toko)\s+/, ''); };
  const vn = vk(vendor);
  const sameV = function (v) { const x = vk(v); return !!x && !!vn && (x === vn || x.indexOf(vn) > -1 || vn.indexOf(x) > -1); };
  const t = tokens_(nama);
  // 1) harga faktur = HNT, atau = HET (faktur yang menulis HET lalu diskon di bawah)
  let rows = pl.filter(function (r) { return sameV(r.vendor) && (Math.abs(r.hnt - harga) <= 1 || (r.het && Math.abs(r.het - harga) <= 1)); });
  // 2) harga tidak cocok → nama harus hampir sama persis
  if (!rows.length) rows = pl.filter(function (r) { return sameV(r.vendor) && sim_(t, tokens_(r.nama)) >= 0.85; });
  if (!rows.length) return null;
  rows.sort(function (a, b) { return sim_(t, tokens_(b.nama)) - sim_(t, tokens_(a.nama)); });
  return rows[0];
}
function titleCase_(s) {
  return rapikanNama_(String(s || '').toLowerCase().replace(/\b([a-z])/g, function (m) { return m.toUpperCase(); }));
}
function rapikanNama_(s) {   // "20 Ml" → "20ml", spasi ganda / di akhir dibuang
  return String(s || '').replace(/(\d+(?:[.,]\d+)?)\s*(ml|gr|g|gram|mg|kg|l|pcs|oz)\b/gi, function (_, n, u) { return n + u.toLowerCase(); })
    .replace(/\s+/g, ' ').trim();
}
function tokens_(s) { return norm_(s).split(' ').filter(function (w) { return w.length > 1; }); }
function norm_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim(); }
function sim_(a, b) {
  if (!a.length || !b.length) return 0;
  let hit = 0;
  a.forEach(function (w) { if (b.some(function (x) { return x === w || (w.length >= 3 && x.indexOf(w) === 0) || (x.length >= 3 && w.indexOf(x) === 0); })) hit++; });
  return hit / Math.max(a.length, Math.min(b.length, a.length + 2));
}
function matchVendor_(list, s) {
  if (!s) return '';
  const t = tokens_(s).filter(function (w) { return ['pt','cv','tbk'].indexOf(w) < 0; });
  let best = '', sc = 0;
  list.forEach(function (v) { const x = sim_(t, tokens_(v)); if (x > sc) { sc = x; best = v; } });
  return sc >= 0.5 ? best : '';
}

/** ================= HELPERS DATA ================= **/
function master_() { return SpreadsheetApp.openById(CFG.MASTER_ID); }
function hmap_(sh) {
  const m = {}; headers_(sh).forEach(function (h, i) { m[h] = i; });
  Object.keys(P).forEach(function (k) { if (m[P[k]] === undefined) throw new Error('Kolom "' + P[k] + '" tidak ada di tab ' + sh.getName()); });
  return m;
}
function headers_(sh) { return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); }); }
const SET_DEF = [
  ['Margin di atas harga beli', 0.10], ['Margin minimal (Rp)', 2000], ['Pembulatan ke atas (Rp)', 500], ['PPN masukan', 0.11],
  ['Nama pricelist', 'Harga Aura'], ['ID eksternal pricelist', '__import__.aura_pricelist_harga_aura'],
  ['Bulan penghapusan sebelum ED', 6], ['Brand pakai Harga Aura', 'Viva']];
function settings_(ss) {           // dibaca menurut label kolom A; baris yang belum ada ditambahkan otomatis
  const sh = ss.getSheetByName(BM.T_SET), m = {};
  sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 2).getValues().forEach(function (r) { if (r[0]) m[String(r[0]).trim()] = r[1]; });
  const add = SET_DEF.filter(function (x) { return !(x[0] in m); });
  if (add.length) { sh.getRange(sh.getLastRow() + 1, 1, add.length, 2).setValues(add); add.forEach(function (x) { m[x[0]] = x[1]; }); }
  return { margin: num_(m['Margin di atas harga beli']) || 0.1, min: num_(m['Margin minimal (Rp)']) || 2000,
           round: num_(m['Pembulatan ke atas (Rp)']) || 500, ppn: num_(m['PPN masukan']) || 0.11,
           pricelistId: String(m['ID eksternal pricelist'] || ''), bulanHapus: num_(m['Bulan penghapusan sebelum ED']),
           brandPL: String(m['Brand pakai Harga Aura'] || '').split(/[,;]/).map(function (x) { return x.trim().toLowerCase(); }).filter(String) };
}
function saran_(beli, het, mm, st) {
  const m = mm == null ? st.margin : mm;
  const x = Math.ceil(Math.max(beli * (1 + m), beli + st.min) / st.round) * st.round;
  return het ? Math.min(het, x) : x;
}
function produkIndex_(ss) {
  const sh = ss.getSheetByName(CFG.SHEET_PRODUK), c = hmap_(sh);
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const o = { bySku: {}, byBarcode: {}, list: [] };
  v.forEach(function (r) {
    const sku = String(r[c[P.sku]] || ''); if (!sku) return;
    const p = { sku: sku, pre: sku.split('-')[0], brand: String(r[c[P.brand]] || ''), nama: String(r[c[P.nama]]), beli: num_(r[c[P.beli]]), tok: tokens_(r[c[P.nama]]) };
    o.bySku[sku] = p; o.list.push(p);
    const bc = String(r[c[P.barcode]] || ''); if (bc) o.byBarcode[bc] = sku;
  });
  return o;
}
function kodeVendor_(ss) {
  const sh = ss.getSheetByName(BM.T_KODE), o = {};
  if (sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(function (r) {
    if (!r[3]) return;
    if (r[1] !== '') o[(r[0] + '|k|' + String(r[1]).trim()).toLowerCase()] = String(r[3]);
    if (r[2] !== '') o[(r[0] + '|n|' + norm_(r[2])).toLowerCase()] = String(r[3]);
  });
  return o;
}
function priceList_(ss) {
  const sh = ss.getSheetByName(BM.T_PRICELIST);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(function (r) { return r[0] && num_(r[2]); })
    .map(function (r) { return { vendor: String(r[0]), nama: String(r[1]), hnt: num_(r[2]), het: num_(r[3]) }; });
}
function vendors_(ss) {
  const set = {};
  [BM.T_KODE, BM.T_RIWAYAT].forEach(function (t) {
    const sh = ss.getSheetByName(t); if (!sh || sh.getLastRow() < 2) return;
    const col = t === BM.T_KODE ? 1 : 2;
    sh.getRange(2, col, sh.getLastRow() - 1, 1).getValues().forEach(function (r) { if (r[0]) set[String(r[0])] = 1; });
  });
  return Object.keys(set).sort();
}
function nextSku_(pre, used) {
  let mx = 0; Object.keys(used).forEach(function (s) {
    const m = s.match(new RegExp('^' + pre + '-(\\d+)$')); if (m) mx = Math.max(mx, Number(m[1]));
  });
  return pre + '-' + ('000' + (mx + 1)).slice(-4);
}
function lotRows_(d, st) {         // sheet Import Lot: lot = ED (ddmmyy), tanggal penghapusan = ED − N bulan
  const out = [['id', 'name', 'product_id', 'expiration_date', 'removal_date']], seen = {};
  d.lines.forEach(function (l) {
    if (l.tanpaEd) return;
    const sku = l.status === 'Baru' ? l.skuBaru : l.sku; if (!sku) return;
    l.ed.forEach(function (e) {
      if (!e.tgl) return;
      const lot = lotName_(e.tgl), k = sku + '|' + lot; if (seen[k]) return; seen[k] = 1;
      out.push(['aura_lot_' + sku.toLowerCase().replace(/-/g, '_') + '_' + lot, lot, sku, e.tgl + ' 00:00:00',
                st.bulanHapus ? minusBulan_(e.tgl, st.bulanHapus) + ' 00:00:00' : '']);
    });
  });
  return out;
}
function minusBulan_(ymd, n) {
  const m = String(ymd).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return Utilities.formatDate(new Date(Number(m[1]), Number(m[2]) - 1 - n, Number(m[3])), 'Asia/Jakarta', 'yyyy-MM-dd');
}
/** Perubahan harga jual produk lama → SELALU lewat pricelist Harga Aura.
 *  Ada ID aturan → aturan lama diperbarui. Belum ada → aturan baru + pengingat di sheet "Hapus Promo Lama". */
function ubahHarga_(st, brand, sku, nama, lama, baru, rid, tmplId, impPL, promo) {
  tmplId = String(tmplId || '').trim();
  if (!tmplId) { promo.push([sku, nama, lama || '', baru, 'ID Odoo produk kosong di Master Data — isi dulu, lalu ubah Harga Tetap di pricelist Harga Aura.']); return; }
  impPL.push(['', rid || '', '1_product', tmplId, 'fixed', baru]);
  if (rid) return;                                   // sudah ikut Harga Aura → cukup diperbarui
  const sudahPL = st.brandPL.indexOf(String(brand || '').trim().toLowerCase()) > -1;
  promo.push([sku, nama, lama || '', baru, sudahPL
    ? 'Aturan baru dibuat (ID aturan belum disinkron). Cek di Harga Aura: bila ada aturan lama untuk produk ini, hapus yang lama.'
    : 'Hapus / arsipkan promo lama produk ini — harga jualnya sekarang ikut pricelist Harga Aura.']);
}

/** Barcode internal EAN-13: 20 + 10 digit urut + digit cek (awalan 20–29 = khusus pemakaian dalam toko, tidak bentrok dengan barcode pabrik). */
function barcodeInternal_(ada) {
  let max = 0;
  Object.keys(ada).forEach(function (b) { if (/^20\d{11}$/.test(b)) max = Math.max(max, Number(b.slice(2, 12))); });
  const body = '20' + String(max + 1).padStart(10, '0');
  let s = 0; body.split('').forEach(function (x, i) { s += Number(x) * (i % 2 ? 3 : 1); });
  return body + ((10 - s % 10) % 10);
}

function tambahBrandPL_(ss, brands) {
  const sh = ss.getSheetByName(BM.T_SET), v = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 2).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim() === 'Brand pakai Harga Aura') {
    const ada = String(v[i][1] || '').split(/[,;]/).map(function (x) { return x.trim(); }).filter(String);
    brands.forEach(function (b) { if (ada.map(function (x) { return x.toLowerCase(); }).indexOf(b.toLowerCase()) < 0) ada.push(b); });
    sh.getRange(i + 2, 2).setValue(ada.join(', '));
    return;
  }
}
function infoFile_(t) {            // jumlah baris per sheet → panduan melewati langkah yang kosong
  return { produk: t.produk.length - 1, het: t.het.length - 1, pl: t.pl.length - 1, manual: t.manual.length - 1,
           po: t.po.length - 1, lot: t.lot.length - 1 };
}

function tglLokal_(ymd) {   // '2026-10-07' → tanggal lokal (bukan UTC) supaya tidak mundur sehari di sheet
  const m = String(ymd || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : new Date();
}
function lotName_(ymd) { // 2029-06-12 → 120629 (konvensi lot = tanggal ED)
  const m = String(ymd).match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? m[3] + m[2] + m[1].slice(2) : String(ymd);
}
function folder_(name) { const it = DriveApp.getFoldersByName(name); return it.hasNext() ? it.next() : DriveApp.createFolder(name); }

/** ================= TAB NOTA ================= **/
function ensureTabs_(ss) {
  if (!ss.getSheetByName(BM.T_NOTA)) {
    const s = ss.insertSheet(BM.T_NOTA);
    s.appendRow(['ID','Diubah','Status','Vendor','No Faktur','Tgl Faktur','Total Nota','Jumlah Baris','File Import','Data (jangan diedit)']);
    s.setFrozenRows(1);
  }
  if (!ss.getSheetByName(BM.T_PRICELIST)) {
    const s = ss.insertSheet(BM.T_PRICELIST);
    s.appendRow(['Vendor (nama di Odoo)','Nama di Price List','Harga Beli / HNT (Rp)','HET (Rp)','Sumber / Periode']);
    s.setFrozenRows(1);
  }
}
function saveNota_(ss, id, d, status) {
  const sh = ss.getSheetByName(BM.T_NOTA);
  const row = [id, new Date(), status, d.vendor, d.noFaktur, d.tanggal, d.totalNota, d.lines.length, d.fileImport || '', JSON.stringify(d)];
  const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
  const i = ids.indexOf(id);
  if (i > -1) sh.getRange(i + 2, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
}
function findNota_(ss, id) {
  const sh = ss.getSheetByName(BM.T_NOTA);
  const v = sh.getRange(2, 1, Math.max(1, sh.getLastRow() - 1), 10).getValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === id) return { status: v[i][2], draft: JSON.parse(v[i][9]) };
  throw new Error('Nota ' + id + ' tidak ditemukan.');
}
function listNota_(ss) {
  const sh = ss.getSheetByName(BM.T_NOTA), last = sh.getLastRow(); if (last < 2) return [];
  const n = Math.min(30, last - 1);                       // hanya 30 nota terakhir yang dibaca
  return sh.getRange(last - n + 1, 1, n, 10).getValues().reverse().map(function (r) {
    let d = {}; try { d = JSON.parse(r[9]); } catch (e) {}
    return { id: r[0], diubah: r[1] ? Utilities.formatDate(new Date(r[1]), 'Asia/Jakarta', 'dd/MM HH:mm') : '',
             status: r[2], vendor: r[3], no: r[4], tgl: String(r[5]), total: r[6], baris: r[7], file: r[8],
             fileInfo: d.fileInfo || null, langkah: d.langkah || {},
             brands: (d.lines || []).map(function (l) {   // Baru → nama brand; Restock → prefix SKU (diterjemahkan di halaman)
               return l.status === 'Baru' ? String((l.baru && l.baru.brand) || '').trim() : String(l.sku || '').split('-')[0];
             }).filter(String) };
  });
}
function include(name) { return HtmlService.createHtmlOutputFromFile(name).getContent(); }
