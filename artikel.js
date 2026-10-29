/* =====================================================================
   artikel.js — list artikel promo & cek stok toko (VLOOKUP di browser)

   Admin : membaca Excel list artikel (banyak sheet) → list siap simpan.
   Toko  : membaca Export by Bin → dicocokkan dengan list artikel bulan itu
           → ringkasan, detail per SKU/bin, dan ekspor Excel.
   File stok toko TIDAK dikirim ke server; semuanya diproses di perangkat.
   ===================================================================== */
(function () {
  'use strict';
  var XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', _xl = null;
  function needXLSX() {
    if (window.XLSX) return Promise.resolve(true);
    return _xl || (_xl = new Promise(function (res) {
      var s = document.createElement('script'); s.src = XLSX_URL;
      s.onload = function () { res(true); }; s.onerror = function () { _xl = null; res(false); };
      document.head.appendChild(s);
    }));
  }
  var normH = function (v) { return String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9]/g, ''); };
  var tx = function (v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); };
  var num = function (v) { if (typeof v === 'number') return v; var n = Number(String(v == null ? '' : v).replace(/,/g, '').replace(/[^0-9.\-]/g, '')); return isNaN(n) ? 0 : n; };
  function normCode(v) {
    if (v == null) return '';
    if (typeof v === 'number') return isFinite(v) ? String(Math.round(v)) : '';
    return String(v).trim().replace(/\.0+$/, '').replace(/\s+/g, ' ');
  }
  /* kode di list → kunci pencocokan.
     12–13 digit = SKU persis · 9 digit = semua warna & ukuran artikel itu · 9 digit + warna = artikel di warna itu */
  function keyOf(code) {
    var m = String(code || '').match(/^(\d{9,13})(?:[\s._\-\/]+([A-Za-z]{2,4}))?\b/);
    if (!m) return null;
    var d = m[1], c = (m[2] || '').toUpperCase();
    if (d.length >= 12) return { t: 'S', k: d };
    if (d.length === 9) return c ? { t: 'GC', k: d + ':' + c } : { t: 'G', k: d };
    return { t: 'S', k: d };
  }
  var isCodeLike = function (v) { return /^\d{9,13}/.test(normCode(v)); };

  /* ------------------------------------------- list artikel (admin) */
  var CODE_H = ['skucode', 'kodesku', 'sku', 'articlenumber', 'articleno', 'article', 'kodeartikel', 'noartikel', 'artikel', 'genericcode', 'kodegeneric', 'generic',
    'materialcode', 'material', 'itemcode', 'kodebarang', 'plu', 'barcode', 'kode', 'code'];
  var NAME_H = ['articledescription', 'skuname', 'namasku', 'namaartikel', 'articlename', 'namabarang', 'namaproduk', 'productname', 'itemname', 'description', 'deskripsi', 'nama', 'name'];
  var COLOR_H = ['warna', 'color', 'colour', 'kodewarna', 'colorcode'];
  /* kolom harga & diskon dikenali dari judulnya (huruf kecil tanpa spasi/simbol) */
  var PPRICE_RE = /^(hargapromo|promoprice|hargadiskon|hargadisc|discprice|discountprice|hargaakhir|hargasetelah|netprice|nettprice|hargafinal|finalprice|specialprice|hargaspesial|hargajualpromo|hargapromosi|aftersdisc|afterdisc)/;
  var NPRICE_RE = /^(harganormal|normalprice|retailprice|hargaretail|srp|rsp|rrp|retail|price|harga|hargajual|hargaawal|originalprice|hargaasli|hargabarang|hargasatuan|unitprice|hargaitem)$|^(harganormal|normalprice|retailprice|hargaretail)/;
  var DISC_RE = /^(disc|diskon|discount|potongan|persen|percent|pct)|(disc|diskon|discount)$/;
  function detectPriceCols(headers, skip) {
    var h = headers.map(normH), ok = function (i) { return !skip || skip.indexOf(i) < 0; }, f = function (re, not) { for (var i = 0; i < h.length; i++) if (ok(i) && h[i] && re.test(h[i]) && (!not || not.indexOf(i) < 0)) return i; return -1; };
    var ppc = f(PPRICE_RE), pc = f(NPRICE_RE, [ppc]), dc = f(DISC_RE, [ppc, pc]);
    return { dc: dc, pc: pc, ppc: ppc };
  }
  var INFO_RE = /disc|diskon|potongan|harga|price|normal|net|promo|warna|color|size|ukuran|mc|kategori|category|gender|brand|keterangan|remark|status|qty|stok/;

  function findHeader(aoa) {
    for (var i = 0; i < Math.min(aoa.length, 30); i++) {
      var h = (aoa[i] || []).map(normH), ci = -1;
      for (var p = 0; p < CODE_H.length && ci < 0; p++) ci = h.indexOf(CODE_H[p]);
      if (ci < 0) continue;
      var ok = 0;
      for (var j = i + 1; j < Math.min(aoa.length, i + 25); j++) if (isCodeLike((aoa[j] || [])[ci])) ok++;
      if (ok) return { row: i, code: ci };
    }
    /* tanpa judul kolom: pilih kolom yang paling banyak berisi kode 9–13 digit */
    var cnt = {}, first = {};
    for (var r = 0; r < Math.min(aoa.length, 300); r++) (aoa[r] || []).forEach(function (v, c) { if (isCodeLike(v)) { cnt[c] = (cnt[c] || 0) + 1; if (first[c] == null) first[c] = r; } });
    var best = -1; Object.keys(cnt).forEach(function (c) { if (best < 0 || cnt[c] > cnt[best]) best = +c; });
    if (best < 0) return null;
    return { row: first[best] - 1, code: best };
  }
  function readListSheets(wb) {
    var X = window.XLSX, out = [];
    wb.SheetNames.forEach(function (sn) {
      var aoa = X.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '', blankrows: false });
      if (!aoa.length) return;
      var H = findHeader(aoa);
      if (!H) { out.push({ sheet: sn, skip: true, reason: 'tidak ada kolom kode artikel/SKU' }); return; }
      var width = 0; aoa.slice(Math.max(H.row, 0), Math.max(H.row, 0) + 50).forEach(function (r) { width = Math.max(width, r.length); });
      var hdr = [];
      for (var c = 0; c < width; c++) { var t = H.row >= 0 ? tx((aoa[H.row] || [])[c]) : ''; hdr.push(t || ('Kolom ' + String.fromCharCode(65 + (c % 26)) + (c >= 26 ? Math.floor(c / 26) : ''))); }
      var nh = hdr.map(normH), nameCol = -1, colorCol = -1;
      for (var p = 0; p < NAME_H.length && nameCol < 0; p++) { var k = nh.indexOf(NAME_H[p]); if (k >= 0 && k !== H.code) nameCol = k; }
      for (p = 0; p < COLOR_H.length && colorCol < 0; p++) { var q = nh.indexOf(COLOR_H[p]); if (q >= 0 && q !== H.code) colorCol = q; }
      var body = aoa.slice(H.row + 1), coded = body.filter(function (r) { return isCodeLike(r[H.code]); });
      var info = [], skipCol = function (i) { return i === H.code || i === nameCol || i === colorCol || /^Kolom /.test(hdr[i]) || /^(no|nomor|nourut|urut|number)$/.test(nh[i]); };
      hdr.forEach(function (h, i) { if (!skipCol(i) && INFO_RE.test(nh[i])) info.push(i); });
      hdr.forEach(function (h, i) { if (info.length < 6 && !skipCol(i) && info.indexOf(i) < 0 && coded.some(function (r) { return tx(r[i]); })) info.push(i); });
      info = info.slice(0, 8).sort(function (a, b) { return a - b; });
      var PC = detectPriceCols(hdr, [H.code, nameCol, colorCol]);
      out.push({ sheet: sn, headers: hdr, hdrRow: H.row, codeCol: H.code, nameCol: nameCol, colorCol: colorCol, infoCols: info, body: body, count: coded.length,
        discCol: PC.dc, priceCol: PC.pc, ppriceCol: PC.ppc, dfix: '' });
    });
    return out;
  }
  function cellOut(v) {
    if (typeof v === 'number') return Math.round(v * 10000) / 10000;
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return tx(v);
  }
  /* sheet terbaca → list siap simpan: rows = [kode, nama, info...] */
  function buildList(S, name, promoId) {
    var seen = {}, rows = [], cols = S.infoCols.slice();
    [S.discCol, S.priceCol, S.ppriceCol].forEach(function (c) { if (c != null && c >= 0 && cols.indexOf(c) < 0 && c !== S.codeCol && c !== S.nameCol) cols.push(c); });
    cols.sort(function (a, b) { return a - b; });
    S.body.forEach(function (r) {
      var code = normCode(r[S.codeCol]); if (!/^\d{9,13}/.test(code)) return;
      var col = S.colorCol >= 0 ? tx(r[S.colorCol]).toUpperCase() : '';
      if (/^\d{9}$/.test(code) && /^[A-Z]{2,4}$/.test(col)) code = code + '.' + col;
      if (seen[code]) return; seen[code] = 1;
      rows.push([code, S.nameCol >= 0 ? tx(r[S.nameCol]) : ''].concat(cols.map(function (i) { return cellOut(r[i]); })));
    });
    var at = function (c) { return c >= 0 ? cols.indexOf(c) : -1; };
    return { name: name || S.sheet, sheet: S.sheet, promoId: promoId || '', cols: cols.map(function (i) { return S.headers[i]; }), rows: rows,
      dc: at(S.discCol), pc: at(S.priceCol), ppc: at(S.ppriceCol), dfix: parseDisc(S.dfix) || 0 };
  }
  function parseMoney(v) {
    if (typeof v === 'number') return v > 0 ? Math.round(v) : 0;
    var s = String(v == null ? '' : v).replace(/rp\.?/ig, '').replace(/\s/g, '');
    if (!s) return 0;
    if (/^\d{1,3}([.,]\d{3})+([.,]\d{1,2})?$/.test(s)) s = s.replace(/[.,]\d{1,2}$/, function (m) { return m.length === 4 ? m : ''; }).replace(/[.,]/g, '');
    var n = parseFloat(s.replace(',', '.'));
    return isNaN(n) || n <= 0 ? 0 : Math.round(n);
  }
  /* 0.5 → 50 · "50%" → 50 · 50 → 50 · "Disc 30%" → 30 */
  function parseDisc(v) {
    if (v == null || v === '') return 0;
    if (typeof v === 'number') { var x = v > 0 && v <= 1 ? v * 100 : v; return x > 0 && x < 100 ? Math.round(x * 10) / 10 : 0; }
    var str = String(v), m = str.replace(',', '.').match(/(\d+(?:\.\d+)?)/);
    if (!m) return 0;
    var n = +m[1]; if (n > 0 && n <= 1 && str.indexOf('%') < 0) n *= 100;
    return n > 0 && n < 100 ? Math.round(n * 10) / 10 : 0;
  }
  var r100 = function (x) { return Math.round(x / 100) * 100; };
  function listPriceCols(l) {
    if (l._pcx) return l._pcx;
    var P = l.dc != null ? { dc: l.dc, pc: l.pc, ppc: l.ppc } : detectPriceCols(l.cols || []);   // list lama: ditebak dari judul kolom
    return (l._pcx = P);
  }
  /* diskon & harga untuk satu baris list. normal = harga normal, promo = harga setelah diskon */
  function priceOf(l, r, katNormal) {
    var P = listPriceCols(l), d = P.dc >= 0 ? parseDisc(r[P.dc + 2]) : 0, n = P.pc >= 0 ? parseMoney(r[P.pc + 2]) : 0, pp = P.ppc >= 0 ? parseMoney(r[P.ppc + 2]) : 0, src = n ? 'list' : '';
    if (!d && l.dfix) d = +l.dfix;
    if (!n && katNormal) { n = katNormal; src = 'katalog'; }
    if (n && pp && !d && pp < n) d = Math.round((1 - pp / n) * 1000) / 10;
    if (n && d && !pp) pp = r100(n * (1 - d / 100));
    if (pp && n && pp >= n) pp = 0;
    return { disc: d, normal: n, promo: pp, src: src };
  }
  /* harga terbaik per SKU di hasil cek (diskon terbesar di antara list yang cocok) */
  function attachPrices(R, lists, katFn) {
    R.found.forEach(function (x) {
      var kn = katFn ? katFn(x.it) : 0, best = null;
      x.hits.forEach(function (h) {
        var l = lists[h[0]], p = priceOf(l, l.rows[h[1]], kn); p.list = l.name;
        if (!best || p.disc > best.disc || (p.disc === best.disc && p.promo && !best.promo)) best = p;
      });
      x.pr = best;
    });
    return R;
  }
  function keyStats(rows) {
    var s = { S: 0, G: 0, GC: 0 };
    rows.forEach(function (r) { var k = keyOf(r[0]); if (k) s[k.t]++; });
    return s;
  }

  /* ------------------------------------------------- Export by Bin (toko) */
  function readByBin(wb) {
    var X = window.XLSX;
    for (var si = 0; si < wb.SheetNames.length; si++) {
      var aoa = X.utils.sheet_to_json(wb.Sheets[wb.SheetNames[si]], { header: 1, raw: true, defval: '', blankrows: false });
      for (var i = 0; i < Math.min(aoa.length, 20); i++) {
        var h = (aoa[i] || []).map(normH), C = {};
        ['skucode', 'skuname', 'genericcode', 'bincode', 'zonecode', 'mclv2', 'mclv3', 'mclv4', 'brand', 'bincategory', 'totalqty', 'availableqty', 'location'].forEach(function (k) { C[k] = h.indexOf(k); });
        if (C.skucode < 0 || C.totalqty < 0) continue;
        var items = {}, loc = '', rows = 0, pcs = 0;
        for (var r = i + 1; r < aoa.length; r++) {
          var R = aoa[r] || [], sku = normCode(R[C.skucode]); if (!sku) continue;
          var q = num(R[C.totalqty]), zone = tx(R[C.zonecode]), cat = tx(R[C.bincategory]), bin = tx(R[C.bincode]);
          if (!loc && C.location >= 0) loc = tx(R[C.location]);
          var it = items[sku];
          if (!it) {
            var nm = tx(R[C.skuname]), parts = nm.split(','), g = normCode(R[C.genericcode]);
            if (!/^\d{9}$/.test(g)) g = /^\d{12,13}$/.test(sku) ? sku.slice(0, 9) : sku;
            it = items[sku] = { sku: sku, gen: g, name: nm, base: tx(parts[0]), color: tx(parts[1]).toUpperCase(), size: tx(parts.slice(2).join(',')),
              mc2: tx(R[C.mclv2]), mc3: tx(R[C.mclv3]), mc4: tx(R[C.mclv4]), brand: tx(R[C.brand]), disp: 0, wh: 0, dmg: 0, oth: 0, bins: [] };
          }
          var grp = /damage/i.test(zone) || /damage/i.test(cat) ? 'dmg' : /sales|pos|display/i.test(zone) || /^S[A-Z]-/i.test(bin) ? 'disp' : /warehouse/i.test(zone) ? 'wh' : 'oth';
          it[grp] += q; it.bins.push({ bin: bin, zone: zone, qty: q, g: grp }); rows++; pcs += q;
        }
        if (!rows) throw new Error('File Export by Bin tidak berisi baris stok.');
        var m = loc.match(/^(\d{3,6})\s*-\s*(.+)$/);
        return { items: items, rows: rows, pcs: pcs, loc: loc, siteCode: m ? m[1] : '', siteName: m ? m[2].replace(/\s+/g, ' ') : loc, sheet: wb.SheetNames[si] };
      }
    }
    throw new Error('Kolom SKU Code dan Total Qty tidak ditemukan. Pastikan ini file Export by Bin dari sistem.');
  }

  /* VLOOKUP: tiap SKU di stok dicocokkan ke semua list (SKU persis → artikel+warna → artikel) */
  function match(stock, lists) {
    var idx = { S: {}, G: {}, GC: {} }, rowHit = lists.map(function (l) { return new Array(l.rows.length).fill(0); });
    lists.forEach(function (l, li) { l.rows.forEach(function (r, ri) { var k = keyOf(r[0]); if (k) (idx[k.t][k.k] = idx[k.t][k.k] || []).push([li, ri]); }); });
    var found = [];
    Object.keys(stock.items).forEach(function (sku) {
      var it = stock.items[sku], seen = {}, hits = [];
      [idx.S[sku], idx.GC[it.gen + ':' + it.color], idx.G[it.gen]].forEach(function (arr) {
        (arr || []).forEach(function (h) { if (seen[h[0]]) return; seen[h[0]] = 1; hits.push(h); rowHit[h[0]][h[1]] += it.disp + it.wh; });
      });
      if (!hits.length) return;
      var st = it.disp > 0 ? 'disp' : it.wh > 0 ? 'wh' : it.dmg > 0 ? 'dmg' : 'nol';
      found.push({ it: it, hits: hits, st: st });
    });
    found.sort(function (a, b) { return (a.hits[0][0] - b.hits[0][0]) || (b.it.wh + b.it.disp) - (a.it.wh + a.it.disp); });
    var per = lists.map(function (l, li) {
      var f = found.filter(function (x) { return x.hits.some(function (h) { return h[0] === li; }); });
      var o = { name: l.name, promoId: l.promoId, listN: l.rows.length, sku: f.length, disp: 0, wh: 0, dmg: 0, belum: 0, kosong: 0 };
      f.forEach(function (x) { o.disp += x.it.disp; o.wh += x.it.wh; o.dmg += x.it.dmg; if (x.st === 'wh') o.belum++; });
      o.kosong = rowHit[li].filter(function (v) { return !v; }).length;
      return o;
    });
    var tot = { sku: found.length, disp: 0, wh: 0, dmg: 0, belum: 0, allSku: Object.keys(stock.items).length };
    found.forEach(function (x) { tot.disp += x.it.disp; tot.wh += x.it.wh; tot.dmg += x.it.dmg; if (x.st === 'wh') tot.belum++; });
    return { found: found, per: per, tot: tot, rowHit: rowHit };
  }

  /* ------------------------------------------------------- ekspor Excel */
  function sheetName(s, used) {
    var n = String(s).replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 28).trim() || 'List', b = n, i = 2;
    while (used[n.toLowerCase()]) n = b.slice(0, 26) + ' ' + (i++);
    used[n.toLowerCase()] = 1; return n;
  }
  var ST_TXT = { disp: 'Terdisplay', wh: 'Di gudang, belum display', dmg: 'Damage saja', nol: 'Stok 0' };
  function exportResult(R, stock, lists, meta) {
    var X = window.XLSX, wb = X.utils.book_new(), used = {};
    var add = function (aoa, name, widths) { var ws = X.utils.aoa_to_sheet(aoa); if (widths) ws['!cols'] = widths.map(function (w) { return { wch: w }; }); X.utils.book_append_sheet(wb, ws, sheetName(name, used)); };
    add([['Cek stok promo'], ['Toko', stock.loc || meta.site], ['Bulan promo', meta.monthLabel], ['File stok', meta.file], ['Diproses', meta.at], [],
      ['List artikel', 'Artikel di list', 'SKU ada di toko', 'Pcs display', 'Pcs gudang', 'Pcs damage', 'SKU belum display', 'Kode list tanpa stok']]
      .concat(R.per.map(function (p) { return [p.name, p.listN, p.sku, p.disp, p.wh, p.dmg, p.belum, p.kosong]; }))
      .concat([['TOTAL', '', R.tot.sku, R.tot.disp, R.tot.wh, R.tot.dmg, R.tot.belum, '']]), 'Ringkasan', [34, 16, 16, 12, 12, 12, 18, 20]);
    var mcN = function (v) { v = tx(v); return !v || v === '-' || /^\d+$/.test(v) ? 'Tanpa MC' : v; }, MC = {};
    R.found.forEach(function (x) {
      var k = [mcN(x.it.mc2), mcN(x.it.mc3), mcN(x.it.mc4)].join('\u0001'), o = MC[k] || (MC[k] = [0, 0, 0, 0]);
      o[0]++; o[1] += x.it.disp; o[2] += x.it.wh; if (x.st === 'wh') o[3]++;
    });
    add([['MC Lv2', 'MC Lv3', 'MC Lv4', 'SKU promo', 'Pcs display', 'Pcs gudang', 'SKU belum display']].concat(Object.keys(MC).sort().map(function (k) { return k.split('\u0001').concat(MC[k]); })),
      'Per MC', [22, 22, 26, 11, 12, 12, 18]);
    var base = ['SKU', 'Generic', 'Nama', 'Warna', 'Size', 'Harga normal', 'Diskon %', 'Harga promo', 'MC Lv2', 'MC Lv3', 'MC Lv4', 'Display', 'Gudang', 'Damage', 'Lainnya', 'Total', 'Status', 'Bin'];
    var rowOf = function (x) { var i = x.it, p = x.pr || {}; return [i.sku, i.gen, i.base, i.color, i.size, p.normal || '', p.disc || '', p.promo || '', i.mc2, i.mc3, i.mc4, i.disp, i.wh, i.dmg, i.oth, i.disp + i.wh + i.dmg + i.oth, ST_TXT[x.st], i.bins.map(function (b) { return b.bin + ' (' + b.qty + ')'; }).join(', ')]; };
    add([['List'].concat(base)].concat(R.found.map(function (x) { return [x.hits.map(function (h) { return lists[h[0]].name; }).join(' | ')].concat(rowOf(x)); })), 'Semua SKU promo', [24, 14, 11, 34, 7, 8, 13, 9, 13, 14, 14, 16, 8, 8, 8, 8, 8, 22, 40]);
    var pb = [['List', 'SKU', 'Nama', 'Bin', 'Zona', 'Qty']];
    R.found.forEach(function (x) { x.it.bins.forEach(function (b) { pb.push([lists[x.hits[0][0]].name, x.it.sku, x.it.name, b.bin, b.zone, b.qty]); }); });
    add(pb, 'Per bin', [24, 14, 40, 14, 12, 7]);
    lists.forEach(function (l, li) {
      var rows = R.found.filter(function (x) { return x.hits.some(function (h) { return h[0] === li; }); });
      add([base.concat(['Kode di list']).concat(l.cols)].concat(rows.map(function (x) { var h = x.hits.filter(function (y) { return y[0] === li; })[0], lr = l.rows[h[1]]; return rowOf(x).concat([lr[0]]).concat(lr.slice(2)); })), l.name);
    });
    var miss = [['List', 'Kode di list', 'Nama']];
    lists.forEach(function (l, li) { l.rows.forEach(function (r, ri) { if (!R.rowHit[li][ri]) miss.push([l.name, r[0], r[1]]); }); });
    add(miss, 'Tidak ada stok', [24, 16, 44]);
    var fn = 'Cek_Stok_Promo_' + (stock.siteCode || meta.site || 'toko') + '_' + meta.month + '.xlsx';
    X.writeFile(wb, fn);
    return fn;
  }
  function exportList(lists, month) {
    var X = window.XLSX, wb = X.utils.book_new(), used = {};
    lists.forEach(function (l) { var ws = X.utils.aoa_to_sheet([['Kode', 'Nama'].concat(l.cols)].concat(l.rows)); ws['!cols'] = [{ wch: 16 }, { wch: 44 }]; X.utils.book_append_sheet(wb, ws, sheetName(l.name, used)); });
    X.writeFile(wb, 'List_Artikel_Promo_' + month + '.xlsx');
  }

  window.Artikel = { needXLSX: needXLSX, readListSheets: readListSheets, buildList: buildList, keyStats: keyStats, keyOf: keyOf,
    readByBin: readByBin, match: match, priceOf: priceOf, attachPrices: attachPrices, listPriceCols: listPriceCols, parseDisc: parseDisc, parseMoney: parseMoney, exportResult: exportResult, exportList: exportList, ST_TXT: ST_TXT };
})();
