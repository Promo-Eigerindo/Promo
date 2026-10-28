/* =====================================================================
   parse.js — membaca file juknis (PDF atau PPTX) di browser, lalu
   mengubah tiap slide menjadi data promo yang bisa diedit admin.

   Keduanya diubah dulu ke bentuk yang sama: daftar potongan teks
   { s: teks, x, y, h } pada kanvas 720 x 405, supaya satu logika baca
   bisa dipakai untuk PDF maupun PPTX.
   ===================================================================== */
(function () {
  'use strict';

  var PDFJS = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
  var PDFJS_W = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  var JSZIP = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';

  function script(url) {
    return new Promise(function (res, rej) {
      var s = document.createElement('script');
      s.src = url; s.onload = res; s.onerror = function () { rej(new Error('Gagal memuat ' + url + '. Cek koneksi internet.')); };
      document.head.appendChild(s);
    });
  }
  async function pdfLib() {
    if (!window.pdfjsLib) { await script(PDFJS); window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_W; }
    return window.pdfjsLib;
  }
  async function zipLib() { if (!window.JSZip) await script(JSZIP); return window.JSZip; }

  /* ------------------------------------------------------------ PDF */
  async function itemsFromPdf(buf) {
    var lib = await pdfLib();
    var pdf = await lib.getDocument({ data: new Uint8Array(buf) }).promise, pages = [];
    for (var p = 1; p <= pdf.numPages; p++) {
      var pg = await pdf.getPage(p), vp = pg.getViewport({ scale: 1 }), tc = await pg.getTextContent();
      pages.push(tc.items.filter(function (i) { return i.str.trim(); }).map(function (i) {
        return { s: i.str.replace(/\s+/g, ' '), x: i.transform[4] / vp.width * 720,
                 y: (vp.height - i.transform[5]) / vp.height * 405, h: i.height / vp.height * 405 };
      }));
    }
    return pages;
  }

  /* ----------------------------------------------------------- PPTX */
  function tag(el, name) { return Array.prototype.slice.call(el.getElementsByTagName(name)); }
  function textOf(node) { return tag(node, 'a:t').map(function (t) { return t.textContent; }).join('').replace(/\s+/g, ' ').trim(); }
  function offOf(el) {
    var n = el;
    while (n && n.nodeType === 1) {
      var o = n.getElementsByTagName ? n.getElementsByTagName('a:off')[0] : null;
      if (o && o.parentNode && o.parentNode.nodeName === 'a:xfrm') return o;
      n = n.parentNode;
    }
    return null;
  }
  function pptItems(xml, CX, CY) {
    var doc = new DOMParser().parseFromString(xml, 'application/xml');
    var sx = 720 / CX, sy = 405 / CY, ptY = 405 / (CY / 12700), out = [];

    tag(doc, 'p:sp').forEach(function (sp) {
      var o = sp.getElementsByTagName('a:off')[0];
      var x = o ? +o.getAttribute('x') * sx : 0, y = o ? +o.getAttribute('y') * sy : 0, dy = 0;
      tag(sp, 'a:p').forEach(function (p) {
        var s = textOf(p);
        var rPr = p.getElementsByTagName('a:rPr')[0];
        var sz = rPr && rPr.getAttribute('sz') ? +rPr.getAttribute('sz') / 100 : 16;
        var h = Math.max(6, sz * ptY);
        if (s) out.push({ s: s, x: x, y: y + dy, h: h });
        dy += h * 1.35;
      });
    });

    tag(doc, 'a:tbl').forEach(function (tb) {
      var o = offOf(tb);
      var x = o ? +o.getAttribute('x') * sx : 0, y = o ? +o.getAttribute('y') * sy : 0, dy = 0;
      tag(tb, 'a:tr').forEach(function (tr) {
        var cells = tag(tr, 'a:tc').map(function (tc) { return textOf(tc); }).filter(function (v, i, a) { return true; });
        var s = cells.join(' | ').replace(/(\s*\|\s*)+$/, '').trim();
        if (s) out.push({ s: s, x: x, y: y + dy, h: 11 });
        dy += 14;
      });
    });
    return out;
  }
  async function itemsFromPptx(buf) {
    var Z = await zipLib(), zip = await Z.loadAsync(buf);
    var presF = zip.file('ppt/presentation.xml');
    if (!presF) throw new Error('File PPTX tidak terbaca (bukan PowerPoint?).');
    var pres = await presF.async('string');
    var m = pres.match(/<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
    var CX = m ? +m[1] : 12192000, CY = m ? +m[2] : 6858000;

    var files = [];
    var relsF = zip.file('ppt/_rels/presentation.xml.rels');
    if (relsF) {
      var rels = await relsF.async('string'), map = {}, re = /<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g, r;
      while ((r = re.exec(rels))) map[r[1]] = r[2];
      var ri = /<p:sldId[^>]*r:id="([^"]+)"/g, s;
      while ((s = ri.exec(pres))) { var t = map[s[1]]; if (t) files.push('ppt/' + String(t).replace(/^\.\.\//, '')); }
    }
    if (!files.length) {
      files = Object.keys(zip.files).filter(function (f) { return /^ppt\/slides\/slide\d+\.xml$/.test(f); })
        .sort(function (a, b) { return +a.match(/\d+/)[0] - +b.match(/\d+/)[0]; });
    }
    var pages = [];
    for (var i = 0; i < files.length; i++) {
      var f = zip.file(files[i]); if (!f) continue;
      pages.push(pptItems(await f.async('string'), CX, CY));
    }
    return pages;
  }

  /* ------------------------------------------------- baca slide → data */
  function lines(items, x0, x1) {
    var L = [];
    items.filter(function (i) { return i.x >= x0 && i.x < x1; })
      .sort(function (a, b) { return a.y - b.y || a.x - b.x; })
      .forEach(function (it) {
        var l = null;
        for (var k = 0; k < L.length; k++) if (Math.abs(L[k].y - it.y) <= 4) { l = L[k]; break; }
        if (l) l.parts.push(it); else L.push({ y: it.y, parts: [it] });
      });
    return L.map(function (l) {
      return { y: l.y, x: Math.min.apply(null, l.parts.map(function (p) { return p.x; })),
        t: l.parts.sort(function (a, b) { return a.x - b.x; }).map(function (p) { return p.s; })
             .join(' ').replace(/\s+/g, ' ').replace(/\s+([,.)])/g, '$1').trim() };
    }).sort(function (a, b) { return a.y - b.y; });
  }

  var BLN = { jan: 1, januari: 1, feb: 2, februari: 2, mar: 3, maret: 3, apr: 4, april: 4, mei: 5, may: 5,
    jun: 6, juni: 6, jul: 7, juli: 7, agu: 8, agt: 8, ags: 8, agustus: 8, aug: 8, sep: 9, sept: 9, september: 9,
    okt: 10, oktober: 10, oct: 10, nov: 11, november: 11, des: 12, desember: 12, dec: 12 };
  var pad = function (n) { return ('0' + n).slice(-2); };
  function dimOf(m) { var a = m.split('-'); return new Date(+a[0], +a[1], 0).getDate(); }

  /* "1 - 31 OKTOBER 2026", "16–18, 23–25 & 30–31 OKT", "s.d. 8 OKTOBER" */
  function parsePeriode(txt, defMonth) {
    var t = String(txt || '').toLowerCase().replace(/[\u2013\u2014]/g, '-');
    var mo = null, yr = null;
    var mm = t.match(/\b(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember|jan|feb|mar|apr|jun|jul|agu|agt|ags|aug|sep|sept|okt|oct|nov|des|dec)\b\.?\s*(20\d{2})?/);
    if (mm) { mo = BLN[mm[1]]; yr = mm[2] ? +mm[2] : null; }
    var dm = defMonth || new Date().toISOString().slice(0, 7);
    if (!mo) mo = +dm.slice(5);
    if (!yr) yr = +dm.slice(0, 4);
    var m = yr + '-' + pad(mo), last = dimOf(m), ranges = [];
    var sd = t.match(/s\.?\s*d\.?\s*(\d{1,2})/);
    if (sd) return { ranges: [[m + '-01', m + '-' + pad(Math.min(last, +sd[1]))]], month: m, exact: true };
    var head = t.split(/\b(januari|februari|maret|april|mei|juni|juli|agustus|september|oktober|november|desember|jan|feb|mar|apr|jun|jul|agu|agt|ags|aug|sep|sept|okt|oct|nov|des|dec)\b/)[0];
    head.split(/,|&|\bdan\b/).forEach(function (g) {
      var r = g.match(/(\d{1,2})\s*-\s*(\d{1,2})/);
      if (r) { ranges.push([m + '-' + pad(+r[1]), m + '-' + pad(Math.min(last, +r[2]))]); return; }
      var one = g.match(/^\s*(\d{1,2})\s*$/);
      if (one && +one[1] >= 1 && +one[1] <= last) ranges.push([m + '-' + pad(+one[1]), m + '-' + pad(+one[1])]);
    });
    if (!ranges.length) return { ranges: [[m + '-01', m + '-' + pad(last)]], month: m, exact: false };
    return { ranges: ranges, month: m, exact: true };
  }

  function catName(c) {
    var s = String(c).toLowerCase();
    if (/regul/.test(s)) return 'Reguler';
    if (/event/.test(s)) return 'Event';
    if (/bank|kartu|cc\b/.test(s)) return 'Banking';
    if (/bazaar|bazar|factory/.test(s)) return 'Bazaar';
    return c.replace(/\s+/g, ' ').trim().replace(/^./, function (x) { return x.toUpperCase(); }).slice(0, 24) || 'Lainnya';
  }

  function readPromo(items, all, cat, defPeriode) {
    var head = items.filter(function (i) { return i.y < 52 && i.x < 520 && i.h >= 10; })
      .sort(function (a, b) { return a.x - b.x; }).map(function (i) { return i.s; }).join(' ').replace(/\s+/g, ' ').trim();
    if (!head) head = (all[0] || {}).t || 'Promo tanpa judul';
    head = head.replace(/syarat\s*&?\s*ketentuan/ig, '').trim();

    var L = lines(items, 0, 365).filter(function (l) { return l.y > 50 && l.y < 398; });
    var R = lines(items, 365, 721).filter(function (l) { return l.y > 50 && l.y < 398; });

    var f = {};
    L.forEach(function (l) {
      var m = l.t.match(/^(Periode|Min\.?\s*Transaksi|Maks\.?\s*Transaksi|Berlaku di|Produk)\s*:\s*(.*)$/i);
      if (m) f[m[1].toLowerCase().replace(/[^a-z]/g, '')] = m[2].trim();
    });
    var kY = null, tY = null;
    L.forEach(function (l) {
      if (/^KETENTUAN:?$/i.test(l.t)) kY = l.y;
      if (/^TIPS\s*KASIR/i.test(l.t)) tY = l.y;
    });
    var ket = kY != null ? L.filter(function (l) { return l.y > kY && (tY == null || l.y < tY - 2); }).map(function (l) { return l.t; })
      : L.filter(function (l) { return !/^(Periode|Min|Maks|Berlaku|Produk)\b/i.test(l.t) && !/^SYARAT/i.test(l.t); }).map(function (l) { return l.t; });
    var tips = tY != null ? L.filter(function (l) { return l.y >= tY - 2; }).map(function (l) { return l.t; }).join(' ').replace(/^TIPS\s*KASIR:?\s*/i, '') : '';

    var ex = [], cur = null;
    R.forEach(function (l) {
      var cm = l.t.match(/^\d?\s*Contoh\s*(\d)\s*[\u2014-]\s*(.*)$/i);
      if (cm) { cur = { n: +cm[1], judul: cm[2].trim(), lines: [] }; ex.push(cur); return; }
      if (!cur || /^\d$/.test(l.t) || /^CONTOH PENERAPAN/i.test(l.t)) return;
      cur.lines.push(l.t);
    });
    var contoh = ex.map(function (e) {
      var pe = -1;
      e.lines.forEach(function (t, i) { if (pe < 0 && /^Pesan:/i.test(t)) pe = i; });
      var body = pe >= 0 ? e.lines.slice(0, pe) : e.lines;
      var pesan = pe >= 0 ? e.lines.slice(pe).join(' ').replace(/^Pesan:\s*/i, '') : '';
      var item = body[0] || '', res = body.slice(1).join(' ');
      var bm = res.match(/(Bayar\s+Rp\s*[\d.]+|\d+\s*[\u00d7x]\s*Rp\s*[\d.]+|Harga normal|Bunga normal)\s*$/i);
      var hasil = bm ? res.slice(0, bm.index).trim() : res, bayar = bm ? bm[1] : '';
      return { n: e.n, judul: e.judul, item: item, hasil: hasil, bayar: bayar, pesan: pesan,
        ok: !/tidak|ditolak|<|harga normal|bunga normal|^bukan|gagal/i.test(hasil + ' ' + e.judul) };
    });

    var per = f.periode || defPeriode || '';
    var P = parsePeriode(per, null);
    var berlaku = f.berlakudi || '';
    return { cat: cat, title: head, periode: per, ranges: P.ranges, exact: P.exact,
      minTrx: f.mintransaksi || '', maxTrx: f.makstransaksi || '', berlaku: berlaku, produk: f.produk || '',
      ket: ket, tips: tips, contoh: contoh, excel: [],
      site: /bazaar|factory/i.test(berlaku) ? 'bazaar' : /selected|terpilih/i.test(berlaku) ? 'selected' : 'reguler' };
  }

  function readMemo(items, all) {
    var tblY = 999;
    all.forEach(function (l) { if (/^Item\b/i.test(l.t) && l.y < tblY) tblY = l.y; });
    var top = lines(items, 0, 365).filter(function (l) { return l.y > 50 && l.y < tblY - 4; })
      .concat(lines(items, 365, 721).filter(function (l) { return l.y > 50 && l.y < tblY - 4; })).map(function (l) { return l.t; });
    var tabel = all.filter(function (l) { return l.y >= tblY - 1 && l.y < 398; }).map(function (l) { return l.t; });
    return { title: (all[0] || {}).t || 'Update promo', isi: top, tabel: tblY < 999 ? tabel : [], date: '', link: '' };
  }

  function parseSlides(pages, fileName) {
    var cat = 'Lainnya', period = '', title = '', promos = [], memos = [], unknown = [];
    pages.forEach(function (items, idx) {
      var all = lines(items, 0, 721), txt = all.map(function (l) { return l.t; }).join('\n'), U = txt.toUpperCase();
      var big = items.filter(function (i) { return i.h >= 20; }).map(function (i) { return i.s.trim(); }).join(' ').trim();

      if (/PETUNJUK\s+TEKNIS/i.test(txt) && !/SYARAT/i.test(U)) {
        title = title || 'Juknis Program Promo';
        var pm = txt.match(/periode\s*:?\s*([^\n]+)/i);
        if (pm) period = pm[1].trim();
        return;
      }
      if (/^PROMO\b/i.test(big) && !/SYARAT|KETENTUAN:/i.test(U)) {
        var c = big.replace(/^PROMO\s*/i, '').trim();
        if (c) cat = catName(c);
        return;
      }
      if (/SYARAT\s*&?\s*KETENTUAN/i.test(U) || /\bKETENTUAN\s*:/i.test(U)) {
        var p = readPromo(items, all, cat, period);
        p.slide = idx + 1;
        promos.push(p);
        return;
      }
      if (/UPDATE\s+PROMO|\bMEMO\b/i.test(U)) {
        var m2 = readMemo(items, all); m2.slide = idx + 1; memos.push(m2);
        return;
      }
      if (txt.trim()) unknown.push({ slide: idx + 1, text: txt.slice(0, 500) });
    });

    var month = parsePeriode(period, null).month;
    if (!period && promos.length && promos[0].ranges[0]) month = promos[0].ranges[0][0].slice(0, 7);
    promos.forEach(function (p, i) { p.id = ''; p.sort = i + 1; });
    return { month: month, title: title || 'Juknis Promo', period: period, file: fileName,
      promos: promos, memos: memos, unknown: unknown };
  }

  async function parseFile(file) {
    var buf = await file.arrayBuffer();
    var pages = /\.pdf$/i.test(file.name) ? await itemsFromPdf(buf)
      : /\.pptx$/i.test(file.name) ? await itemsFromPptx(buf)
      : (function () { throw new Error('Format belum didukung. Gunakan PDF atau PPTX.'); })();
    if (!pages.length) throw new Error('File tidak berisi slide yang bisa dibaca.');
    var r = parseSlides(pages, file.name);
    r.slides = pages.length;
    return r;
  }

  window.Juknis = { parseFile: parseFile, parsePeriode: parsePeriode, catName: catName };
})();
