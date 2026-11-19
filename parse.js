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
        return { s: i.str.replace(/\s+/g, ' '), x: i.transform[4] / vp.width * 720, w: (i.width || 0) / vp.width * 720,
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
      if (o && o.parentNode && /xfrm$/.test(o.parentNode.nodeName)) return o;
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
  /* gabung potongan teks satu baris; bila jaraknya rapat (potongan kata dari PDF) tanpa spasi */
  function joinParts(parts) {
    parts = parts.slice().sort(function (a, b) { return a.x - b.x; });
    var out = '', prev = null;
    parts.forEach(function (p) {
      if (prev) {
        var gap = prev.w ? p.x - (prev.x + prev.w) : 99;
        out += (gap < 0.9 && !/\s$/.test(prev.s) && !/^\s/.test(p.s)) ? '' : ' ';
      }
      out += p.s; prev = p;
    });
    return out;
  }
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
        t: joinParts(l.parts).replace(/\s+/g, ' ').replace(/\s+([,.)])/g, '$1').trim() };
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

  var KNOWN = /regul|event|bank|kartu|bazaar|bazar|factory|\beac\b|adventure club|tambah/i;
  function catName(c) {
    var s = String(c).toLowerCase();
    if (/regul/.test(s)) return 'Reguler';
    if (/event/.test(s)) return 'Event';
    if (/bank|kartu|cc\b/.test(s)) return 'Banking';
    if (/bazaar|bazar|factory/.test(s)) return 'Bazaar';
    if (/\beac\b|adventure club/.test(s)) return 'EAC';
    if (/tambah/.test(s)) return 'Tambahan';
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
    var ket = kY != null ? L.filter(function (l) { return l.y > kY && (tY == null || l.y < tY - 2) && !/^EIGER ADVENTURE\s*\|/i.test(l.t); }).map(function (l) { return l.t; })
      : L.filter(function (l) { return !/^(Periode|Min|Maks|Berlaku|Produk)\b/i.test(l.t) && !/^SYARAT/i.test(l.t); }).map(function (l) { return l.t; });
    var tips = tY != null ? L.filter(function (l) { return l.y >= tY - 2 && !/^EIGER ADVENTURE\s*\|/i.test(l.t); }).map(function (l) { return l.t; }).join(' ').replace(/^[^A-Za-z]*TIPS\s*KASIR\s*:?\s*/i, '').replace(/\s*EIGER ADVENTURE\s*\|.*$/i, '') : '';

    var ex = [], cur = null;
    R.forEach(function (l) {
      var cm = l.t.match(/^\d?\s*Contoh\s*(\d)\s*[\u2014-]\s*(.*)$/i);
      if (cm) { cur = { n: +cm[1], judul: cm[2].trim(), lines: [] }; ex.push(cur); return; }
      if (!cur || /^\d$/.test(l.t) || /^CONTOH PENERAPAN/i.test(l.t)) return;
      var chk = /^\s*\u2705/.test(l.t), t = l.t.replace(EMO, '');
      if (t) cur.lines.push({ t: t, x: l.x, chk: chk });
    });
    var anyChk = ex.some(function (e) { return e.lines.some(function (l) { return l.chk; }); });
    var contoh = ex.map(function (e) {
      var pe = -1;
      e.lines.forEach(function (l, i) { if (pe < 0 && /^Pesan:/i.test(l.t)) pe = i; });
      var bodyL = pe >= 0 ? e.lines.slice(0, pe) : e.lines, body = bodyL.map(function (l) { return l.t; });
      var pesan = pe >= 0 ? e.lines.slice(pe).map(function (l) { return l.t; }).join(' ').replace(/^Pesan:\s*/i, '') : '';
      var item = body[0] || '', res = body.slice(1).join(' ').replace(/\s*\uD83D\uDCB0\s*/g, ' ').trim();
      var bm = res.match(/(Bayar\s+Rp\s*[\d.]+|\d+\s*[\u00d7x]\s*Rp\s*[\d.]+|Harga normal|Bunga normal)\s*$/i);
      var hasil = bm ? res.slice(0, bm.index).trim() : res, bayar = bm ? bm[1] : '';
      return { n: e.n, judul: e.judul, item: item, hasil: hasil, bayar: bayar, pesan: pesan,
        ok: anyChk ? bodyL.slice(1).some(function (l) { return l.chk; })                       // PPTX: baris hasil diawali ✅
          : (bodyL[0] && bodyL[1] && bodyL[0].x && bodyL[1].x) ? bodyL[1].x >= bodyL[0].x - 6     // PDF: baris hasil berhasil menjorok (ada ikon ✅)
          : !/tidak|ditolak|<|harga normal|bunga normal|^bukan|gagal|belum|lewat|di luar|luar jam|kurang|hangus/i.test(hasil + ' ' + e.judul) };
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

  /* emoji di awal baris (🛒 ✅ 💰 💬 ❌ …) dibuang dari teks; ✅ dipakai untuk menandai contoh berhasil */
  var EMO = /^(?:[\uD800-\uDFFF\u2600-\u27BF\uFE0F\u200D]|\s)+/;
  /* ------------------------------------------- slide detail: tiering & teknis eksekusi */
  var DAYS = /^(SENIN|SELASA|RABU|KAMIS|JUMAT|JUM'AT|SABTU|MINGGU|HARI\s*H|H-\d)$/i;
  function isCaps(t) { var L = String(t).replace(/[^A-Za-z]/g, ''); return L.length >= 4 && L.replace(/[^A-Z]/g, '').length / L.length > 0.7; }
  function slideHead(items) {
    var title = items.filter(function (i) { return i.y < 52 && i.x < 480 && i.h >= 10; }).sort(function (a, b) { return a.x - b.x; });
    var sub = items.filter(function (i) { return i.y < 52 && i.x >= 470 && i.h < 10; }).sort(function (a, b) { return a.x - b.x; });
    var lead = lines(items, 0, 721).filter(function (l) { return l.y >= 52 && l.y < 88 && isCaps(l.t); })[0];
    return { title: joinParts(title).replace(/\s+/g, ' ').trim(), sub: joinParts(sub).replace(/\s+/g, ' ').trim(), lead: lead ? lead.t : '' };
  }
  var NOTE_RE = /^([^A-Za-z0-9]{1,4}\s*)?[A-Z][A-Z &/]{2,30}:/;
  function notesOf(items, yFrom) {
    var out = [], cur = null;
    lines(items, 0, 721).filter(function (l) { return l.y > yFrom && l.y < 392 && !/^EIGER ADVENTURE/i.test(l.t); }).forEach(function (l) {
      var t = l.t.replace(/^[^A-Za-z0-9]+(?=[A-Z])/, ''), m = t.match(/^([A-Z][A-Z &/]{2,30}):\s*(.*)$/);
      if (m) { cur = { label: m[1].trim(), text: m[2] }; out.push(cur); }
      else if (cur) cur.text += (cur.text ? ' ' : '') + t;
      else { cur = { label: '', text: t }; out.push(cur); }
    });
    return out;
  }
  function noteStartY(items, yMin) {
    var L = lines(items, 0, 721).filter(function (l) { return l.y > yMin && NOTE_RE.test(l.t); });
    return L.length ? L[0].y - 2 : 390;
  }
  function rowsIn(items, x0, x1, y0, y1) {
    var R = [];
    items.filter(function (i) { return i.x >= x0 && i.x < x1 && i.y >= y0 && i.y < y1; }).sort(function (a, b) { return a.y - b.y || a.x - b.x; })
      .forEach(function (it) { var r = null; for (var k = 0; k < R.length; k++) if (Math.abs(R[k].y - it.y) <= 4) { r = R[k]; break; } if (r) r.p.push(it); else R.push({ y: it.y, p: [it] }); });
    return R.map(function (r) { return { y: r.y, x: Math.min.apply(null, r.p.map(function (p) { return p.x; })), h: Math.max.apply(null, r.p.map(function (p) { return p.h; })), t: joinParts(r.p).replace(/\s+/g, ' ').trim() }; });
  }
  function readTier(items) {
    var H = slideHead(items), labs = items.filter(function (i) { return /^TIER\s*\d+$/i.test(i.s.trim()); }).sort(function (a, b) { return a.x - b.x; });
    var yTop = Math.min.apply(null, labs.map(function (l) { return l.y; })), yN = noteStartY(items, yTop);
    var tiers = labs.map(function (lb, i) {
      var x0 = i ? (labs[i - 1].x + lb.x) / 2 : 0, x1 = i < labs.length - 1 ? (lb.x + labs[i + 1].x) / 2 : 721;
      var R = rowsIn(items, x0, x1, lb.y - 2, yN).filter(function (r) { return !/^TIER\s*\d+$/i.test(r.t); });
      return { label: lb.s.trim().toUpperCase(), lines: R.map(function (r) { return { t: r.t, big: r.h >= 16 }; }) };
    });
    return { type: 'tier', title: H.title, sub: H.sub, lead: H.lead, tiers: tiers, notes: notesOf(items, yN - 1) };
  }
  function readEksekusi(items) {
    var H = slideHead(items), hd = items.filter(function (i) { return DAYS.test(i.s.trim()) && i.h >= 10; }).sort(function (a, b) { return a.x - b.x; });
    var yTop = Math.min.apply(null, hd.map(function (d) { return d.y; })), yN = noteStartY(items, yTop);
    var days = hd.map(function (d, i) {
      var x0 = d.x - 45, x1 = i < hd.length - 1 ? hd[i + 1].x - 45 : 721;
      var num = items.filter(function (it) { return it.x >= x0 && it.x < d.x && Math.abs(it.y - d.y) < 12 && /^\d{1,2}$/.test(it.s.trim()); })[0];
      var subR = rowsIn(items, d.x - 2, x1, d.y + 4, d.y + 26)[0];
      var body = items.filter(function (it) { return it.x >= x0 && it.x < x1 && it.y > (subR ? subR.y + 4 : d.y + 10) && it.y < yN; });
      var tags = body.filter(function (it) { return it.x < d.x + 8 && /^(\d{1,2}[.:]\d{2}|[A-Z]{2,5}|[A-Z]-\d|H[+-]\d)$/.test(it.s.trim()); }).sort(function (a, b) { return a.y - b.y; });
      var txt = rowsIn(body.filter(function (it) { return tags.indexOf(it) < 0; }), x0, x1, 0, 999);
      var entries = tags.map(function (tg, k) {
        var yA = tg.y - 5, yB = k < tags.length - 1 ? tags[k + 1].y - 5 : 999;
        return { tag: tg.s.trim(), text: txt.filter(function (r) { return r.y >= yA && r.y < yB; }).map(function (r) { return r.t; }).join(' ') };
      });
      if (!tags.length && txt.length) entries.push({ tag: '', text: txt.map(function (r) { return r.t; }).join(' ') });
      return { n: num ? num.s.trim() : String(i + 1), day: d.s.trim().toUpperCase(), sub: subR ? subR.t : '', items: entries };
    });
    return { type: 'eksekusi', title: H.title, sub: H.sub, lead: H.lead, days: days, notes: notesOf(items, yN - 1) };
  }
  function slideKind(items) {
    if (items.filter(function (i) { return /^TIER\s*\d+$/i.test(i.s.trim()); }).length >= 2) return 'tier';
    if (items.filter(function (i) { return DAYS.test(i.s.trim()) && i.h >= 10; }).length >= 2 && !items.some(function (i) { return /SYARAT\s*&?\s*KETENTUAN/i.test(i.s); })) return 'eksekusi';
    return '';
  }
  /* ---- cadangan: slide detail bentuk lain → "info tambahan" (tabel / kartu per kolom / daftar) */
  var SKIP_RE = /RINGKASAN\s+PROGRAM|KALENDER\s+PROMO|TIMELINE\s+PROMO|THANK\s*YOU|TERIMA\s*KASIH|DAFTAR\s+ISI/i;
  function readGeneric(items) {
    var H = slideHead(items), leadRow = lines(items, 0, 721).filter(function (l) { return l.y >= 52 && l.y < 88 && isCaps(l.t); })[0];
    var yN = noteStartY(items, 52);
    var body = items.filter(function (i) {
      return i.y >= 52 && i.y < Math.min(yN, 392) && i.s.trim() && !/^EIGER ADVENTURE/i.test(i.s) && !(leadRow && Math.abs(i.y - leadRow.y) <= 4 && H.lead);
    });
    var out = { type: 'info', title: H.title, sub: H.sub, lead: H.lead, notes: notesOf(items, yN - 1) };
    if (!body.length) return out;
    /* tabel dari PPTX: satu item per baris, sel dipisah " | " */
    var piped = body.filter(function (i) { return / \| /.test(i.s); });
    if (piped.length >= 2) {
      out.table = piped.sort(function (a, b) { return a.y - b.y; }).map(function (i) { return i.s.split(/\s*\|\s*/).map(function (c) { return c.trim(); }); });
      body = body.filter(function (i) { return piped.indexOf(i) < 0; });
      if (!body.length) return out;
    }
    /* kolom = rentang x yang terisi teks, dipisah celah kosong */
    var span = body.map(function (i) { var w = i.w || Math.min(720 - i.x, i.s.length * i.h * 0.5); return { a: i.x, b: i.x + w }; })
      .filter(function (r) { return r.b - r.a < 430; }).sort(function (p, q) { return p.a - q.a; });
    var cols = [];
    span.forEach(function (r) { var c = cols[cols.length - 1]; if (c && r.a <= c.b + 12) c.b = Math.max(c.b, r.b); else cols.push({ a: r.a, b: r.b }); });
    if (cols.length > 8 || cols.length < 2) cols = [{ a: 0, b: 721 }];
    var colOf = function (i) { for (var k = cols.length - 1; k >= 0; k--) if (i.x >= cols[k].a - 6) return k; return 0; };
    var rowsAll = rowsIn(body, 0, 721, 0, 999);
    if (cols.length >= 2 && !out.table) {
      var shared = rowsAll.filter(function (r) {
        var ks = {}; body.forEach(function (i) { if (Math.abs(i.y - r.y) <= 4) ks[colOf(i)] = 1; }); return Object.keys(ks).length >= 2;
      }).length;
      /* kalimat yang menyambung ke baris berikut (diawali huruf kecil) = kartu, bukan tabel */
      var wrapped = body.some(function (i) { return /^[a-z]/.test(i.s.trim()) && body.some(function (j) { return j !== i && colOf(j) === colOf(i) && j.y < i.y && i.y - j.y < 22 && !/[.:!?]$/.test(j.s.trim()); }); });
      if (shared / rowsAll.length >= 0.6 && rowsAll.length >= 2 && !wrapped) {
        out.table = rowsAll.map(function (r) {
          var cells = cols.map(function () { return []; });
          body.forEach(function (i) { if (Math.abs(i.y - r.y) <= 4) cells[colOf(i)].push(i); });
          return cells.map(function (c) { return joinParts(c).replace(/\s+/g, ' ').trim(); });
        });
        return out;
      }
    }
    out.cols = cols.map(function (c, k) {
      var mine = body.filter(function (i) { return colOf(i) === k; });
      return rowsIn(mine, 0, 721, 0, 999).map(function (r) { return { t: r.t, big: r.h >= 14 }; });
    }).filter(function (c) { return c.length; });
    return out;
  }
  var STOP = ['promo', 'disc', 'diskon', 'tiering', 'teknis', 'eksekusi', 'free', 'all', 'items', 'item', 'dan', 'the', 'untuk', 'timeline', 'konten', 'live', 'store'];
  function toks(t) { return String(t || '').toLowerCase().replace(/[^a-z0-9%,]+/g, ' ').split(' ').filter(function (w) { return w.length >= 2 && STOP.indexOf(w) < 0; }); }
  /* tempelkan slide detail ke promo induknya: cocokkan nama, kalau tidak ketemu pakai promo sebelumnya */
  function attachExtras(promos, extras) {
    extras.forEach(function (ex) {
      var want = toks(ex.sub + ' ' + ex.title.replace(/^(teknis eksekusi|tiering|detail)\s*[—–-]?\s*/i, '')), best = null, bs = 0;
      promos.forEach(function (p) { var pt = toks(p.title), sc = want.filter(function (w) { return pt.indexOf(w) >= 0; }).length; if (sc > bs) { bs = sc; best = p; } });
      if (!best || bs < 2) {
        if (ex.type === 'info') { ex.to = ''; return; }                     // info tambahan tanpa promo yang mirip → info umum bulan ini
        var prev = promos.filter(function (p) { return p.slide < ex.slide; }); best = prev[prev.length - 1] || best;
      }
      if (best) { (best.extra = best.extra || []).push(ex); ex.to = best.title; }
    });
  }
  function parseSlides(pages, fileName) {
    var cat = 'Lainnya', period = '', title = '', promos = [], memos = [], unknown = [], extras = [];
    pages.forEach(function (items, idx) {
      var all = lines(items, 0, 721), txt = all.map(function (l) { return l.t; }).join('\n'), U = txt.toUpperCase();
      var big = items.filter(function (i) { return i.h >= 20; }).map(function (i) { return i.s.trim(); }).join(' ').trim();

      var kind = slideKind(items);
      if (kind) { var ex = kind === 'tier' ? readTier(items) : readEksekusi(items); ex.slide = idx + 1; extras.push(ex); return; }
      if (/PETUNJUK\s+TEKNIS/i.test(txt) && !/SYARAT/i.test(U)) {
        var cv = txt.replace(/\s+/g, ' ').match(/Petunjuk\s+Teknis\s+Promo\s+([A-Za-z+& ]{3,20})/i);   // "Petunjuk Teknis Promo Tambahan"
        if (cv && !/^program/i.test(cv[1].trim()) && KNOWN.test(cv[1])) cat = catName(cv[1].trim());
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
        var ft = txt.match(/JUKNIS\s+PROMO[^|\n]*\|\s*([^|\n]+?)\s*$/im);                          // footer "… | PROMO TAMBAHAN"
        var p = readPromo(items, all, ft && KNOWN.test(ft[1]) ? catName(ft[1]) : cat, period);
        p.slide = idx + 1;
        promos.push(p);
        return;
      }
      if (/UPDATE\s+PROMO|\bMEMO\b/i.test(U)) {
        var m2 = readMemo(items, all); m2.slide = idx + 1; memos.push(m2);
        return;
      }
      if (txt.trim() && !SKIP_RE.test(txt) && all.length >= 2) { var gx = readGeneric(items); gx.slide = idx + 1; if (gx.title || (gx.cols || []).length || gx.table) { extras.push(gx); return; } }
      if (txt.trim()) unknown.push({ slide: idx + 1, text: txt.slice(0, 500) });
    });

    var month = parsePeriode(period, null).month;
    if (!period && promos.length && promos[0].ranges[0]) month = promos[0].ranges[0][0].slice(0, 7);
    attachExtras(promos, extras);
    promos.forEach(function (p, i) { p.id = ''; p.sort = i + 1; });
    return { month: month, title: title || 'Juknis Promo', period: period, file: fileName,
      promos: promos, memos: memos, unknown: unknown, extras: extras };
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

  window.Juknis = { parseFile: parseFile, parsePeriode: parsePeriode, catName: catName,
    _raw: async function (file) { var b = await file.arrayBuffer(); return /\.pdf$/i.test(file.name) ? itemsFromPdf(b) : itemsFromPptx(b); } };
})();
