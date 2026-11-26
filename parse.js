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
      var its = tc.items.filter(function (i) { return i.str.trim(); }).map(function (i) {
        return { s: i.str.replace(/\s+/g, ' '), x: i.transform[4] / vp.width * 720, w: (i.width || 0) / vp.width * 720,
                 y: (vp.height - i.transform[5]) / vp.height * 405, h: i.height / vp.height * 405 };
      });
      try {
        (await pg.getAnnotations()).forEach(function (a) {
          if (a.subtype !== 'Link' || !a.url || !a.rect) return;
          var x0 = a.rect[0] / vp.width * 720 - 1, x1 = a.rect[2] / vp.width * 720 + 1, yt = (vp.height - a.rect[3]) / vp.height * 405 - 2, yb = (vp.height - a.rect[1]) / vp.height * 405 + 2;
          its.forEach(function (i) { var cx = i.x + (i.w || 0) / 2; if (cx >= x0 && cx <= x1 && i.y >= yt && i.y - i.h <= yb) (i.links = i.links || []).push({ t: i.s.trim(), u: a.url }); });
        });
      } catch (e) {}
      pages.push(its);
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
  function pptItems(xml, CX, CY, rel) {
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
        var lk = [];
        tag(p, 'a:r').forEach(function (r) {
          var hl = r.getElementsByTagName('a:hlinkClick')[0], id = hl && (hl.getAttribute('r:id') || hl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'));
          var t = textOf(r); if (id && rel && rel[id] && t) lk.push({ t: t, u: rel[id] });
        });
        if (s) out.push(lk.length ? { s: s, x: x, y: y + dy, h: h, links: lk } : { s: s, x: x, y: y + dy, h: h });
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
  /* ------------------------------------------------ bentuk slide (untuk tampilan "seperti PPT") */
  /* setiap shape → [x, y, w, h, geom, fill, line, paras, anchor, radius]; paras = [[align, [[teks, gaya, pt, warna, link]]]] */
  var SCHEME = { lt1: 'FFFFFF', bg1: 'FFFFFF', dk1: '000000', tx1: '000000', lt2: 'E7E6E6', dk2: '44546A' };
  function kids(el, name) { return Array.prototype.filter.call(el.childNodes, function (n) { return n.nodeName === name; }); }
  function kid(el, name) { return el ? kids(el, name)[0] || null : null; }
  function colorOf(el) {                       // a:solidFill → 'RRGGBB' / 'RRGGBBAA'
    if (!el) return '';
    var c = kid(el, 'a:srgbClr') || kid(el, 'a:schemeClr'); if (!c) return '';
    var v = c.nodeName === 'a:srgbClr' ? c.getAttribute('val') : SCHEME[c.getAttribute('val')] || '';
    if (!v) return '';
    var al = kid(c, 'a:alpha'); if (al) { var a = Math.round(+al.getAttribute('val') / 100000 * 255); if (a < 250) v += ('0' + a.toString(16)).slice(-2); }
    return v.toUpperCase();
  }
  function parasOf(txBody, rel) {
    if (!txBody) return [];
    var lstSz = null, out = [];
    kids(txBody, 'a:p').forEach(function (p) {
      var pPr = kid(p, 'a:pPr'), al = pPr && pPr.getAttribute('algn'), runs = [];
      Array.prototype.forEach.call(p.childNodes, function (r) {
        if (r.nodeName === 'a:br') { runs.push(['\n', '', 0, '', '']); return; }
        if (r.nodeName !== 'a:r' && r.nodeName !== 'a:fld') return;
        var t = kid(r, 'a:t'); t = t ? t.textContent : ''; if (!t) return;
        var pr = kid(r, 'a:rPr'), st = '', sz = 0, col = '', ln = '';
        if (pr) {
          if (pr.getAttribute('b') === '1') st += 'b';
          if (pr.getAttribute('i') === '1') st += 'i';
          if (pr.getAttribute('u') && pr.getAttribute('u') !== 'none') st += 'u';
          sz = pr.getAttribute('sz') ? +pr.getAttribute('sz') / 100 : 0;
          col = colorOf(kid(pr, 'a:solidFill'));
          var hl = kid(pr, 'a:hlinkClick'), id = hl && (hl.getAttribute('r:id') || hl.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id'));
          if (id && rel && rel[id]) ln = rel[id];
        }
        runs.push([t, st, sz, col, ln]);
      });
      var end = kid(p, 'a:endParaRPr');
      if (!runs.length) runs.push(['', '', end && end.getAttribute('sz') ? +end.getAttribute('sz') / 100 : 0, '', '']);
      out.push([al === 'ctr' ? 'c' : al === 'r' ? 'r' : al === 'just' ? 'j' : '', runs]);
    });
    while (out.length && out[out.length - 1][1].every(function (r) { return !r[0].trim(); })) out.pop();
    return out;
  }
  function pptShapes(xml, CX, CY, rel) {
    var doc = new DOMParser().parseFromString(xml, 'application/xml'), K = 720 / CX, out = [];
    var tree = doc.getElementsByTagName('p:spTree')[0]; if (!tree) return out;
    function walk(node, T) {                                   // T = transform grup: x' = ox + (x - cx) * s
      Array.prototype.forEach.call(node.childNodes, function (el) {
        var n = el.nodeName;
        if (n === 'p:grpSp') {
          var gx = kid(kid(el, 'p:grpSpPr'), 'a:xfrm');
          if (!gx) return walk(el, T);
          var o = kid(gx, 'a:off'), e = kid(gx, 'a:ext'), co = kid(gx, 'a:chOff'), ce = kid(gx, 'a:chExt');
          var sx = ce && +ce.getAttribute('cx') ? +e.getAttribute('cx') / +ce.getAttribute('cx') : 1, sy = ce && +ce.getAttribute('cy') ? +e.getAttribute('cy') / +ce.getAttribute('cy') : 1;
          var ox = +o.getAttribute('x'), oy = +o.getAttribute('y'), cx = co ? +co.getAttribute('x') : ox, cy = co ? +co.getAttribute('y') : oy;
          walk(el, function (x, y, w, h) { var a = [ox + (x - cx) * sx, oy + (y - cy) * sy, w * sx, h * sy]; return T ? T.apply(null, a) : a; });
          return;
        }
        if (n !== 'p:sp' && n !== 'p:cxnSp' && n !== 'p:graphicFrame') return;
        var spPr = kid(el, 'p:spPr'), xf = n === 'p:graphicFrame' ? kid(el, 'p:xfrm') : kid(spPr, 'a:xfrm');
        if (!xf) return;
        var off = kid(xf, 'a:off'), ext = kid(xf, 'a:ext'); if (!off || !ext) return;
        var b = [+off.getAttribute('x'), +off.getAttribute('y'), +ext.getAttribute('cx'), +ext.getAttribute('cy')];
        if (T) b = T(b[0], b[1], b[2], b[3]);
        var x = b[0] * K, y = b[1] * K, w = b[2] * K, h = b[3] * K;
        if (x + w < 1 || y + h < 1 || x > 719 || y > 404) return;                // di luar slide
        var R = function (v) { return Math.round(v * 10) / 10; };
        if (n === 'p:graphicFrame') {                                             // tabel asli → sel-sel
          var tbl = el.getElementsByTagName('a:tbl')[0]; if (!tbl) return;
          var cols = tag(tbl, 'a:gridCol').map(function (g) { return +g.getAttribute('w') * K; }), TR = [], th = 0;
          var tp = kid(tbl, 'a:tblPr'), fRow = tp && tp.getAttribute('firstRow') === '1', bRow = tp && tp.getAttribute('bandRow') === '1';
          tag(tbl, 'a:tr').forEach(function (tr, ri) {
            var rh = +tr.getAttribute('h') * K, cells = []; th += rh;
            kids(tr, 'a:tc').forEach(function (tc) {
              if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') { cells.push(0); return; }
              var pr = kid(tc, 'a:tcPr'), ps = parasOf(kid(tc, 'a:txBody'), rel), cf = pr ? colorOf(kid(pr, 'a:solidFill')) : '';
              var lnB = pr && kid(pr, 'a:lnB'), lc = lnB ? colorOf(kid(lnB, 'a:solidFill')) : '';
              if (!cf && !(pr && kid(pr, 'a:noFill'))) {                         // gaya tabel bawaan → header navy, baris selang-seling
                if (fRow && ri === 0) { cf = '073763'; ps.forEach(function (q) { q[1].forEach(function (r) { if (!r[3]) r[3] = 'FFFFFF'; if (r[1].indexOf('b') < 0) r[1] += 'b'; }); }); }
                else cf = bRow && ri % 2 === 0 ? 'F1F5F9' : 'FFFFFF';
              }
              cells.push([cf, lc || 'E2E8F0', ps, +tc.getAttribute('gridSpan') || 1, +tc.getAttribute('rowSpan') || 1, pr && pr.getAttribute('anchor') === 'ctr' ? 'c' : '']);
            });
            TR.push([R(rh), cells]);
          });
          out.push([R(x), R(y), R(w), R(Math.max(h, th)), 'T', '', '', { c: cols.map(R), r: TR, hd: fRow ? 1 : 0 }, '', 0]);
          return;
        }
        var g = kid(spPr, 'a:prstGeom'), geom = g ? g.getAttribute('prst') : '';
        var G = n === 'p:cxnSp' ? 'l' : geom === 'ellipse' ? 'e' : geom === 'roundRect' ? 'R' : geom === 'rightArrow' || geom === 'chevron' || geom === 'triangle' ? 'a' : 'r';
        var fill = kid(spPr, 'a:noFill') ? '' : colorOf(kid(spPr, 'a:solidFill'));
        var lnE = kid(spPr, 'a:ln'), line = lnE && !kid(lnE, 'a:noFill') ? colorOf(kid(lnE, 'a:solidFill')) : '';
        var lw = lnE && lnE.getAttribute('w') ? Math.max(0.5, +lnE.getAttribute('w') / 12700) : 0.75;
        var rad = 0; if (G === 'R') { var gd = g.getElementsByTagName('a:gd')[0], adj = gd ? +(String(gd.getAttribute('fmla')).replace(/\D/g, '')) : 16667; rad = R(Math.min(w, h) * adj / 100000); }
        var tx = kid(el, 'p:txBody'), bp = tx && kid(tx, 'a:bodyPr'), anc = bp && bp.getAttribute('anchor');
        var ps = parasOf(tx, rel), hasT = ps.some(function (p) { return p[1].some(function (r) { return r[0].trim(); }); });
        if (!hasT) ps = [];
        if (!fill && !line && !ps.length) return;                                 // tidak terlihat
        var ins = bp ? [bp.getAttribute('lIns'), bp.getAttribute('tIns'), bp.getAttribute('rIns'), bp.getAttribute('bIns')].map(function (v, i) { return v == null ? (i % 2 ? 3.6 : 7.2) : R(+v / 12700); }) : 0;
        var sh = [R(x), R(y), R(w), R(h), G, G === 'l' ? '' : fill, line, ps, anc === 'ctr' ? 'c' : anc === 'b' ? 'b' : '', rad];
        if (line && Math.abs(lw - 0.75) > 0.2) sh.push(R(lw)); else sh.push(0);
        if (ins && ins.join() !== '7.2,3.6,7.2,3.6') sh.push(ins);
        out.push(sh);
      });
    }
    walk(tree, null);
    return out;
  }
  /* buang kepala slide (pita judul) & kaki slide; sisanya disimpan ringkas */
  function layoutOf(shapes) {
    if (!shapes || !shapes.length) return null;
    var top = 0, bot = 405;
    shapes.forEach(function (s) { if (s[1] < 4 && s[3] < 70 && s[2] > 600 && s[5]) top = Math.max(top, s[1] + s[3]); });          // pita judul
    shapes.forEach(function (s) { if (s[1] < top + 8 && s[1] >= top - 1 && s[3] <= 8 && s[2] > 600) top = Math.max(top, s[1] + s[3]); }); // garis aksen di bawahnya
    shapes.forEach(function (s) { if (s[1] > 360 && s[2] > 600 && s[5] && s[3] < 40) bot = Math.min(bot, s[1]); });                // pita kaki
    var body = shapes.filter(function (s) { return s[1] >= top - 1 && s[1] + s[3] <= bot + 1 && !(s[1] < top + 2 && s[3] < 2); });
    if (body.filter(function (s) { return s[4] === 'T' || s[7].length; }).length < 2) return null;
    var lo = Math.min.apply(null, body.map(function (s) { return s[1]; })), hi = Math.max.apply(null, body.map(function (s) { return s[1] + s[3]; }));
    return { y0: Math.max(top, lo - 8), y1: Math.min(bot, hi + 10), s: body };            // potong ruang kosong atas/bawah
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
      var rf = zip.file(files[i].replace(/slides\/(slide\d+\.xml)$/, 'slides/_rels/$1.rels')), rel = {};
      if (rf) { var rx = await rf.async('string'), rr, re2 = /<Relationship\b[^>]*>/g; while ((rr = re2.exec(rx))) { var idm = rr[0].match(/Id="([^"]+)"/), tm = rr[0].match(/Target="([^"]+)"/); if (idm && tm && /^https?:|^mailto:/i.test(tm[1])) rel[idm[1]] = tm[1].replace(/&amp;/g, '&'); } }
      var xml = await f.async('string'), its = pptItems(xml, CX, CY, rel);
      try { its.lay = layoutOf(pptShapes(xml, CX, CY, rel)); } catch (e) { its.lay = null; }
      pages.push(its);
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
      var ic = (l.t.match(/^([^A-Za-z0-9]+)(?=[A-Z])/) || [])[1] || '', t = l.t.replace(/^[^A-Za-z0-9]+(?=[A-Z])/, ''), m = t.match(/^([A-Z][A-Z &/]{2,30}):\s*(.*)$/);
      ic = ic.replace(/\s+/g, '');
      if (m) { cur = { label: m[1].trim(), text: m[2] }; if (/[\uD800-\uDFFF\u2600-\u27BF]/.test(ic)) cur.icon = ic; out.push(cur); }
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
  /* ---- alur langkah: lingkaran nomor 1,2,3,… sejajar dalam satu baris → kartu langkah */
  function stepNums(items) {
    var nums = items.filter(function (i) { return /^\d{1,2}$/.test(i.s.trim()) && i.y > 52 && i.h >= 9; });
    var best = null;
    nums.forEach(function (a) {
      var row = nums.filter(function (b) { return Math.abs(b.y - a.y) <= 8; }).sort(function (p, q) { return p.x - q.x; });
      var ok = row.length >= 3 && row.length <= 8 && row.every(function (r, k) { return +r.s.trim() === k + 1; });
      if (ok && (!best || row.length > best.length)) best = row;
    });
    return best;
  }
  function readSteps(items) {
    var H = slideHead(items), N = stepNums(items), yTop = N[0].y, yN = noteStartY(items, yTop);
    var steps = N.map(function (n, k) {
      var x0 = n.x - 12, x1 = k < N.length - 1 ? N[k + 1].x - 12 : 721;
      var R = rowsIn(items.filter(function (i) { return i !== n && /[A-Za-z0-9]/.test(i.s) && !/^EIGER ADVENTURE/i.test(i.s); }), x0, x1, n.y - 4, yN)
        .filter(function (r) { return !/^\d{1,2}$/.test(r.t); });
      var hT = R.length ? R[0].h : 0, minH = Math.min.apply(null, R.map(function (r) { return r.h; }).concat([99]));
      var nt = 1; if (hT > minH + 0.8) { while (nt < R.length && Math.abs(R[nt].h - hT) < 0.8) nt++; }
      return { n: n.s.trim(), title: R.slice(0, nt).map(function (r) { return r.t; }).join(' '), text: R.slice(nt).map(function (r) { return r.t; }).join(' ') };
    });
    return { type: 'steps', title: H.title, sub: H.sub, lead: H.lead, steps: steps, notes: notesOf(items, yN - 1) };
  }
  /* semua link di slide (teks → URL), dipakai untuk membuat teks yang sama bisa diklik di web */
  function slideLinks(items) {
    var seen = {}, out = [];
    items.forEach(function (i) { (i.links || []).forEach(function (l) { var t = l.t.replace(/^[^A-Za-z0-9]+/, '').trim(); if (t && !seen[t + l.u]) { seen[t + l.u] = 1; out.push({ t: t, u: l.u }); } }); });
    return out;
  }
  function slideKind(items) {
    if (items.filter(function (i) { return /^TIER\s*\d+$/i.test(i.s.trim()); }).length >= 2) return 'tier';
    if (items.filter(function (i) { return DAYS.test(i.s.trim()) && i.h >= 10; }).length >= 2 && !items.some(function (i) { return /SYARAT\s*&?\s*KETENTUAN/i.test(i.s); })) return 'eksekusi';
    var tx = items.map(function (i) { return i.s; }).join(' ');
    if (!SKIP_RE.test(tx) && !/SYARAT\s*&?\s*KETENTUAN/i.test(tx) && stepNums(items)) return 'steps';
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
        if (ex.type === 'info' || ex.type === 'steps') { ex.to = ''; return; }                     // info tambahan tanpa promo yang mirip → info umum bulan ini
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
      if (kind) { var ex = kind === 'tier' ? readTier(items) : kind === 'steps' ? readSteps(items) : readEksekusi(items); ex.slide = idx + 1; ex.links = slideLinks(items); if (items.lay) ex.lay = items.lay; extras.push(ex); return; }
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
      if (txt.trim() && !SKIP_RE.test(txt) && all.length >= 2) { var gx = readGeneric(items); gx.slide = idx + 1; gx.links = slideLinks(items); if (items.lay) gx.lay = items.lay; if (gx.title || (gx.cols || []).length || gx.table) { extras.push(gx); return; } }
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
