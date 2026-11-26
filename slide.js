/* =====================================================================
   Tampilan slide "seperti PPT"
   - PC / layar lebar : posisi, warna & ukuran persis seperti slide (diskalakan)
   - HP               : isi slide yang sama disusun ulang ke bawah agar mudah dibaca
   Data: e.lay = { y0, y1, s: [[x, y, w, h, geom, fill, line, paras, anchor, radius, lineW, inset?]] }
   ===================================================================== */
(function () {
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function col(v) {
    if (!v) return '';
    if (v.length === 8) return 'rgba(' + parseInt(v.slice(0, 2), 16) + ',' + parseInt(v.slice(2, 4), 16) + ',' + parseInt(v.slice(4, 6), 16) + ',' + (parseInt(v.slice(6), 16) / 255).toFixed(2) + ')';
    return '#' + v;
  }
  function S(a) { return { x: a[0], y: a[1], w: a[2], h: a[3], g: a[4], fill: a[5], line: a[6], ps: a[4] === 'T' ? [] : a[7] || [], tb: a[4] === 'T' ? a[7] : null, anc: a[8], rad: a[9] || 0, lw: a[10] || 0.75, ins: a[11] || [7.2, 3.6, 7.2, 3.6] }; }
  var URL_RE = /(https?:\/\/[^\s<]+)/g;
  function runHtml(r, unit, minPx) {
    var t = r[0]; if (t === '\n') return '<br>';
    var st = '', sz = r[2];
    if (r[1].indexOf('b') >= 0) st += 'font-weight:700;';
    if (r[1].indexOf('i') >= 0) st += 'font-style:italic;';
    if (r[1].indexOf('u') >= 0) st += 'text-decoration:underline;';
    if (sz) st += 'font-size:' + unit(sz) + ';';
    if (r[3]) st += 'color:' + col(r[3]) + ';';
    var h = esc(t);
    if (r[4]) return '<a href="' + esc(r[4]) + '" target="_blank" rel="noopener" class="xlink" style="' + st + '">' + (/^\s*https?:\/\//.test(t) && t.length > 40 ? 'Buka link ↗' : h) + '</a>';
    h = h.replace(URL_RE, function (u) { return '<a href="' + u + '" target="_blank" rel="noopener" class="xlink">' + (/sharepoint|onedrive|google\./i.test(u) ? 'Buka link ↗' : u) + '</a>'; });
    return st ? '<span style="' + st + '">' + h + '</span>' : h;
  }
  function parasHtml(ps, unit) {
    return ps.map(function (p) {
      var al = p[0] === 'c' ? 'center' : p[0] === 'r' ? 'right' : p[0] === 'j' ? 'justify' : '';
      var inner = p[1].map(function (r) { return runHtml(r, unit); }).join('');
      var big = p[1].reduce(function (m, r) { return Math.max(m, r[2] || 0); }, 0);
      return '<p' + (al ? ' style="text-align:' + al + '"' : '') + '>' + (inner || '&nbsp;') + '</p>';
    }).join('');
  }
  /* tabel asli PowerPoint */
  function tableHtml(T, unit, cls) {
    var tot = T.c.reduce(function (m, v) { return m + v; }, 0) || 1;
    return '<table class="' + cls + '"><colgroup>' + T.c.map(function (w) { return '<col style="width:' + (w / tot * 100).toFixed(2) + '%">'; }).join('') + '</colgroup>'
      + T.r.map(function (row) {
        return '<tr' + '' + '>' + row[1].map(function (c) {
          if (!c) return '';
          return '<td' + (c[3] > 1 ? ' colspan="' + c[3] + '"' : '') + (c[4] > 1 ? ' rowspan="' + c[4] + '"' : '') + ' style="' + (c[0] ? 'background:' + col(c[0]) + ';' : '') + 'border-color:' + col(c[1] || 'E2E8F0') + ';vertical-align:' + (c[5] === 'c' ? 'middle' : 'top') + '">' + parasHtml(c[2], unit) + '</td>';
        }).join('') + '</tr>';
      }).join('') + '</table>';
  }
  function tableFlow(T) {
    var nc = T.c.length;
    if (nc <= 3) return '<div class="fn"><div class="fnt">' + tableHtml(T, px, 'slt-m') + '</div></div>';
    var heads = T.hd && T.r[0] ? T.r[0][1].map(function (c) { return c ? c[2].map(function (p) { return p[1].map(function (r) { return r[0]; }).join(''); }).join(' ') : ''; }) : null;
    return T.r.slice(heads ? 1 : 0).map(function (row) {                 // banyak kolom → satu kartu per baris
      return '<div class="fn fn-box" style="background:' + col((row[1][0] || [])[0] || 'FFFFFF') + ';border:1px solid #E2E8F0">' + row[1].map(function (c, j) {
        return c ? (heads && heads[j] ? '<small class="fn-lbl">' + esc(heads[j]) + '</small>' : '') + '<div class="tx">' + parasHtml(c[2], px) + '</div>' : '';
      }).join('') + '</div>';
    }).join('');
  }
  function hasText(s) { return s.ps.some(function (p) { return p[1].some(function (r) { return r[0] && r[0].trim(); }); }); }

  /* ------------------------------------------------------------ tampilan PC (persis) */
  function absHtml(L) {
    var y0 = L.y0, H = Math.max(40, L.y1 - L.y0);
    var u = function (pt) { return (pt / 7.2).toFixed(3) + 'cqw'; };
    var h = '<div class="sla" style="aspect-ratio:720/' + Math.round(H) + '"><div class="sla-in">';
    L.s.forEach(function (a) {
      var s = S(a);
      var st = 'left:' + (s.x / 7.2).toFixed(3) + '%;top:' + ((s.y - y0) / H * 100).toFixed(3) + '%;width:' + (s.w / 7.2).toFixed(3) + '%;height:' + (s.h / H * 100).toFixed(3) + '%;';
      if (s.g === 'T') { h += '<div class="sh" style="' + st + 'padding:0;font-size:' + u(11) + '">' + tableHtml(s.tb, u, 'slt') + '</div>'; return; }
      if (s.g === 'l') { st += 'border-top:' + u(s.lw) + ' solid ' + col(s.line || '94A3B8') + ';height:0;'; h += '<div class="sh" style="' + st + '"></div>'; return; }
      if (s.fill) st += 'background:' + col(s.fill) + ';';
      if (s.line) st += 'border:' + u(s.lw) + ' solid ' + col(s.line) + ';';
      if (s.g === 'e') st += 'border-radius:50%;';
      else if (s.rad) st += 'border-radius:' + u(s.rad) + ';';
      if (s.g === 'a') { h += '<div class="sh" style="' + st.replace(/background:[^;]+;/, '') + 'background:' + col(s.fill || s.line || '94A3B8') + ';clip-path:polygon(0 0,100% 50%,0 100%);border:0"></div>'; return; }
      st += 'padding:' + (s.ps.length ? s.ins.map(u).join(' ') : '0') + ';justify-content:' + (s.anc === 'c' ? 'center' : s.anc === 'b' ? 'flex-end' : 'flex-start') + ';';
      st += 'font-size:' + u(11) + ';';
      h += '<div class="sh" style="' + st + '">' + (s.ps.length ? '<div class="tx">' + parasHtml(s.ps, u) + '</div>' : '') + '</div>';
    });
    return h + '</div></div>';
  }

  /* ------------------------------------------------------------ tampilan HP (disusun ulang) */
  function inside(c, p) {                           // c di dalam p: ≥85% luas c tertutup p
    if (c === p || (isBox(c) && c.w * c.h > p.w * p.h)) return false;        // teks lepas boleh sedikit lebih besar dari kotaknya
    var ow = Math.min(c.x + c.w, p.x + p.w) - Math.max(c.x, p.x), oh = Math.min(c.y + c.h, p.y + p.h) - Math.max(c.y, p.y);
    if (ow <= 0 || oh <= 0) return false;
    return ow * oh >= 0.8 * Math.max(1, c.w * c.h) || (c.x >= p.x - 1.5 && c.y >= p.y - 1.5 && c.x + c.w <= p.x + p.w + 1.5 && c.y + c.h <= p.y + p.h + 1.5);
  }
  function isBox(s) { return (s.fill || s.line) && s.g !== 'l' && s.g !== 'a'; }
  function buildTree(list) {
    list.forEach(function (s, i) { s.i = i; s.kids = []; });
    list.forEach(function (c) {
      var best = null;
      list.forEach(function (p) {
        if (!isBox(p) || !inside(c, p)) return;
        if (c.w * c.h === p.w * p.h && c.i < p.i) return;           // bentuk sama besar: yang di bawah jadi induk
        if (!best || p.w * p.h < best.w * best.h) best = p;
      });
      c.par = best;
    });
    var roots = [];
    list.forEach(function (s) { (s.par ? s.par.kids : roots).push(s); });
    return roots;
  }
  function rows(items) {
    items = items.slice().sort(function (a, b) { return a.y - b.y || a.x - b.x; });
    var R = [];
    items.forEach(function (s) {
      var r = R.filter(function (r) { var ov = Math.min(r.b, s.y + s.h) - Math.max(r.t, s.y); return ov > 0.45 * Math.min(s.h, r.b - r.t); })[0];
      if (r) { r.it.push(s); r.t = Math.min(r.t, s.y); r.b = Math.max(r.b, s.y + s.h); } else R.push({ t: s.y, b: s.y + s.h, it: [s] });
    });
    R.forEach(function (r) {                       // dalam satu baris: item yang bertumpuk vertikal jadi satu kolom
      var C = [];
      r.it.sort(function (a, b) { return a.x - b.x || a.y - b.y; }).forEach(function (s) {
        var c = C.filter(function (c) { var ov = Math.min(c.r, s.x + s.w) - Math.max(c.l, s.x); return ov > 0.4 * Math.min(s.w, c.r - c.l); })[0];
        if (c) { c.it.push(s); c.l = Math.min(c.l, s.x); c.r = Math.max(c.r, s.x + s.w); } else C.push({ l: s.x, r: s.x + s.w, it: [s] });
      });
      C.forEach(function (c) { c.it.sort(function (a, b) { return a.y - b.y; }); });
      r.cols = C;
    });
    return R;
  }
  var px = function (pt) { return Math.max(11.5, Math.min(21, pt * 1.33)).toFixed(1) + 'px'; };
  function small(s) { return s.w <= 70 && s.h <= 40; }
  function plainText(s) { return s.ps.map(function (p) { return p[1].map(function (r) { return r[0]; }).join(''); }).join(' ').trim(); }

  /* deteksi tabel: kotak-kotak bersaudara dengan susunan kolom yang sama → di HP tiap sel diberi label kolom */
  function tableHeads(sibs) {
    var boxes = sibs.filter(function (s) { return isBox(s) && s.kids.length >= 2 && s.kids.every(function (k) { return !k.kids.length && hasText(k); }); });
    var sig = function (b) { return b.kids.slice().sort(function (a, c) { return a.x - c.x; }).map(function (k) { return Math.round(k.x / 6); }).join(','); };
    var groups = {};
    boxes.forEach(function (b) { var k = sig(b); (groups[k] = groups[k] || []).push(b); });
    var map = {};
    Object.keys(groups).forEach(function (k) {
      var g = groups[k].sort(function (a, b) { return a.y - b.y; }); if (g.length < 3) return;
      var hd = g[0], heads = hd.kids.slice().sort(function (a, b) { return a.x - b.x; }).map(plainText);
      g.slice(1).forEach(function (b) { map[b.i] = heads; });
      hd.isHead = 1;
    });
    return map;
  }
  function node(s, label) {
    var st = '', cls = 'fn';
    if (s.g === 'T') return tableFlow(s.tb);
    if (s.g === 'l') return '<div class="fn-l" style="border-top:1.5px solid ' + col(s.line || 'CBD5E1') + '"></div>';
    if (s.fill && s.g !== 'a') st += 'background:' + col(s.fill) + ';';
    if (s.line && s.g !== 'a') st += 'border:' + Math.max(1, Math.min(2.5, s.lw * 1.3)).toFixed(1) + 'px solid ' + col(s.line) + ';';
    if (s.g === 'a') return '<div class="fn-a" style="background:' + col(s.fill || s.line || '94A3B8') + '"></div>';
    if (s.g === 'e') { cls += ' fn-e'; st += 'width:' + Math.round(Math.max(28, s.w * 1.35)) + 'px;height:' + Math.round(Math.max(28, s.h * 1.35)) + 'px;'; }
    else if (s.rad) st += 'border-radius:' + Math.round(Math.min(14, s.rad * 1.33)) + 'px;';
    if (!hasText(s) && !s.kids.length) {                                  // hiasan: garis aksen / kotak kecil
      if (s.h <= 5) return '<div class="fn-bar" style="background:' + col(s.fill || s.line) + '"></div>';
      if (s.g === 'e') return '<div class="' + cls + '" style="' + st + '"></div>';
      if (s.w <= 10) return '<div class="fn-dot" style="background:' + col(s.fill || s.line) + ';height:' + Math.round(s.h * 1.2) + 'px"></div>';
    }
    if (s.fill || s.line) cls += ' fn-box';
    if (s.anc === 'c' && (s.fill || s.line)) cls += ' fn-mid';
    var body = (label ? '<small class="fn-lbl">' + esc(label) + '</small>' : '') + (s.ps.length ? '<div class="tx">' + parasHtml(s.ps, px) + '</div>' : '') + (s.kids.length ? flow(s.kids) : '');
    return '<div class="' + cls + '" style="' + st + '">' + body + '</div>';
  }
  /* susun ulang untuk HP dengan "XY-cut": potong di celah kosong.
     celah vertikal lebar (≥12pt) = kolom; selain itu potong per baris; kolom ditampilkan berdampingan bila muat, selain itu ditumpuk */
  function gaps(items, ax) {
    var p = ax === 'x' ? ['x', 'w'] : ['y', 'h'];
    var iv = items.map(function (s) { return [s[p[0]], s[p[0]] + s[p[1]]]; }).sort(function (a, b) { return a[0] - b[0]; });
    var out = [], end = iv[0][1];
    for (var i = 1; i < iv.length; i++) {
      if (iv[i][0] >= end - 0.5) out.push({ at: (end + iv[i][0]) / 2, w: iv[i][0] - end });
      end = Math.max(end, iv[i][1]);
    }
    return out;
  }
  function split(items, ax, cuts) {
    var p = ax === 'x' ? ['x', 'w'] : ['y', 'h'], parts = cuts.map(function () { return []; }).concat([[]]);
    items.forEach(function (s) { var c = s[p[0]] + s[p[1]] / 2, k = 0; while (k < cuts.length && c > cuts[k].at) k++; parts[k].push(s); });
    return parts.filter(function (q) { return q.length; });
  }
  function xycut(items) {
    if (items.length <= 1) return { leaf: items[0] };
    var vg = gaps(items, 'x'), hg = gaps(items, 'y'), best = null;
    vg.forEach(function (g) { var sc = g.w; if (!best || sc > best.sc) best = { ax: 'x', g: g, sc: sc }; });
    hg.forEach(function (g) { var sc = g.w * 2.5 + 0.01; if (!best || sc >= best.sc) best = { ax: 'y', g: g, sc: sc }; });   // baris diutamakan
    if (!best) return { v: items.slice().sort(function (a, b) { return a.y - b.y || a.x - b.x; }).map(function (s) { return { leaf: s }; }) };
    var k = best.ax === 'x' ? 'h' : 'v', parts = split(items, best.ax, [best.g]).map(xycut), flat = [];
    parts.forEach(function (q) { if (q[k]) flat = flat.concat(q[k]); else flat.push(q); });          // gabungkan potongan searah
    var o = {}; o[k] = flat; return o;
  }
  function ext(n) {                                 // lebar & sifat "kecil" sebuah bagian
    if (n.leaf) return { l: n.leaf.x, r: n.leaf.x + n.leaf.w, sm: small(n.leaf) };
    var e = (n.h || n.v).map(ext);
    return { l: Math.min.apply(null, e.map(function (q) { return q.l; })), r: Math.max.apply(null, e.map(function (q) { return q.r; })), sm: e.every(function (q) { return q.sm; }) };
  }
  function render(n) {
    if (n.leaf) return n.leaf.isHead ? '' : node(n.leaf);
    if (n.v) return n.v.map(render).join('');
    var parts = n.h.filter(function (q) { return render(q); }), W = 330;
    if (parts.length === 1) return render(parts[0]);
    var E = parts.map(ext), need = E.reduce(function (m, e) { return m + (e.sm ? Math.max(30, (e.r - e.l) * 1.3) : Math.max(110, (e.r - e.l) * 0.72)) + 8; }, 0);
    if (need > W) {
      var g2 = parts.length >= 3 && E.every(function (e) { return !e.sm && e.r - e.l <= 170; });     // kolom ramping (mis. tier) → 2 per baris
      return '<div class="fr ' + (g2 ? 'g2' : 'v') + '">' + parts.map(function (q) { return g2 ? '<div class="fc">' + render(q) + '</div>' : render(q); }).join('') + '</div>';
    }
    return '<div class="fr h">' + parts.map(function (q, i) {
      var e = E[i];
      return '<div class="fc' + (e.sm ? ' sm' : '') + '" style="flex:' + (e.sm ? '0 0 auto' : Math.round(e.r - e.l) + ' 1 0') + '">' + render(q) + '</div>';
    }).join('') + '</div>';
  }
  function flow(items) { return render(xycut(items)); }
  /* sel tabel: induknya baris tabel → tambahkan label kolom */
  function flowHtml(L) {
    var list = L.s.map(S);
    var roots = buildTree(list);
    // label kolom tabel: tandai di setiap baris
    function mark(arr) {
      var heads = tableHeads(arr);
      arr.forEach(function (b) {
        if (heads[b.i]) b.kids.slice().sort(function (a, c) { return a.x - c.x; }).forEach(function (k, j) { k.lbl = heads[b.i][j]; });
        mark(b.kids);
      });
    }
    mark(roots);
    var html = flow(roots);
    return '<div class="slf">' + html + '</div>';
  }
  /* node() memakai label dari mark() */
  var _node = node;
  node = function (s) { return _node(s, s.lbl); };

  function html(L, opt) {
    if (!L || !L.s || !L.s.length) return '';
    return '<div class="slv">' + absHtml(L) + flowHtml(L)
      + '<button class="sl-tg" type="button" data-sltg>Lihat seperti slide asli</button></div>';
  }
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('[data-sltg]'); if (!b) return;
    var v = b.closest('.slv'), on = v.classList.toggle('asli');
    b.textContent = on ? 'Kembali ke tampilan HP' : 'Lihat seperti slide asli';
  });
  window.SlideView = { html: html };
})();
