/* =====================================================================
   foto.js — foto produk per artikel+warna

   Admin : membaca file link foto (Seller SKU, Color, Desc, Gambar1–8),
           diringkas per artikel+warna (mis. 910000058.BLK) lalu dibagi
           ke 100 kelompok supaya toko cukup mengambil yang dibutuhkan.
   Toko  : galeri foto yang bisa diperbesar (geser, ketuk dua kali untuk
           zoom, arahkan kursor untuk kaca pembesar di PC).
   ===================================================================== */
(function () {
  'use strict';
  var BASES = ['https://id-live.slatic.net/p/', 'https://sg-live-01.slatic.net/p/', 'https://id-live-01.slatic.net/p/'];
  var CHUNK = 40000;
  var normH = function (v) { return String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9]/g, ''); };
  var tx = function (v) { return String(v == null ? '' : v).trim(); };

  /* https://id-live.slatic.net/p/<32hex>.jpg  →  "0<32hex>" (hemat ±40 karakter per foto) */
  function enc(u) {
    u = tx(u);
    for (var i = 0; i < BASES.length; i++) {
      if (u.indexOf(BASES[i]) === 0) { var m = /^([0-9a-f]{32})\.jpg$/.exec(u.slice(BASES[i].length)); if (m) return i + m[1]; }
    }
    return u;
  }
  function url(s, bases, size) {
    bases = bases || BASES;
    var m = /^(\d)([0-9a-f]{32})$/.exec(s || ''), u = m ? (bases[+m[1]] || BASES[0]) + m[2] + '.jpg' : String(s || '');
    if (size && /slatic\.net\/p\/[0-9a-f]{32}\.jpg$/.test(u)) u += '_' + size + 'x' + size + 'q80.jpg';
    return u;
  }
  var bucketOf = function (key) { return String(key).slice(0, 9).slice(-2); };

  function readBook(wb) {
    var X = window.XLSX, best = null;
    wb.SheetNames.forEach(function (sn) {
      if (best) return;
      var aoa = X.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: '', blankrows: false }), tanggal = '';
      for (var i = 0; i < Math.min(aoa.length, 15); i++) {
        var row = aoa[i] || [], h = row.map(normH);
        var td = h.indexOf('tanggalupdate'); if (td >= 0) { var v = row[td + 1]; tanggal = v instanceof Date ? v.toISOString().slice(0, 10) : tx(v).slice(0, 10); }
        var cs = h.findIndex(function (x) { return x === 'sellersku' || x === 'skucode' || x === 'sku'; });
        var imgs = []; h.forEach(function (x, j) { if (/^(gambar|image|foto|img|picture)\d*$/.test(x)) imgs.push(j); });
        if (cs < 0 || !imgs.length) continue;
        var cc = h.findIndex(function (x) { return x === 'color' || x === 'warna' || x === 'colour'; });
        var cd = h.findIndex(function (x) { return x === 'desc' || x === 'description' || x === 'skuname' || x === 'nama'; });
        var map = {}, nSku = 0, nImg = 0, skipped = 0, sample = [], derived = 0;
        for (var r = i + 1; r < aoa.length; r++) {
          var R = aoa[r] || [], sku = tx(typeof R[cs] === 'number' ? Math.round(R[cs]) : R[cs]).replace(/\.0+$/, '');
          if (!/^\d{9,13}$/.test(sku)) { if (sku) skipped++; continue; }
          var list = imgs.map(function (j) { return tx(R[j]); }).filter(function (u) { return /^https?:\/\//i.test(u); });
          if (!list.length) { skipped++; continue; }
          nSku++;
          var key = cc >= 0 ? tx(R[cc]).toUpperCase().replace(/\s+/g, '') : '';
          if (!/^\d{9}\.[A-Z0-9]{2,5}$/.test(key)) {
            var col = cd >= 0 ? tx(String(R[cd]).split(',')[1]).toUpperCase() : '';
            key = sku.slice(0, 9) + '.' + (col || 'NOC'); derived++;
          }
          if (map[key]) continue;
          map[key] = list.map(enc); nImg += list.length;
          if (sample.length < 12 && r % 97 === 3) sample.push([key, cd >= 0 ? tx(R[cd]) : '', list[0]]);
        }
        if (!nSku) continue;
        best = { sheet: sn, tanggal: tanggal, map: map, nSku: nSku, nKey: Object.keys(map).length, nImg: nImg, skipped: skipped, derived: derived, sample: sample };
      }
    });
    if (!best) throw new Error('Kolom Seller SKU dan Gambar1… tidak ditemukan. Gunakan format file link foto yang biasa.');
    return best;
  }
  /* peta foto → baris [kelompok, bagian, potongan JSON] untuk disimpan */
  function toRows(map) {
    var B = {};
    Object.keys(map).forEach(function (k) { var b = bucketOf(k); (B[b] = B[b] || {})[k] = map[k]; });
    var rows = [];
    Object.keys(B).sort().forEach(function (b) {
      var s = JSON.stringify(B[b]);
      for (var i = 0, p = 0; i < s.length; i += CHUNK, p++) rows.push([b, p, s.slice(i, i + CHUNK)]);
    });
    return rows;
  }

  /* ----------------------------------------------------- galeri zoom */
  var LB = null;
  function open(urls, start, info) {
    close();
    if (!urls || !urls.length) return;
    info = info || {};
    var el = document.createElement('div');
    el.className = 'lb'; el.setAttribute('role', 'dialog'); el.setAttribute('aria-modal', 'true');
    el.innerHTML = '<div class="lb-box"><div class="lb-top"><span class="lb-c"></span><button class="lb-x" aria-label="Tutup">×</button></div>'
      + '<div class="lb-stage"><img class="lb-img" alt=""><button class="lb-nav l" aria-label="Foto sebelumnya">‹</button><button class="lb-nav r" aria-label="Foto berikutnya">›</button>'
      + '<span class="lb-hint">Ketuk dua kali untuk zoom</span></div>'
      + '<div class="lb-th"></div>'
      + '<div class="lb-info">' + (info.title ? '<b>' + info.title + '</b>' : '') + (info.sub ? '<small>' + info.sub + '</small>' : '') + (info.note ? '<em>' + info.note + '</em>' : '') + '</div></div>';
    document.body.appendChild(el);
    document.body.classList.add('lb-open');
    LB = { el: el, urls: urls, i: start || 0, z: 1 };
    el.querySelector('.lb-th').innerHTML = urls.map(function (u, k) { return '<button data-k="' + k + '"><img loading="lazy" src="' + u.thumb + '" alt=""></button>'; }).join('');
    show(LB.i);
    var stage = el.querySelector('.lb-stage'), img = el.querySelector('.lb-img');
    el.addEventListener('click', function (e) {
      if (e.target.closest('.lb-x') || e.target === el) { close(); return; }
      var t = e.target.closest('[data-k]'); if (t) { show(+t.dataset.k); return; }
      if (e.target.closest('.lb-nav.l')) { show(LB.i - 1); return; }
      if (e.target.closest('.lb-nav.r')) { show(LB.i + 1); return; }
    });
    /* PC: kaca pembesar mengikuti kursor */
    var fine = window.matchMedia && window.matchMedia('(pointer:fine)').matches;
    if (fine) {
      stage.addEventListener('mousemove', function (e) { if (e.target.closest('.lb-nav')) { zoom(1); return; } var r = stage.getBoundingClientRect(); zoom(2.4, (e.clientX - r.left) / r.width * 100, (e.clientY - r.top) / r.height * 100); });
      stage.addEventListener('mouseleave', function () { zoom(1); });
    }
    /* HP: geser untuk ganti foto, ketuk dua kali untuk zoom, geser saat zoom untuk melihat detail */
    var sx = 0, sy = 0, moved = false, lastTap = 0;
    stage.addEventListener('touchstart', function (e) { var t = e.touches[0]; sx = t.clientX; sy = t.clientY; moved = false; }, { passive: true });
    stage.addEventListener('touchmove', function (e) {
      var t = e.touches[0]; if (Math.abs(t.clientX - sx) > 8 || Math.abs(t.clientY - sy) > 8) moved = true;
      if (LB.z > 1) { e.preventDefault(); var r = stage.getBoundingClientRect(); zoom(LB.z, (t.clientX - r.left) / r.width * 100, (t.clientY - r.top) / r.height * 100); }
    }, { passive: false });
    stage.addEventListener('touchend', function (e) {
      var t = e.changedTouches[0], dx = t.clientX - sx;
      if (LB.z === 1 && moved && Math.abs(dx) > 45 && Math.abs(dx) > Math.abs(t.clientY - sy)) { show(LB.i + (dx < 0 ? 1 : -1)); return; }
      if (!moved) {
        var now = Date.now();
        if (now - lastTap < 320) { var r = stage.getBoundingClientRect(); if (LB.z > 1) zoom(1); else zoom(2.5, (t.clientX - r.left) / r.width * 100, (t.clientY - r.top) / r.height * 100); lastTap = 0; }
        else lastTap = now;
      }
    });
    img.addEventListener('error', function () { if (img.dataset.full && img.src !== img.dataset.full) img.src = img.dataset.full; });
    document.addEventListener('keydown', onKey);
  }
  function show(i) {
    if (!LB) return;
    var n = LB.urls.length; LB.i = (i + n) % n;
    var img = LB.el.querySelector('.lb-img'), u = LB.urls[LB.i];
    zoom(1);
    img.dataset.full = u.full; img.src = u.big || u.full;
    LB.el.querySelector('.lb-c').textContent = (LB.i + 1) + ' / ' + n;
    LB.el.querySelectorAll('.lb-th button').forEach(function (b, k) { b.classList.toggle('on', k === LB.i); if (k === LB.i) b.scrollIntoView({ block: 'nearest', inline: 'center' }); });
    LB.el.classList.toggle('single', n < 2);
  }
  function zoom(z, x, y) {
    if (!LB) return;
    var img = LB.el.querySelector('.lb-img'); LB.z = z;
    img.style.transformOrigin = (x == null ? 50 : x) + '% ' + (y == null ? 50 : y) + '%';
    img.style.transform = z > 1 ? 'scale(' + z + ')' : '';
    LB.el.classList.toggle('zoomed', z > 1);
  }
  function onKey(e) { if (!LB) return; if (e.key === 'Escape') close(); else if (e.key === 'ArrowRight') show(LB.i + 1); else if (e.key === 'ArrowLeft') show(LB.i - 1); }
  function close() { if (!LB) return; LB.el.remove(); LB = null; document.body.classList.remove('lb-open'); document.removeEventListener('keydown', onKey); }
  /* daftar kode foto → objek URL untuk galeri */
  function gallery(codes, bases) { return (codes || []).map(function (c) { return { thumb: url(c, bases, 120), big: url(c, bases, 720), full: url(c, bases) }; }); }

  window.Foto = { BASES: BASES, enc: enc, url: url, bucketOf: bucketOf, readBook: readBook, toRows: toRows, open: open, close: close, gallery: gallery };
})();
