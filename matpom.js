/* =====================================================================
   Monitoring pemasangan materi promo
   - Toko   : lapor per materi → foto (dikompres otomatis di HP) atau "Tidak pasang" + alasan
   - HO/Admin/toko : rekap % per toko → branch → regional, foto bukti, export Excel
   Dipakai oleh index.html (toko & HO) dan admin.html.  MP.setup({...}) dipanggil oleh halaman.
   Elemen <div data-mpstore>, <div data-mpmon>, <div data-mpban> otomatis diisi saat muncul.
   ===================================================================== */
(function () {
  var C = null;                                   // konfigurasi dari halaman
  var MYST = null, MINE = {}, MINE_AT = {}, REKAP = {}, IMG = {}, BUSY = {}, REASONS = [], PUB = true;
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); };
  var MONL = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];
  function fmtM(m) { return m ? MONL[+m.slice(5, 7) - 1] + ' ' + m.slice(0, 4) : ''; }
  function fmtD(d) { return d ? (+d.slice(8, 10)) + ' ' + MON[+d.slice(5, 7) - 1] + ' ' + d.slice(0, 4) : ''; }
  function fmtT(iso) { if (!iso) return ''; var d = new Date(iso); if (isNaN(d)) return String(iso).slice(0, 10); return d.getDate() + ' ' + MON[d.getMonth()] + ' ' + ('0' + d.getHours()).slice(-2) + '.' + ('0' + d.getMinutes()).slice(-2); }
  function today() { var d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
  function daysTo(d) { if (!d) return null; return Math.round((new Date(d + 'T00:00:00') - new Date(today() + 'T00:00:00')) / 86400000); }
  function pct(x) { return x == null || isNaN(x) ? '–' : Math.round(x * 100) + '%'; }
  function tone(x) { return x == null ? 'n' : x >= 0.9 ? 'g' : x >= 0.7 ? 'a' : 'r'; }
  function svg(n) { return window.svg ? window.svg(n) : ''; }
  function camps() { return (C && C.camps() || []); }
  function active() { return camps().filter(function (c) { return c.aktif; }); }
  function role() { return C ? C.role() : 'site'; }

  /* ------------------------------------------------ kompres foto di HP (±1280px, JPEG ±150 KB) */
  function loadImg(file) {
    return new Promise(function (res, rej) {
      var u = URL.createObjectURL(file), im = new Image();
      im.onload = function () { res({ im: im, u: u }); }; im.onerror = function () { URL.revokeObjectURL(u); rej(new Error('File bukan gambar.')); };
      im.src = u;
    });
  }
  async function compress(file, stamp) {
    var o = await loadImg(file), im = o.im, W0 = im.naturalWidth, H0 = im.naturalHeight, out = null;
    var tries = [[1280, 0.72], [1280, 0.6], [1024, 0.58], [900, 0.5]];
    for (var i = 0; i < tries.length; i++) {
      var s = Math.min(1, tries[i][0] / Math.max(W0, H0)), W = Math.round(W0 * s), H = Math.round(H0 * s);
      var cv = document.createElement('canvas'); cv.width = W; cv.height = H;
      var x = cv.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, W, H); x.drawImage(im, 0, 0, W, H);
      if (stamp) {                                  // cap waktu & No Site di foto (bukti asli, bukan foto lama)
        var fs = Math.max(14, Math.round(W / 42)); x.font = '700 ' + fs + 'px Arial, sans-serif';
        var tw = x.measureText(stamp).width; x.fillStyle = 'rgba(0,0,0,.55)'; x.fillRect(0, H - fs * 1.9, tw + fs * 1.4, fs * 1.9);
        x.fillStyle = '#fff'; x.fillText(stamp, fs * 0.7, H - fs * 0.62);
      }
      out = cv.toDataURL('image/jpeg', tries[i][1]);
      if (out.length * 0.75 < 300 * 1024) break;
    }
    URL.revokeObjectURL(o.u);
    return { data: out, kb: Math.round(out.length * 0.75 / 1024), orig: Math.round(file.size / 1024) };
  }

  /* ------------------------------------------------ foto: thumbnail Drive, cadangan lewat server */
  function thumbUrl(id, big) { return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w' + (big ? 1600 : 360); }
  function imgTag(id, big, local) {
    if (local) return '<img src="' + local + '" alt="Foto materi">';
    if (IMG[id]) return '<img src="' + IMG[id] + '" alt="Foto materi">';
    return '<img data-mpimg="' + esc(id) + '" src="' + (PUB ? thumbUrl(id, big) : 'data:image/gif;base64,R0lGODlhAQABAAAAACw=') + '" alt="Foto materi" loading="lazy" referrerpolicy="no-referrer">';
  }
  function fetchImg(el) {
    var id = el.getAttribute('data-mpimg'); if (!id || el._mpf) return; el._mpf = 1;
    if (IMG[id]) { el.src = IMG[id]; return; }
    C.api('mpImg', { id: id }).then(function (r) { IMG[id] = r.data; document.querySelectorAll('img[data-mpimg="' + id + '"]').forEach(function (x) { x.src = r.data; }); }).catch(function () { el.alt = 'Foto tidak bisa dimuat'; });
  }
  document.addEventListener('error', function (e) { var t = e.target; if (t && t.tagName === 'IMG' && t.hasAttribute('data-mpimg')) fetchImg(t); }, true);
  function hydrateImgs(root) { if (!PUB) (root || document).querySelectorAll('img[data-mpimg]').forEach(fetchImg); }

  /* ------------------------------------------------ hitung status */
  var NZ = function (v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().toUpperCase(); };
  /* grup media mengikuti urutan materi: Offline, Online, Media lokal, KOL, … (bebas diisi admin) */
  function groupsOf(c) { var g = []; (c.items || []).forEach(function (x) { var n = x.grp || 'Offline'; if (g.indexOf(n) < 0) g.push(n); }); return g; }
  /* materi berlaku untuk toko? scope kosong = semua; "REG 03, JATIM 01, 2038" = regional / branch / No Site */
  function inScope(it, st) {
    if (!it.scope) return true; if (!st) return false;
    var me = [String(st[0]).trim(), NZ(st[2]), NZ(st[3])];
    return String(it.scope).split(',').some(function (t) { t = NZ(t); return t && me.indexOf(t) >= 0; });
  }
  function itemsFor(c, st, g) { return (c.items || []).filter(function (x) { return inScope(x, st) && (!g || (x.grp || 'Offline') === g); }); }
  function storeScore(c, map, st) {                // map: item → row · st: [code, name, branch, regional]
    var r = { g: {}, all: [0, 0], lapor: 0, no: 0, need: 0 };
    groupsOf(c).forEach(function (g) { r.g[g] = [0, 0]; });
    itemsFor(c, st).forEach(function (it) {
      var row = map && map[it.k], ok = row && row.status === 'ok' ? 1 : 0, g = it.grp || 'Offline';
      r.g[g][1]++; r.g[g][0] += ok; r.all[1]++; r.all[0] += ok; r.need++; if (row) r.lapor++; if (row && row.status === 'no') r.no++;
    });
    return r;
  }
  function shortG(g) { var m = { OFFLINE: 'Off', ONLINE: 'On', 'MEDIA LOKAL': 'Lokal' }; return m[NZ(g)] || String(g).split(' ')[0].slice(0, 6); }
  function ratio(a) { return a[1] ? a[0] / a[1] : null; }

  /* ================================================= TOKO: lapor pemasangan */
  function loadMine(cid, force) {
    if (!force && MINE[cid] && Date.now() - MINE_AT[cid] < 120000) return Promise.resolve(MINE[cid]);
    return C.api('mpMine', { cid: cid }).then(function (r) {
      var m = {}; (r.rows || []).forEach(function (x) { m[x[1]] = { status: x[2], reason: x[3], note: x[4], file: x[5], at: x[6] }; });
      if (r.reasons) REASONS = r.reasons; PUB = r.pub !== false;
      if (r.st) MYST = r.st;
      MINE[cid] = m; MINE_AT[cid] = Date.now(); return m;
    });
  }
  function storeCard(c) {
    var m = MINE[c.id] || {}, sc = storeScore(c, m, MYST), dl = daysTo(c.deadline), tot = sc.need;
    if (!tot) return '<section class="panel mp-c" data-mpc="' + esc(c.id) + '" hidden></section>';      // tidak ada materi untuk toko ini
    var dTxt = c.deadline ? 'Batas lapor ' + fmtD(c.deadline) + (dl == null ? '' : dl > 1 ? ' · ' + dl + ' hari lagi' : dl === 1 ? ' · besok' : dl === 0 ? ' · hari ini' : ' · lewat ' + (-dl) + ' hari') : 'Lapor setelah materi terpasang';
    var done = sc.lapor === tot;
    var h = '<section class="panel mp-c' + (done ? ' done' : '') + '" data-mpc="' + esc(c.id) + '">'
      + '<div class="mp-h"><div><small class="' + (dl != null && dl <= 1 && !done ? 'late' : '') + '">' + esc(dTxt) + '</small><b>' + esc(c.name) + '</b>'
      + (c.note ? '<p>' + esc(c.note) + '</p>' : '') + '</div>'
      + '<div class="mp-ring" style="--p:' + (tot ? sc.lapor / tot * 100 : 0) + '"><span><b>' + sc.lapor + '/' + tot + '</b><small>dilapor</small></span></div></div>';
    groupsOf(c).forEach(function (g) {
      var L = itemsFor(c, MYST, g); if (!L.length) return;
      h += '<h4 class="mp-g">' + g + '</h4>' + L.map(function (it) { return itemRow(c, it, m[it.k]); }).join('');
    });
    h += '<div class="mp-foot">' + (done ? '<span class="mp-ok">✓ Semua materi sudah dilaporkan. Terima kasih!</span>' : '<span class="small muted">Foto otomatis dikecilkan sebelum dikirim, jadi hemat kuota.</span>')
      + '<button class="btn sm" data-mpmonopen="' + esc(c.id) + '">Lihat rekap branch</button></div></section>';
    return h;
  }
  function itemRow(c, it, row) {
    var key = c.id + '|' + it.k, b = BUSY[key];
    var file = '<label class="btn primary sm mp-cam">' + svg('image') + (row && row.status === 'ok' ? 'Ganti foto' : 'Terpasang · foto') + '<input type="file" accept="image/*" data-mpfile="' + esc(key) + '" hidden></label>';
    var h = '<div class="mp-i' + (row ? ' s-' + row.status : '') + '" data-mpi="' + esc(key) + '"><div class="mp-l"><b>' + esc(it.label) + '</b>';
    if (b && b.t === 'up') return h + '<small><span class="spin"></span>Mengirim foto (' + b.kb + ' KB, asli ' + b.orig + ' KB)…</small></div><div class="mp-th">' + imgTag('', 0, b.data) + '</div></div>';
    if (b && b.t === 'form') {
      return h + '</div><div class="mp-no">'
        + '<select class="inp" data-mpreason>' + '<option value="">— Pilih alasan tidak pasang —</option>' + REASONS.map(function (r) { return '<option' + (row && row.reason === r ? ' selected' : '') + '>' + esc(r) + '</option>'; }).join('') + '</select>'
        + '<input class="inp" data-mpnote maxlength="200" placeholder="Keterangan (wajib bila Lainnya)" value="' + esc(row && row.note || '') + '">'
        + '<div class="mp-a"><button class="btn primary sm" data-mpnosave="' + esc(key) + '">Simpan</button><button class="btn sm" data-mpcancel="' + esc(key) + '">Batal</button></div></div></div>';
    }
    if (!row) return h + '<small class="muted">Belum dilaporkan</small></div><div class="mp-a">' + file + '<button class="btn sm" data-mpnoform="' + esc(key) + '">Tidak pasang</button></div></div>';
    if (row.status === 'ok') return h + '<small class="ok">✓ Terpasang · ' + esc(fmtT(row.at)) + '</small><div class="mp-a">' + file + '<button class="btn sm ghost" data-mpclear="' + esc(key) + '">Batalkan</button></div></div>'
      + '<button class="mp-th" data-mpview="' + esc(row.file) + '" data-cap="' + esc(it.label) + '">' + imgTag(row.file, 0, row.local) + '</button></div>';
    return h + '<small class="bad">✕ Tidak dipasang: ' + esc(row.reason) + (row.note ? ' — ' + esc(row.note) : '') + '</small><div class="mp-a">' + file + '<button class="btn sm" data-mpnoform="' + esc(key) + '">Ubah alasan</button></div></div></div>';
  }
  function paintStore(cid) {
    document.querySelectorAll('[data-mpc="' + cid + '"]').forEach(function (el) {
      var c = camps().filter(function (x) { return x.id === cid; })[0]; if (c) { el.outerHTML = storeCard(c); }
    });
    hydrateImgs(); paintBanners();
  }
  function mountStore(el) {
    var L = active();
    if (!L.length) { el.innerHTML = ''; return; }
    el.innerHTML = '<div class="mp-wrap"><div class="ph"><h2>Lapor pemasangan materi</h2></div>' + L.map(function (c) { return '<section class="panel mp-c" data-mpc="' + esc(c.id) + '"><span class="spin"></span>Memuat…</section>'; }).join('') + '</div>';
    L.forEach(function (c) { loadMine(c.id).then(function () { paintStore(c.id); }).catch(function (e) { var x = el.querySelector('[data-mpc="' + c.id + '"]'); if (x) x.innerHTML = '<div class="empty">Gagal memuat: ' + esc(e.message || e) + '</div>'; }); });
  }
  /* banner di beranda toko: masih ada materi yang belum dilapor */
  function paintBanners() {
    document.querySelectorAll('[data-mpban]').forEach(function (el) {
      var L = active(), need = 0, near = null;
      L.forEach(function (c) { var m = MINE[c.id]; if (!m) return; var s = storeScore(c, m, MYST), left = s.need - s.lapor; if (left > 0) { need += left; if (!near || (c.deadline && c.deadline < near.deadline)) near = c; } });
      if (!need) { el.innerHTML = ''; return; }
      var dl = near && daysTo(near.deadline);
      el.innerHTML = '<button class="mp-ban" data-mpgo="1"><span class="mp-bi">' + svg('image') + '</span><span><b>' + need + ' materi promo belum dilaporkan</b><small>' + esc(near ? near.name : '') + (dl != null ? ' · batas ' + (dl > 0 ? dl + ' hari lagi' : dl === 0 ? 'hari ini' : 'sudah lewat') : '') + '</small></span><span class="go">' + svg('next') + '</span></button>';
    });
  }
  function mountBanner(el) {
    if (role() !== 'site') { el.innerHTML = ''; return; }
    var L = active(); if (!L.length) return;
    Promise.all(L.map(function (c) { return loadMine(c.id).catch(function () {}); })).then(paintBanners);
  }

  /* ------------------------------------------------ aksi toko */
  document.addEventListener('change', async function (e) {
    var t = e.target; if (!t.matches || !t.matches('input[data-mpfile]')) return;
    var f = t.files && t.files[0]; if (!f) return;
    var p = t.getAttribute('data-mpfile').split('|'), cid = p[0], k = p[1], key = cid + '|' + k;
    var me = C.me(), stamp = (me.code || '') + ' · ' + new Date().toLocaleString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    try {
      var z = await compress(f, stamp);
      BUSY[key] = { t: 'up', data: z.data, kb: z.kb, orig: z.orig }; paintStore(cid);
      var r = await C.post({ a: 'mpUp', cid: cid, item: k, data: z.data });
      MINE[cid] = MINE[cid] || {}; MINE[cid][k] = { status: 'ok', file: r.file, at: new Date().toISOString(), local: z.data };
      IMG[r.file] = z.data; C.toast('Foto terkirim (' + z.kb + ' KB).');
    } catch (er) { C.toast('Gagal mengirim foto: ' + (er.message || er.code || er) + '. Coba lagi.', 6000); }
    delete BUSY[key]; paintStore(cid);
  });
  document.addEventListener('click', async function (e) {
    var t = e.target.closest && e.target.closest('[data-mpnoform],[data-mpcancel],[data-mpnosave],[data-mpclear],[data-mpview],[data-mpgo],[data-mpmonopen]'); if (!t) return;
    var a = t.dataset;
    if (a.mpgo) { C.go('materi'); return; }
    if (a.mpmonopen) { C.openMonitor(a.mpmonopen); return; }
    if (a.mpview) { viewer(a.mpview, a.cap || ''); return; }
    var key = a.mpnoform || a.mpcancel || a.mpnosave || a.mpclear, cid = key.split('|')[0], k = key.split('|')[1];
    if (a.mpnoform) { if (!REASONS.length) await loadMine(cid, true); BUSY[key] = { t: 'form' }; paintStore(cid); return; }
    if (a.mpcancel) { delete BUSY[key]; paintStore(cid); return; }
    if (a.mpnosave) {
      var box = t.closest('.mp-i'), rs = box.querySelector('[data-mpreason]').value, nt = box.querySelector('[data-mpnote]').value.trim();
      if (!rs) { C.toast('Pilih alasannya dulu.'); return; }
      if (rs === 'Lainnya' && !nt) { C.toast('Tulis keterangan untuk alasan Lainnya.'); box.querySelector('[data-mpnote]').focus(); return; }
      t.disabled = true;
      try { await C.post({ a: 'mpNo', cid: cid, item: k, reason: rs, note: nt }); MINE[cid][k] = { status: 'no', reason: rs, note: nt, at: new Date().toISOString() }; delete BUSY[key]; C.toast('Alasan tersimpan.'); }
      catch (er) { t.disabled = false; C.toast('Gagal: ' + (er.message || er.code), 6000); return; }
      paintStore(cid); return;
    }
    if (a.mpclear) {
      if (!confirm('Batalkan laporan materi ini? Foto akan dihapus.')) return;
      t.disabled = true;
      try { await C.post({ a: 'mpClear', cid: cid, item: k }); delete MINE[cid][k]; C.toast('Laporan dibatalkan.'); } catch (er) { C.toast('Gagal: ' + (er.message || er.code), 6000); }
      paintStore(cid);
    }
  });

  /* ------------------------------------------------ penampil foto */
  function viewer(id, cap) {
    var ov = document.createElement('div'); ov.className = 'mp-ov';
    ov.innerHTML = '<div class="mp-ovb"><div class="mp-ovh"><span>' + esc(cap) + '</span><button aria-label="Tutup" data-mpx>' + svg('x') + '</button></div>'
      + '<div class="mp-ovi">' + (IMG[id] ? '<img src="' + IMG[id] + '" alt="">' : imgTag(id, 1)) + '</div>'
      + (C.role() !== 'site' ? '<a class="small" href="https://drive.google.com/file/d/' + esc(id) + '/view" target="_blank" rel="noopener">Buka di Google Drive ↗</a>' : '') + '</div>';
    document.body.appendChild(ov); hydrateImgs(ov);
    ov.addEventListener('click', function (e) { if (e.target === ov || e.target.closest('[data-mpx]')) ov.remove(); });
  }

  /* ================================================= MONITORING (rekap %) */
  var MON_ST = { cid: '', f: 'all', q: '', open: {} };
  function loadRekap(cid, force) {
    if (!force && REKAP[cid] && Date.now() - REKAP[cid].t < 60000) return Promise.resolve(REKAP[cid]);
    return C.api('mpRekap', { cid: cid }).then(function (r) {
      if (r.reasons) REASONS = r.reasons; PUB = r.pub !== false;
      var by = {}; (r.rows || []).forEach(function (x) { (by[x[0]] = by[x[0]] || {})[x[1]] = { status: x[2], reason: x[3], note: x[4], file: x[5], at: x[6] }; });
      REKAP[cid] = { t: Date.now(), camp: r.camp, stores: r.stores || [], by: by, folder: r.folder || '' };
      return REKAP[cid];
    });
  }
  function build(R) {                              // → { nat, regs: [{name, sc, branches:[{name, sc, stores:[...]}]}] }
    var c = R.camp, GR = groupsOf(c), regs = {};
    var add = function (t, s) { GR.forEach(function (g) { t.g[g][0] += s.g[g][0]; t.g[g][1] += s.g[g][1]; }); t.all[0] += s.all[0]; t.all[1] += s.all[1]; t.n++; if (s.lapor) t.lapor++; if (s.lapor === s.need) t.full++; };
    var blank = function () { var b = { g: {}, all: [0, 0], n: 0, lapor: 0, full: 0 }; GR.forEach(function (g) { b.g[g] = [0, 0]; }); return b; }, nat = blank();
    R.stores.forEach(function (s) {
      if (/own\s*platform|e-?commerce/i.test(s[3] + ' ' + s[2])) return;   // e-commerce tidak dimonitor
      var nz = function (v) { return String(v || '').replace(/\s+/g, ' ').trim().toUpperCase().replace(/^REQ\s*(\d+)$/, 'REG $1'); }, sc = storeScore(c, R.by[s[0]], s), rg = nz(s[3]) || 'TANPA REGIONAL', br = nz(s[2]) || 'TANPA BRANCH';
      var G = regs[rg] = regs[rg] || { name: rg, sc: blank(), br: {} }, B = G.br[br] = G.br[br] || { name: br, sc: blank(), stores: [] };
      if (!sc.need) return;                                              // tidak ada materi untuk toko ini
      B.stores.push({ code: s[0], name: s[1], st: s, sc: sc, map: R.by[s[0]] || {} }); add(B.sc, sc); add(G.sc, sc); add(nat, sc);
    });
    var srt = function (a, b) { var ra = /^REG/i.test(a.name) ? 0 : 1, rb = /^REG/i.test(b.name) ? 0 : 1; return ra - rb || String(a.name).localeCompare(String(b.name), 'id', { numeric: true }); };   // REG 01..05 dulu, lainnya di bawah
    return { groups: GR, nat: nat, regs: Object.keys(regs).map(function (k) { var G = regs[k]; G.branches = Object.keys(G.br).map(function (b) { return G.br[b]; }).sort(srt); return G; }).sort(srt) };
  }
  function cell(v) { return '<td class="mp-p t-' + tone(v) + '"><span style="--w:' + (v == null ? 0 : Math.min(100, v * 100)) + '%">' + pct(v) + '</span></td>'; }
  function gcells(sc) { return Object.keys(sc.g).map(function (g) { return cell(sc.g[g][1] ? ratio(sc.g[g]) : null); }).join('') + cell(ratio(sc.all)); }
  function rowHtml(cls, label, sc, extra) {
    return '<tr class="' + cls + '"' + (extra || '') + '><th>' + label + '</th><td class="mp-n">' + sc.lapor + '/' + sc.n + '</td>' + gcells(sc) + '</tr>';
  }
  function storeLine(c, s, mine) {
    var chips = itemsFor(c, s.st).map(function (it) {
      var r = s.map[it.k];
      if (!r) return '<span class="mp-ch n" title="' + esc(it.label) + ': belum lapor">–</span>';
      if (r.status === 'ok') return '<button class="mp-ch g" data-mpview="' + esc(r.file) + '" data-cap="' + esc(s.code + ' ' + s.name + ' · ' + it.label) + '" title="' + esc(it.label) + ': terpasang">✓</button>';
      return '<span class="mp-ch r" title="' + esc(it.label + ': ' + r.reason + (r.note ? ' — ' + r.note : '')) + '" data-mpwhy="' + esc(it.label + ': ' + r.reason + (r.note ? ' — ' + r.note : '')) + '">✕</span>';
    }).join('');
    var bm = role() === 'ho' || role() === 'admin';
    return '<tr class="mp-st' + (mine ? ' me' : '') + (s.sc.lapor ? '' : ' belum') + '"><th><b>' + esc(s.code) + '</b> ' + esc(String(s.name).replace(/^EIGER\s+/i, '')) + '<div class="mp-chs">' + chips
      + (bm ? '<button class="mp-close" data-mpclose="' + esc(s.code) + '" data-nm="' + esc(s.name) + '" title="Tandai toko ini sudah tutup">Toko tutup?</button>' : '') + '</div></th>'
      + '<td class="mp-n">' + s.sc.lapor + '/' + s.sc.need + '</td>' + gcells(s.sc) + '</tr>';
  }
  function monHtml(R) {
    var c = R.camp, T = build(R), me = C.me(), isSite = role() === 'site', f = MON_ST.f, q = MON_ST.q.toLowerCase();
    var keep = function (s) {
      if (q && (s.code + ' ' + s.name).toLowerCase().indexOf(q) < 0) return false;
      if (f === 'belum') return !s.sc.lapor;
      if (f === 'kurang') return ratio(s.sc.all) < 1;
      if (f === 'tidak') return s.sc.no > 0;
      return true;
    };
    var N = T.nat, reasons = {};
    T.regs.forEach(function (G) { G.branches.forEach(function (B) { B.stores.forEach(function (s) { Object.keys(s.map).forEach(function (k) { var r = s.map[k]; if (r.status === 'no') reasons[r.reason] = (reasons[r.reason] || 0) + 1; }); }); }); });
    var opts = camps().map(function (x) { return '<option value="' + esc(x.id) + '"' + (x.id === c.id ? ' selected' : '') + '>' + esc(x.name) + ' · ' + fmtM(x.month) + (x.aktif ? '' : ' (ditutup)') + '</option>'; }).join('');
    var h = '<div class="mp-mon">'
      + '<div class="mp-bar"><select class="inp" data-mpsel>' + opts + '</select>'
      + (!isSite ? '<button class="btn sm" data-mpexp>' + svg('download') + 'Export Excel</button>' + (R.folder ? '<a class="btn sm" href="' + esc(R.folder) + '" target="_blank" rel="noopener">' + svg('folder') + 'Folder foto</a>' : '') : '')
      + '<button class="btn sm ghost" data-mpref title="Muat ulang">' + svg('refresh') + '</button></div>'
      + '<div class="mp-kpi" style="--n:' + Math.min(6, T.groups.length + 2) + '"><div class="k t-' + tone(ratio(N.all)) + '"><small>Pemasangan nasional</small><b>' + pct(ratio(N.all)) + '</b></div>'
      + T.groups.map(function (g) { var v = ratio(N.g[g]); return '<div class="k t-' + tone(v) + '"><small>' + esc(g) + '</small><b>' + pct(v) + '</b></div>'; }).join('')
      + '<div class="k"><small>Toko sudah lapor</small><b>' + N.lapor + '<span>/' + N.n + '</span></b></div></div>'
      + '<p class="small muted" style="margin:6px 2px 10px">Materi: ' + (c.items || []).map(function (x) { return esc(x.label) + ' <i>(' + esc(x.grp) + (x.scope ? ' · khusus ' + esc(x.scope) : '') + ')</i>'; }).join(' · ') + (c.deadline ? ' · batas lapor ' + fmtD(c.deadline) : '') + '. Persentase = materi terpasang ÷ materi yang wajib.</p>'
      + '<div class="mp-fl"><div class="seg">' + [['all', 'Semua'], ['belum', 'Belum lapor'], ['kurang', 'Belum 100%'], ['tidak', 'Ada yg tidak pasang']].map(function (x) { return '<button data-mpf="' + x[0] + '" class="' + (f === x[0] ? 'on' : '') + '">' + x[1] + '</button>'; }).join('') + '</div>'
      + '<input class="inp" type="search" data-mpq placeholder="Cari No Site / nama toko" value="' + esc(MON_ST.q) + '"></div>'
      + '<div class="mp-tw"><table class="mp-t"><thead><tr><th>Regional / Branch / Toko</th><th>Lapor</th>' + T.groups.map(function (g) { return '<th><span class="lg">' + esc(g) + '</span><span class="sm">' + esc(shortG(g)) + '</span></th>'; }).join('') + '<th>ACH</th></tr></thead><tbody>'
      + rowHtml('nat', 'NASIONAL', N);
    T.regs.forEach(function (G) {
      h += rowHtml('reg', esc(G.name), G.sc);
      G.branches.forEach(function (B) {
        var key = G.name + '|' + B.name, mineB = isSite && B.stores.some(function (s) { return s.code === String(me.code); });
        var shown = B.stores.filter(keep), open = MON_ST.open[key] != null ? MON_ST.open[key] : (mineB || !!q || f !== 'all');
        if ((q || f !== 'all') && !shown.length) return;
        h += rowHtml('br' + (open ? ' open' : '') + (mineB ? ' me' : ''), '<span class="mp-car">' + svg('next') + '</span>' + esc(B.name), B.sc, ' data-mpbr="' + esc(key) + '"');
        if (open) h += shown.map(function (s) { return storeLine(c, s, isSite && s.code === String(me.code)); }).join('') || '<tr class="mp-st"><th colspan="' + (T.groups.length + 3) + '" class="muted">Tidak ada toko sesuai filter.</th></tr>';
      });
    });
    h += '</tbody></table></div>';
    var RK = Object.keys(reasons).sort(function (a, b) { return reasons[b] - reasons[a]; });
    if (RK.length) h += '<div class="mp-rs"><h4>Alasan tidak pasang</h4>' + RK.map(function (k) { return '<div><span>' + esc(k) + '</span><b>' + reasons[k] + '</b></div>'; }).join('') + '</div>';
    h += '<p class="small muted mp-leg"><span class="mp-ch g">✓</span> terpasang (ketuk untuk lihat foto) <span class="mp-ch r">✕</span> tidak dipasang (ketuk untuk alasan) <span class="mp-ch n">–</span> belum lapor</p></div>';
    return h;
  }
  function mountMonitor(el, cid) {
    var L = camps();
    if (!L.length) { el.innerHTML = '<div class="empty">Belum ada monitoring pemasangan materi. Admin membuatnya di halaman admin → Monitoring materi.</div>'; return; }
    MON_ST.cid = cid || el.getAttribute('data-cid') || MON_ST.cid || (active()[0] || L[0]).id;
    if (!L.some(function (c) { return c.id === MON_ST.cid; })) MON_ST.cid = L[0].id;
    el.innerHTML = '<div class="empty"><span class="spin"></span>Memuat rekap…</div>';
    loadRekap(MON_ST.cid).then(function (R) { el.innerHTML = monHtml(R); hydrateImgs(el); })
      .catch(function (e) { el.innerHTML = '<div class="empty">Gagal memuat rekap: ' + esc(e.message || e) + '</div>'; });
  }
  function repaintMon(force) {
    document.querySelectorAll('[data-mpmon]').forEach(function (el) {
      if (force) delete REKAP[MON_ST.cid];
      var R = REKAP[MON_ST.cid]; if (R && !force) { var y = window.scrollY; el.innerHTML = monHtml(R); window.scrollTo(0, y); hydrateImgs(el); } else mountMonitor(el, MON_ST.cid);
    });
  }
  document.addEventListener('click', function (e) {
    var t = e.target.closest && e.target.closest('[data-mpbr],[data-mpf],[data-mpref],[data-mpexp],[data-mpwhy],[data-mpclose]'); if (!t) return;
    if (t.dataset.mpclose) {
      e.stopPropagation();
      var note = prompt('Tandai toko ' + t.dataset.mpclose + ' ' + t.dataset.nm + ' SUDAH TUTUP?\n\nToko tidak bisa login lagi dan tidak dihitung di monitoring. Admin bisa mengaktifkan kembali.\n\nKeterangan (mis. tutup sejak Okt 2026):', '');
      if (note === null) return;
      t.disabled = true;
      C.post({ a: 'siteClose', code: t.dataset.mpclose, note: note.trim() }).then(function () { C.toast('Toko ' + t.dataset.mpclose + ' ditandai tutup.'); REKAP = {}; repaintMon(true); })
        .catch(function (er) { t.disabled = false; C.toast('Gagal: ' + (er.message || er.code), 6000); });
      return;
    }
    if (t.dataset.mpwhy) { C.toast(t.dataset.mpwhy, 5000); return; }
    if (t.dataset.mpbr) { var k = t.dataset.mpbr, cur = t.classList.contains('open'); MON_ST.open[k] = !cur; repaintMon(); return; }
    if (t.dataset.mpf) { MON_ST.f = t.dataset.mpf; MON_ST.open = {}; repaintMon(); return; }
    if (t.dataset.mpref != null) { repaintMon(true); return; }
    if (t.dataset.mpexp != null) { var R = REKAP[MON_ST.cid]; if (R) exportX(R).catch(function (er) { C.toast(er.message || 'Gagal membuat Excel.', 6000); }); }
  });
  document.addEventListener('change', function (e) { var t = e.target; if (t.matches && t.matches('[data-mpsel]')) { MON_ST.cid = t.value; MON_ST.open = {}; repaintMon(); } });
  var qT = null;
  document.addEventListener('input', function (e) {
    var t = e.target; if (!t.matches || !t.matches('[data-mpq]')) return;
    clearTimeout(qT); qT = setTimeout(function () { MON_ST.q = t.value; MON_ST.open = {}; repaintMon(); var n = document.querySelector('[data-mpq]'); if (n) { n.focus(); n.setSelectionRange(n.value.length, n.value.length); } }, 300);
  });

  /* ------------------------------------------------ export Excel (tanpa gambar → ringan; foto berupa link Drive) */
  var _ex = null;
  function needExcelJS() {
    if (window.ExcelJS) return Promise.resolve(true);
    var U = ['https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js', 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js'];
    return _ex || (_ex = new Promise(function (res) {
      var i = 0; (function next() { if (i >= U.length) { _ex = null; res(false); return; } var s = document.createElement('script'); s.src = U[i++]; s.onload = function () { res(!!window.ExcelJS); }; s.onerror = next; document.head.appendChild(s); })();
    }));
  }
  async function exportX(R) {
    C.toast('Menyiapkan Excel…', 8000);
    if (!(await needExcelJS())) throw new Error('Pembuat Excel gagal dimuat. Cek koneksi internet.');
    var c = R.camp, T = build(R), wb = new window.ExcelJS.Workbook(), NAVY = 'FF073763';
    var LINE = { style: 'thin', color: { argb: 'FFB8C2D0' } }, B4 = { top: LINE, left: LINE, bottom: LINE, right: LINE };
    var FILL = { g: 'FFD9F2E3', a: 'FFFFF1CC', r: 'FFFCDADA', n: 'FFF1F5F9' };
    var hdr = function (row) { row.eachCell(function (x) { x.font = { bold: true, color: { argb: 'FFFFFFFF' }, size: 9.5 }; x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }; x.border = B4; x.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; }); };
    var pc = function (x, v) { x.value = v == null ? null : v; x.numFmt = '0%'; x.alignment = { horizontal: 'center', vertical: 'middle' }; x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: FILL[tone(v)] } }; };
    var setup = function (ws, rows, orient) { ws.pageSetup = { paperSize: 9, orientation: orient || 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.3, right: 0.3, top: 0.45, bottom: 0.5, header: 0.2, footer: 0.25 }, printTitlesRow: rows }; ws.headerFooter = { oddFooter: '&L&8 ' + c.name.replace(/&/g, '&&') + '&R&8 Hal &P / &N' }; };
    var dt = new Date().toLocaleString('id-ID');
    /* 1. rekap */
    var ws = wb.addWorksheet('Rekap', { views: [{ state: 'frozen', ySplit: 5, showGridLines: false }] });
    var GR = T.groups, NG = GR.length;
    ws.columns = [{ width: 16 }, { width: 18 }, { width: 10 }, { width: 11 }, { width: 10 }].concat(GR.map(function () { return { width: 11 }; })).concat([{ width: 10 }]);
    ws.addRow(['MONITORING PEMASANGAN MATERI PROMO']).font = { bold: true, size: 14, color: { argb: NAVY } };
    ws.addRow([c.name + ' · ' + fmtM(c.month) + (c.deadline ? ' · batas lapor ' + fmtD(c.deadline) : '')]).font = { bold: true, size: 11 };
    ws.addRow(['Materi: ' + (c.items || []).map(function (x) { return x.label + ' (' + x.grp + (x.scope ? ', khusus ' + x.scope : '') + ')'; }).join(', ') + '   ·   Data per ' + dt]).font = { size: 9, italic: true, color: { argb: 'FF64748B' } };
    ws.addRow([]);
    hdr(ws.addRow(['REGIONAL', 'BRANCH', 'JUMLAH TOKO', 'SUDAH LAPOR', '% LAPOR'].concat(GR.map(function (g) { return g.toUpperCase(); })).concat(['ACH'])));
    var put = function (a, b, sc, lvl) {
      var r = ws.addRow([a, b, sc.n, sc.lapor]); pc(r.getCell(5), sc.n ? sc.lapor / sc.n : null); GR.forEach(function (g, i) { pc(r.getCell(6 + i), sc.g[g][1] ? ratio(sc.g[g]) : null); }); pc(r.getCell(6 + NG), ratio(sc.all));
      r.eachCell({ includeEmpty: true }, function (x, n) { x.border = B4; if (n <= 4) { x.alignment = { horizontal: n <= 2 ? 'left' : 'center' }; if (lvl < 2) x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: lvl === 0 ? 'FFDCE4EF' : 'FFEEF2F8' } }; } x.font = { bold: lvl < 2, size: lvl === 0 ? 10.5 : 9.5 }; });
    };
    put('NASIONAL', '', T.nat, 0);
    T.regs.forEach(function (G) { put(G.name, '', G.sc, 1); G.branches.forEach(function (B) { put('', B.name, B.sc, 2); }); });
    setup(ws, '5:5', 'portrait');
    /* 2. detail toko */
    var IT = c.items || [], d = wb.addWorksheet('Detail toko', { views: [{ state: 'frozen', ySplit: 6, xSplit: 4, showGridLines: false }] });
    var nIt = IT.length, P0 = 5 + nIt, F0 = P0 + NG + 1;                        // kolom persen & foto
    d.columns = [{ width: 8 }, { width: 13 }, { width: 8 }, { width: 30 }].concat(IT.map(function () { return { width: 16 }; })).concat(GR.map(function () { return { width: 8 }; })).concat([{ width: 8 }]).concat(IT.map(function () { return { width: 11 }; }));
    d.addRow(['DETAIL PEMASANGAN PER TOKO · ' + c.name]).font = { bold: true, size: 13, color: { argb: NAVY } };
    d.addRow(['Status: Terpasang / Tidak dipasang (alasan) / Belum lapor. Kolom "Foto" berisi link ke Google Drive.   ·   ' + dt]).font = { size: 9, italic: true, color: { argb: 'FF64748B' } };
    d.addRow([]);
    var g1 = ['', '', '', ''].concat(IT.map(function (x) { return String(x.grp).toUpperCase(); })).concat(GR.map(function () { return '%'; })).concat(['%']).concat(IT.map(function () { return 'FOTO'; }));
    var r1 = d.addRow(g1); hdr(r1);
    var r2 = d.addRow(['REG', 'BRANCH', 'SITE', 'NAMA TOKO'].concat(IT.map(function (x) { return x.label.toUpperCase(); })).concat(GR.map(function (g) { return g.toUpperCase(); })).concat(['ACH']).concat(IT.map(function (x) { return x.label; }))); hdr(r2); r2.height = 30;
    d.addRow([]);
    T.regs.forEach(function (G) {
      G.branches.forEach(function (B) {
        B.stores.forEach(function (s) {
          var vals = [G.name, B.name, s.code, s.name].concat(IT.map(function (it) { if (!inScope(it, s.st)) return 'Tidak wajib'; var r = s.map[it.k]; return !r ? 'Belum lapor' : r.status === 'ok' ? 'Terpasang' : 'Tidak: ' + r.reason + (r.note ? ' (' + r.note + ')' : ''); }));
          var row = d.addRow(vals.concat(GR.map(function () { return null; })).concat([null]).concat(IT.map(function () { return ''; })));
          row.font = { size: 9 };
          row.eachCell({ includeEmpty: true }, function (x, n) { x.border = B4; x.alignment = { vertical: 'middle', wrapText: n > 4 && n <= 4 + nIt, horizontal: n === 3 ? 'center' : n > 4 && n <= 4 + nIt ? 'center' : 'left' }; });
          IT.forEach(function (it, j) {
            var r = s.map[it.k], x = row.getCell(5 + j);
            if (!inScope(it, s.st)) { x.font = { size: 8.5, color: { argb: 'FF94A3B8' } }; return; }
            x.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: !r ? FILL.n : r.status === 'ok' ? FILL.g : FILL.r } };
            if (r && r.status === 'ok' && r.file) { var lk = row.getCell(F0 + j); lk.value = { text: 'Lihat foto', hyperlink: 'https://drive.google.com/file/d/' + r.file + '/view' }; lk.font = { size: 9, color: { argb: 'FF1155CC' }, underline: true }; }
          });
          GR.forEach(function (g, i) { pc(row.getCell(P0 + i), s.sc.g[g][1] ? ratio(s.sc.g[g]) : null); }); pc(row.getCell(P0 + NG), ratio(s.sc.all));
        });
      });
    });
    setup(d, '4:5');
    /* 3. alasan */
    var a = wb.addWorksheet('Alasan tidak pasang', { views: [{ showGridLines: false }] });
    a.columns = [{ width: 12 }, { width: 14 }, { width: 8 }, { width: 32 }, { width: 22 }, { width: 34 }, { width: 30 }];
    a.addRow(['ALASAN TIDAK PASANG · ' + c.name]).font = { bold: true, size: 13, color: { argb: NAVY } };
    a.addRow([]);
    hdr(a.addRow(['REG', 'BRANCH', 'SITE', 'NAMA TOKO', 'MATERI', 'ALASAN', 'KETERANGAN']));
    T.regs.forEach(function (G) { G.branches.forEach(function (B) { B.stores.forEach(function (s) { IT.forEach(function (it) { var r = s.map[it.k]; if (r && r.status === 'no') { var w = a.addRow([G.name, B.name, s.code, s.name, it.label, r.reason, r.note]); w.font = { size: 9 }; w.eachCell({ includeEmpty: true }, function (x) { x.border = B4; x.alignment = { wrapText: true, vertical: 'top' }; }); } }); }); }); });
    setup(a, '3:3');
    var buf = await wb.xlsx.writeBuffer(), fn = 'Monitoring_Materi_' + c.name.replace(/[^\w]+/g, '_').slice(0, 40) + '_' + today() + '.xlsx';
    var bl = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), ln = document.createElement('a');
    ln.href = URL.createObjectURL(bl); ln.download = fn; document.body.appendChild(ln); ln.click(); setTimeout(function () { URL.revokeObjectURL(ln.href); ln.remove(); }, 1500);
    C.toast('Tersimpan: ' + fn, 5000);
  }

  /* ------------------------------------------------ pasang otomatis saat elemen muncul */
  function scan() {
    if (!C) return;
    document.querySelectorAll('[data-mpstore]:not([data-mpdone])').forEach(function (el) { el.setAttribute('data-mpdone', 1); mountStore(el); });
    document.querySelectorAll('[data-mpmon]:not([data-mpdone])').forEach(function (el) { el.setAttribute('data-mpdone', 1); mountMonitor(el); });
    document.querySelectorAll('[data-mpban]:not([data-mpdone])').forEach(function (el) { el.setAttribute('data-mpdone', 1); mountBanner(el); });
  }
  if (window.MutationObserver) new MutationObserver(function () { clearTimeout(scan._t); scan._t = setTimeout(scan, 20); }).observe(document.documentElement, { childList: true, subtree: true });
  window.MP = {
    setup: function (cfg) { C = cfg; scan(); },
    openCamp: function (cid) { MON_ST.cid = cid; MON_ST.open = {}; },
    reset: function () { MINE = {}; MINE_AT = {}; REKAP = {}; },
    compress: compress, build: build, storeScore: storeScore, exportX: exportX
  };
})();
