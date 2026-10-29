# Promo Eigerindo

Halaman promo nasional: toko masuk dengan **No Site**, lalu melihat promo yang sedang berjalan, kalender bulanan, syarat & ketentuan, contoh penerapan, memo update, dan link Excel list artikel.

Admin cukup mengunggah file juknis (PDF atau PPTX). Isinya dibaca otomatis, ditampilkan untuk dikoreksi, lalu diterbitkan ke seluruh toko.

---

## Isi repo

| File | Isinya |
|---|---|
| `index.html` | Halaman untuk toko (login No Site) |
| `admin.html` | Halaman admin (unggah juknis, koreksi, terbitkan, kelola toko) |
| `parse.js` | Pembaca file juknis PDF & PPTX di browser |
| `app.css` | Tampilan bersama (mobile: app bar + bottom nav, desktop: sidebar) |
| `icons.js` | Ikon yang dipakai kedua halaman |
| `artikel.js` | Baca Excel list artikel, baca Export by Bin, pencocokan (VLOOKUP), dan ekspor Excel |
| `img/logo.png`, `img/logo-white.png` | Logo Eigerindo MPI (latar transparan, versi biru & putih) |
| `config.js` | Satu-satunya file yang perlu kamu ubah setelah deploy |
| `Code.gs` | Backend Apps Script (disalin ke script.google.com, bukan ke GitHub Pages) |

---

## Pemasangan (sekali saja, sekitar 15 menit)

### 1. Backend — Apps Script + Spreadsheet

1. Buka [script.google.com](https://script.google.com) → **New project**.
2. Hapus isi `Code.gs` bawaan, tempel seluruh isi `Code.gs` dari repo ini.
3. Jalankan fungsi **`setup`** sekali (pilih `setup` di dropdown → Run → izinkan akses).
   Buka **Execution log**: di situ ada link spreadsheet database yang baru dibuat. Simpan linknya.
4. **Project Settings → Script properties → Add script property**
   - `ADMIN_PW` → password admin yang kamu mau (ganti dari bawaan `eiger123`).
5. **Deploy → New deployment → Web app**
   - Description: bebas
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Deploy → salin URL yang berakhiran `/exec`.

> Setiap kali `Code.gs` diubah, ulangi: **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**. URL-nya tetap sama.

### 2. Frontend — GitHub Pages

1. Buat repo baru, unggah semua file termasuk folder `img/` (struktur foldernya harus sama).
2. Buka `config.js`, ganti `API` dengan URL `/exec` tadi.
3. **Settings → Pages → Source: Deploy from a branch → `main` / root** → Save.
4. Tunggu ±1 menit. Alamatnya jadi:
   - Toko: `https://<user>.github.io/<repo>/`
   - Admin: `https://<user>.github.io/<repo>/admin.html`

### 3. Daftar toko nasional

Buka `admin.html` → masuk → tab **Daftar toko** → tempel daftar dari Excel, satu baris per toko:

```
2278;EIGER Adventure Store Jogja;Jateng 01
2279;EIGER Solo Paragon;Jateng 01
```

Kolom kedua dan ketiga boleh dikosongkan (`2278` saja juga bisa). Pemisah boleh `;`, tab, atau koma — jadi tinggal copy-paste langsung dari Excel.

---

## Alur kerja tiap bulan

1. Buka `admin.html`, tab **Unggah & terbitkan**.
2. Seret file juknis (PDF atau PPTX).
3. Periksa hasil bacanya. Tiap promo bisa diubah: judul, kategori, periode, syarat, ketentuan, tips kasir, contoh penerapan.
4. Tempel **link Excel list artikel** di promo yang membutuhkannya (tombol Ubah → bagian paling bawah).
5. Isi **link file juknis** (Drive/SharePoint) supaya toko bisa membuka slide aslinya.
6. Tekan **Terbitkan ke toko**. Selesai — seluruh toko langsung melihat versi baru tanpa perlu diberi tahu.

Ada memo update di tengah bulan? Unggah file memonya, atau tekan **Edit bulan ini** di tab *Promo tersimpan*, tambahkan memo, lalu terbitkan ulang.

---

## Yang perlu diketahui soal pembacaan otomatis

Pembacaan otomatis mengandalkan tata letak slide juknis: judul promo di atas, blok **Syarat & Ketentuan** di kolom kiri, **Contoh Penerapan** di kolom kanan, dan halaman kategori bertuliskan “PROMO …”. Selama template HO tidak berubah, hasilnya rapi.

Begitu templatenya berubah, sebagian field bisa kosong atau masuk ke tempat yang salah. Karena itu alurnya selalu **baca → koreksi → terbitkan**, bukan terbit langsung. Kalau satu slide gagal terbaca sama sekali, tekan **Tambah promo** dan isi manual — halaman toko tetap tampil normal.

Slide yang tidak dikenali ditampilkan apa adanya lewat tombol *Lihat isinya*, jadi kamu bisa menyalin teksnya tanpa membuka file aslinya.

---

## Keamanan & privasi

- Login toko hanya No Site, tanpa password — cukup untuk membatasi akses ke daftar toko resmi, bukan untuk data rahasia. Halaman ini memang hanya berisi informasi promo.
- Halaman admin dikunci password terpisah yang disimpan di Script Properties, tidak di repo.
- File juknis dibaca di browser admin; yang dikirim ke spreadsheet hanya hasil teksnya.
- Daftar toko yang sedang membuka halaman terlihat di header (angka) dan lengkap di tab **Toko online & log** pada halaman admin.

## Batasan yang diketahui

- Satu bulan hanya menyimpan satu versi juknis. Menerbitkan ulang akan mengganti isi bulan tersebut (riwayatnya tercatat di sheet `Log`).
- File juknis dan Excel tidak disimpan di spreadsheet — yang disimpan linknya. Unggah filenya ke Drive atau SharePoint dulu, lalu tempel linknya.
- Apps Script punya kuota harian. Dengan ±400 toko yang membuka beberapa kali sehari masih aman, tapi kalau nanti terasa lambat, naikkan jeda `ping` di `index.html` (sekarang 60 detik).

---

## List artikel promo & cek stok toko

**Admin** → menu *List artikel promo*: pilih bulan, unggah Excel list artikel. Satu file boleh berisi banyak sheet; tiap sheet menjadi satu list. Untuk tiap sheet bisa diatur nama list, kolom kode, kolom nama, kolom warna, kolom info yang ditampilkan ke toko, dan promo yang dikaitkan (terisi otomatis dari kemiripan nama sheet dengan judul promo). Mengunggah ulang untuk bulan yang sama akan menggantikan list lama.

Cara kode dicocokkan:
- 12–13 digit → SKU persis
- 9 digit → artikel, semua warna & ukuran
- 9 digit + kolom Warna (atau `910005461.BLK`) → artikel di warna itu saja

**Toko** → menu *Cek stok promo*: unggah file Export by Bin dari sistem. Hasilnya: SKU promo yang ada di toko, pcs terdisplay vs gudang, SKU yang belum dipajang beserta lokasi bin, dan kode list yang tidak ada stoknya. Bisa diekspor ke Excel (Ringkasan, Semua SKU promo, Per bin, satu sheet per list, Tidak ada stok).

File Export by Bin **tidak dikirim ke server**; semuanya diproses di perangkat toko.
