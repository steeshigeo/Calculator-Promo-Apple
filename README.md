# Kalkulator Promo Apple

Kalkulator promo statis (HTML) yang datanya diperbarui otomatis dari file master Excel / Google Sheets,
dengan pola yang sama seperti Dashtp3staff:

```
File master (.xlsx)                 ← tim trainer edit di sini
   │  GitHub Action (tiap 30 menit + tombol "Run workflow")
   ▼
sync_onedrive.py  → baca 6 sheet → data.json  (di-commit otomatis bila ada perubahan)
   │  push → Vercel / hosting statis ter-deploy ulang
   ▼
index.html membaca data.json  (dicek ulang tiap 5 menit selama sesi aktif)
```

| File | Fungsi |
|---|---|
| `index.html` | Kalkulator. Membaca `data.json`, menyimpan salinan di perangkat (jalan offline dengan data terakhir). |
| `data.json` | Hasil parsing master. **Jangan diedit manual**, akan ditimpa otomatis. |
| `sync_onedrive.py` | Unduh master → validasi → tulis `data.json`. |
| `.github/workflows/onedrive_sync.yml` | Jadwal otomatis + tombol manual. |
| `requirements.txt` | `openpyxl`, `requests`. |

## Pasang (sekali saja)

1. **Buat repo GitHub — pilih Private.** `data.json` berisi harga, nilai trade-in, dan harga provider/Qoala.
   Di repo publik semuanya bisa dibaca siapa pun.
2. Upload semua file di folder ini ke repo (termasuk folder `.github`).
3. **Secret link master**: *Settings → Secrets and variables → Actions → New repository secret*
   - Name: `ONEDRIVE_DIRECT_URL`
   - Value: link file master. Yang didukung:
     - link Google Sheets biasa (`https://docs.google.com/spreadsheets/d/…`, dibagikan "siapa saja yang memiliki link"), otomatis di-export ke xlsx
     - link share OneDrive / SharePoint, atau direct link `.xlsx`
4. **Izin bot**: *Settings → Actions → General → Workflow permissions → Read and write permissions*.
5. **Jalankan pertama kali**: tab *Actions → Sync Promo Data → Run workflow*. Pastikan hijau, lalu cek `data.json` ikut ter-commit.
6. **Deploy**: import repo ke Vercel (atau Netlify / Cloudflare Pages). Setiap commit bot akan men-deploy ulang otomatis.
   Bila memakai GitHub Pages (repo privat butuh paket berbayar), cek setelah run pertama bahwa situs ikut ter-update.

## Cara update harga / promo

1. Edit file master seperti biasa.
2. Tunggu maksimal ±30 menit, **atau** buka *Actions → Sync Promo Data → Run workflow* untuk langsung.
3. Perangkat staff memuat data baru otomatis (maksimal 5 menit) atau lewat tombol ↻ Refresh.

Bila sheet salah format, workflow **gagal (merah)** dengan pesan jelas dan `data.json` lama tetap dipakai — kalkulator tidak rusak.
Pesan umum:

| Pesan | Artinya |
|---|---|
| `Header sheet 'Price List' tidak ditemukan` | Nama kolom diubah. Butuh: SAP Article, SAPDescription, Category, Normal Price, Promotion Price |
| `Sheet 'BNPL' tidak ditemukan` | Nama tab diubah. Nama tab harus persis: Price List, Promo Berjalan, BNPL, Provider, Qoala Protection, Trade in |
| `File bukan .xlsx` / `bukan .xlsx (halaman login)` | Link tidak publik atau bukan file xlsx |
| `Jumlah produk turun drastis` | Sheet terpotong / terhapus sebagian; update ditolak |

Di log run juga muncul **peringatan** (kuning) yang perlu dirapikan di sheet, misalnya varian yang harganya beda tapi
label RAM/Storage sama (hanya satu yang bisa dipilih di kalkulator) dan promo yang periodenya sudah lewat.

## Aturan format master

- **Price List**: judul model = baris dengan **kolom A saja** terisi (kolom B dan C kosong). Baris tanpa judul sendiri ikut judul di atasnya,
  jadi varian berbeda model (mis. iPhone 15 vs iPhone 15 Plus) harus punya baris judul masing-masing.
- **Promo Berjalan**: kolom `Scheme` hanya `Direct Discount` atau `CB by Billing`. Kolom `Periode` kosong = ikut baris di atasnya.
  Tanggal berakhir dibaca dari teks Periode (mis. `24 April - 13 September`, `Hingga 31 Desember`); promo yang sudah lewat otomatis
  disembunyikan di kalkulator. Periode yang tidak terbaca tetap ditampilkan.
- **Kartu kredit**: tenor cicilan 0% ada di `CARD_MATRIX` pada `sync_onedrive.py` (bukan di sheet). Bank yang muncul di sheet Promo
  tapi belum ada di `CARD_MATRIX` ditambahkan otomatis dengan tenor 3/6/12 dan diberi peringatan di log.

## Uji di komputer sendiri

```bash
pip install -r requirements.txt
python sync_onedrive.py --file "Calculator Promo.xlsx"     # menghasilkan data.json
python -m http.server 8000                                  # buka http://localhost:8000
```
(Membuka `index.html` langsung dengan klik ganda tidak bisa memuat `data.json` karena aturan browser; pakai server lokal di atas.)

## Catatan keamanan

Password kalkulator dan admin ada di dalam `index.html`, jadi hanya membatasi akses kasual, bukan pengaman data.
Data di `data.json` bisa diakses siapa pun yang tahu URL situsnya. Bila perlu pembatasan sungguhan, aktifkan
proteksi akses di hosting (mis. Vercel Password Protection atau Cloudflare Access).

## Kuota GitHub Actions

Repo privat paket gratis mendapat 2.000 menit/bulan. Jadwal 30 menit memakai ±1.500 menit (tiap run dibulatkan 1 menit).
Jangan dipercepat menjadi 15 menit di repo privat karena melewati kuota.
