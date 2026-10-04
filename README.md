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
| `index.html` | Kalkulator (desain Liquid Glass, terang/gelap, tampilan ponsel & desktop). Membaca `data.json`, menyimpan salinan di perangkat (jalan offline dengan data terakhir). |
| `data.json` | Hasil parsing master. **Jangan diedit manual**, akan ditimpa otomatis. |
| `sync_onedrive.py` | Unduh master → validasi → tulis `data.json`. |
| `.github/workflows/onedrive_sync.yml` | Jadwal otomatis + tombol manual. |
| `access.json` | Hash password staff & admin (PBKDF2, tanpa password asli). **Upload sekali, jangan ditimpa lagi**; diubah otomatis lewat halaman Admin. |
| `update_access.py` + `.github/workflows/access_update.yml` | Workflow "Update Access": menulis hash password baru ke `access.json`. |
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

## Varian / Warna & laporan perubahan harga

- Satu **tipe + storage** bisa punya banyak baris di sheet (warna, atau iPhone 15 vs iPhone 15 Plus yang berbagi judul). Kalkulator menampilkan dropdown
  **Varian / Warna** berisi semua baris itu (dengan harganya bila berbeda), jadi harga di baris mana pun yang kamu ubah bisa dilihat dan dipilih.
  Sales talk menyebut varian yang dipilih. Disarankan tetap menambah baris judul sendiri (mis. "iPhone 15 Plus") agar pilihan lebih ringkas.
- Log Action sekarang memuat **laporan perubahan**: setiap harga yang berubah ditulis, mis.
  `[iPhone] iPhone 15 128GB Pink: promo Rp15.499.000 -> Rp9.999.000`, plus ringkasan jumlah. Bila log hanya menulis
  "Tidak ada perubahan data", artinya `data.json` sudah sama dengan file master saat itu.
- **Jangan menimpa `data.json` dan `access.json`** di repo saat meng-update file lain; keduanya dibuat/diubah otomatis oleh Action.

## Menu, halaman, Run & Refresh

- **Sidebar (☰)**: Home (kalkulator, halaman awal), Guide (panduan pemakaian), About (deskripsi & manfaat untuk staff, dengan statistik data), Admin, pengaturan Tampilan (Auto/Terang/Gelap) dan Getaran, serta **Contributors**.
- **Run (▶)**: memicu workflow `onedrive_sync.yml` (repo `steeshigeo/Calculator-Promo-Apple`) dan menampilkan popup
  *"Sync in progress, wait ~20s then tap Refresh."* dengan hitung mundur di home (di teks status dan lencana tombol Refresh).
  Bila sudah ada run yang berjalan, Run tidak memicu ganda. Jeda 45 detik antar-Run.
- **Refresh (↻)**: memuat `data.json` terbaru ke kalkulator (berdenyut saat hitung mundur selesai). Bila belum berubah, tunggu ±1 menit
  (deploy situs) lalu tap lagi.
- Token GitHub dibaca dari **sel A2** Google Sheet token (lewat JSONP) dan hanya disimpan di memori halaman, tidak di storage.

**Token yang aman** (siapa pun yang memegang link Google Sheet token bisa membacanya):

- Pakai **fine-grained personal access token**, hanya repo `Calculator-Promo-Apple`, izin **Actions: Read and write** (Metadata: Read otomatis), dengan masa berlaku.
- Dampak terburuk bila bocor: memicu workflow berulang (kuota menit), menghapus log run, **atau mereset password lewat workflow Update Access**.
- Opsi lebih aman: pindahkan pemicu ke Vercel Serverless Function (token di Environment Variable, tidak sampai ke browser).

## Password staff & admin (sinkron ke semua staff)

- Password disimpan sebagai **hash PBKDF2-SHA256** (salt acak, 150.000 iterasi) di `access.json`, bukan teks biasa di kode. Semua perangkat membaca file yang sama.
- **Awal pemasangan**: `access.json` berisi password lama (staff `maptech`, admin `digiceria`, username admin `MAPTECH`). Segera ganti lewat Admin.
- **Mengganti**: sidebar → Admin → masukkan kredensial admin → isi password baru → *Terapkan ke semua staff*. Browser mengirim **hanya salt + hash**
  ke workflow `Update Access`, yang meng-commit `access.json`. Dalam ±1–2 menit semua staff memakai password baru; staff yang sedang login diminta masuk ulang.
- Password admin juga bisa dipakai untuk masuk sebagai staff (untuk pemulihan bila password staff terlupa).
- Ini **gerbang ringan**, bukan keamanan data: `data.json` tetap bisa dibuka siapa pun yang tahu URL situsnya. Untuk pembatasan sungguhan pakai proteksi akses hosting (Vercel Password Protection / Cloudflare Access).
- Butuh HTTPS (Vercel otomatis) karena pemeriksaan hash memakai Web Crypto.

## Tampilan & perangkat

- Desktop (layar ≥ 1100 px) memakai tata letak penuh dua kolom: kartu langkah di kiri, Rincian Simulasi dan Sales Talk menempel di kanan.
- **Getaran**: Android memakai Vibration API. **iOS: sejak iOS 26.5 Apple menambal getar lewat script**; yang masih berfungsi hanya sentuhan langsung ke switch asli.
  Karena itu di iPhone/iPad kalkulator memasang switch `<input type="checkbox" switch>` transparan tepat di atas tab, menu, Run, dan Refresh, sehingga jari menyentuh switch itu.
  Perlu **Settings → Sounds & Haptics → System Haptics** aktif. Saklar *Getaran* di sidebar memakai switch bawaan iOS, jadi mengetuknya sekaligus menjadi tes getar. Laptop/desktop tidak punya motor getar.
- Promo yang namanya memuat **Samsung** diabaikan (`IGNORE_PROMO_KEYWORDS` di `sync_onedrive.py` dan `index.html`).
- Dropdown Trade In menampilkan nama berhuruf kapital di awal (Samsung, Galaxy A54 5G); `iPhone` dan `iPad` ditulis khusus. Hanya tampilan.

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
| `Gagal mengunduh file master … HTTP 401/403` | Link OneDrive tidak terbuka untuk umum. Lihat "Link OneDrive ditolak (401)" di bawah |
| `File bukan .xlsx` / `halaman login` | Link tidak publik atau bukan file xlsx |
| `Jumlah produk turun drastis` | Sheet terpotong / terhapus sebagian; update ditolak |

Di log run juga muncul **peringatan** (kuning) yang perlu dirapikan di sheet, misalnya varian yang harganya beda tapi
label RAM/Storage sama (hanya satu yang bisa dipilih di kalkulator) dan promo yang periodenya sudah lewat.

## Link OneDrive ditolak (401)

GitHub Action berjalan tanpa login, jadi link master harus bisa dibuka **tanpa akun**:

1. Buka file master di OneDrive → **Bagikan** → klik ikon pengaturan link.
2. Pilih **Siapa saja yang memiliki link**, izin **Dapat melihat**. Pilihan "Orang tertentu" atau "Orang di organisasi Anda" akan selalu 401.
3. **Salin link baru**, lalu ganti isi secret `ONEDRIVE_DIRECT_URL` (link lama tidak ikut berubah izinnya).
4. Jalankan ulang *Run workflow*.

Akun kantor/sekolah (SharePoint) sering memblokir "Siapa saja" oleh admin. Kalau pilihan itu tidak muncul, pakai link
Google Sheets (dibagikan "Siapa saja yang memiliki link" → Viewer) di secret yang sama; script otomatis mengubahnya ke export xlsx.
Log run akan menyebut host dan kode HTTP tiap percobaan (tanpa membuka isi link-nya).

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

Lihat bagian *Password staff & admin* dan *Token yang aman* di atas. Ringkas: password hanya gerbang akses ringan, data harga tetap publik bagi yang tahu URL-nya,
dan token GitHub di Google Sheet harus fine-grained dengan izin seminimal mungkin.

## Kuota GitHub Actions

Repo privat paket gratis mendapat 2.000 menit/bulan. Jadwal 30 menit memakai ±1.500 menit (tiap run dibulatkan 1 menit).
Jangan dipercepat menjadi 15 menit di repo privat karena melewati kuota.
