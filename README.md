# Sistem Portal Ujian Google Form Anticurang & Monitoring (Arsitektur SIGURU v3.0)

Aplikasi serverless berbasis **Google Apps Script**, **Google Sheets Database**, dan **Google Drive Storage** yang responsif untuk **Desktop & Mobile (`100dvh`)**, dengan pemisahan total antara **Web Publik (Siswa)** dan **Dashboard Kontrol (Admin / Pengawas)**.

---

## Daftar Fitur Utama

1. **Server-Side Template Inclusion (Scriptlet Pattern)**: `StyleCSS.html`, `Sidebar.html`, `Header.html`, `AppJS.html`.
2. **Multi-Page Serverless Web App (MPA)** via Query Parameter Routing (`?page=portal|monitoring|laporan|kelola-soal|data-mapel|data-murid|pengaturan`) dengan navigasi induk `window.top.location.href` anti blank putih.
3. **Web Publik Jadwal per Hari & Mode Anticurang Multi-User** (`Portal.html`):
   - Tampilan bersih khusus siswa (tanpa sidebar/login admin).
   - Kolom **Kelas** peserta pada modal konfirmasi otomatis terisi sesuai jadwal soal.
   - Kunci layar penuh (*Fullscreen*), deteksi pindah tab/jendela, blokir klik kanan/salin-tempel/DevTools, audio peringatan, overlay kunci pengawas, serta pencatatan realtime ke `Log_Monitoring`.
4. **Pemisahan Jam Buka vs Durasi Soal (Opsi B)**:
   - `Waktu Mulai` & `Waktu Selesai` mengatur kapan soal dapat dibuka dan batas terakhir siswa boleh membukanya.
   - Kolom **H (`Durasi (Menit)`)** menentukan durasi pengerjaan penuh dalam menit (misal `30`) terhitung sejak siswa membuka soal, tetap berjalan penuh meskipun melewati `Waktu Selesai` (kosongkan jika fleksibel).
5. **Master Data Mapel & Data Murid** (`?page=data-mapel` & `?page=data-murid`):
   - Kelola data mata pelajaran (`DataMapel.html`) dan data murid/peserta didik (`DataMurid.html`).
   - Form Tambah/Edit Soal dilengkapi nomor otomatis yang menghitung kelanjutan data di database, serta input Kelas dan Mapel yang dapat dipilih maupun diketik (searchable datalist).
   - Pada modal siswa, Nama Lengkap Siswa dapat dipilih via daftar maupun diketik untuk mencari nama secara instan, dan NIS otomatis terisi sesuai siswa terpilih.
6. **Three-Tier / Two-Tier Access Control**:
   - **Publik / Siswa**, **Pengawas Ruang** (Monitoring, Laporan, dan PIN buka kunci peserta), dan **Administrator** (akses penuh termasuk Kelola Link Soal, Data Mapel & Kelas, Hapus Log, dan Pengaturan Sistem).
7. **SheetJS Excel Processing & Dynamic Adaptive Print / PDF Engine** (`Laporan.html`):
   - Ekspor `.xls` (BIFF8) di Monitoring, Laporan, dan Kelola Soal; Impor Excel jadwal soal; Smart Kop Divider otomatis untuk gambar maupun teks.
8. **Dual-Mode Logo & Kop Surat Ingestion** (`Pengaturan.html`): Mode URL Google Drive/Gambar vs Mode Unggah Berkas ke Google Drive.

---

## URL Deployment Saat Ini & Hak Akses

- **Project Apps Script**: `10om5OlL4AEtNkp5O5OHu9-5v8z38eyhhxWIytGeAmakklYii7-MYFT-g`
- **Link Web Publik (Khusus Siswa)**:
  `https://script.google.com/macros/s/AKfycbxw20eF7UAvQv4zJJlMds32mUjU42KInSNiLgGKuZz9_sF76sOH4xhQEcvHkg9V-rO9/exec`
- **Link Dashboard Admin / Pengawas**:
  `https://script.google.com/macros/s/AKfycbxw20eF7UAvQv4zJJlMds32mUjU42KInSNiLgGKuZz9_sF76sOH4xhQEcvHkg9V-rO9/exec?page=admin` (atau `?page=pengawas`)
- **PIN Login Admin**: Tersimpan di Database Sheet `Pengaturan` kolom `PIN Admin` (`G2`, default awal `123456`) — akses penuh Admin.
- **PIN Login & Buka Kunci Pengawas**: Tersimpan di Database Sheet `Pengaturan` kolom `PIN Pengawas` (`F2`, default awal `654321`) — akses pantau Monitoring, cetak Laporan, dan buka kunci layar siswa. Keduanya dapat diubah langsung dari menu **Pengaturan Sistem** maupun sel Spreadsheet.

---

## Panduan Pemasangan & Pindah ke Database Spreadsheet Baru (Tanpa Edit Kode)

Aplikasi dirancang agar pengguna non-koding dapat mengganti database Spreadsheet atau menyalin proyek tanpa perlu mengubah kode di `Code.gs`:

### Cara 1 — Ganti Spreadsheet Langsung dari Menu Pengaturan Sistem
Gunakan cara ini bila Anda sudah memiliki Web App terdeploy dan ingin memindahkan datanya ke file Google Spreadsheet baru:
1. Buat file Google Spreadsheet baru di Google Drive (pastikan akun pemilik Web App memiliki akses **Editor** ke Spreadsheet tersebut).
2. Salin **URL lengkap Spreadsheet** (misal `https://docs.google.com/spreadsheets/d/1abc.../edit`) atau **ID Spreadsheet**-nya saja.
3. Buka Dashboard Admin (`?page=pengaturan`), masuk dengan **PIN Admin**.
4. Pada panel **Koneksi Database & Link Web App**, tempel URL/ID tersebut ke kolom **Ganti Link / ID Spreadsheet Baru (Opsional)**.
5. Klik **Simpan Pengaturan** (sistem akan memvalidasi akses lalu otomatis membuat tabel database yang belum ada di Spreadsheet baru). Anda juga dapat menekan tombol **Siapkan Struktur Database** kapan saja untuk memastikan ke-5 tab sheet tersedia lengkap.

### Cara 2 — Duplikasi Spreadsheet (*Make a Copy* / Container-Bound)
Gunakan cara ini bila proyek Apps Script menempel di file Google Spreadsheet utama dan Anda ingin menggandakannya untuk sekolah/ruang lain:
1. Di Google Spreadsheet, pilih **File $\rightarrow$ Buat salinan (Make a copy)**.
2. Di Spreadsheet hasil salinan, buka **Ekstensi $\rightarrow$ Apps Script**.
3. Klik **Terapkan (Deploy) $\rightarrow$ Deployment baru**:
   - Jenis: **Aplikasi web**
   - Jalankan sebagai: **Saya (Me)**
   - Yang memiliki akses: **Siapa saja (Anyone)**
4. Selama kolom *Ganti Link / ID Spreadsheet Baru* di Pengaturan dibiarkan kosong, aplikasi otomatis mendeteksi **`SpreadsheetApp.getActiveSpreadsheet()`** sehingga langsung membaca Spreadsheet hasil duplikasi tanpa mengedit `Code.gs`.
5. Buka URL deployment baru dengan tambahan `?page=admin`, login menggunakan PIN awal **`123456`**, lalu masuk ke **Pengaturan Sistem** dan klik **Siapkan Struktur Database** bila diperlukan.

---

## Lokasi Penyimpanan Konfigurasi Koneksi (`Script Properties`) & Cara Reset

Beberapa konfigurasi tingkat sistem **tidak disimpan di dalam Google Sheet**, melainkan di memori internal Google Apps Script (**`PropertiesService.getScriptProperties()`**), antara lain:
- `CUSTOM_SPREADSHEET_ID`: Link/ID Spreadsheet pengganti jika diatur dari menu Pengaturan.
- `CUSTOM_DEPLOY_URL`: Link deploy cadangan (opsional).
- `ADMIN_PIN`, `MAX_VIOLATIONS`, `REQUIRE_FULLSCREEN`: Cache konfigurasi keamanan & PIN Admin.

**Mengapa `CUSTOM_SPREADSHEET_ID` tidak disimpan di Sheet?**
Karena aplikasi harus mengetahui ID Spreadsheet tujuan terlebih dahulu sebelum bisa membuka file Spreadsheet mana pun (`getSpreadsheet_()`).

**Urutan Prioritas Koneksi (`Code.gs`)**:
- **Database (`getSpreadsheet_()`)**:
  1. `CUSTOM_SPREADSHEET_ID` di *Script Properties* (jika diisi lewat menu Pengaturan).
  2. `SpreadsheetApp.getActiveSpreadsheet()` (otomatis aktif saat file Spreadsheet diduplikasi).
  3. Fallback `SPREADSHEET_ID` default di `Code.gs`.
- **URL Aplikasi (`getAppUrl()`)**:
  1. `CUSTOM_DEPLOY_URL` di *Script Properties* (jika diisi manual oleh Admin).
  2. `ScriptApp.getService().getUrl()` (otomatis mendeteksi URL deployment yang sedang berjalan).
  3. Fallback `FIXED_DEPLOY_URL` di `Code.gs`.

**Cara Reset Jika Salah Memasukkan ID Spreadsheet**:
Jika Anda salah menempel ID Spreadsheet baru (atau salah memasukkan ID yang tidak bisa diakses) hingga dashboard gagal memuat:
1. Buka Editor Google Apps Script $\rightarrow$ klik ikon gerigi **Setelan Proyek (Project Settings)** di sidebar kiri.
2. Gulir ke bagian **Properti Skrip (Script Properties)** $\rightarrow$ klik **Edit properti skrip**.
3. Hapus atau kosongkan nilai **`CUSTOM_SPREADSHEET_ID`** (dan **`CUSTOM_DEPLOY_URL`** bila perlu), lalu simpan.

---

## Catatan URL Google Workspace vs URL Publik Umum

- Saat menggunakan akun Google Workspace sekolah (misal `@smkn1klk.sch.id`), fungsi `ScriptApp.getService().getUrl()` terkadang mengembalikan URL berawalan domain seperti:
  `https://script.google.com/a/<domain-sekolah>/macros/s/.../exec`
- Jika peserta ujian membuka Web Publik menggunakan akun **`@gmail.com` pribadi**, link berawalan `/a/<domain>/` dapat memunculkan layar login Google Workspace sekolah.
- **Solusi**: Gunakan URL universal tanpa `/a/<domain>/` (`https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec`) saat membagikan ke siswa, dan tempel URL universal tersebut pada kolom **Link Deploy Web App Cadangan (Opsional)** di **Pengaturan Sistem** agar tombol *Buka Web Publik* & *Salin Link Web Publik* di Sidebar selalu mengarah ke URL universal tersebut.

---

## Struktur Lengkap 5 Sheet Database

Seluruh HTML antarmuka pengguna sengaja **tidak menampilkan** nama sheet dan posisi kolom teknis demi keamanan. Berikut referensi struktur tabel di Google Spreadsheet untuk keperluan Admin:

1. **`DB_Anticurang`** (Jadwal & Link Soal)
   - Kolom A: `No`
   - Kolom B: `Mapel`
   - Kolom C: `Kelas`
   - Kolom D: `Link GFrom`
   - Kolom E: `Tanggal` (format `YYYY-MM-DD`, kosongkan jika fleksibel)
   - Kolom F: `Waktu Mulai` (`HH:MM`, jam mulai soal boleh dibuka)
   - Kolom G: `Waktu Selesai` (`HH:MM`, batas akhir soal boleh dibuka)
   - Kolom H: `Durasi (Menit)` (cukup isi angka menit misal `30`, kosongkan jika durasi fleksibel)

2. **`DB_Mapel`** (Master Mata Pelajaran untuk dropdown form soal)
   - Kolom A: `No`
   - Kolom B: `Nama Mapel`

3. **`DB_Murid`** (Master Data Siswa / Murid)
   - Kolom A: `No`
   - Kolom B: `Nama Murid`
   - Kolom C: `NIS/NISN`
   - Kolom D: `Kelas`
   *(Catatan: Opsi pilihan kelas pada form Tambah/Edit Soal otomatis dihimpun dari data kelas unik yang ada di sheet ini).*

4. **`Pengaturan`** (Identitas Web, Kop, Drive & PIN)
   - Kolom A: `Nama Sekolah`
   - Kolom B: `Logo`
   - Kolom C: `Kop`
   - Kolom D: `Folder ID Drive` (opsional, tempat simpan unggahan Logo/Kop)
   - Kolom E: `Sumber Data` (teks keterangan sumber data pada cetakan Laporan PDF)
   - Kolom F: `PIN Pengawas` (`F2`, default `654321`)
   - Kolom G: `PIN Admin` (`G2`, default `123456`)
   - Kolom H: `Batas Pelanggaran` (`H2`, default `3` - batas maksimal pelanggaran sebelum layar soal terkunci, isi `0` untuk peringatan saja)

5. **`Log_Monitoring`** (Riwayat Sesi & Pelanggaran Anticurang)
   - Kolom A–M: `ID Sesi`, `Waktu Akses`, `Nama Siswa`, `NIS / Peserta`, `Kelas Siswa`, `No Soal`, `Mapel`, `Kelas Jadwal`, `Tanggal Ujian`, `Status`, `Jumlah Pelanggaran`, `Catatan Pelanggaran`, `Waktu Selesai`.

---

## Alur Penggunaan Singkat

- **Siswa**: Buka link utama (`/exec`) $\rightarrow$ pilih jadwal berstatus *Sedang Berlangsung* $\rightarrow$ pada jendela konfirmasi:
  * **Kelas** terisi otomatis & terkunci sesuai kelas jadwal ujian.
  * **Nama Siswa** dipilih melalui dropdown yang otomatis memfilter siswa khusus pada kelas tersebut.
  * **NIS / No. Peserta** otomatis terisi (`readOnly`) mengikuti murid yang dipilih.
  * Centang persetujuan $\rightarrow$ klik **Buka Soal Anticurang** dan kerjakan pada mode Layar Penuh hingga menekan **Kirim / Submit** pada Google Form lalu klik **Selesai Ujian**.
- **Pengawas Ruang**: Buka link `?page=pengawas` (atau `?page=admin`) $\rightarrow$ login memakai **PIN Pengawas** $\rightarrow$ pantau status pengerjaan & pelanggaran secara berkala, cetak/ekspor rekap, atau masukkan **PIN Pengawas** pada layar siswa jika peserta terkunci.
- **Administrator**: Login memakai **PIN Admin** $\rightarrow$ kelola **Data Mapel** dan **Data Murid** (tambah/edit/hapus atau impor Excel siswa per kelas), tambah/impor link soal pada **Kelola Link Soal**, pantau **Dashboard Monitoring**, cetak **Laporan PDF**, dan atur identitas/PIN/koneksi database di **Pengaturan Sistem**.
