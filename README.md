# Sistem Portal Ujian Google Form Mode Anticurang & Dashboard Monitoring Real-Time

Aplikasi ujian daring (*online assessment*) serverless berbasis **Google Apps Script**, **Google Sheets Database**, dan **Google Drive Storage** yang responsif untuk **Desktop, Tablet & Mobile (`100dvh`)**. Mengintegrasikan Google Form ke dalam lingkungan proteksi ujian (*anti-cheat lockdown*), pemantauan kehadiran dan pelanggaran peserta secara real-time, serta pemisahan hak akses antara **Web Publik (Siswa)** dan **Dashboard Pengawas / Administrator**.

---

## Daftar Fitur Utama

1. **Multi-Page Serverless Web App (MPA) Architecture**:
   - Routing dinamis melalui parameter URL query (`?page=portal|monitoring|laporan|kelola-soal|data-mapel|data-murid|pengaturan`).
   - Navigasi antar halaman menggunakan `window.top.location.href` untuk menjamin transisi mulus dan bebas dari masalah *blank screen* pada iframe Google Apps Script.
2. **Server-Side Template Inclusion (Scriptlet Pattern)**:
   - Modularitas kode UI terbagi rapi ke dalam komponen `StyleCSS.html`, `Sidebar.html`, `Header.html`, dan `AppJS.html` yang di-inject di server-side via `HtmlService.createTemplateFromFile()`.
3. **Web Publik Jadwal Ujian & Searchable Filter Kelas (`Portal.html`)**:
   - Tampilan bersih khusus siswa tanpa sidebar maupun menu login admin.
   - Filter tab status otomatis aktif pada tab **Sedang Berlangsung** saat halaman pertama kali dibuka.
   - Filter Kelas dilengkapi **Searchable Combobox** yang dapat diketik secara bebas (*type to search*) maupun dipilih dari menu dropdown untuk mempercepat pencarian jadwal di sekolah besar.
   - Modal konfirmasi ujian otomatis mengunci pilihan Kelas sesuai jadwal soal, serta menyediakan pencarian instan nama siswa yang otomatis mengisi NIS/Nomor Peserta secara aman (`readOnly`).
4. **Anti-Cheat Lockdown & Violation Engine**:
   - **Kunci Layar Penuh (*Fullscreen Enforcement*)**: Memaksa siswa tetap di layar penuh saat mengerjakan soal.
   - **Deteksi Berpindah Tab/Aplikasi**: Mengawasi peristiwa `blur` dan `visibilitychange` jendela peramban.
   - **Proteksi Akses**: Memblokir klik kanan (*context menu*), tombol F12 / pintasan Developer Tools, serta fungsi salin-tempel (*copy-paste*).
   - **Watermark Dinamis Anti-Screenshot**: Menampilkan identitas peserta, mapel, dan label perlindungan berlatar transparan di atas lembar soal.
   - **Audio Alarm Peringatan**: Membunyikan nada sirine peringatan saat terjadi pelanggaran.
   - **Overlay Kunci Pengawas**: Jika batas toleransi pelanggaran tercapai, layar terkunci permanen dan hanya dapat dibuka oleh Pengawas Ruang menggunakan PIN Pengawas 6 digit.
   - **Pencatatan Real-Time**: Setiap pelanggaran dan waktu mulai/selesai dicatat langsung ke sheet `Log_Monitoring`.
5. **Graceful Time-Up Protection (Penanganan Waktu Habis & Submit Google Form)**:
   - Dilengkapi *countdown timer* interaktif pada topbar ujian.
   - **Pengingat Sisa Waktu 2 Menit & 1 Menit**: Ketika sisa waktu mencapai 2 menit dan 1 menit, sistem memunculkan popup pengingat ramah disertai nada chime lembut dan indikator timer berkedip. Notifikasi ini dilindungi secara khusus sehingga kemunculan maupun klik tombol murid **dikecualikan 100% dari hitungan pelanggaran**.
   - Ketika waktu ujian habis, sistem memunculkan dialog peringatan interaktif dan memberikan tombol pintas untuk langsung fokus ke lembar Google Form agar siswa dapat segera menekan tombol **Kirim / Submit** sebelum sesi ditutup otomatis.
6. **Pemisahan Jam Buka vs Durasi Soal (Opsi B)**:
   - `Waktu Mulai` & `Waktu Selesai` mengatur jendela waktu kapan siswa diperbolehkan membuka soal.
   - Kolom **H (`Durasi (Menit)`)** menentukan batas waktu pengerjaan perorangan dalam menit (misal `30`) yang mulai dihitung mundur saat siswa mengklik *Buka Soal*, tetap berjalan penuh meskipun melewati `Waktu Selesai` (kosongkan jika durasi fleksibel).
7. **Master Data Mapel & Data Murid** (`?page=data-mapel` & `?page=data-murid`):
   - Manajemen CRUD data mata pelajaran (`DB_Mapel`) dan data siswa (`DB_Murid`).
   - Form Tambah/Edit Soal dilengkapi penomoran otomatis yang menghitung kelanjutan data di database, serta input Kelas dan Mapel yang dapat dipilih maupun diketik (searchable datalist).
   - Impor dan ekspor data siswa maupun jadwal soal via format file Excel (`.xlsx` / `.xls`).
8. **Dashboard Monitoring Real-Time & Ikon Tombol Aksi Modern** (`Monitoring.html` & `KelolaSoal.html`):
   - Kartu metrik live: Total Sesi, Sedang Aktif, Terkunci, dan Selesai.
   - Pemantauan tabel status siswa lengkap dengan jumlah dan catatan pelanggaran secara berkala.
   - Seluruh tombol aksi tabel (**Edit** dan **Hapus**) telah diperbarui menggunakan **SVG Icon modern** yang responsif, intuitif, dan hemat ruang.
9. **Laporan Pelaksanaan & Cetak PDF Dinamis** (`Laporan.html`):
   - Filter cetak berdasarkan tanggal, mata pelajaran, dan kelas.
   - Input nama pengawas dan NIP dinamis: Saat diisi, teks nama dan NIP tampil rapi tanpa garis bawah potong; saat dikosongkan, garis titik-titik tanda tangan otomatis muncul.
   - Smart Kop Surat: Mendukung mode gambar logo instansi maupun susunan teks kop surat otomatis dengan garis pembatas dinamis.
10. **Akselerasi Cache (`CacheService`) & Portabilitas Database Dinamis**:
    - Cache memori server 60 detik untuk jadwal dan daftar murid, memangkas latensi pembacaan data hingga **< 30 milidetik** saat ratusan siswa mengakses secara serentak.
    - Database Spreadsheet dapat diganti kapan saja dari menu **Pengaturan Sistem** tanpa perlu mengubah kode sumber `Code.gs`. Mendukung pula mekanisme *Make a copy* (Container-Bound) secara otomatis.
11. **Multi-Tier Access Control**:
    - **Publik / Siswa**: Mengakses Portal Jadwal Ujian dan mengerjakan soal.
    - **Pengawas Ruang**: Akses pemantauan Monitoring, cetak Laporan PDF, dan input PIN buka kunci layar peserta ujian.
    - **Administrator**: Akses penuh ke seluruh fitur termasuk Kelola Soal, Data Mapel, Data Murid, Hapus Log Riwayat, dan Pengaturan Sistem.

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

## Kapasitas Pengguna & Performa Skala Massal (Concurrent Users)

Secara arsitektur sistem, aplikasi ujian anticurang ini dirancang untuk menangani **100 hingga 500+ siswa sekaligus (serentak)** dengan sangat lancar.

### 1. Faktor Arsitektur Pendukung

- **Pengerjaan Soal Ditangani Langsung oleh Google Form**: Lembar soal dan proses pengiriman jawaban (*submit*) diproses langsung oleh infrastruktur **Google Cloud / Google Form**. Google Form sendiri sanggup menerima ribuan jawaban secara bersamaan tanpa membebani kuota eksekusi Google Apps Script.
- **Optimasi Cache Berkecepatan Tinggi (*CacheService*)**: Jadwal ujian dan data murid telah di-cache di memori server selama 60 detik. Ketika ratusan siswa membuka halaman portal bersamaan, data disajikan dalam **< 30 milidetik** tanpa membebani kuota baca Google Spreadsheet.
- **Beban Aplikasi Web Sangat Ringan**: Aplikasi web hanya memproses pencatatan saat siswa pertama kali menekan tombol *Buka Soal*, mendeteksi pelanggaran (jika berpindah tab), dan mencatat waktu selesai ujian.

### 2. Batasan Teknis & Kuota Resmi Google (Google Apps Script Quotas)

| Kategori | Akun Gmail Gratis (@gmail.com) | Akun Google Workspace / Belajar.id |
| :--- | :--- | :--- |
| **Eksekusi Simultan (Tepat di Detik yang Sama)** | Maksimal ~30 proses bersamaan | Maksimal ~30 proses bersamaan |
| **Batas Kapasitas Baris Spreadsheet** | Hingga 10.000.000 sel data | Hingga 10.000.000 sel data |
| **Kapasitas Responden Google Form** | Ratusan ribu responden | Ratusan ribu responden |
| **Rekomendasi Jumlah Siswa per Sesi** | **100 – 300 siswa per gelombang** | **300 – 500+ siswa per gelombang** |

> [!NOTE]
> Batas **30 proses simultan** Google Apps Script mengukur aksi yang terjadi **pada milidetik yang sama persis**. Selama siswa mengklik tombol masuk ujian secara wajar (tersebar dalam rentang 1–5 menit), ratusan siswa tidak akan mengalami kendala antrean (*queue*).

### 3. Tips & Rekomendasi Pelaksanaan Ujian Massal

1. **Beri Jeda Waktu Buka Jadwal**: Atur jam buka portal 5–10 menit sebelum ujian dimulai agar siswa masuk secara bertahap dan tidak menekan tombol *Buka Soal* serempak di detik yang sama.
2. **Pengaturan Google Form**: Di Google Form, aktifkan **"Batasi ke 1 tanggapan"** jika ingin memastikan setiap siswa hanya bisa mengirim jawaban satu kali.
3. **Pemeliharaan Data Berkala**: Setelah selesai satu periode ujian (misal setelah selesai PTS/PAS), Anda dapat mengosongkan riwayat di menu *Dashboard Monitoring* (atau ekspor ke Excel terlebih dahulu) agar Google Spreadsheet tetap ringan dan responsif.

---

## Alur Penggunaan Singkat

- **Siswa**: Buka link utama (`/exec`) $\rightarrow$ pilih jadwal berstatus *Sedang Berlangsung* $\rightarrow$ pada jendela konfirmasi:
  - **Kelas** terisi otomatis & terkunci sesuai kelas jadwal ujian.
  - **Nama Siswa** dipilih melalui dropdown yang otomatis memfilter siswa khusus pada kelas tersebut.
  - **NIS / No. Peserta** otomatis terisi (`readOnly`) mengikuti murid yang dipilih.
  - Centang persetujuan $\rightarrow$ klik **Buka Soal Anticurang** dan kerjakan pada mode Layar Penuh hingga menekan **Kirim / Submit** pada Google Form lalu klik **Selesai Ujian**.
- **Pengawas Ruang**: Buka link `?page=pengawas` (atau `?page=admin`) $\rightarrow$ login memakai **PIN Pengawas** $\rightarrow$ pantau status pengerjaan & pelanggaran secara berkala, cetak/ekspor rekap, atau masukkan **PIN Pengawas** pada layar siswa jika peserta terkunci.
- **Administrator**: Login memakai **PIN Admin** $\rightarrow$ kelola **Data Mapel** dan **Data Murid** (tambah/edit/hapus atau impor Excel siswa per kelas), tambah/impor link soal pada **Kelola Link Soal**, pantau **Dashboard Monitoring**, cetak **Laporan PDF**, dan atur identitas/PIN/koneksi database di **Pengaturan Sistem**.
