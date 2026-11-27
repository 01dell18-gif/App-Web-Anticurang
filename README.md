# Sistem Portal Ujian Google Form Anticurang & Monitoring (Arsitektur SIGURU v3.0)

Aplikasi serverless berbasis **Google Apps Script**, **Google Sheets Database**, dan **Google Drive Storage** yang mengadopsi arsitektur dari `Teknik pembuatan.md`:

1. **Server-Side Template Inclusion (Scriptlet Pattern)**: `StyleCSS.html`, `Sidebar.html`, `Header.html`, `AppJS.html`.
2. **Multi-Page Serverless Web App (MPA)** via Query Parameter Routing (`?page=portal|monitoring|laporan|kelola-soal|data-master|pengaturan`) dengan navigasi induk `window.top.location.href` anti blank putih.
3. **Mobile-First Off-Canvas Drawer & Closed-by-Default Accordion Navigation** + Icon-Only Navbar Actions (`Refresh`, `Clear Cache`, `Login/Logout`) + Real-Time Hardware Clock & Time-Based Greeting.
4. **Two-Tier Access Control (Publik / Peserta vs Admin)** dengan proteksi PIN & session persistence.
5. **Web Publik Jadwal per Hari & Mode Anticurang Multi-User**: Input identitas siswa, Layar Penuh (Fullscreen), deteksi pindah tab/jendela, blokir klik kanan/copy-paste/DevTools, timer mundur, serta pencatatan langsung ke `Log_Monitoring`.
6. **SheetJS Excel Processing**: Ekspor `.xls` (BIFF8) di Monitoring & Laporan, serta Impor & Ekspor Excel jadwal link soal di Kelola Soal.
7. **Dynamic Adaptive Print & PDF Engine** (`Laporan.html`): Smart Kop Divider (otomatis mendeteksi gambar vs teks kop) dan Dynamic Source Data Injection.
8. **Dual-Mode Logo & Kop Surat Ingestion** (`Pengaturan.html`): Mode URL Drive/Gambar vs Mode Unggah Berkas ke Google Drive.

## URL & Database (2 Link Terpisah)

- **Spreadsheet**: `1xZTTYRaCQQqjSI40XX9lFpz7lRIDxoAkpgE5TW72rZA` (`DB_Anticurang` kolom A–H termasuk `Durasi (Menit)`, `DB_Mapel` kolom `No, Nama Mapel`, `DB_Kelas` kolom `No, Kelas`, `Pengaturan`, `Log_Monitoring`)
- **Master Data Dropdown (`DB_Mapel` & `DB_Kelas`)**: Dikelola melalui menu admin **Data Mapel & Kelas** (`?page=data-master`) dan otomatis mengisi dropdown `<select>` pada form **Tambah/Edit Link Google Form**, sementara kolom **Kelas** peserta di Web Publik terisi otomatis sesuai kelas soal di `DB_Anticurang`.
- **Jendela Buka vs Durasi Soal (`DB_Anticurang`)**: Kolom `Waktu Mulai` & `Waktu Selesai` mengatur kapan soal dapat dibuka dan batas terakhir dibukanya; Kolom **H** (`Durasi (Menit)`, cukup isi angka menit misal `30`) menentukan lama pengerjaan penuh siswa setelah membuka soal (tetap penuh meski melewati `Waktu Selesai`).
- **Project Apps Script**: `10om5OlL4AEtNkp5O5OHu9-5v8z38eyhhxWIytGeAmakklYii7-MYFT-g`
- **Link Web Publik (Khusus Siswa)**: `https://script.google.com/macros/s/AKfycbxw20eF7UAvQv4zJJlMds32mUjU42KInSNiLgGKuZz9_sF76sOH4xhQEcvHkg9V-rO9/exec`
- **Link Dashboard Admin / Pengawas**: `https://script.google.com/macros/s/AKfycbxw20eF7UAvQv4zJJlMds32mUjU42KInSNiLgGKuZz9_sF76sOH4xhQEcvHkg9V-rO9/exec?page=admin` (atau `?page=pengawas`)
- **PIN Login Admin**: Tersimpan di Database Sheet `Pengaturan` kolom `PIN Admin` (`G2`, default awal `123456`) — akses penuh Admin (Monitoring, Laporan, Kelola Link Soal, Pengaturan)
- **PIN Login & Buka Kunci Pengawas**: Tersimpan di Database Sheet `Pengaturan` kolom `PIN Pengawas` (`F2`, default awal `654321`) — dapat dilihat/diubah di *Pengaturan Sistem* atau langsung di sel Spreadsheet; bila login pakai PIN Pengawas hanya bisa memantau Monitoring & Cetak Laporan, serta membuka kunci siswa yang terkunci
