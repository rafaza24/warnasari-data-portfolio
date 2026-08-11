# 📱 Warnasari Data - Aplikasi Pendataan Dusun Digital

![Flutter](https://img.shields.io/badge/Flutter-02569B?style=for-the-badge&logo=flutter&logoColor=white)
![Dart](https://img.shields.io/badge/Dart-0175C2?style=for-the-badge&logo=dart&logoColor=white)
![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Android%20%7C%20Web-green?style=for-the-badge&logo=android&logoColor=white)
![Status](https://img.shields.io/badge/Status-Completed-brightgreen?style=for-the-badge)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)

---

## 📌 Tentang Proyek

**Warnasari Data** adalah platform digital komprehensif yang dikembangkan dalam rangka **Program KKN (Kuliah Kerja Nyata) KKN 06 Desa Warnasari Ikopin University** untuk mendukung digitalisasi administrasi kependudukan di tingkat dusun/desa.

Aplikasi ini dibangun untuk menggantikan proses pencatatan manual warga yang memakan waktu dan rawan kesalahan, dengan solusi digital yang terintegrasi, efisien, dan mudah digunakan oleh perangkat desa.

> **Konteks:** Digunakan aktif oleh petugas RT/RW di Dusun 1 Warnasari untuk pengelolaan data kependudukan secara digital dan terpusat oleh tim **KKN 06 Desa Warnasari Ikopin University**.

---

## 📸 Tangkapan Layar Aplikasi (Showcase)

| Halaman Login | Dashboard Utama | Scan KK (AI Extraction) |
|:---:|:---:|:---:|
| ![Login](screenshot_aplikasi_android/raw/01_halaman_login.png) | ![Dashboard](screenshot_aplikasi_android/raw/02_dashboard_statistik.png) | ![Scan KK AI](screenshot_aplikasi_android/raw/04_scan_kk_ai.jpeg) |

| Rekapitulasi Laporan | Grafik Demografi | Hak Akses & Pengguna |
|:---:|:---:|:---:|
| ![Rekapitulasi](screenshot_aplikasi_android/raw/05_rekapitulasi_laporan.png) | ![Visualisasi Grafik](screenshot_aplikasi_android/raw/06_visualisasi_grafik.jpeg) | ![Hak Akses](screenshot_aplikasi_android/raw/07_manajemen_pengguna.png) |

---

## ✨ Fitur Utama Aplikasi

- **Dashboard Statistik & Demografi:** Menampilkan ringkasan populasi, rasio jenis kelamin, dan statistik demografi secara realtime.
- **Smart Search & Filter RW/RT:** Pencarian data warga cepat berbasis NIK, nama, serta filter tingkatan RT/RW.
- **Manajemen Mutasi Warga:** Pelaporan warga datang, pindah, meninggal dunia, serta pendaftaran bayi baru lahir.
- **Smart Scan KK (AI Extraction):** Fitur pemindaian dan ekstraksi otomatis data Kartu Keluarga untuk mempercepat proses input data.
- **Cetak Laporan & Dokumen:** Fitur eksport cetak laporan bulanan, buku induk warga, dan surat keterangan secara langsung dari aplikasi.
- **Manajemen Pengguna & Hak Akses:** Pengaturan akun pengguna dengan sistem role berdasarkan tingkatan (Admin, Petugas RT, Petugas RW).

---

## 🛠️ Arsitektur & Teknologi

| Komponen | Teknologi | Keterangan |
| :--- | :--- | :--- |
| **Mobile App (Frontend)** | Flutter & Dart | Antarmuka aplikasi mobile Android/Web yang responsif dan fleksibel |
| **Web Portal & Dashboard** | HTML5, CSS3, Chart.js (`Index.html`) | Dashboard antarmuka web interaktif yang di-render oleh Google Apps Script |
| **Backend API & Database** | Google Apps Script (`Code.gs`) | REST API Serverless yang terhubung dengan Google Sheets sebagai database |
| **Konfigurasi GAS** | `appsscript.json` | Manifest hak akses dan konfigurasi lingkungan Google Apps Script |

---

## 🚀 Cara Menjalankan Proyek

### 1. Jalankan Aplikasi Mobile (Flutter)

```bash
# Masuk ke direktori Flutter
cd flutter_app

# Install dependency
flutter pub get

# Jalankan aplikasi di Emulator / Device Android
flutter run
```

### 2. Deploy Backend Google Apps Script

1. Buka [Google Apps Script](https://script.google.com) dan buat project baru.
2. Salin isi file `Code.gs` ke editor GAS.
3. Salin isi file `Index.html` ke file HTML baru di editor GAS.
4. Salin isi `appsscript.json` ke bagian **Manifest** project.
5. Klik **Deploy → New Deployment → Web App**.
6. Atur **Execute as: Me** dan **Who has access: Anyone**.
7. Salin URL deployment yang dihasilkan dan masukkan ke variabel `_initialUrl` di `flutter_app/lib/main.dart`.

---

## 👤 Developer

<table>
  <tr>
    <td align="center">
      <b>RafazaTech</b><br>
      <i>Full Stack Developer — Flutter, Google Apps Script & Web Automation</i><br><br>
      <a href="https://github.com/rafaza24">
        <img src="https://img.shields.io/badge/GitHub-rafaza24-181717?style=for-the-badge&logo=github&logoColor=white" alt="GitHub"/>
      </a>
    </td>
  </tr>
</table>

---

## 📝 Lisensi & Hak Cipta

© 2024 **RafazaTech** • **KKN 06 Desa Warnasari Ikopin University**.  
Dikembangkan untuk mendukung digitalisasi dan administrasi pendataan warga dusun secara transparan dan efisien.
