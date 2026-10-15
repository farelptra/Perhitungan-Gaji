# 💼 Kalkulator Gaji dan Lembur Karyawan (PP No. 35 Tahun 2021)

Aplikasi web interaktif berbasis **Streamlit** untuk menghitung upah lembur karyawan secara otomatis, akurat, dan transparan sesuai ketentuan ketenagakerjaan resmi Republik Indonesia (**Peraturan Pemerintah No. 35 Tahun 2021**, khusus **Sistem 5 Hari Kerja / 40 Jam per Minggu**). Dilengkapi fitur ekspor laporan ke format **Excel (.xlsx)** standar akuntansi dan siap dideploy secara gratis ke **Streamlit Community Cloud**.

---

## 📌 Fitur Utama

1. **Parameter Gaji Fleksibel**:
   - Input Gaji Pokok / UMK bulanan (Default: `Rp 5.247.870`).
   - Perhitungan otomatis **Upah Per Jam** dengan rumus baku Depnaker:  
     $$\text{Upah Per Jam} = \frac{\text{Gaji Pokok}}{173}$$
2. **Data Editor Interaktif (`st.data_editor`)**:
   - Mendukung penambahan, pengubahan, dan penghapusan baris jadwal kehadiran secara dinamis.
   - Kolom: `Tanggal`, `Hari` (otomatis diselaraskan ke Bahasa Indonesia), `Keterangan` (Hari Kerja / Hari Libur / Weekend), `Jam Masuk` (HH:MM), `Jam Keluar` (HH:MM), dan `Catatan`.
3. **Logika Perhitungan Sesuai PP 35/2021 (Sistem 5 Hari Kerja)**:
   - **Hari Kerja (Senin - Jumat)**: Jam kerja standar `08:00 - 17:00`. Waktu kerja sebelum 08:00 dan sesudah 17:00 dihitung sebagai lembur.
     - *Jam ke-1*: Pengali **1.5x**
     - *Jam ke-2 dan seterusnya*: Pengali **2.0x**
   - **Hari Libur (Sabtu, Minggu, Hari Libur Resmi)**:
     - Total jam kerja aktual **langsung dipotong 1 jam untuk istirahat**.
     - *8 jam pertama*: Pengali **2.0x**
     - *Jam ke-9*: Pengali **3.0x**
     - *Jam ke-10 dan seterusnya*: Pengali **4.0x**
4. **Dashboard KPI & Summary Metrics**:
   - `⏱️ Total Jam Lembur`
   - `💵 Total Uang Lembur`
   - `💰 Gaji Pokok (UMK)`
   - `💳 Take Home Pay (Kotor)`
5. **Ekspor Excel (.xlsx) Berkualitas Tinggi**:
   - Dibangun dengan modul `io.BytesIO` dan mesin `openpyxl`.
   - Tata letak resmi berkas penggajian: Kop judul, metadata karyawan, tabel bergaris dengan header *Royal Navy*, formula otomatis Excel (`=SUM(...)`), dan format mata uang `Rp #,##0`.
6. **Penanganan Error Tangguh (Try-Except)**:
   - Validasi format jam (`HH:MM`), rentang jam 00:00 - 23:59, dan shift lintas tengah malam (*cross-midnight*).
   - Pesan kesalahan yang informatif dengan nomor baris spesifik.

---

## 📁 Struktur Direktori

```text
Hitung Gaji/
├── .streamlit/
│   └── config.toml          # Konfigurasi tema dan visual Streamlit
├── app.py                   # Seluruh kode aplikasi dan logika perhitungan
├── requirements.txt         # Dependensi Python untuk Streamlit Cloud
└── README.md                # Dokumentasi dan panduan deployment
```

---

## 🚀 Panduan Menjalankan Secara Lokal

### 1. Prasyarat
Pastikan Python 3.9+ telah terpasang di komputer Anda.

### 2. Instal Dependensi
Buka terminal / PowerShell di direktori proyek dan jalankan:
```bash
pip install -r requirements.txt
```

### 3. Jalankan Aplikasi
```bash
streamlit run app.py
```
Aplikasi akan otomatis terbuka di browser pada alamat `http://localhost:8501`.

---

## 🌐 Panduan Deploy Gratis ke Streamlit Cloud

Aplikasi ini 100% kompatibel dan siap dideploy tanpa biaya melalui **Streamlit Community Cloud**:

1. **Upload ke GitHub**:
   - Buat repositori baru di akun GitHub Anda (misal: `kalkulator-lembur-pp35`).
   - Unggah berkas `app.py`, `requirements.txt`, `.streamlit/config.toml`, dan `README.md`.
2. **Kunjungi Streamlit Community Cloud**:
   - Buka [share.streamlit.io](https://share.streamlit.io/) dan login menggunakan akun GitHub Anda.
3. **Deploy App**:
   - Klik tombol **"New app"**.
   - Pilih repositori yang baru saja dibuat.
   - Branch: `main` (atau `master`).
   - Main file path: `app.py`.
   - Klik **"Deploy!"**.
4. Dalam 1-2 menit, aplikasi web Anda aktif dan dapat dibagikan kepada HRD, karyawan, atau tim operasional melalui URL publik.

---

## ⚖️ Referensi Hukum
- **Peraturan Pemerintah Republik Indonesia Nomor 35 Tahun 2021** tentang Perjanjian Kerja Waktu Tertentu, Alih Daya, Waktu Kerja dan Waktu Istirahat, dan Pemutusan Hubungan Kerja (Pasal 31 & Pasal 32).
