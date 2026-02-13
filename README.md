# Dashboard Sampling Audit (Pilih Sampel Pemeriksaan)

Aplikasi ini adalah alat bantu berbasis web yang dikembangkan menggunakan Python dan Streamlit untuk membantu auditor dalam menentukan jumlah sampel dan memilih sampel pemeriksaan secara statistik. Aplikasi ini dirancang untuk mempermudah proses uji petik (sampling) dalam pemeriksaan keuangan.

## Cara Kerja Aplikasi

Secara umum, alur kerja aplikasi ini terdiri dari empat langkah utama:

1.  **Upload Data**: Pengguna mengunggah data populasi yang akan diperiksa dalam format Excel (`.xlsx`) atau CSV (`.csv`).
2.  **Pemetaan Kolom**: Aplikasi membaca data dan meminta pengguna untuk memilih kolom yang akan dijadikan referensi:
    *   **Kolom ID**: Identitas unik per baris (misal: Nomor Bukti, No. Invoice).
    *   **Kolom Nilai**: Nilai numerik/moneter yang akan diuji (misal: Nilai Transaksi, Harga Perolehan).
3.  **Perhitungan Jumlah Sampel**:
    *   Pengguna memilih metodologi statistik yang diinginkan (MUS, Unstratified MPU, Stratified MPU, atau Difference Estimation).
    *   Pengguna memasukkan parameter audit seperti *Confidence Level*, *Salah Saji Tertoleransi (SST)*, dan estimasi lainnya.
    *   Aplikasi secara otomatis menghitung jumlah sampel ($n$) yang direkomendasikan berdasarkan rumus standar audit.
4.  **Pemilihan Sampel (Eksekusi)**:
    *   Berdasarkan jumlah sampel ($n$) yang dihitung, pengguna memilih teknik pengambilan sampel (contoh: Acak Sederhana, Sistematis, PPS, atau Analisis Hukum Benford).
    *   Aplikasi menghasilkan tabel daftar sampel terpilih yang dapat diunduh untuk keperluan Kertas Kerja Pemeriksaan (KKP).

## Fitur dan Kelebihan

*   **Otomatisasi Perhitungan Rumus**: Tidak perlu menghitung manual menggunakan kalkulator atau rumus Excel yang rumit. Aplikasi mendukung metode standar seperti:
    *   *Monetary Unit Sampling (MUS)*
    *   *Mean Per Unit (MPU)* (Stratified & Unstratified)
*   **Fleksibilitas Teknik Sampling**: Mendukung berbagai teknik pengambilan sampel:
    *   *Simple Random Sampling*
    *   *Systematic Sampling*
    *   *Probability Proportional to Size (PPS)*
    *   *Stratified Top Value* (Prioritas nilai terbesar)
*   **Analisis Anomali (Benford's Law)**: Fitur khusus untuk mendeteksi anomali data berdasarkan digit awal angka, berguna untuk mendeteksi potensi *fraud* atau ketidakwajaran data.
*   **Antarmuka User-Friendly**: Menggunakan antarmuka visual (GUI) web yang mudah digunakan tanpa perlu koding.
*   **Transparansi**: Parameter perhitungan (seperti *Reliability Factor, Expansion Factor*) terlihat jelas dan sesuai kaidah juknis pemeriksaan.

## Keterbatasan

*   **Kualitas Data Input**: Akurasi hasil sangat bergantung pada kebersihan dan kelengkapan data yang diupload. Data dengan format nilai yang tidak standar atau mengandung *missing values* mungkin perlu dibersihkan terlebih dahulu.
*   **Asumsi Statistik**: Beberapa metode (seperti MPU) mengasumsikan distribusi data tertentu (misal: normal). Pengguna perlu memahami karakteristik datanya agar metode yang dipilih tepat.
*   **Input Manual**: Untuk metode Stratifikasi, penentuan batas strata (*threshold*) masih memerlukan input pertimbangan professional (*professional judgment*) dari auditor, tidak sepenuhnya otomatis ditentukan oleh mesin.
*   **Penyederhanaan**: Metode *Difference/Ratio Estimation* yang diimplementasikan menggunakan pendekatan varians yang disederhanakan.

## Instalasi

### Prasyarat
- Python 3.8 atau lebih baru
- pip (Python package installer)

### Langkah Instalasi
1. Clone atau unduh repository ini
2. Buka terminal dan navigasi ke folder proyek
3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Cara Menjalankan

Jalankan perintah berikut di terminal:

```bash
streamlit run app.py
```

Aplikasi akan terbuka di browser default Anda di alamat `http://localhost:8501`

## Struktur File

```
├── app.py                    # File utama aplikasi Streamlit
├── calculations.py           # Modul perhitungan statistik & sampling
├── selections.py             # Modul teknik pengambilan sampel
├── requirements.txt          # Daftar dependencies
└── README.md                 # Dokumentasi ini
```

## Dependencies

Aplikasi ini memerlukan dependencies berikut:
- **streamlit** (>=1.52.0) - Framework web application
- **pandas** (>=2.3.0) - Data manipulation dan analisis
- **numpy** (>=2.4.0) - Numerical computing
- **chardet** (>=5.2.0) - Deteksi encoding file
- **openpyxl** (>=3.1.0) - Generate laporan Excel
- **python-docx** (>=1.2.0) - Generate laporan Word

Untuk detail lengkap, lihat file `requirements.txt`

## Modul Utama

### app.py
File utama yang menjalankan aplikasi Streamlit. Fitur utama:
- Upload dan preprocessing data (CSV/Excel)
- Deteksi encoding dan delimiter otomatis
- Konversi format Rupiah ke nilai numerik
- Interface untuk perhitungan sampel dan pemilihan
- Export hasil dalam format Excel dan Word

**Fungsi Kunci:**
- `detect_encoding()` - Deteksi encoding file
- `detect_csv_delimiter()` - Deteksi delimiter CSV
- `convert_rupiah_to_numeric()` - Konversi rupiah ke angka
- `generate_laporan_xlsx()` - Buat laporan Excel
- `generate_laporan_docx()` - Buat laporan Word

### calculations.py
Modul perhitungan metodologi statistik sampling.

**Metode yang Didukung:**
1. **Monetary Unit Sampling (MUS)**
   - Formula: n = (NB × RF) / (SST - (DSS × EF))
   - Referensi: Juknis Hal 47

2. **Unstratified Mean Per Unit (MPU)**
   - Formula: n = ((UR × SD × N) / A)²
   - Referensi: Juknis Hal 45

3. **Stratified Mean Per Unit (MPU)**
   - Formula: n = (Σ(Ni × Si))² / ((SST/Ur)² + Σ(Ni × Si²))
   - Dengan FPC (Finite Population Correction)

**Komponen Perhitungan:**
- Reliability Factor (RF) - Berdasarkan confidence level (90%, 95%, 99%)
- Expansion Factor (EF) - Berdasarkan jumlah error yang diharapkan
- Upper Result Limit (UR) - Koefisien untuk MPU

### selections.py
Modul teknik pengambilan sampel.

**Teknik Sampling yang Tersedia:**
1. **Simple Random Sampling** - Acak sederhana tanpa bobot
2. **Probability Proportional to Size (PPS)** - Berdasarkan nilai rupiah (wajib untuk MUS)
3. **Systematic Sampling** - Interval sampling yang konsisten
4. **Random Systematic Sampling** - Interval dengan lompatan acak
5. **Stratified Top Value** - Prioritas 20% nilai terbesar + 80% random
6. **Benford's Law Analysis** - Deteksi anomali berdasarkan digit pertama

## Contoh Workflow

1. **Upload Data**
   - File CSV atau Excel dengan minimal 2 kolom (ID dan Nilai)
   - Sistem otomatis mendeteksi encoding dan delimiter

2. **Pilih Kolom**
   - Kolom ID: Nomor Bukti, Invoice No, etc.
   - Kolom Nilai: Nilai Transaksi, Amount, etc.

3. **Hitung Sampel**
   - Pilih metodologi (MUS / MPU Stratified / MPU Unstratified)
   - Input parameter: Confidence Level, SST, estimasi lainnya
   - Sistem menghitung jumlah sampel optimal (n)

4. **Ambil Sampel**
   - Pilih teknik sampling
   - Sistem memberikan daftar item yang akan diperiksa
   - Download laporan dalam format Excel atau Word

## Parameter Audit

### Confidence Level
- 90% (RF: 2.40, UR: 1.65)
- 95% (RF: 3.00, UR: 1.96) - Default
- 99% (RF: 3.70, UR: 2.58)

### SST (Salah Saji Tertoleransi)
Nilai rupiah maksimal yang dapat ditoleransi dalam populasi

### DSS (Desirable Sampling Sensitivity) - Untuk MUS
Nilai rupiah yang diharapkan untuk ditemukan kesalahan

### Expected Standard Deviation - Untuk MPU
Estimasi standar deviasi populasi nilai

## Output yang Dihasilkan

### Laporan Excel (.xlsx)
- **Ringkasan**: Metadata perhitungan dan parameter
- **Detail Sampel**: Daftar lengkap item yang dipilih untuk diperiksa
- **Statistik**: Perbandingan statistik populasi vs sampel

### Laporan Word (.docx)
- Ringkasan eksekutif
- Metadata dan parameter perhitungan
- Statistik deskriptif
- Detail sampel (maks 300 baris)

## Catatan Penting

1. **Data Quality**: Akurasi hasil bergantung pada kualitas data input. Pastikan data sudah bersih dan konsisten.

2. **Format Nilai**: Aplikasi dapat mendeteksi format Rupiah (1.250.000,00) dan konversi otomatis ke numerik.

3. **Missing Values**: Item dengan nilai kosong (NaN) akan ditangani sesuai dengan metode yang dipilih.

4. **Professional Judgment**: Pemilihan metodologi dan parameter harus tetap melibatkan pertimbangan profesional auditor.

## License

Lihat file LICENSE untuk informasi lisensi

## Kontribusi

Untuk laporan bug atau saran fitur, silakan buat issue atau hubungi tim pengembang.
