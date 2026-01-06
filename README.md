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

## Cara Menjalankan

Pastikan Python dan *library* yang dibutuhkan sudah terinstal. Jalankan perintah berikut di terminal:

```bash
streamlit run app.py
```
