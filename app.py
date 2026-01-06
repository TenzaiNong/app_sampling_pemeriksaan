import math
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

# --- IMPORT MODUL SENDIRI ---
import calculations as calc
import selections as sel

st.set_page_config(page_title="Dashboard Sampling Audit", layout="wide")
st.title("🕵️ Dashboard Uji Petik Pemeriksaan Keuangan")
st.markdown("---")

# 1. UPLOAD
st.sidebar.header("1. Upload Data")
uploaded_file = st.sidebar.file_uploader("Upload Tabel Data (Excel/CSV)",
                                         type=['xlsx', 'csv'])

if uploaded_file is not None:
    # Load Data
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.success(f"Data dimuat: {len(df)} baris.")

    # 2. PEMETAAN
    col1, col2 = st.columns(2)
    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    all_cols = df.columns.tolist()

    with col1:
        id_col = st.selectbox("Kolom ID Sampel", all_cols)
    with col2:
        value_col = st.selectbox("Kolom Nilai Rupiah", numeric_cols)

    total_nilai_buku = df[value_col].sum()
    st.info(f"💰 Total Nilai Buku: Rp {total_nilai_buku:,.2f}")

    # 3. METODE HITUNG JUMLAH SAMPEL
    st.markdown("---")
    st.subheader("Metode Penentuan Jumlah Sampel")

    metode_sampling = st.selectbox("Pilih Metode", [
        "Monetary Unit Sampling (MUS)", "Unstratified Mean Per Unit (MPU)",
        "Stratified Mean Per Unit (MPU)", "Difference/Ratio Estimation"
    ])

    # Parameter Input UI
    col_in1, col_in2, col_in3 = st.columns(3)
    with col_in1:
        confidence = st.selectbox("Confidence Level", [90, 95, 99], index=1)
    with col_in2:
        sst = st.number_input("Salah Saji Tertoleransi (SST)",
                              value=float(total_nilai_buku) * 0.05)

    n_res = 0
    error_msg = None

    # --- PANGGIL FUNGSI DARI MODUL CALCULATIONS ---
    if metode_sampling == "Monetary Unit Sampling (MUS)":
        with col_in3:
            dss = st.number_input("Dugaan Salah Saji (DSS)", value=0.0)
        n_res, error_msg = calc.calculate_mus(total_nilai_buku, confidence,
                                              sst, dss)

    elif metode_sampling == "Unstratified Mean Per Unit (MPU)":
        with col_in3:
            sd = st.number_input("Estimasi Standar Deviasi", value=100000.0)
        n_res, error_msg = calc.calculate_mpu_unstratified(
            len(df), confidence, sst, sd)

    elif metode_sampling == "Stratified Mean Per Unit (MPU)":
        st.info(
            "ℹ️ Tentukan batas nilai (threshold) untuk membagi populasi menjadi beberapa kelompok (Strata)."
        )

        # 1. Input Batas Strata
        col_strata1, col_strata2 = st.columns([2, 1])
        with col_strata1:
            # User memasukkan batas angka dipisah koma
            thresholds_input = st.text_input(
                "Masukkan Batas Nilai Rupiah (pisahkan dengan koma)",
                value="100000000, 500000000",
                help=
                "Contoh: 100000000, 500000000 artinya akan ada 3 strata: <100Jt, 100Jt-500Jt, dan >500Jt"
            )

        # 2. Proses Pembagian Strata (Binning)
        try:
            # Bersihkan input dan urutkan
            thresholds = sorted([
                float(x.strip()) for x in thresholds_input.split(',')
                if x.strip()
            ])

            # Tambahkan -Infinity dan +Infinity untuk menangkap semua data
            bins = [-float('inf')] + thresholds + [float('inf')]

            # Buat Label Strata
            labels = []
            for i in range(len(bins) - 1):
                lower = bins[i]
                upper = bins[i + 1]
                if lower == -float('inf'):
                    labels.append(f"< {upper:,.0f}")
                elif upper == float('inf'):
                    labels.append(f"> {lower:,.0f}")
                else:
                    labels.append(f"{lower:,.0f} - {upper:,.0f}")

            # Lakukan Segmentasi Data
            df['Strata'] = pd.cut(df[value_col], bins=bins, labels=labels)

            # Hitung Statistik per Strata (Count & StdDev)
            strata_stats = df.groupby('Strata')[value_col].agg(
                ['count', 'std', 'mean']).reset_index()

            # Tampilkan Tabel Preview Strata
            st.write("Current Population Distribution per Strata:")

            # Siapkan data untuk perhitungan (User bisa edit SD jika mau)
            strata_summary = []

            # Gunakan Data Editor agar user bisa sesuaikan Estimasi SD per strata
            # Default SD diisi dengan SD aktual populasi (jika ada) atau 0
            strata_stats['std'] = strata_stats['std'].fillna(0)

            edited_strata = st.data_editor(
                strata_stats,
                column_config={
                    "Strata":
                    "Kelompok Nilai",
                    "count":
                    "Jumlah Populasi (N)",
                    "mean":
                    "Rata-rata Nilai",
                    "std":
                    st.column_config.NumberColumn(
                        "Estimasi SD (Bisa Diedit)",
                        help="Standar Deviasi untuk perhitungan MPU",
                        required=True,
                        format="%.2f")
                },
                disabled=["Strata", "count", "mean"],  # Kunci kolom ini
                hide_index=True,
                use_container_width=True)

            # Konversi hasil editor ke format list untuk dikirim ke calculations.py
            for index, row in edited_strata.iterrows():
                strata_summary.append({
                    'strata': row['Strata'],
                    'count': row['count'],
                    'std_dev': row['std']
                })

            # Hitung N Total
            n_res, error_msg = calc.calculate_mpu_stratified(
                strata_summary, confidence, sst)

            # Tampilkan Alokasi Sampel (Saran Neyman Allocation Sederhana)
            if n_res > 0:
                st.markdown("#### 📊 Alokasi Sampel per Strata (Disesuaikan)")
                st.caption(
                    "Menggunakan Neyman Allocation dengan pembatasan maksimal populasi."
                )

                total_weight = sum(
                    [s['count'] * s['std_dev'] for s in strata_summary])

                final_allocations = []
                allocation_text = ""

                # 1. INISIALISASI DICTIONARY PENAMPUNG
                allocation_dict = {
                }  # <--- TAMBAHAN PENTING (Wadah untuk menyimpan jatah)

                for s in strata_summary:
                    weight = s['count'] * s['std_dev']

                    if total_weight > 0:
                        raw_allocation = (weight / total_weight) * n_res
                        n_teoritis = math.ceil(raw_allocation)
                    else:
                        n_teoritis = 0

                    # Logika Clamping (Sensus)
                    if n_teoritis >= s['count']:
                        n_final_strata = s['count']
                        keterangan = "✅ **Sensus** (Ambil Semua)"
                    else:
                        n_final_strata = n_teoritis
                        keterangan = ""

                    final_allocations.append(n_final_strata)

                    # 2. SIMPAN JATAH STRATA KE DICTIONARY
                    allocation_dict[s['strata']] = int(
                        n_final_strata)  # <--- TAMBAHAN PENTING

                    # Tampilkan output
                    allocation_text += f"- **{s['strata']}**: {n_final_strata} sampel (Populasi: {s['count']}) {keterangan}\n"

                st.markdown(allocation_text)

                # Hitung ulang total
                n_adjusted_total = sum(final_allocations)

                st.info(
                    f"💡 Total Sampel yang disarankan setelah penyesuaian populasi: **{n_adjusted_total}** item."
                )

                # Update n_res agar masuk ke kotak input bawah
                n_res = n_adjusted_total

                # 3. SIMPAN KE SESSION STATE AGAR BISA DIBACA TOMBOL GENERATE
                st.session_state[
                    'allocation_dict'] = allocation_dict  # <--- TAMBAHAN PENTING
                st.session_state[
                    'df_stratified'] = df  # <--- TAMBAHAN PENTING (Menyimpan DF yang punya kolom 'Strata')

        except ValueError:
            st.error(
                "Format batas nilai salah. Pastikan hanya angka dan koma.")
            n_res = 0

    elif metode_sampling == "Difference/Ratio Estimation":
        with col_in3:
            var = st.number_input("Estimasi Varians", value=1000000.0)
        n_res, error_msg = calc.calculate_difference_ratio(
            len(df), confidence, sst, var)

    # Tampilkan Hasil N
    if error_msg:
        st.error(error_msg)
        n_final = 0
    else:
        st.metric("Jumlah Sampel Disarankan (n)", f"{n_res} Item")
        n_final = st.number_input("Jumlah Sampel Final",
                                  min_value=1,
                                  max_value=len(df),
                                  value=int(n_res) if n_res > 0 else 30)

    # 4. TEKNIK EKSEKUSI
    st.markdown("---")
    st.subheader("Teknik Pemilihan Sampel")

    teknik = st.selectbox("Teknik Pemilihan", [
        "Acak Sederhana", "PPS (Wajib untuk MUS)", "Sistematis",
        "Sistematis Acak", "Stratifikasi (Top Value)",
        "Benford's Law (Anomali)"
    ])

    if st.button("🚀 Generate Sampel"):
        sampled_df = pd.DataFrame()

        # --- LOGIKA BARU: CEK APAKAH PAKAI STRATIFIED MPU ---
        if metode_sampling == "Stratified Mean Per Unit (MPU)" and 'allocation_dict' in st.session_state:
            st.info("Menggunakan pemilihan terdistribusi sesuai Strata...")

            # Ambil DF yang sudah ada label Strata-nya
            df_to_use = st.session_state.get('df_stratified', df)
            alloc_data = st.session_state['allocation_dict']

            # Panggil fungsi baru di selections.py
            sampled_df = sel.select_stratified_distributed(
                df_to_use,
                alloc_data,
                teknik,  # Teknik pilihan (misal: Sistematis)
                value_col)

        else:
            # --- LOGIKA LAMA (UNTUK METODE NON-STRATIFIED) ---
            if teknik == "Acak Sederhana":
                sampled_df = sel.select_simple_random(df, n_final)
            elif teknik == "PPS (Wajib untuk MUS)":
                sampled_df = sel.select_pps(df, n_final, value_col)
            elif teknik == "Sistematis":
                sampled_df = sel.select_systematic(df, n_final)
            elif teknik == "Sistematis Acak":
                sampled_df = sel.select_random_systematic(df, n_final)
            elif teknik == "Stratifikasi (Top Value)":
                sampled_df = sel.select_stratified_top_value(
                    df, n_final, value_col)
            elif teknik == "Benford's Law (Anomali)":
                sampled_df = sel.select_benford_anomaly(df, n_final, value_col)

        # Output
        if not sampled_df.empty:
            st.success(f"Terpilih {len(sampled_df)} sampel.")

            # Tampilkan rincian per strata (opsional, untuk verifikasi)
            if 'Strata' in sampled_df.columns:
                st.write("Rincian Sampel per Strata:")
                st.write(sampled_df['Strata'].value_counts())

            st.dataframe(sampled_df)

        # Download Excel
        buff = BytesIO()
        with pd.ExcelWriter(buff, engine='openpyxl') as writer:
            sampled_df.to_excel(writer, index=False)
        st.download_button("Download Excel", buff.getvalue(), "sampel.xlsx")

    else:
        st.warning("Tidak ada sampel yang terpilih. Cek parameter.")

else:
    st.info("Silakan upload data di sidebar.")
