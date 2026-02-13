"""
Dashboard Belanja/Lainnya - Sampling dengan berbagai metode
"""

import streamlit as st
import pandas as pd
import numpy as np
import math
from io import BytesIO

import calculations as calc
import selections as sel
from helpers import (detect_csv_delimiter, convert_rupiah_to_numeric, 
                      generate_laporan_xlsx, generate_laporan_docx)


def dashboard_belanja():
    """
    Main dashboard untuk sampling Belanja/Lainnya
    """
    # === PANDUAN METODE UJI PETIK ===
    with st.expander("📚 Panduan Metode Uji Petik", expanded=False):
        st.markdown("""
        ### Panduan Penentuan Metode Uji Petik
        
        Dalam menentukan metode uji petik, pemeriksa pada umumnya mempertimbangkan dua aspek utama, yaitu **nilai transaksi** dan **jumlah transaksi**. 
        
        - **Nilai transaksi** merupakan nilai uang yang diakui dan dicatat dalam laporan keuangan sebagai akibat dari suatu transaksi
        - **Jumlah transaksi** adalah banyaknya kejadian transaksi yang dicatat dalam suatu periode tertentu
        
        Berdasarkan kedua pertimbangan tersebut, pemeriksa dapat menentukan metode uji petik statistik yang sesuai dengan karakteristik populasi yang diperiksa:
        
        ---
        
        #### 1️⃣ **Monetary Unit Sampling (MUS)**
        
        Metode ini dapat digunakan apabila pemeriksa melakukan uji petik yang **berfokus pada nilai transaksi**. 
        
        Dalam metode ini, transaksi dengan nilai yang lebih besar memiliki peluang lebih tinggi untuk terpilih sebagai sampel.  
        
        ---
        
        #### 2️⃣ **Unstratified Mean Per Unit (MPU)**
        
        Metode ini dapat digunakan apabila pemeriksa melakukan uji petik yang **berfokus pada jumlah transaksi**, tanpa mempertimbangkan perbedaan nilai transaksi. 
        
        Metode ini sesuai diterapkan pada populasi dengan karakteristik nilai transaksi yang **relatif homogen** dan tidak memiliki rentang nilai yang terlalu besar.
        
        ---
        
        #### 3️⃣ **Stratified Mean Per Unit (MPU)**
        
        Metode ini dapat digunakan apabila pemeriksa **mempertimbangkan jumlah transaksi sekaligus nilai transaksi** dalam penentuan sampel. 
        
        Metode ini sesuai untuk populasi dengan karakteristik nilai transaksi yang **heterogen** serta memiliki **rentang nilai yang besar**, sehingga diperlukan pengelompokan (stratifikasi) untuk meningkatkan efektivitas dan representativitas sampel.
        
        """)

    # 1. UPLOAD
    st.sidebar.header("1. Upload Data")
    uploaded_file = st.sidebar.file_uploader("Upload Tabel Data (Excel/CSV/Parquet)",
                                             type=['xlsx', 'csv', 'parquet'])

    if uploaded_file is not None:
        # Load Data
        if uploaded_file.name.endswith('.csv'):
            try:
                # Deteksi delimiter dan encoding
                delimiter, encoding = detect_csv_delimiter(uploaded_file)
                df = pd.read_csv(uploaded_file, sep=delimiter, encoding=encoding)
                
                # 🔥 KONVERSI FORMAT RUPIAH KE NUMERIK
                df = convert_rupiah_to_numeric(df)
                
                st.info(f"📌 Delimiter: **'{delimiter}'** | Encoding: **{encoding}**")
            except Exception as e:
                st.error(f"❌ Gagal membaca file CSV: {e}")
                st.stop()
        
        elif uploaded_file.name.endswith('.parquet'):
            try:
                df = pd.read_parquet(uploaded_file)
                st.info("✅ File Parquet berhasil dibaca")
            except Exception as e:
                st.error(f"❌ Gagal membaca file Parquet: {e}")
                st.stop()
        
        else:  # Excel (.xlsx)
            try:
                df = pd.read_excel(uploaded_file)
                st.info("✅ File Excel berhasil dibaca")
            except Exception as e:
                st.error(f"❌ Gagal membaca file Excel: {e}")
                st.stop()

        st.success(f"Data dimuat: {len(df)} baris.")
        
        # Konversi kolom dengan format Rupiah ke numerik
        try:
            df = convert_rupiah_to_numeric(df)
        except Exception:
            print("⚠️ Gagal menjalankan convert_rupiah_to_numeric; melewatkan konversi.")
            pass
        
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
            "Stratified Mean Per Unit (MPU)"
        ])

        # Parameter Input UI
        col_in1, col_in2, col_in3, col_in4 = st.columns(4)
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
            with col_in4:
                expansion = st.selectbox("Expansion Factor", [1, 5, 10, 15, 20, 25, 30, 37], index=1)
            n_res, error_msg = calc.calculate_mus(total_nilai_buku, confidence, sst, dss, expansion)

        elif metode_sampling == "Unstratified Mean Per Unit (MPU)":
            with col_in3:
                current_std_dev = df[value_col].std()
                
                st.info(f"💡 SD Populasi saat ini: {current_std_dev:,.2f}")
                
                sd = st.number_input(
                    "Estimasi Standar Deviasi (SD)", 
                    value=float(current_std_dev),
                    min_value=0.0,
                    format="%.2f",
                    help="Default diisi dengan SD Populasi saat ini (sesuai saran Piloting Sample). Bisa diubah jika menggunakan data tahun lalu."
                )

            n_res, error_msg = calc.calculate_mpu_unstratified(
                len(df), confidence, sst, sd
            )

        elif metode_sampling == "Stratified Mean Per Unit (MPU)":
            st.info("ℹ️ Stratifikasi otomatis menggunakan metode Kuantil (Membagi populasi sama rata).")
            
            n_bins = st.slider("Bagi populasi menjadi berapa bagian? (Default: 4 - Strata)", min_value=3, max_value=10, value=4)
            
            try:
                df['Strata'], bin_edges = pd.qcut(df[value_col], q=n_bins, retbins=True, duplicates='drop', precision=0)
                
                cat_codes = df['Strata'].cat.categories
                new_labels = []
                
                for cat in cat_codes:
                    lbl = f"{cat.left:,.0f} s.d {cat.right:,.0f}"
                    new_labels.append(lbl)
                
                df['Strata'] = df['Strata'].cat.rename_categories(new_labels)
                
                strata_stats = df.groupby('Strata', observed=True)[value_col].agg(['count', 'std', 'mean']).reset_index()
                
                st.write("📊 Distribusi Populasi per Strata:")
                
                strata_stats['std'] = strata_stats['std'].fillna(0)
                
                edited_strata = st.data_editor(
                    strata_stats,
                    column_config={
                        "Strata": "Range Nilai (Rupiah)",
                        "count": "Populasi (N)",
                        "mean": "Rata-rata",
                        "std": st.column_config.NumberColumn(
                            "Estimasi SD",
                            help="Standar Deviasi (Bisa diedit)",
                            required=True,
                            format="%.2f"
                        )
                    },
                    disabled=["Strata", "count", "mean"],
                    hide_index=True,
                    use_container_width=True
                )
                
                strata_summary = []
                
                for index, row in edited_strata.iterrows():
                    strata_summary.append({
                        'strata': row['Strata'],
                        'count': row['count'],
                        'std_dev': row['std']
                    })
                
                n_res, error_msg = calc.calculate_mpu_stratified(strata_summary, confidence, sst)
                
                if n_res > 0:
                    st.markdown("#### 📍 Alokasi Sampel")
                    
                    total_weight = sum([s['count'] * s['std_dev'] for s in strata_summary])
                    allocation_dict_default = {} 
                    allocation_list_default = []
                    
                    for s in strata_summary:
                        weight = s['count'] * s['std_dev']
                        
                        if total_weight > 0:
                            raw_allocation = (weight / total_weight) * n_res
                            n_teoritis = math.ceil(raw_allocation)
                        else:
                            n_teoritis = 0
                        
                        if n_teoritis >= s['count']:
                            n_final_strata = s['count']
                        else:
                            n_final_strata = n_teoritis
                        
                        allocation_list_default.append(n_final_strata)
                        allocation_dict_default[s['strata']] = int(n_final_strata)
                    
                    if 'allocation_adjustments' not in st.session_state:
                        st.session_state['allocation_adjustments'] = allocation_list_default
                    
                    st.write("🎯 **Sesuaikan alokasi sampel per Strata (opsional)**:")
                    
                    allocation_editor_data = []
                    for idx, s in enumerate(strata_summary):
                        allocation_editor_data.append({
                            "Strata": s['strata'],
                            "Populasi (N)": s['count'],
                            "Alokasi Otomatis": allocation_dict_default[s['strata']],
                            "Alokasi Manual": st.session_state['allocation_adjustments'][idx] 
                                if idx < len(st.session_state['allocation_adjustments']) 
                                else allocation_dict_default[s['strata']]
                        })
                    
                    edited_allocation_raw = st.data_editor(
                        allocation_editor_data,
                        column_config={
                            "Strata": "Range Nilai (Rupiah)",
                            "Populasi (N)": st.column_config.NumberColumn(
                                "Populasi",
                                disabled=True
                            ),
                            "Alokasi Otomatis": st.column_config.NumberColumn(
                                "Alokasi Otomatis",
                                disabled=True,
                                help="Berdasarkan rumus proporsi"
                            ),
                            "Alokasi Manual": st.column_config.NumberColumn(
                                "Alokasi Manual",
                                help="Sesuaikan jumlah sampel untuk strata ini",
                                required=True,
                                min_value=0
                            )
                        },
                        disabled=["Strata", "Populasi (N)", "Alokasi Otomatis"],
                        hide_index=True,
                        use_container_width=True
                    )
                    
                    edited_allocation = pd.DataFrame(edited_allocation_raw)
                    
                    allocation_dict = {}
                    final_allocations = []
                    
                    for idx, row in edited_allocation.iterrows():
                        strata_name = row['Strata']
                        n_adjusted = int(row['Alokasi Manual'])
                        pop_strata = row['Populasi (N)']
                        
                        if n_adjusted > pop_strata:
                            n_adjusted = pop_strata
                            st.warning(f"⚠️ Alokasi untuk '{strata_name}' disesuaikan menjadi {pop_strata} (tidak boleh melebihi populasi)")
                        
                        allocation_dict[strata_name] = n_adjusted
                        final_allocations.append(n_adjusted)
                    
                    n_adjusted_total = sum(final_allocations)
                    st.info(f"💡 Total Sampel (Sesuai Alokasi Manual): **{n_adjusted_total}** item.")
                    n_res = n_adjusted_total
                    
                    st.session_state['allocation_dict'] = allocation_dict
                    st.session_state['df_stratified'] = df
                    st.session_state['allocation_adjustments'] = final_allocations

            except Exception as e:
                st.error(f"Gagal membagi kuartil otomatis. Data mungkin terlalu sedikit atau seragam. Error: {e}")
                n_res = 0

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

        generate_btn = st.button("🚀 Generate Sampel")

        if generate_btn:
            sampled_df = pd.DataFrame()

            if metode_sampling == "Stratified Mean Per Unit (MPU)" and 'allocation_dict' in st.session_state:
                st.info("Menggunakan pemilihan terdistribusi sesuai Strata...")
                
                df_to_use = st.session_state.get('df_stratified', df)
                alloc_data = st.session_state['allocation_dict']

                sampled_df = sel.select_stratified_distributed(
                    df_to_use,
                    alloc_data,
                    teknik, 
                    value_col)

            else:
                if teknik == "Acak Sederhana":
                    sampled_df = sel.select_simple_random(df, n_final)
                elif teknik == "PPS (Wajib untuk MUS)":
                    sampled_df = sel.select_pps(df, n_final, value_col)
                elif teknik == "Sistematis":
                    sampled_df = sel.select_systematic(df, n_final)
                elif teknik == "Sistematis Acak":
                    sampled_df = sel.select_random_systematic(df, n_final)
                elif teknik == "Stratifikasi (Top Value)":
                    sampled_df = sel.select_stratified_top_value(df, n_final, value_col)
                elif teknik == "Benford's Law (Anomali)":
                    sampled_df = sel.select_benford_anomaly(df, n_final, value_col)
            
            st.session_state['sampled_df'] = sampled_df
            if 'report_docx_bytes' in st.session_state:
                del st.session_state['report_docx_bytes']

        if 'sampled_df' in st.session_state and not st.session_state['sampled_df'].empty:
            
            current_sampled_df = st.session_state['sampled_df']
            
            st.success(f"Terpilih {len(current_sampled_df)} sampel.")

            display_df = current_sampled_df.copy()
            try:
                if value_col in display_df.columns:
                    total_sampled_value = display_df[value_col].sum()
                    total_nilai_buku = df[value_col].sum()
                    pct_nilai_sampel = (total_sampled_value / total_nilai_buku * 100) if total_nilai_buku > 0 else 0
                    col_metric1, col_metric2 = st.columns(2)
                    with col_metric1:
                        st.metric("Total Nilai Sampel", f"Rp {total_sampled_value:,.2f}")
                    with col_metric2:
                        st.metric("Persentase Nilai Sampel", f"{pct_nilai_sampel:.2f}%")
                else:
                    st.info("Kolom Nilai Rupiah tidak ditemukan dalam data sampel; total tidak ditampilkan.")
            except Exception as e:
                st.warning(f"Gagal menghitung total nilai sampel: {e}")

            if 'Strata' in display_df.columns:
                st.write("Rincian Sampel per Strata:")
                st.write(display_df['Strata'].value_counts())

            st.dataframe(display_df)
            
            col_btn1, col_btn2, col_btn3 = st.columns(3)
            
            with col_btn1:
                buff_sampel = BytesIO()
                with pd.ExcelWriter(buff_sampel, engine='openpyxl') as writer:
                    current_sampled_df.to_excel(writer, index=False, sheet_name='Sampel')
                buff_sampel.seek(0)
                st.download_button(
                    label="📊 Download Sampel (.xlsx)",
                    data=buff_sampel.getvalue(),
                    file_name="sampel.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col_btn2:
                buff_laporan = generate_laporan_xlsx(
                    df,
                    current_sampled_df,
                    metode_sampling,
                    teknik,
                    confidence,
                    sst,
                    value_col,
                    n_final
                )
                st.download_button(
                    label="📋 Generate Laporan Lengkap (.xlsx)",
                    data=buff_laporan.getvalue(),
                    file_name=f"Laporan_Sampling_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with col_btn3:
                def on_generate_docx_inline():
                    try:
                        sdf = st.session_state['sampled_df']
                        buf = generate_laporan_docx(
                            df_original=df,
                            sampled_df=sdf,
                            metode_sampling=metode_sampling,
                            teknik=teknik,
                            confidence=confidence,
                            sst=sst,
                            value_col=value_col,
                            n_final=n_final
                        )
                        st.session_state['report_docx_bytes'] = buf.getvalue()
                    except Exception as exc:
                        st.error(f"❌ Gagal membuat .docx: {exc}")

                st.button("📄 Generate Laporan (.docx)", on_click=on_generate_docx_inline, key="btn_docx_inline")
                
                if 'report_docx_bytes' in st.session_state and st.session_state['report_docx_bytes']:
                    st.success("✅ Laporan .docx siap!")
                    st.download_button(
                        label="⬇️ Download File (.docx)",
                        data=st.session_state['report_docx_bytes'],
                        file_name=f"Laporan_Sampling_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="dl_docx_inline"
                    )

        elif 'sampled_df' in st.session_state and st.session_state['sampled_df'].empty:
            st.warning("Tidak ada sampel yang terpilih. Cek parameter.")

    else:
        st.info("Silakan upload data di sidebar.")
