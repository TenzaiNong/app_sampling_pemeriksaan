"""
Dashboard Utama - Sampling Audit Keuangan
Menggabungkan analisis Belanja dan Pendapatan
"""

import streamlit as st
import pandas as pd
from io import BytesIO

# --- IMPORT MODUL SENDIRI ---
import pendapatan_analyzer as pend_analyzer
from helpers import (generate_laporan_pendapatan_xlsx, generate_laporan_pendapatan_docx, 
                      generate_template_pendapatan)
from belanja_dashboard import dashboard_belanja


st.set_page_config(page_title="Dashboard Sampling Audit", layout="wide")
st.title("🕵️ Dashboard Uji Petik Pemeriksaan Keuangan")
st.markdown("---")

# === PILIHAN JENIS SAMPLING ===
jenis_sampling = st.selectbox(
    "📋 Pilih Jenis Sampling:",
    ["Belanja/Lainnya", "Pendapatan"],
    help="Pilih jenis sampling yang akan dianalisis"
)

st.markdown("---")

# === CONDITIONAL LOGIC BERDASARKAN JENIS SAMPLING ===
if jenis_sampling == "Pendapatan":
    # ===== DASHBOARD PENDAPATAN =====
    st.subheader("📊 Analisis Anomali Pendapatan Wajib Pajak")
    st.write("Unggah data pelaporan untuk mengidentifikasi WP dengan pelaporan identik atau variasi rendah (<=20%).")
    
    # Download Template
    st.markdown("---")
    st.subheader("📥 Download Template Kertas Kerja")
    st.write("Silakan download template di bawah sebagai format standard untuk upload data:")
    
    template_buff = generate_template_pendapatan()
    st.download_button(
        label="📄 Download Template Pendapatan",
        data=template_buff.getvalue(),
        file_name="Template_Pendapatan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.markdown("---")
    st.subheader("📤 Upload Data Pendapatan")
    
    # Upload File
    uploaded_file_pend = st.file_uploader(
        "Pilih file (Excel, CSV, atau Parquet)",
        type=['xlsx', 'csv', 'parquet'],
        key='pend_uploader'
    )
    
    if uploaded_file_pend:
        # Load data berdasarkan ekstensi
        try:
            if uploaded_file_pend.name.endswith('.csv'):
                df_pend = pd.read_csv(uploaded_file_pend)
            elif uploaded_file_pend.name.endswith('.parquet'):
                df_pend = pd.read_parquet(uploaded_file_pend)
            else:
                df_pend = pd.read_excel(uploaded_file_pend)
            
            st.success(f"✅ Data berhasil dimuat: {len(df_pend)} baris")
            
            # Preview Data
            with st.expander("👁️ Preview Data", expanded=False):
                st.dataframe(df_pend.head())
            
            # Tentukan kolom bulan yang tersedia
            bulan_list = ['JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI',
                          'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER']
            available_bulan = [b for b in bulan_list if b in df_pend.columns]
            
            if available_bulan:
                st.info(f"📅 Kolom bulan terdeteksi: {', '.join(available_bulan)}")
                
                # Jalankan analisis
                if st.button("🔍 Lakukan Analisis Anomali", key="btn_analisis_pend"):
                    with st.spinner("Menganalisis anomali pendapatan..."):
                        anomali_results = pend_analyzer.detect_anomali_pendapatan(df_pend, available_bulan)
                    
                    st.session_state['anomali_results'] = anomali_results
                    st.session_state['df_pendapatan'] = df_pend
                    st.session_state['bulan_cols_pend'] = available_bulan
                
                # Tampilkan hasil jika sudah ada
                if 'anomali_results' in st.session_state:
                    anomali_results = st.session_state['anomali_results']
                    df_pend = st.session_state['df_pendapatan']
                    available_bulan = st.session_state.get('bulan_cols_pend', available_bulan)
                    
                    st.markdown("---")
                    st.subheader("📈 Hasil Analisis")
                    
                    # Hitung statistik untuk display
                    stat_pend = pend_analyzer.hitung_statistik_pendapatan(df_pend, available_bulan)
                    
                    # Metrics
                    col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
                    with col_m1:
                        st.metric("Total WP", stat_pend['total_wp'])
                    with col_m2:
                        st.metric("WP dengan Anomali", len(anomali_results))
                    with col_m3:
                        pct_anomali = (len(anomali_results) / stat_pend['total_wp'] * 100) if stat_pend['total_wp'] > 0 else 0
                        st.metric("% Anomali", f"{pct_anomali:.2f}%")
                    with col_m4:
                        st.metric("Total Pendapatan", f"Rp {stat_pend['total_pendapatan']:,.0f}")
                    with col_m5:
                        anomalous_total = sum([item.get('total_realisasi', 0) for item in anomali_results]) if anomali_results else 0
                        anomalous_pct = (anomalous_total / stat_pend['total_pendapatan'] * 100) if stat_pend['total_pendapatan'] > 0 else 0
                        st.metric("Total Realisasi Anomali", f"Rp {anomalous_total:,.0f}", delta=f"{anomalous_pct:.2f}%")
                    
                    # Daftar Anomali
                    st.subheader("📋 Daftar WP dengan Anomali")
                    if anomali_results:
                        df_anomali_display = pd.DataFrame(anomali_results)[['nomor', 'nama_wp', 'npwpd', 'jenis_anomali', 'rata_rata', 'total_realisasi']]
                        df_anomali_display = df_anomali_display.rename(columns={
                            'nomor': 'No',
                            'nama_wp': 'Nama WP',
                            'npwpd': 'NPWPD',
                            'jenis_anomali': 'Jenis Anomali',
                            'rata_rata': 'Rata-rata Pendapatan',
                            'total_realisasi': 'Total Realisasi'
                        })
                        # Format angka
                        if 'Total Realisasi' in df_anomali_display.columns:
                            df_anomali_display['Total Realisasi'] = df_anomali_display['Total Realisasi'].apply(lambda x: f"Rp {x:,.0f}" if pd.notna(x) else "Rp 0")
                        if 'Rata-rata Pendapatan' in df_anomali_display.columns:
                            df_anomali_display['Rata-rata Pendapatan'] = df_anomali_display['Rata-rata Pendapatan'].apply(lambda x: f"Rp {x:,.2f}" if pd.notna(x) else "-")
                        st.dataframe(df_anomali_display, use_container_width=True)
                        
                        # Download Laporan
                        st.markdown("---")
                        st.subheader("📥 Download Laporan")
                        
                        col_dl1, col_dl2 = st.columns(2)
                        
                        with col_dl1:
                            laporan_xlsx_pend = generate_laporan_pendapatan_xlsx(
                                df_pend, anomali_results, available_bulan
                            )
                            st.download_button(
                                label="📊 Download Laporan (.xlsx)",
                                data=laporan_xlsx_pend.getvalue(),
                                file_name=f"Laporan_Anomali_Pendapatan_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                        
                        with col_dl2:
                            laporan_docx_pend = generate_laporan_pendapatan_docx(
                                df_pend, anomali_results, available_bulan
                            )
                            st.download_button(
                                label="📄 Download Laporan (.docx)",
                                data=laporan_docx_pend.getvalue(),
                                file_name=f"Laporan_Anomali_Pendapatan_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                    else:
                        st.success("✅ Tidak ditemukan anomali berdasarkan kriteria analisis.")
            else:
                st.error("❌ Kolom bulan (JANUARI - DESEMBER) tidak ditemukan dalam file. Silakan gunakan template yang tersedia.")
        
        except Exception as e:
            st.error(f"❌ Gagal membaca file: {e}")

else:
    # ===== DASHBOARD BELANJA (ORIGINAL) =====
    dashboard_belanja()
