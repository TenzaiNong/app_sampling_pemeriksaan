"""
Modul untuk analisis anomali data pendapatan WP.
Mendeteksi pelaporan identik dan variasi rendah (<=5%) antar bulan.
"""

import pandas as pd
import numpy as np


def detect_anomali_pendapatan(df, bulan_cols):
    """
    Deteksi anomali pada data pendapatan WP.
    
    Kriteria:
    1. Pelaporan Identik: Nilai sama setiap bulan
    2. Variasi Rendah: Perubahan antar bulan <= 10%
    
    Args:
        df: DataFrame dengan data WP
        bulan_cols: List kolom bulan yang akan dianalisis
        
    Returns:
        List berisi dict dengan informasi anomali setiap WP
    """
    results = []

    for index, row in df.iterrows():
        # Ambil nilai bulan yang ada (anggap NaN sebagai 0 untuk total)
        vals = []
        for col in bulan_cols:
            val = row.get(col, 0)
            if pd.isna(val):
                continue
            # Keep zeros for total calculation but skip zeros for variation checks
            vals.append(float(val))

        # Skip jika data kurang dari 2 bulan non-NaN
        non_na_vals = [v for v in vals if v != None]
        if len(non_na_vals) < 2:
            continue

        adalah_anomali = False
        jenis_anomali = []

        # Hitung total realisasi (jumlah seluruh bulan, NaN=0)
        total_realisasi = 0.0
        for col in bulan_cols:
            v = row.get(col, 0)
            try:
                if pd.isna(v):
                    v = 0
                total_realisasi += float(v)
            except Exception:
                # jika bukan numeric, anggap 0
                continue

        # Kriteria 1: Pelaporan Identik (semua non-NaN sama)
        nonzero_vals_for_check = [float(row[col]) for col in bulan_cols if pd.notna(row.get(col, None)) and row.get(col, None) != 0]
        if len(nonzero_vals_for_check) >= 2 and len(set(nonzero_vals_for_check)) == 1:
            adalah_anomali = True
            jenis_anomali.append("Pelaporan Identik (0.00%)")

        # Kriteria 2: Variasi Rendah (<= 5%) - tampilkan persentase variasi aktual
        if len(nonzero_vals_for_check) >= 2 and not (len(set(nonzero_vals_for_check)) == 1):
            variasi_rendah_pct = None
            for i in range(len(nonzero_vals_for_check) - 1):
                prev = nonzero_vals_for_check[i]
                curr = nonzero_vals_for_check[i+1]
                if prev != 0:
                    perubahan = abs(curr - prev) / prev
                    if perubahan <= 0.05:
                        # ambil pertama kali ditemukan variasi rendah sebagai representative
                        variasi_rendah_pct = perubahan
                        break

            if variasi_rendah_pct is not None:
                adalah_anomali = True
                pct_display = variasi_rendah_pct * 100
                jenis_anomali.append(f"Variasi Rendah: {pct_display:.2f}%")

        if adalah_anomali:
            results.append({
                'nomor': row.get('NOMOR', index + 1),
                'nama_wp': row.get('NAMA WP', ''),
                'npwpd': row.get('NPWPD', ''),
                'jenis_anomali': ' | '.join(jenis_anomali),
                'bulan_terisi': len([c for c in bulan_cols if pd.notna(row.get(c, None)) and row.get(c, None) != 0]),
                'rata_rata': np.mean(nonzero_vals_for_check) if nonzero_vals_for_check else 0,
                'min': np.min(nonzero_vals_for_check) if nonzero_vals_for_check else 0,
                'max': np.max(nonzero_vals_for_check) if nonzero_vals_for_check else 0,
                'std_dev': np.std(nonzero_vals_for_check) if len(nonzero_vals_for_check) > 1 else 0,
                'total_realisasi': total_realisasi,
            })

    return results


def hitung_statistik_pendapatan(df, bulan_cols):
    """
    Hitung statistik pendapatan keseluruhan.
    
    Args:
        df: DataFrame dengan data WP
        bulan_cols: List kolom bulan
        
    Returns:
        Dict dengan statistik keseluruhan
    """
    # Gabungkan semua nilai bulan
    all_values = []
    for col in bulan_cols:
        vals = df[col].dropna()
        all_values.extend(vals[vals != 0].tolist())
    
    if not all_values:
        return {
            'total_wp': len(df),
            'rata_rata': 0,
            'median': 0,
            'min': 0,
            'max': 0,
            'std_dev': 0,
            'total_pendapatan': 0
        }
    
    total_pendapatan = sum(all_values)
    
    return {
        'total_wp': len(df),
        'rata_rata': np.mean(all_values),
        'median': np.median(all_values),
        'min': np.min(all_values),
        'max': np.max(all_values),
        'std_dev': np.std(all_values),
        'total_pendapatan': total_pendapatan
    }
