"""
Modul untuk analisis anomali data pendapatan WP.
Mendeteksi pelaporan identik dan variasi rendah (<=20%) antar bulan.
"""

import pandas as pd
import numpy as np


def detect_anomali_pendapatan(df, bulan_cols):
    """
    Deteksi anomali pada data pendapatan WP.
    
    Kriteria:
    1. Pelaporan Identik: Nilai sama setiap bulan
    2. Variasi Rendah: Perubahan antar bulan <= 20%
    
    Args:
        df: DataFrame dengan data WP
        bulan_cols: List kolom bulan yang akan dianalisis
        
    Returns:
        List berisi dict dengan informasi anomali setiap WP
    """
    results = []
    
    for index, row in df.iterrows():
        # Ambil nilai bulan yang ada
        vals = []
        for col in bulan_cols:
            val = row[col]
            # Skip nilai yang NaN atau 0
            if pd.notna(val) and val != 0:
                vals.append(float(val))
        
        # Skip jika data kurang dari 2 bulan
        if len(vals) < 2:
            continue
        
        adalah_anomali = False
        jenis_anomali = []
        variasi_min = float('inf')
        
        # Kriteria 1: Pelaporan Identik
        if len(set(vals)) == 1:
            adalah_anomali = True
            jenis_anomali.append("Pelaporan Identik (0.00%)")
        
        # Kriteria 2: Variasi Rendah (<= 20%)
        if len(vals) >= 2 and not (len(set(vals)) == 1):  # Jika tidak semua identik
            variasi_rendah_pct = None
            
            # Hitung semua perubahan persentase antar bulan
            for i in range(len(vals) - 1):
                if vals[i] != 0:
                    perubahan = abs(vals[i+1] - vals[i]) / vals[i]
                    variasi_min = min(variasi_min, perubahan)
                    
                    # Catat persentase variasi terendah yang <= 20%
                    if perubahan <= 0.20 and variasi_rendah_pct is None:
                        variasi_rendah_pct = perubahan
            
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
                'bulan_terisi': len(vals),
                'rata_rata': np.mean(vals),
                'min': np.min(vals),
                'max': np.max(vals),
                'std_dev': np.std(vals) if len(vals) > 1 else 0,
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
