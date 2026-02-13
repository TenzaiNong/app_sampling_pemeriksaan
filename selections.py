import pandas as pd
import numpy as np
import math


def select_simple_random(df, n, random_state=42):
    """Acak Sederhana"""
    if n > len(df): n = len(df)
    return df.sample(n=n, random_state=random_state)


def select_pps(df, n, value_col, random_state=42):
    """
    Probability Proportional to Size (PPS).
    Wajib untuk MUS. Juknis Hal 54.
    """
    if n > len(df): n = len(df)
    # Bobot berdasarkan nilai rupiah (negatif dijadikan 0 agar tidak error)
    weights = df[value_col].apply(lambda x: max(x, 0))

    # Jika total weight 0, fallback ke random biasa
    if weights.sum() == 0:
        return df.sample(n=n, random_state=random_state)

    return df.sample(n=n, weights=weights, random_state=random_state)


def select_systematic(df, n):
    """
    Sistematis (Interval). Juknis Hal 53.
    """
    if n <= 0: return pd.DataFrame()
    if n > len(df): return df

    interval = len(df) // n
    start = np.random.randint(0, interval + 1)
    return df.iloc[start::interval].head(n)


def select_random_systematic(df, n):
    """
    Sistematis Acak (Interval + Jumps). Juknis Hal 53 (Poin 4).
    """
    if n <= 0: return pd.DataFrame()

    interval = len(df) // n
    if interval == 0: return df  # Populasi < n

    indices = []
    current_idx = np.random.randint(0, interval)

    for _ in range(n):
        if current_idx < len(df):
            indices.append(current_idx)
            # Lompatan acak antar interval
            current_idx += np.random.randint(1, max(interval * 2, 2))

    return df.iloc[indices]


def select_stratified_top_value(df, n, value_col):
    """
    Stratifikasi Sederhana (Top Value + Random).
    Mengambil item nilai terbesar sebagai prioritas.
    """
    if n > len(df): return df

    n_top = int(n * 0.2)  # 20% sampel adalah Top Value
    n_rand = n - n_top

    df_sorted = df.sort_values(by=value_col, ascending=False)
    top_strata = df_sorted.head(n_top)
    remaining = df_sorted.iloc[n_top:]
    random_strata = remaining.sample(n=n_rand)

    return pd.concat([top_strata, random_strata])


def select_benford_anomaly(df, n, value_col):
    """
    Benford's Law Analysis.
    Memilih item yang digit pertamanya (7,8,9) mencurigakan.
    """
    temp_df = df.copy()
    # Ambil digit pertama
    temp_df['first_digit'] = temp_df[value_col].astype(str).str.lstrip(
        '0').str[:1]

    # Fokus pada digit 7, 8, 9 (High Risk jika frekuensi tidak wajar)
    suspicious_digits = ['7', '8', '9']
    suspicious_df = temp_df[temp_df['first_digit'].isin(suspicious_digits)]

    if len(suspicious_df) == 0:
        return pd.DataFrame()  # Tidak ada anomali ditemukan

    if len(suspicious_df) >= n:
        sampled = suspicious_df.sample(n=n)
    else:
        sampled = suspicious_df  # Ambil semua yang mencurigakan

    # Bersihkan kolom helper
    if 'first_digit' in sampled.columns:
        del sampled['first_digit']

    return sampled


def select_stratified_distributed(df, allocation_dict, technique_name,
                                  value_col):
    """
    Fungsi Wrapper: Memilih sampel secara terpisah untuk setiap Strata
    berdasarkan jatah (allocation) yang sudah dihitung.
    """
    sampled_parts = []

    # allocation_dict format: {'< 100 Juta': 22, '> 100 Juta': 5}

    # Loop untuk setiap strata yang ada di data
    # Pastikan df memiliki kolom 'Strata'
    if 'Strata' not in df.columns:
        return pd.DataFrame()  # Error safety

    for stratum_name, n_target in allocation_dict.items():
        # Ambil data hanya untuk strata ini
        df_strata = df[df['Strata'] == stratum_name].copy()

        if len(df_strata) == 0 or n_target <= 0:
            continue

        # Terapkan teknik yang dipilih user ke sub-populasi ini
        res = pd.DataFrame()

        if technique_name == "Acak Sederhana":
            res = select_simple_random(df_strata, n_target)

        elif technique_name == "PPS (Wajib untuk MUS)":
            res = select_pps(df_strata, n_target, value_col)

        elif technique_name == "Sistematis":
            res = select_systematic(df_strata, n_target)

        elif technique_name == "Sistematis Acak":
            res = select_random_systematic(df_strata, n_target)

        elif technique_name == "Benford's Law (Anomali)":
            # Untuk Benford, biasanya tidak pakai target n spesifik per strata,
            # tapi kita coba paksa ambil n_target jika ada yang suspect
            res = select_benford_anomaly(df_strata, n_target, value_col)

        else:
            # Default fallback
            res = select_simple_random(df_strata, n_target)

        sampled_parts.append(res)

    # Gabungkan kembali semua hasil
    if sampled_parts:
        result = pd.concat(sampled_parts).reset_index(drop=True)

        # Pastikan kolom Strata ditampilkan dengan jelas (konversi ke string)
        if 'Strata' in result.columns:
            # Convert categories or intervals to readable strings
            try:
                result['Strata'] = result['Strata'].astype(str)
            except Exception:
                # fallback: cast element-wise
                result['Strata'] = result['Strata'].apply(lambda x: str(x) if pd.notna(x) else '')

        return result
    else:
        return pd.DataFrame()
