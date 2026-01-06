import math


def get_reliability_factor(confidence_level):
    """
    Mengambil Reliability Factor (RF) untuk MUS (Zero Errors).
    Referensi: Juknis Hal 48 (Tabel 1.4).
    """
    lookup = {90: 2.31, 95: 3.00, 99: 4.61}
    return lookup.get(confidence_level, 3.00)


def get_expansion_factor(confidence_level):
    """
    Mengambil Expansion Factor (EF) untuk MUS.
    Referensi: Juknis Hal 48.
    """
    lookup = {90: 1.5, 95: 1.6, 99: 1.9}
    return lookup.get(confidence_level, 1.6)


def get_ur_coefficient(confidence_level):
    """
    Koefisien UR untuk metode MPU.
    Referensi: Juknis Hal 45 (Tabel Konversi).
    """
    lookup = {90: 1.65, 95: 1.96, 99: 2.58}
    return lookup.get(confidence_level, 1.96)


# --- FUNGSI UTAMA PERHITUNGAN N ---


def calculate_mus(total_nilai_buku, confidence_level, sst, dss):
    """
    Rumus Monetary Unit Sampling (MUS).
    Juknis Hal 47 (Poin 5b).
    n = (NB * RF) / (SST - (DSS * EF))
    """
    rf = get_reliability_factor(confidence_level)
    ef = get_expansion_factor(confidence_level)

    denominator = sst - (dss * ef)
    if denominator <= 0:
        return 0, "SST terlalu kecil dibandingkan DSS (Denominator <= 0)"

    n = (total_nilai_buku * rf) / denominator
    return math.ceil(n), None


def calculate_mpu_unstratified(population_size, confidence_level, sst,
                               std_dev_est):
    """
    Rumus Unstratified Mean Per Unit (MPU).
    Juknis Hal 45 (Poin 1b).
    n = ((UR * SD * N) / A)^2
    """
    ur = get_ur_coefficient(confidence_level)
    try:
        n = ((ur * std_dev_est * population_size) / sst)**2
        return math.ceil(n), None
    except ZeroDivisionError:
        return 0, "SST tidak boleh 0"


def calculate_difference_ratio(population_size, confidence_level, sst,
                               est_variance):
    """
    Rumus estimasi untuk Difference/Ratio Estimation.
    (Penyederhanaan berbasis varians).
    """
    ur = get_ur_coefficient(confidence_level)
    try:
        n = ((ur * population_size * math.sqrt(est_variance)) / sst)**2
        return math.ceil(n), None
    except ZeroDivisionError:
        return 0, "SST tidak boleh 0"


def calculate_mpu_stratified(strata_summary, confidence_level, sst):
    """
    Rumus Stratified MPU (Mean Per Unit).
    Referensi Logika: Juknis Hal 45 (modifikasi untuk multi-strata).
    Formula: n_total = ( (Ur * Sum(Ni * SDi)) / SST )^2
    """
    ur = get_ur_coefficient(confidence_level)

    # Hitung Sum(Ni * SDi)
    sum_ni_sdi = 0
    for item in strata_summary:
        ni = item['count']  # Jumlah populasi di strata i
        sdi = item['std_dev']  # Standar Deviasi estimasi di strata i
        sum_ni_sdi += (ni * sdi)

    try:
        # Hitung n total
        n_total = ((ur * sum_ni_sdi) / sst)**2
        return math.ceil(n_total), None
    except ZeroDivisionError:
        return 0, "SST tidak boleh 0"
