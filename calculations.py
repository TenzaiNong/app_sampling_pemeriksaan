import math


def get_reliability_factor(confidence_level):
    """
    Mengambil Reliability Factor (RF) untuk MUS (Zero Errors).
    Referensi: Juknis Hal 28 (Tabel 1.4).
    """
    lookup = {90: 2.40, 95: 3.00, 99: 3.70}
    return lookup.get(confidence_level, 3.00)


def get_expansion_factor(expansion_factor):
    """
    Mengambil Expansion Factor (EF) untuk MUS.
    Referensi: Juknis Hal 36.
    """
    lookup = {1: 1.9, 5: 1.6, 10: 1.5, 15: 1.4, 20: 1.3, 25: 1.25, 30: 1.2, 37: 1.15}
    return lookup.get(expansion_factor, 1.5)


def get_ur_coefficient(confidence_level):
    """
    Koefisien UR untuk metode MPU.
    Referensi: Juknis Hal 45 (Tabel Konversi).
    """
    lookup = {90: 1.65, 95: 1.96, 99: 2.58}
    return lookup.get(confidence_level, 1.96)


# --- FUNGSI UTAMA PERHITUNGAN N ---


def calculate_mus(total_nilai_buku, confidence_level, sst, dss, expansion_factor):
    """
    Rumus Monetary Unit Sampling (MUS).
    Juknis Hal 47 (Poin 5b).
    n = (NB * RF) / (SST - (DSS * EF))
    """
    rf = get_reliability_factor(confidence_level)
    ef = get_expansion_factor(expansion_factor)

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
        ns = ((ur * std_dev_est * population_size) / sst)**2
        n = ns/(1 + (ns / population_size))
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
    Rumus Stratified MPU (Mean Per Unit) dengan FPC Presisi.
    Formula: n = (Sum(Ni * Si))^2 / ( (SST/Ur)^2 + Sum(Ni * Si^2) )
    """
    ur = get_ur_coefficient(confidence_level)

    # Inisialisasi KEDUA variabel penampung
    sum_ni_sdi = 0
    sum_ni_sdi2 = 0  # <--- PERBAIKAN 1: Harus di-init nol dulu

    for item in strata_summary:
        ni = item['count']
        sdi = item['std_dev']
        
        # Hitung komponen
        sum_ni_sdi += (ni * sdi)          # Untuk Pembilang
        sum_ni_sdi2 += (ni * (sdi ** 2))  # Untuk Penyebut (Koreksi Populasi)

    try:
        # Hitung Varians yang Ditoleransi (V)
        if ur == 0: return 0, "Confidence Level error"
        v_allowed = (sst / ur) ** 2

        # PERBAIKAN 2: Rumus Pembilang harus dikuadratkan totalnya
        numerator = sum_ni_sdi ** 2       
        
        # Rumus Penyebut
        denominator = v_allowed + sum_ni_sdi2

        if denominator == 0:
            return 0, "Denominator 0, cek parameter SST"

        n_total = numerator / denominator
        return math.ceil(n_total), None

    except ZeroDivisionError:
        return 0, "SST tidak boleh 0"
