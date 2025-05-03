import pandas as pd
import numpy as np
import openpyxl

def baca_data(nama_file):
    """Membaca data dari file Excel."""
    try:
        df = pd.read_excel(nama_file)
        return df
    except FileNotFoundError:
        print(f"Error: File '{nama_file}' tidak ditemukan.")
        return None

def fuzzifikasi_servis(kualitas):
    """Fungsi keanggotaan untuk kualitas servis."""
    rendah = max(0, min(1, (50 - kualitas) / 20))
    sedang = max(0, min(1, (kualitas - 30) / 30, (80 - kualitas) / 30))
    tinggi = max(0, min(1, (kualitas - 60) / 20))
    return {"rendah": rendah, "sedang": sedang, "tinggi": tinggi}

def fuzzifikasi_harga(harga):
    """Fungsi keanggotaan untuk harga."""
    murah = max(0, min(1, (40000 - harga) / 10000))
    sedang = max(0, min(1, (harga - 30000) / 10000, (50000 - harga) / 10000))
    mahal = max(0, min(1, (harga - 40000) / 10000))
    return {"murah": murah, "sedang": sedang, "mahal": mahal}

def inferensi(fuzzifikasi_servis, fuzzifikasi_harga):
    """Menerapkan aturan inferensi Fuzzy."""
    rules = {
        ("rendah", "murah"): "buruk",
        ("rendah", "sedang"): "buruk",
        ("rendah", "mahal"): "buruk",
        ("sedang", "murah"): "cukup",
        ("sedang", "sedang"): "baik",
        ("sedang", "mahal"): "cukup",
        ("tinggi", "murah"): "baik",
        ("tinggi", "sedang"): "sangat_baik",
        ("tinggi", "mahal"): "baik",
    }
    hasil_inferensi = {"buruk": 0, "cukup": 0, "baik": 0, "sangat_baik": 0}
    for (servis, harga), kualitas in rules.items():
        derajat_keanggotaan = min(fuzzifikasi_servis[servis], fuzzifikasi_harga[harga])
        hasil_inferensi[kualitas] = max(hasil_inferensi[kualitas], derajat_keanggotaan)
    return hasil_inferensi

def defuzzifikasi(hasil_inferensi):
    """Melakukan defuzzifikasi menggunakan metode Centroid of Area."""
    # Representasi himpunan fuzzy output (skor kelayakan)
    skor = np.arange(0, 101, 1)
    buruk = np.maximum(0, 1 - skor / 50)
    cukup = np.maximum(0, np.minimum(skor / 30, (80 - skor) / 30))
    baik = np.maximum(0, np.minimum((skor - 40) / 30, (90 - skor) / 30))
    sangat_baik = np.maximum(0, (skor - 70) / 30)

    # Agregasi output fuzzy sets berdasarkan derajat keanggotaan hasil inferensi
    aggregated = np.maximum.reduce([
        np.minimum(hasil_inferensi["buruk"], buruk),
        np.minimum(hasil_inferensi["cukup"], cukup),
        np.minimum(hasil_inferensi["baik"], baik),
        np.minimum(hasil_inferensi["sangat_baik"], sangat_baik)
    ])

    # Hitung centroid
    numerator = np.sum(skor * aggregated)
    denominator = np.sum(aggregated)

    if denominator == 0:
        return 0  # Handle kasus jika tidak ada area fuzzy yang aktif
    return numerator / denominator

def main():
    """Fungsi utama untuk menjalankan sistem Fuzzy Logic."""
    nama_file_input = "restoran.xlsx"
    data_restoran = baca_data(nama_file_input)

    if data_restoran is None:
        return

    hasil_evaluasi = []
    for index, row in data_restoran.iterrows():
        id_pelanggan = row['id Pelanggan']
        kualitas_servis = row['Pelayanan']
        harga = row['harga']

        fuzzy_servis = fuzzifikasi_servis(kualitas_servis)
        fuzzy_harga = fuzzifikasi_harga(harga)
        hasil_inferensi = inferensi(fuzzy_servis, fuzzy_harga)
        skor_kelayakan = defuzzifikasi(hasil_inferensi)

        hasil_evaluasi.append({
            'Id Pelanggan': id_pelanggan,
            'Kualitas Servis': kualitas_servis,
            'Harga': harga,
            'Skor Kelayakan': skor_kelayakan
        })

    # Urutkan restoran berdasarkan skor kelayakan tertinggi
    restoran_terbaik = sorted(hasil_evaluasi, key=lambda x: x['Skor Kelayakan'], reverse=True)[:5]

    # Simpan output ke file Excel
    nama_file_output = "peringkat.xlsx"
    df_output = pd.DataFrame(restoran_terbaik)
    try:
        df_output.to_excel(nama_file_output, index=False)
        print(f"\nHasil peringkat 5 restoran terbaik telah disimpan ke '{nama_file_output}'")
        print("\nDaftar 5 restoran terbaik:")
        for restoran in restoran_terbaik:
            print(f"ID: {restoran['Id Pelanggan']}, Servis: {restoran['Kualitas Servis']}, Harga: {restoran['Harga']:.2f}, Skor: {restoran['Skor Kelayakan']:.2f}")
    except Exception as e:
        print(f"Error saat menyimpan ke file Excel: {e}")

if __name__ == "__main__":
    main()