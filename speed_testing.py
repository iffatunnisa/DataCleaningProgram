"""Cleaning data Speed Testing.

Modul ini membaca data dari sheet Speed_Testing pada file Excel,
memilih kolom yang diperlukan, menormalisasi nilai throughput dan
koordinat, lalu menyimpan hasilnya ke file CSV.

Tahapan pemrosesan:
1. Memeriksa keberadaan file input.
2. Membaca sheet Speed_Testing dengan header pada baris kedua.
3. Memvalidasi keberadaan seluruh kolom wajib.
4. Memilih kolom yang diperlukan.
5. Mengonversi throughput menjadi angka.
6. Mengonversi Latitude dan Longitude menjadi numerik.
7. Menghapus baris yang tidak memiliki koordinat valid.
8. Menyimpan hasil cleaning dalam format CSV UTF-8 dengan BOM.
"""

import tempfile
from pathlib import Path

import pandas as pd


# =========================================================
# KONFIGURASI KOLOM
# =========================================================

# Daftar kolom yang wajib tersedia pada sheet Speed_Testing.
# Hanya kolom-kolom ini yang disertakan dalam file output.
REQUIRED_COLUMNS = [
    "Home Operator",
    "Application Throughput DL",
    "Application Throughput UL",
    "Latitude",
    "Longitude",
]


# =========================================================
# UTILITAS PEMROSESAN FILE
# =========================================================

def get_file_path(file):
    """Mengambil path file dari input Gradio atau path biasa.

    Args:
        file: Path berupa string, objek pathlib.Path, atau objek
            file yang menyediakan atribut name.

    Returns:
        str: Path file yang dapat digunakan untuk membaca Excel.
    """
    if isinstance(file, (str, Path)):
        return str(file)

    return file.name


# =========================================================
# NORMALISASI NILAI NUMERIK
# =========================================================

def clean_number(value):
    """Mengonversi nilai throughput menjadi angka float.

    Nilai kosong dikembalikan sebagai None. Spasi dihapus dan
    koma diubah menjadi titik jika koma merupakan satu-satunya
    pemisah desimal. Nilai yang tidak dapat dikonversi juga
    dikembalikan sebagai None.

    Args:
        value: Nilai throughput dari data mentah.

    Returns:
        float | None: Nilai numerik hasil konversi atau None
        jika nilai kosong maupun tidak valid.
    """
    if pd.isna(value):
        return None

    value = str(value).strip().replace(" ", "")

    if not value:
        return None

    if "," in value and "." not in value:
        value = value.replace(",", ".")

    try:
        return float(value)
    except (ValueError, TypeError):
        return None


# =========================================================
# PROSES UTAMA CLEANING SPEED TESTING
# =========================================================

def process_file(file):
    """Membersihkan data Speed Testing dari file Excel.

    Fungsi ini merupakan handler utama yang dipanggil oleh
    antarmuka Gradio.

    Input:
        file: File Excel yang diunggah pengguna.

    Output:
        tuple:
            - Path file CSV hasil cleaning atau None jika gagal.
            - Pesan status proses untuk ditampilkan di antarmuka.

    Catatan:
        Baris hanya dihapus jika Latitude atau Longitude tidak
        valid. Nilai throughput yang kosong atau tidak valid
        tidak menjadi alasan untuk menghapus baris.
    """

    # -----------------------------------------------------
    # 1. VALIDASI FILE INPUT
    # -----------------------------------------------------
    if file is None:
        return None, "❌ Silakan upload file Excel terlebih dahulu."

    # -----------------------------------------------------
    # 2. MEMBACA SHEET EXCEL
    # -----------------------------------------------------
    # Header berada pada baris kedua Excel, sehingga digunakan
    # header=1 karena indeks baris pandas dimulai dari nol.
    try:
        df = pd.read_excel(
            get_file_path(file),
            sheet_name="Speed_Testing",
            header=1,
        )
    except ValueError:
        return None, "❌ Sheet 'Speed_Testing' tidak ditemukan."
    except Exception as exc:
        return None, f"❌ Gagal membaca file: {exc}"

    # -----------------------------------------------------
    # 3. VALIDASI KOLOM WAJIB
    # -----------------------------------------------------
    # Proses dihentikan jika satu atau lebih kolom wajib tidak
    # ditemukan pada data yang dibaca.
    missing = [
        col for col in REQUIRED_COLUMNS
        if col not in df.columns
    ]

    if missing:
        return None, (
            "❌ Kolom berikut tidak ditemukan:\n\n"
            + "\n".join(f"- {col}" for col in missing)
        )

    # -----------------------------------------------------
    # 4. MEMILIH KOLOM YANG DIPERLUKAN
    # -----------------------------------------------------
    # Salinan dibuat agar normalisasi tidak mengubah DataFrame
    # asli yang dibaca dari file Excel.
    df_clean = df[REQUIRED_COLUMNS].copy()

    # -----------------------------------------------------
    # 5. NORMALISASI THROUGHPUT DOWNLOAD DAN UPLOAD
    # -----------------------------------------------------
    # Nilai throughput dikonversi menjadi float. Nilai kosong
    # atau tidak valid akan menjadi None.
    for col in (
        "Application Throughput DL",
        "Application Throughput UL",
    ):
        df_clean[col] = df_clean[col].apply(clean_number)

    # -----------------------------------------------------
    # 6. NORMALISASI KOORDINAT
    # -----------------------------------------------------
    # Nilai Latitude dan Longitude yang tidak dapat dikonversi
    # menjadi angka akan diubah menjadi NaN.
    for col in ("Latitude", "Longitude"):
        df_clean[col] = pd.to_numeric(
            df_clean[col],
            errors="coerce",
        )

    # -----------------------------------------------------
    # 7. FILTERING KOORDINAT TIDAK VALID
    # -----------------------------------------------------
    # Menghitung jumlah baris sebelum filtering untuk mengetahui
    # berapa baris yang dihapus karena koordinat tidak valid.
    before = len(df_clean)

    df_clean = df_clean.dropna(
        subset=["Latitude", "Longitude"]
    ).copy()

    removed = before - len(df_clean)

    # -----------------------------------------------------
    # 8. MENYIMPAN HASIL KE FILE CSV
    # -----------------------------------------------------
    # File sementara dibuat dengan nama berawalan
    # SpeedTesting_Cleaned_ dan ekstensi .csv.
    # delete=False memungkinkan file digunakan setelah blok
    # with selesai dijalankan.
    try:
        with tempfile.NamedTemporaryFile(
            prefix="SpeedTesting_Cleaned_",
            suffix=".csv",
            delete=False,
        ) as temp_file:
            output_file = temp_file.name

        # index=False mencegah indeks DataFrame ditambahkan
        # sebagai kolom baru. utf-8-sig membantu kompatibilitas
        # saat file CSV dibuka menggunakan Microsoft Excel.
        df_clean.to_csv(
            output_file,
            index=False,
            encoding="utf-8-sig",
        )

    except Exception as exc:
        return None, f"❌ Gagal menyimpan CSV: {exc}"

    # -----------------------------------------------------
    # 9. MENGEMBALIKAN FILE DAN STATUS PROSES
    # -----------------------------------------------------
    return output_file, (
        "✅ Speed Testing berhasil diproses!\n\n"
        f"Baris hasil: {len(df_clean)}\n"
        f"Jumlah kolom: {len(df_clean.columns)}\n"
        f"Baris tanpa koordinat valid yang dihapus: {removed}\n\n"
        f"File output: {output_file}"
    )
