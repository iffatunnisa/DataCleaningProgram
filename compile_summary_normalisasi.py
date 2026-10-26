```python
"""Normalisasi nilai Compile Summary tanpa mengubah style sel Excel.

Modul ini bertanggung jawab untuk:
1. Menentukan tipe data berdasarkan nama kolom.
2. Membuat schema secara otomatis dari header.
3. Mengonversi nilai teks, tanggal, integer, dan float.
4. Mengubah nilai "No Data" menjadi 0 pada kolom numerik.
5. Memvalidasi seluruh nilai sebelum menerapkan perubahan.
6. Mempertahankan style sel Excel selama normalisasi.

Catatan:
- Schema dibentuk dari header, tanpa menggunakan file JSON.
- Kolom yang belum terdaftar akan menghasilkan error.
- Jika terdapat nilai yang gagal dikonversi, perubahan nilai
  tidak diterapkan ke worksheet.
- Style sel, seperti fill, border, alignment, dan number_format,
  tidak diatur ulang oleh proses normalisasi ini.
"""

from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import math

from openpyxl.utils.datetime import from_excel
from openpyxl.cell.cell import MergedCell


# =========================================================
# KONFIGURASI TIPE DATA
# =========================================================

# Kolom yang harus dipertahankan sebagai teks.
# Nilai seperti kode wilayah dan kategori tidak dikonversi
# menjadi angka meskipun bentuk nilainya menyerupai numerik.
TEXT_COLUMNS = {
    "Provinsi", "Kode Kabupaten Kota", "Kabupaten Kota", "Jenis Tes",
    "Collection", "Kecamatan", "Desa", "Lokasi Pengukuran", "Event",
    "PMT UPT", "Operator",
    "SIGS 5G_Category", "SIGQ 5G_Category", "SIGQ 5G_Category_2",
    "SIGS 4G_Category", "SIGQ 4G_Category", "SIGQ 4G_Category_2",
    "SIGQ 2G_Category",
}

# Kolom yang nilainya harus dinormalisasi menjadi tanggal.
DATE_COLUMNS = {"Tanggal Pengukuran"}

# Nama kolom yang tipe datanya ditentukan secara eksplisit.
# Kode Provinsi menggunakan integer, sedangkan kolom lainnya
# menggunakan float.
NUMERIC_EXACT = {
    "Kode Provinsi", "Lat", "Long", "Distance (Km)",
    "BROWSING_RTT (", "ST_Round Trip Time (RTT) (ms)",
    "ST_Round Trip Time (RTT) (ms)_2",
}

# Awalan nama kolom yang menandakan bahwa kolom tersebut numerik.
# Aturan ini digunakan untuk menangani banyak variasi kolom
# pengukuran tanpa harus mendaftarkan setiap nama satu per satu.
NUMERIC_PREFIXES = (
    'FTP DL_', 'FTP UL_', 'CAP DL_', 'CAP UL_', 'ST_',
    'HTTP DL_', 'HTTP UL_', 'BROWSING_', 'PING_', 'YT_',
    'IG POST PIC_', 'IG POST VID_', 'WA CALL_', 'WA MSG_',
    'SMS ONNET_', 'SMS OFFNET_', 'VCALL ONNET_', 'VCALL OFFNET_',
    'AVG VOICE_', 'SIGS 5G_', 'SIGQ 5G_', 'SIGS 4G_',
    'SIGQ 4G_', 'SIGS 2G_', 'SIGQ 2G_',
    'BAD SIGS 5G_', 'BAD SIGQ 5G_',
    'BAD SIGS 4G_', 'BAD SIGQ 4G_',
    'BAD SIGS 2G_', 'BAD SIGQ 2G_',
    'DETAIL Va5G_', 'DETAIL4G_', 'DETAIL2G_',
)

# Placeholder yang dianggap sebagai nilai kosong.
# "No Data" sengaja tidak dimasukkan karena memiliki aturan
# khusus pada konversi kolom numerik.
MISSING = {"", "n/a", "na", "none", "null", "nan", "-"}


# =========================================================
# PEMERIKSAAN NILAI
# =========================================================

def kosong(value):
    """Memeriksa apakah nilai kosong atau placeholder.

    Args:
        value: Nilai sel Excel yang akan diperiksa.

    Returns:
        bool: True jika nilai None atau teksnya termasuk
        placeholder dalam MISSING.
    """
    return (
        value is None
        or (
            isinstance(value, str)
            and value.strip().casefold() in MISSING
        )
    )


# =========================================================
# PENENTUAN TIPE DATA
# =========================================================

def jenis_kolom(name):
    """Menentukan tipe data berdasarkan nama kolom.

    Urutan pemeriksaan:
    1. Nama berakhiran _Category atau _Kategori dianggap teks.
    2. Kolom tanggal diperiksa melalui DATE_COLUMNS.
    3. Kolom teks diperiksa melalui TEXT_COLUMNS.
    4. Kolom numerik diperiksa melalui NUMERIC_EXACT dan
       NUMERIC_PREFIXES.
    5. Kolom yang tidak dikenali menghasilkan ValueError.

    Args:
        name: Nama kolom pada header Excel.

    Returns:
        str: Tipe data berupa text, date, integer, atau float.

    Raises:
        ValueError: Jika kolom belum memiliki aturan tipe data.
    """
    name = str(name).strip()

    if name.casefold().endswith(("_category", "_kategori")):
        return "text"

    if name in DATE_COLUMNS:
        return "date"

    if name in TEXT_COLUMNS:
        return "text"

    if name in NUMERIC_EXACT or name.startswith(NUMERIC_PREFIXES):
        return "integer" if name == "Kode Provinsi" else "float"

    raise ValueError(
        f"Kolom belum memiliki tipe dalam schema: {name!r}"
    )


def buat_schema_dari_header(headers):
    """Membentuk schema tipe data berdasarkan daftar header.

    Nama header dibersihkan dari spasi di awal dan akhir.
    Header kosong dan duplikat ditolak agar setiap kolom
    memiliki aturan normalisasi yang jelas.

    Args:
        headers: Daftar nama kolom dari worksheet.

    Returns:
        dict: Pemetaan nama kolom ke tipe data.

    Raises:
        ValueError: Jika header kosong atau terdapat duplikat.
        ValueError: Jika ada kolom yang belum terdaftar
        pada aturan jenis_kolom().
    """
    headers = [
        str(name).strip() if name is not None else ""
        for name in headers
    ]

    if not headers or not all(headers):
        raise ValueError("Header kosong.")

    if len(set(headers)) != len(headers):
        raise ValueError("Header duplikat.")

    return {
        name: jenis_kolom(name)
        for name in headers
    }


# =========================================================
# KONVERSI NILAI NUMERIK
# =========================================================

def angka(value):
    """Mengonversi nilai menjadi Decimal.

    Mendukung angka dengan koma desimal dan beberapa format
    pemisah ribuan, misalnya:
    - 12,5 menjadi 12.5
    - 1.234,5 menjadi 1234.5
    - 1,234.5 menjadi 1234.5
    - 25% menjadi 25

    Simbol persen dihapus tanpa membagi nilai dengan 100.
    Boolean dan angka non-finite seperti NaN atau Infinity
    ditolak.

    Args:
        value: Nilai numerik yang akan dikonversi.

    Returns:
        Decimal: Nilai numerik hasil konversi.

    Raises:
        ValueError: Jika nilai berupa boolean, format tidak
        valid, atau hasilnya bukan angka finite.
    """
    if isinstance(value, bool):
        raise ValueError("Boolean bukan angka.")

    if isinstance(value, (int, float, Decimal)):
        n = Decimal(str(value))
    else:
        s = (
            str(value)
            .strip()
            .replace("\u00a0", "")
            .replace(" ", "")
        )

        if s.endswith("%"):
            s = s[:-1]

        # Jika koma dan titik muncul bersamaan, pemisah paling
        # kanan dianggap sebagai pemisah desimal.
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")

        # Jika hanya koma yang muncul, koma dianggap desimal.
        elif "," in s:
            s = s.replace(",", ".")

        try:
            n = Decimal(s)
        except InvalidOperation as exc:
            raise ValueError(
                "Format numerik tidak valid."
            ) from exc

    if not n.is_finite():
        raise ValueError("Angka tidak finite.")

    return n


# =========================================================
# NORMALISASI NILAI BERDASARKAN TIPE DATA
# =========================================================

def konversi(value, tipe):
    """Mengonversi nilai sel berdasarkan tipe dalam schema.

    Aturan utama:
    - None tetap None.
    - Kolom teks tetap berupa teks.
    - "No Data" menjadi 0 pada kolom integer dan float.
    - Placeholder lain menjadi None.
    - Tanggal dikonversi dari objek Python, serial Excel,
      atau format teks yang didukung.
    - Integer harus berupa bilangan bulat.
    - Float harus menghasilkan angka finite.

    Fungsi ini hanya menghasilkan nilai baru dan tidak
    mengubah worksheet atau style sel secara langsung.

    Args:
        value: Nilai sel yang akan dikonversi.
        tipe: Tipe data tujuan dari schema.

    Returns:
        Nilai hasil normalisasi sesuai tipe data.

    Raises:
        ValueError: Jika nilai tidak sesuai tipe tujuan.
    """
    if value is None:
        return None

    text_value = str(value).strip()

    # Kolom teks tetap berupa teks, termasuk nilai No Data.
    if tipe == "text":
        return text_value

    # No Data menjadi 0 hanya untuk kolom numerik.
    if (
        tipe in ("integer", "float")
        and text_value.casefold() == "no data"
    ):
        return 0

    # Placeholder kosong dinormalisasi menjadi None.
    if kosong(value):
        return None

    # -----------------------------------------------------
    # KONVERSI TANGGAL
    # -----------------------------------------------------
    if tipe == "date":
        if isinstance(value, bool):
            raise ValueError("Boolean bukan tanggal.")

        if isinstance(value, datetime):
            return value.date()

        if isinstance(value, date):
            return value

        # Angka diperlakukan sebagai serial tanggal Excel.
        if isinstance(value, (int, float)):
            converted = from_excel(value)

            if isinstance(converted, datetime):
                return converted.date()

            return converted

        # Format tanggal teks yang didukung.
        for fmt in (
            '%Y-%m-%d',  # 2026-10-05
            '%d/%m/%Y',  # 04/03/2026
            '%d-%m-%Y',  # 05-10-2026
            '%d-%b-%Y',  # 05-Oct-2026
            '%d-%B-%Y',  # 05-October-2026
            '%Y/%m/%d',  # 2026/10/05
        ):
            try:
                return datetime.strptime(text_value, fmt).date()
            except ValueError:
                continue

        raise ValueError(
            "Tanggal tidak sesuai format yang didukung."
        )

    # -----------------------------------------------------
    # KONVERSI INTEGER DAN FLOAT
    # -----------------------------------------------------
    n = angka(value)

    if tipe == "integer":
        if n != n.to_integral_value():
            raise ValueError("Angka bukan bilangan bulat.")

        return int(n)

    if tipe == "float":
        result = float(n)

        if not math.isfinite(result):
            raise ValueError("Angka terlalu besar.")

        return result

    raise ValueError(f"Tipe data tidak dikenal: {tipe!r}")


# =========================================================
# NORMALISASI WORKSHEET
# =========================================================

def normalisasi_sheet(ws, schema):
    """Menormalisasi nilai worksheet tanpa mengubah style sel.

    Proses dilakukan dalam dua tahap:

    1. Validasi:
       - Membaca dan memeriksa header.
       - Memastikan seluruh kolom tersedia dalam schema.
       - Mengonversi nilai setiap sel dan menyimpan hasilnya
         sementara dalam daftar updates.
       - Mengumpulkan kesalahan konversi tanpa menerapkan
         perubahan pada worksheet.

    2. Penerapan:
       - Jika validasi berhasil, seluruh nilai hasil konversi
         diterapkan ke sel yang bersangkutan.
       - Jika terdapat kesalahan, ValueError diberikan dan
         perubahan nilai tidak diterapkan.

    Sel MergedCell dilewati karena bukan sel utama yang dapat
    menerima nilai. Fungsi hanya mengubah cell.value, bukan
    properti style seperti fill, border, alignment, atau
    number_format.

    Args:
        ws: Worksheet openpyxl yang akan dinormalisasi.
        schema: Pemetaan nama kolom ke tipe data.

    Returns:
        dict: Ringkasan jumlah baris, kolom, dan kolom tambahan.

    Raises:
        ValueError: Jika header kosong, duplikat, tidak terdaftar
        dalam schema, atau terdapat nilai yang gagal dikonversi.
    """

    # -----------------------------------------------------
    # 1. MEMBACA DAN MEMVALIDASI HEADER
    # -----------------------------------------------------
    headers = [
        str(ws.cell(1, col).value).strip()
        if ws.cell(1, col).value is not None
        else ""
        for col in range(1, ws.max_column + 1)
    ]

    if (
        not headers
        or not all(headers)
        or len(set(headers)) != len(headers)
    ):
        raise ValueError("Header kosong atau duplikat.")

    # Pastikan setiap nama kolom telah terdaftar dalam schema.
    extras = [
        name for name in headers
        if name not in schema
    ]

    if extras:
        raise ValueError(
            "Kolom belum masuk schema: " + ", ".join(extras)
        )

    # Menampung perubahan sementara dan kesalahan konversi.
    # Worksheet belum diubah selama tahap validasi.
    updates = []
    errors = []

    # -----------------------------------------------------
    # 2. VALIDASI DAN KONVERSI SEL
    # -----------------------------------------------------
    # Baris pertama berisi header, sehingga pemrosesan data
    # dimulai dari baris kedua.
    for row_num in range(2, ws.max_row + 1):
        for col_num, name in enumerate(headers, start=1):
            cell = ws.cell(row_num, col_num)

            # Sel bagian dari merged range yang bukan sel utama
            # tidak dapat menerima nilai, sehingga dilewati.
            if isinstance(cell, MergedCell):
                continue

            try:
                new_value = konversi(
                    cell.value,
                    schema[name],
                )
                updates.append((cell, new_value))

            except (ValueError, TypeError, OverflowError) as exc:
                errors.append(
                    f"Baris {row_num}, kolom {name}: "
                    f"{cell.value!r} -> {exc}"
                )

    # Jika ada nilai yang gagal dikonversi, hentikan proses
    # sebelum menerapkan perubahan apa pun ke worksheet.
    if errors:
        raise ValueError(
            f"{len(errors)} nilai gagal dikonversi:\n"
            + "\n".join(errors[:25])
        )

    # -----------------------------------------------------
    # 3. MENERAPKAN NILAI HASIL NORMALISASI
    # -----------------------------------------------------
    # Hanya cell.value yang diperbarui. Style sel tidak disentuh.
    for cell, new_value in updates:
        cell.value = new_value

    # -----------------------------------------------------
    # 4. MENGEMBALIKAN RINGKASAN PEMROSESAN
    # -----------------------------------------------------
    return {
        "baris": max(0, ws.max_row - 1),
        "kolom": len(headers),
        "kolom_ditambahkan": [],
    }