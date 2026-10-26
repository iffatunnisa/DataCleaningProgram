
"""
compile_summary.py

Modul untuk membersihkan struktur header dan memvalidasi kolom pada
file Compile Summary sebelum proses normalisasi tipe data.

Alur utama:
1. Membaca file Excel yang diunggah.
2. Menggabungkan dua baris header menjadi satu baris.
3. Menyingkat nama header berdasarkan kamus yang telah ditentukan.
4. Menghapus kolom nomor urut dan kolom pemisah yang tidak diperlukan.
5. Memvalidasi keberadaan dan posisi kolom wajib.
6. Menyisipkan kolom yang belum tersedia dengan nilai default 0.
7. Menormalisasi tipe data melalui modul compile_summary_normalisasi.
8. Menyimpan hasil cleaning ke file Excel baru.
9. Menampilkan jumlah baris, jumlah kolom, dan runtime pemrosesan.

Catatan:
- File sumber tidak ditimpa; hasil disimpan sebagai file baru.
- Aturan posisi kolom diterapkan untuk menjaga konsistensi struktur data.
- Normalisasi tipe data ditangani oleh compile_summary_normalisasi.py.
"""

import time
import re
import tempfile
from pathlib import Path
from copy import copy

from openpyxl import load_workbook

from compile_summary_normalisasi import (
    buat_schema_dari_header,
    normalisasi_sheet,
)


# =============================================================
# 1. KONFIGURASI PEMETAAN HEADER
# =============================================================

# Memetakan nama kelompok header dari file sumber menjadi nama singkat
# yang konsisten. Beberapa variasi ejaan sengaja dicantumkan untuk
# mengakomodasi perbedaan penulisan pada file Compile Summary.
HEADER_MAP = {
    "FTP DL": "FTP DL",
    "FTP UL": "FTP UL",
    "Capacity DL": "CAP DL",
    "Capacity UL": "CAP UL",
    "SpeedTest": "ST",
    "Speed Test": "ST",
    "HTTP DL Transfer": "HTTP DL",
    "HTTP UL Transfer": "HTTP UL",
    "Browsing": "BROWSING",
    "Ping Test": "PING",
    "Youtube": "YT",
    "Instagram Post Picture": "IG POST PIC",
    "Instagram Post Video": "IG POST VID",
    "Whatsapp Call": "WA CALL",
    "Whatsapp Message": "WA MSG",
    "SMS Onnet": "SMS ONNET",
    "SMS Offnet": "SMS OFFNET",
    "Voice Call Onnet": "VCALL ONNET",
    "Voice Call Offnet": "VCALL OFFNET",
    "Average Voice Call": "AVG VOICE",
    "Signal Strenght 5G": "SIGS 5G",
    "Signal Strength 5G": "SIGS 5G",
    "Signal Qualty 5G": "SIGQ 5G",
    "Signal Quality 5G": "SIGQ 5G",
    "Signal Strenght 4G": "SIGS 4G",
    "Signal Strength 4G": "SIGS 4G",
    "Signal Qualty 4G": "SIGQ 4G",
    "Signal Quality 4G": "SIGQ 4G",
    "Signal Strenght 2G": "SIGS 2G",
    "Signal Strength 2G": "SIGS 2G",
    "Signal Qualty 2G": "SIGQ 2G",
    "Signal Quality 2G": "SIGQ 2G",
    "Bad Signal Strenght 5G": "BAD SIGS 5G",
    "Bad Signal Strength 5G": "BAD SIGS 5G",
    "Bad Signal Qualty 5G": "BAD SIGQ 5G",
    "Bad Signal Quality 5G": "BAD SIGQ 5G",
    "Bad Signal Strenght 4G": "BAD SIGS 4G",
    "Bad Signal Strength 4G": "BAD SIGS 4G",
    "Bad Signal Qualty 4G": "BAD SIGQ 4G",
    "Bad Signal Quality 4G": "BAD SIGQ 4G",
    "Bad Signal Strenght 2G": "BAD SIGS 2G",
    "Bad Signal Strength 2G": "BAD SIGS 2G",
    "Bad Signal Qualty 2G": "BAD SIGQ 2G",
    "Bad Signal Quality 2G": "BAD SIGQ 2G",
    "Detail Value 5G": "DETAIL Va5G",
    "Detail Value 4G": "DETAIL4G",
    "Detail Value 2G": "DETAIL2G",
}

# Memproses nama kelompok yang lebih panjang terlebih dahulu.
# Ini membantu menghindari pencocokan yang kurang tepat apabila
# nama header memiliki bagian teks yang sama.
SORTED_GROUPS = sorted(
    HEADER_MAP,
    key=len,
    reverse=True,
)


# =============================================================
# 2. FUNGSI BANTU UNTUK PEMBERSIHAN HEADER
# =============================================================

def get_file_path(file):
    """
    Mengambil path file dari input Gradio atau path biasa.

    Args:
        file: Path berbentuk string, objek Path, atau objek file Gradio.

    Returns:
        Path file dalam bentuk string.
    """
    if isinstance(file, (str, Path)):
        return str(file)

    return file.name


def normalize_text(value):
    """
    Membersihkan teks header dari spasi dan karakter yang tidak diperlukan.

    Penanganan:
    - Mengganti non-breaking space dan line break.
    - Menggabungkan spasi berulang.
    - Mengubah nilai kosong, 'nan', dan 'none' menjadi string kosong.
    - Mengabaikan header otomatis seperti 'Unnamed: ...'.

    Args:
        value: Nilai header dari worksheet.

    Returns:
        Teks header yang telah dibersihkan.
    """
    if value is None:
        return ""

    value = str(value).replace("\xa0", " ").replace("\n", " ")
    value = re.sub(r"\s+", " ", value).strip()

    if value.casefold() in ("", "nan", "none"):
        return ""

    if value.startswith("Unnamed:"):
        return ""

    return value


def abbreviate_parent(parent):
    """
    Menyingkat nama kelompok header berdasarkan HEADER_MAP.

    Pencocokan tidak membedakan huruf besar dan kecil. Jika nama kelompok
    tidak ditemukan dalam kamus, teks aslinya tetap digunakan.
    """
    parent = normalize_text(parent)

    for original in SORTED_GROUPS:
        if parent.casefold() == original.casefold():
            return HEADER_MAP[original]

    return parent


def abbreviate_child(child, parent=""):
    """
    Membersihkan dan menyingkat nama subheader atau metrik.

    Langkah utama:
    1. Membersihkan teks subheader.
    2. Menghapus awalan kelompok yang sudah direpresentasikan oleh header
       induk atau nama kelompok lain yang dikenal.
    3. Menyeragamkan istilah, satuan, dan ejaan tertentu.
    4. Menghapus karakter yang tidak didukung oleh format nama kolom.

    Args:
        child: Nama subheader dari baris kedua.
        parent: Nama header induk dari baris pertama.

    Returns:
        Nama subheader yang telah dinormalisasi.
    """
    child = normalize_text(child)
    parent = normalize_text(parent)

    if not child:
        return ""

    # Daftar awalan yang mungkin sudah terwakili oleh nama kelompok.
    prefixes = [parent] if parent else []

    prefixes += [
        "IG Post Picture",
        "IG Post Video",
        "Instagram Post Picture",
        "Instagram Post Video",
        "FTP DL",
        "FTP UL",
        "Capacity DL",
        "Capacity UL",
        "Speed Test",
        "SpeedTest",
        "HTTP DL",
        "HTTP UL",
        "PING",
        "Youtube",
        "Whatsapp Call",
        "Whatsapp Message",
        "WA Call",
        "WA Message",
        "Browsing",
    ]

    # Mengutamakan awalan terpanjang agar bagian teks yang lebih spesifik
    # dihapus sebelum awalan yang lebih pendek.
    for prefix in sorted(set(prefixes), key=len, reverse=True):
        if prefix and child.casefold().startswith(prefix.casefold()):
            child = child[len(prefix):].strip(" _-")

    # Menyeragamkan ejaan dan penulisan satuan metrik.
    replacements = [
        (r"\bAttemps\b", "Attempts"),
        (r"\bSuccess Attemps\b", "Success Attempts"),
        (r"\bAverage\b", "Avg"),
        (r"\bMaximum\b", "Max"),
        (r"\bMinimum\b", "Min"),
        (r"\bStrenght\b", "Strength"),
        (r"\(\s*Mbps\s*\)|\bMbps\b", "(Mbps)"),
        (r"\(\s*ms\s*\)|\bms\b", "(ms)"),
        (r"\(\s*s\s*\)|\bs\b", "(s)"),
        (r"\(\s*%\s*\)|%", "(%)"),
    ]

    for pattern, replacement in replacements:
        child = re.sub(
            pattern,
            replacement,
            child,
            flags=re.IGNORECASE,
        )

    # Menyeragamkan kapitalisasi istilah teknis tanpa mengubah istilahnya.
    for term in (
        "MOS", "RTT", "RSRP", "RSRQ", "SINR", "TTFP", "CST",
        "Onnet", "Offnet", "Category", "Quality", "Strength",
        "Attempts", "Success", "Median", "Duration", "Speed",
        "Latency", "Sample",
    ):
        child = re.sub(
            rf"\b{re.escape(term)}\b",
            term,
            child,
            flags=re.IGNORECASE,
        )

    # Mempertahankan karakter yang diperbolehkan dalam nama kolom.
    child = re.sub(r"[^A-Za-z0-9 (%)-]+", " ", child)
    child = re.sub(r"\s+", " ", child).strip()

    return child


def make_unique_columns(columns):
    """
    Membuat nama kolom unik apabila terdapat nama yang sama.

    Nama pertama dipertahankan, sedangkan kemunculan berikutnya diberi
    akhiran _2, _3, dan seterusnya.

    Contoh:
        RTT, RTT -> RTT, RTT_2

    Catatan:
        Penomoran dilakukan berdasarkan nama yang sama persis setelah
        proses pembersihan header.
    """
    result = []
    counts = {}

    for name in columns:
        base = name or "Column"
        counts[base] = counts.get(base, 0) + 1

        if counts[base] == 1:
            result.append(base)
        else:
            result.append(f"{base}_{counts[base]}")

    return result


# =============================================================
# 3. PENCARIAN BARIS TERAKHIR YANG BERISI DATA
# =============================================================

def cari_baris_terakhir_data(ws, header_row=1):
    """
    Mencari baris terakhir yang memiliki setidaknya satu nilai data.

    Pemeriksaan dilakukan dari baris paling bawah ke atas. Baris yang
    hanya memiliki format Excel, tetapi tidak memiliki nilai, diabaikan.

    Args:
        ws: Worksheet yang sedang diproses.
        header_row: Nomor baris header yang tidak dihitung sebagai data.

    Returns:
        Nomor baris terakhir yang berisi data. Jika tidak ada data,
        fungsi mengembalikan nomor baris header.

    Catatan:
        Fungsi ini digunakan untuk membatasi pengisian nilai default
        agar tidak menulis angka 0 sampai ke baris kosong akibat format
        worksheet yang masih tersisa.
    """
    for row_idx in range(ws.max_row, header_row, -1):
        for col_idx in range(1, ws.max_column + 1):
            value = ws.cell(row_idx, col_idx).value

            if value is not None and str(value).strip() != "":
                return row_idx

    return header_row


# =============================================================
# 4. VALIDASI DAN PENYISIPAN KOLOM WAJIB
# =============================================================

def validasi_dan_sisipkan_kolom(ws, headers):
    """
    Memastikan kelengkapan dan posisi kolom wajib pada worksheet.

    Aturan struktur:
    1. Kode Provinsi harus tepat sebelum Provinsi.
    2. RTT pertama harus tepat sebelum ST_UL Success (%).
    3. RTT kedua (_2) harus tepat sebelum ST_Rank Avg DL Speed (Mbps).
    4. Kolom wajib yang hilang ditambahkan dengan nilai default 0.
    5. Jika hanya ada satu RTT tanpa akhiran _2 dan posisinya tepat
       sebelum ST_Rank Avg DL Speed (Mbps), kolom itu diperlakukan
       sebagai RTT kedua.
    6. Nilai pada kolom yang sudah tersedia harus dipertahankan.
    7. Jika posisi atau jumlah header tidak sesuai aturan, proses
       dihentikan dengan ValueError untuk mencegah hasil yang keliru.

    Args:
        ws: Worksheet Excel yang sedang diproses.
        headers: Daftar nama kolom yang sesuai dengan worksheet.

    Returns:
        Daftar header setelah validasi dan penyisipan.

    Raises:
        ValueError: Jika header acuan hilang, duplikat, atau posisi
                    kolom tidak sesuai aturan yang ditentukan.
    """
    headers = list(headers)

    # Nama kolom yang diwajibkan dan digunakan sebagai patokan posisi.
    KODE_PROVINSI = "Kode Provinsi"
    PROVINSI = "Provinsi"

    RTT = "ST_Round Trip Time (RTT) (ms)"
    RTT_2 = "ST_Round Trip Time (RTT) (ms)_2"

    ANCHOR_RTT_1 = "ST_UL Success (%)"
    ANCHOR_RTT_2 = "ST_Rank Avg DL Speed (Mbps)"

    # ---------------------------------------------------------
    # Helper: mencari posisi header
    # ---------------------------------------------------------

    def cari_posisi(nama):
        """Mengembalikan semua indeks header yang cocok dengan nama."""
        return [
            i for i, header in enumerate(headers)
            if header == nama
        ]

    def posisi_tunggal(nama):
        """
        Mengambil indeks header yang harus muncul tepat satu kali.

        Indeks yang dikembalikan berbasis 0, sesuai indeks list Python.
        """
        posisi = cari_posisi(nama)

        if len(posisi) != 1:
            raise ValueError(
                f"Header {nama!r} harus tersedia tepat satu kali, "
                f"tetapi ditemukan {len(posisi)} kali."
            )

        return posisi[0]

    # ---------------------------------------------------------
    # Helper: menyisipkan kolom sebelum header acuan
    # ---------------------------------------------------------

    def sisipkan_sebelum(nama_baru, anchor):
        """
        Menambahkan kolom baru tepat sebelum header acuan.

        Selain memperbarui daftar header, fungsi ini juga:
        - Menyisipkan kolom pada worksheet.
        - Menyalin format dari kolom acuan.
        - Menyalin lebar kolom jika tersedia.
        - Mengisi nilai default 0 pada baris yang memiliki data.

        Kolom yang dibuat merupakan kolom baru; nilai default tidak
        ditulis ke kolom lama yang telah berisi data.
        """
        posisi_anchor = posisi_tunggal(anchor)

        # List menggunakan indeks mulai 0, sedangkan worksheet Excel
        # menggunakan nomor kolom mulai 1.
        nomor_kolom = posisi_anchor + 1

        # Menyisipkan kolom baru. Kolom yang ada di sebelah kanan
        # akan bergeser satu posisi.
        ws.insert_cols(nomor_kolom, 1)

        # Kolom acuan berpindah ke kanan setelah penyisipan.
        # Formatnya digunakan sebagai dasar format kolom baru.
        kolom_sumber = nomor_kolom + 1

        for row_idx in range(1, ws.max_row + 1):
            sumber = ws.cell(row_idx, kolom_sumber)
            tujuan = ws.cell(row_idx, nomor_kolom)

            if sumber.has_style:
                tujuan._style = copy(sumber._style)

            if sumber.number_format:
                tujuan.number_format = sumber.number_format

            if sumber.alignment:
                tujuan.alignment = copy(sumber.alignment)

            if sumber.protection:
                tujuan.protection = copy(sumber.protection)

        # Menyalin lebar kolom dari kolom acuan agar tampilan Excel
        # tetap konsisten setelah kolom baru ditambahkan.
        huruf_sumber = ws.cell(
            1, kolom_sumber
        ).column_letter

        huruf_tujuan = ws.cell(
            1, nomor_kolom
        ).column_letter

        lebar = ws.column_dimensions[huruf_sumber].width

        if lebar is not None:
            ws.column_dimensions[huruf_tujuan].width = lebar

        # Memperbarui daftar header agar tetap sinkron dengan worksheet.
        headers.insert(posisi_anchor, nama_baru)
        ws.cell(1, nomor_kolom).value = nama_baru

        # Isi nilai default hanya sampai baris terakhir yang benar-benar
        # berisi data, bukan sampai baris maksimum akibat format Excel.
        baris_terakhir = cari_baris_terakhir_data(ws)

        for row_idx in range(2, baris_terakhir + 1):
            ws.cell(row_idx, nomor_kolom).value = 0

    # ---------------------------------------------------------
    # 4.1 Validasi Kode Provinsi dan Provinsi
    # ---------------------------------------------------------

    # Provinsi menjadi patokan posisi Kode Provinsi sehingga harus ada
    # tepat satu kali sebelum kolom baru dapat disisipkan.
    posisi_tunggal(PROVINSI)

    posisi_kode = cari_posisi(KODE_PROVINSI)

    if len(posisi_kode) > 1:
        raise ValueError(
            f"Header {KODE_PROVINSI!r} ditemukan lebih dari satu kali."
        )

    # Jika belum tersedia, buat Kode Provinsi tepat sebelum Provinsi.
    if not posisi_kode:
        sisipkan_sebelum(KODE_PROVINSI, PROVINSI)

    # Jika sudah tersedia, validasi posisinya tanpa menimpa nilainya.
    if posisi_tunggal(KODE_PROVINSI) + 1 != posisi_tunggal(PROVINSI):
        raise ValueError(
            f"Posisi {KODE_PROVINSI!r} tidak sesuai standar. "
            f"Kolom harus tepat sebelum {PROVINSI!r}."
        )

    # ---------------------------------------------------------
    # 4.2 Pastikan header acuan RTT tersedia
    # ---------------------------------------------------------

    # Kedua header ini diperlukan untuk menentukan posisi standar RTT.
    # Jika salah satunya tidak ada atau duplikat, proses dihentikan.
    posisi_tunggal(ANCHOR_RTT_1)
    posisi_tunggal(ANCHOR_RTT_2)

    # ---------------------------------------------------------
    # 4.3 Identifikasi RTT pertama dan RTT kedua
    # ---------------------------------------------------------

    posisi_rtt = cari_posisi(RTT)
    posisi_rtt_2 = cari_posisi(RTT_2)

    if len(posisi_rtt) > 1:
        raise ValueError(
            f"Header {RTT!r} ditemukan lebih dari satu kali."
        )

    if len(posisi_rtt_2) > 1:
        raise ValueError(
            f"Header {RTT_2!r} ditemukan lebih dari satu kali."
        )

    # Penanganan kasus khusus:
    # Jika hanya ada satu header RTT dan posisinya tepat sebelum kolom
    # Rank Avg DL, RTT tersebut dianggap sebagai RTT kedua.
    #
    # Nama header diubah menjadi RTT_2, tetapi nilai sel tidak diubah.
    if (
        len(posisi_rtt) == 1
        and len(posisi_rtt_2) == 0
        and posisi_rtt[0] == posisi_tunggal(ANCHOR_RTT_2) - 1
    ):
        indeks_header = posisi_rtt[0]
        nomor_kolom = indeks_header + 1

        headers[indeks_header] = RTT_2
        ws.cell(1, nomor_kolom).value = RTT_2

    # ---------------------------------------------------------
    # 4.4 Pastikan RTT pertama tersedia pada posisi standar
    # ---------------------------------------------------------

    # RTT pertama harus berada tepat sebelum ST_UL Success (%).
    # Jika tidak tersedia, kolom baru dibuat dengan nilai default 0.
    if not cari_posisi(RTT):
        sisipkan_sebelum(RTT, ANCHOR_RTT_1)

    # ---------------------------------------------------------
    # 4.5 Pastikan RTT kedua tersedia pada posisi standar
    # ---------------------------------------------------------

    # RTT kedua harus berada tepat sebelum ST_Rank Avg DL Speed (Mbps).
    # Jika tidak tersedia, kolom baru dibuat dengan nilai default 0.
    if not cari_posisi(RTT_2):
        sisipkan_sebelum(RTT_2, ANCHOR_RTT_2)

    # ---------------------------------------------------------
    # 4.6 Validasi posisi akhir kolom RTT
    # ---------------------------------------------------------

    # Pemeriksaan akhir memastikan kedua RTT benar-benar berada di
    # posisi yang ditentukan, bukan hanya tersedia dalam worksheet.
    if posisi_tunggal(RTT) + 1 != posisi_tunggal(ANCHOR_RTT_1):
        raise ValueError(
            f"Posisi {RTT!r} tidak sesuai standar. "
            f"Kolom harus tepat sebelum {ANCHOR_RTT_1!r}. "
            "File tidak diproses untuk menghindari perubahan "
            "data yang tidak aman."
        )

    if posisi_tunggal(RTT_2) + 1 != posisi_tunggal(ANCHOR_RTT_2):
        raise ValueError(
            f"Posisi {RTT_2!r} tidak sesuai standar. "
            f"Kolom harus tepat sebelum {ANCHOR_RTT_2!r}. "
            "File tidak diproses untuk menghindari perubahan "
            "data yang tidak aman."
        )

    # ---------------------------------------------------------
    # 4.7 Sinkronisasi header dan worksheet
    # ---------------------------------------------------------

    # Jumlah header harus sama dengan jumlah kolom worksheet. Pemeriksaan
    # ini mencegah daftar header yang tidak sinkron dengan data Excel.
    if len(headers) != ws.max_column:
        raise ValueError(
            "Jumlah header tidak sama dengan jumlah kolom worksheet. "
            "Proses dihentikan untuk mencegah pergeseran data."
        )

    # Tulis ulang header berdasarkan daftar final setelah seluruh
    # validasi dan penyisipan selesai.
    for nomor_kolom, nama_header in enumerate(headers, start=1):
        ws.cell(1, nomor_kolom).value = nama_header

    return headers


# =============================================================
# 5. PROSES UTAMA PEMBERSIHAN COMPILE SUMMARY
# =============================================================

def process_summary(file):
    """
    Memproses file Compile Summary dari awal hingga menghasilkan Excel baru.

    Tahapan pemrosesan:
    1. Membaca file Excel.
    2. Membersihkan dan menggabungkan header.
    3. Menghapus kolom yang tidak diperlukan.
    4. Memvalidasi dan menyisipkan kolom wajib.
    5. Membuat schema normalisasi.
    6. Menormalisasi tipe data.
    7. Menyimpan hasil dan mencatat runtime.

    Args:
        file: File yang diunggah melalui Gradio atau path file Excel.

    Returns:
        Tuple (output_file, message):
        - output_file: Path file hasil jika pemrosesan berhasil.
        - message: Informasi hasil pemrosesan dan runtime, atau pesan error.
    """
    if file is None:
        return None, "❌ Silakan upload file Compile Summary terlebih dahulu."

    wb = None
    waktu_mulai = time.perf_counter()
    durasi = {}

    try:
        # -----------------------------------------------------
        # Tahap 1: Membaca file Excel
        # -----------------------------------------------------

        t0 = time.perf_counter()

        input_path = get_file_path(file)
        wb = load_workbook(input_path)

        durasi["Baca file Excel"] = time.perf_counter() - t0

        # Gunakan sheet Compile_Summary jika tersedia. Jika tidak,
        # gunakan worksheet aktif sebagai sumber data.
        ws = (
            wb["Compile_Summary"]
            if "Compile_Summary" in wb.sheetnames
            else wb.active
        )

        # Format sumber diharapkan memiliki dua baris header.
        if ws.max_row < 2:
            raise ValueError("File harus memiliki dua baris header.")

        # -----------------------------------------------------
        # Tahap 2: Cleaning header
        # -----------------------------------------------------

        t0 = time.perf_counter()

        max_col = ws.max_column

        # Ambil kedua baris header sebelum worksheet diubah.
        parent_headers = [
            ws.cell(1, col).value
            for col in range(1, max_col + 1)
        ]
        child_headers = [
            ws.cell(2, col).value
            for col in range(1, max_col + 1)
        ]

        # Pada header gabungan Excel, nilai kelompok biasanya hanya
        # berada di sel pertama. Sebarkan nilai tersebut ke seluruh
        # kolom yang termasuk dalam merged range header.
        for merged_range in list(ws.merged_cells.ranges):
            if (
                merged_range.min_row == 1
                and merged_range.max_row == 1
                and merged_range.max_col > merged_range.min_col
            ):
                parent_value = ws.cell(
                    1, merged_range.min_col
                ).value

                for col in range(
                    merged_range.min_col,
                    merged_range.max_col + 1,
                ):
                    parent_headers[col - 1] = parent_value

        # Gabungkan header induk dan subheader menjadi satu nama.
        # Nama kelompok dan metrik dinormalisasi secara terpisah agar
        # aturan singkatan dapat diterapkan secara konsisten.
        new_headers = []

        for index, (parent, child) in enumerate(
            zip(parent_headers, child_headers),
            start=1,
        ):
            parent_text = normalize_text(parent)
            child_text = normalize_text(child)

            short_parent = abbreviate_parent(parent_text)
            short_child = abbreviate_child(child_text, parent_text)

            if short_parent and short_child:
                name = f"{short_parent}_{short_child}"
            elif short_child:
                name = short_child
            elif short_parent:
                name = short_parent
            else:
                # Gunakan nama sementara untuk header yang kosong.
                name = f"Column_{index}"

            # Nama kolom yang dimulai dengan angka diberi awalan agar
            # lebih konsisten ketika digunakan sebagai nama field.
            if name[0].isdigit():
                name = f"col_{name}"

            new_headers.append(name)

        # Beri akhiran pada nama yang berulang agar setiap kolom unik.
        new_headers = make_unique_columns(new_headers)

        # Simpan style header sebelum baris kedua dihapus.
        # Style subheader diprioritaskan jika memiliki fill.
        header_styles = []

        for col in range(1, max_col + 1):
            parent_cell = ws.cell(1, col)
            child_cell = ws.cell(2, col)

            if child_cell.fill.fill_type:
                header_styles.append(copy(child_cell._style))
            else:
                header_styles.append(copy(parent_cell._style))

        # Lepaskan merge agar baris header dan kolom dapat diubah.
        # Merge pada area data dicatat terlebih dahulu agar dapat
        # dipulihkan setelah penghapusan baris dan kolom.
        data_merges = []

        for merged_range in list(ws.merged_cells.ranges):
            if merged_range.min_row <= 2:
                # Merge yang menyentuh dua baris header tidak dipulihkan
                # karena header akan diganti menjadi satu baris.
                ws.unmerge_cells(str(merged_range))
            else:
                # Baris data bergeser satu posisi ke atas setelah
                # baris header kedua dihapus.
                data_merges.append((
                    merged_range.min_row - 1,
                    merged_range.min_col,
                    merged_range.max_row - 1,
                    merged_range.max_col,
                ))
                ws.unmerge_cells(str(merged_range))

        # Tulis nama header gabungan dan terapkan kembali style-nya.
        for col, name in enumerate(new_headers, start=1):
            cell = ws.cell(1, col)
            cell.value = name
            cell._style = copy(header_styles[col - 1])

        # Hapus baris kedua karena header telah digabungkan.
        ws.delete_rows(2, 1)

        # Hapus kolom nomor urut dan kolom pemisah yang berdiri sendiri.
        # Penghapusan dilakukan dari kanan ke kiri agar indeks kolom
        # yang belum diproses tidak berubah akibat pergeseran.
        columns_to_remove = [
            index
            for index, name in enumerate(new_headers, start=1)
            if name.strip().casefold() in {"no", "no.", "-"}
        ]

        for col_idx in sorted(columns_to_remove, reverse=True):
            ws.delete_cols(col_idx, 1)

        # Sinkronkan daftar header dengan kolom yang masih tersisa.
        removed_set = set(columns_to_remove)

        new_headers = [
            name
            for index, name in enumerate(new_headers, start=1)
            if index not in removed_set
        ]

        # Pulihkan merge pada area data selama merge tersebut tidak
        # bersinggungan dengan kolom yang dihapus.
        for min_row, min_col, max_row, max_col_merge in data_merges:
            if any(
                min_col <= col <= max_col_merge
                for col in columns_to_remove
            ):
                continue

            # Hitung jumlah kolom yang dihapus di sebelah kiri merge.
            # Posisi merge digeser ke kiri sesuai jumlah tersebut.
            shift = sum(
                col < min_col
                for col in columns_to_remove
            )

            ws.merge_cells(
                start_row=min_row,
                start_column=min_col - shift,
                end_row=max_row,
                end_column=max_col_merge - shift,
            )

        durasi["Cleaning header"] = time.perf_counter() - t0

        # -----------------------------------------------------
        # Tahap 3: Validasi dan penyisipan kolom wajib
        # -----------------------------------------------------

        t0 = time.perf_counter()

        # Memastikan Kode Provinsi dan kedua kolom RTT tersedia pada
        # posisi standar. Kolom yang hilang akan disisipkan sesuai aturan.
        new_headers = validasi_dan_sisipkan_kolom(
            ws, new_headers
        )

        durasi["Validasi dan sisip kolom"] = (
            time.perf_counter() - t0
        )

        # -----------------------------------------------------
        # Tahap 4: Pembuatan schema normalisasi
        # -----------------------------------------------------

        t0 = time.perf_counter()

        # Schema dibuat dari nama header final dan digunakan oleh modul
        # normalisasi untuk menentukan tipe data masing-masing kolom.
        schema = buat_schema_dari_header(new_headers)

        durasi["Pembuatan schema"] = time.perf_counter() - t0

        # -----------------------------------------------------
        # Tahap 5: Normalisasi tipe data
        # -----------------------------------------------------

        t0 = time.perf_counter()

        # Proses normalisasi dilakukan di modul terpisah agar aturan
        # tipe data tidak bercampur dengan logika pembersihan header.
        hasil = normalisasi_sheet(ws, schema)

        durasi["Normalisasi data"] = time.perf_counter() - t0

        # -----------------------------------------------------
        # Tahap 6: Menyimpan hasil cleaning
        # -----------------------------------------------------

        t0 = time.perf_counter()

        # Buat file output baru agar file sumber tidak tertimpa.
        # delete=False memungkinkan file digunakan setelah blok selesai.
        with tempfile.NamedTemporaryFile(
            prefix="Compile_Summary_Cleaned_",
            suffix=".xlsx",
            delete=False,
        ) as temp_file:
            output_file = temp_file.name

        wb.save(output_file)

        durasi["Penyimpanan file"] = time.perf_counter() - t0

        # Total runtime mencakup proses backend yang diukur di fungsi ini,
        # dari awal pemrosesan hingga file output selesai disimpan.
        runtime_total = time.perf_counter() - waktu_mulai

        # -----------------------------------------------------
        # Tahap 7: Menyusun informasi hasil dan runtime
        # -----------------------------------------------------

        rincian_runtime = "\n".join(
            f"• {nama}: {detik:.2f} detik"
            for nama, detik in durasi.items()
        )

        message = (
            "✅ Compile Summary berhasil diproses!\n\n"
            f"Jumlah baris data: {hasil['baris']}\n"
            f"Jumlah kolom hasil: {hasil['kolom']}\n\n"
            "⏱️ RUNTIME PEMROSESAN\n"
            f"Total runtime backend: {runtime_total:.2f} detik\n"
            f"{rincian_runtime}\n\n"
            "✔ Header digabungkan dan disingkat.\n"
            "✔ Kolom No, No., dan - dihapus.\n"
            "✔ Kolom wajib divalidasi dan disisipkan jika hilang.\n"
            "✔ Tipe data dinormalisasi berdasarkan nama kolom.\n"
            "✔ No Data pada kolom numerik diubah menjadi 0.\n"
        )

        return output_file, message

    except Exception as exc:
        # Jika terjadi kesalahan, tampilkan penyebab dan waktu yang telah
        # berlalu hingga kegagalan terjadi.
        runtime_gagal = time.perf_counter() - waktu_mulai

        return None, (
            f"❌ Gagal memproses Compile Summary: {exc}\n\n"
            f"Waktu hingga gagal: {runtime_gagal:.2f} detik"
        )

    finally:
        # Menutup workbook untuk melepaskan resource file setelah proses
        # selesai, baik berhasil maupun gagal.
        if wb is not None:
            wb.close()
