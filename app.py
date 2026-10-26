
"""
Antarmuka utama aplikasi Data Cleaning - Balmon Aceh.

Deskripsi:
    Aplikasi berbasis Gradio untuk membantu proses pembersihan data
    Speed Testing dan Compile Summary.

Fitur utama:
    1. Speed Testing
       - Menerima file Excel mentah.
       - Memproses data melalui fungsi process_file().
       - Menyediakan hasil cleaning dalam format CSV.

    2. Compile Summary
       - Menerima file Excel dengan dua baris header.
       - Memproses data melalui fungsi process_summary().
       - Menyediakan hasil cleaning dalam format Excel (.xlsx).

Struktur program:
    - Import library dan fungsi pemrosesan.
    - Konfigurasi CSS untuk tampilan antarmuka.
    - Pembuatan komponen aplikasi menggunakan Gradio Blocks.
    - Pengaturan event tombol untuk menjalankan proses cleaning.
    - Peluncuran aplikasi.

Catatan pemeliharaan:
    - Proses cleaning Speed Testing berada di speed_testing.py.
    - Proses cleaning Compile Summary berada di compile_summary.py.
    - Perubahan logika cleaning sebaiknya dilakukan pada modul
      pemrosesan terkait, bukan langsung pada antarmuka ini.
    - Pastikan nama fungsi, urutan input, dan output pada event
      tombol tetap sesuai dengan fungsi pemrosesan masing-masing.
    - File keluaran harus sesuai dengan format yang ditentukan
      pada komponen download.
"""

import gradio as gr

from speed_testing import process_file
from compile_summary import process_summary


# =========================================================
# KONFIGURASI TAMPILAN
# =========================================================
# CSS digunakan untuk mengatur lebar konten, perataan judul,
# warna teks pendukung, dan tampilan footer.
# Pengaturan ini hanya memengaruhi tampilan antarmuka,
# bukan proses pembersihan data.

CSS = """
.gradio-container {
    max-width: 1050px !important;
    margin: auto;
}

.title {
    text-align: center;
    margin-bottom: 5px;
}

.subtitle {
    text-align: center;
    color: #6b7280;
    margin-bottom: 25px;
}

.footer {
    text-align: center;
    color: #9ca3af;
    font-size: 12px;
    margin-top: 25px;
}
"""


# =========================================================
# PEMBUATAN ANTARMUKA UTAMA
# =========================================================
# gr.Blocks digunakan sebagai wadah utama aplikasi.
# Judul yang ditampilkan pada browser adalah
# "Data Cleaning - Balmon Aceh".

with gr.Blocks(
    title="Data Cleaning - Balmon Aceh",
) as demo:

    # -----------------------------------------------------
    # HEADER APLIKASI
    # -----------------------------------------------------
    # Menampilkan judul dan deskripsi singkat aplikasi.

    gr.HTML(
        """
        <div class="title">
            <h1>Data Cleaning</h1>
        </div>
        <div class="subtitle">
            Pengolahan data Speed Testing dan Compile Summary
        </div>
        """
    )

    # =====================================================
    # TAB 1: SPEED TESTING
    # =====================================================
    # Tab ini digunakan untuk memproses file Excel mentah
    # Speed Testing menjadi file CSV hasil cleaning.
    #
    # Fungsi pemrosesan:
    #     process_file() dari speed_testing.py
    #
    # Input:
    #     File Excel berformat .xlsx, .xlsm, atau .xls.
    #
    # Output:
    #     File CSV dan informasi status pemrosesan.
    #
    # Catatan:
    #     Detail pembacaan sheet, normalisasi data, dan
    #     pembuatan CSV ditangani oleh process_file().

    with gr.Tab("Speed Testing"):

        gr.Markdown(
            """
            ### Cleaning Data Speed Testing

            Upload file Excel mentah. Sistem membaca sheet
            `Speed_Testing` dengan header pada baris kedua,
            lalu menghasilkan file CSV.
            """
        )

        with gr.Row():

            # Area input file dan tombol pemrosesan.
            with gr.Column():

                speed_input = gr.File(
                    label="Upload File Raw Speed Testing",
                    file_types=[".xlsx", ".xlsm", ".xls"],
                    type="filepath",
                )

                speed_button = gr.Button(
                    "⚙️ Proses Speed Testing",
                    variant="primary",
                )

            # Area informasi status dan unduhan hasil.
            with gr.Column():

                speed_status = gr.Textbox(
                    label="Status Speed Testing",
                    placeholder="Hasil proses akan muncul di sini...",
                    lines=7,
                    interactive=False,
                )

                speed_download = gr.File(
                    label="Download CSV Speed Testing",
                    file_types=[".csv"],
                )

        # Menghubungkan tombol dengan fungsi pemrosesan.
        #
        # inputs:
        #     speed_input berisi path file yang diunggah.
        #
        # outputs:
        #     speed_download menerima file hasil.
        #     speed_status menerima informasi proses.
        #
        # Urutan output harus sesuai dengan nilai yang
        # dikembalikan oleh process_file().

        speed_button.click(
            fn=process_file,
            inputs=[speed_input],
            outputs=[speed_download, speed_status],
        )

    # =====================================================
    # TAB 2: COMPILE SUMMARY
    # =====================================================
    # Tab ini digunakan untuk memproses file Excel Compile
    # Summary dan menghasilkan file Excel hasil cleaning.
    #
    # Fungsi pemrosesan:
    #     process_summary() dari compile_summary.py
    #
    # Input:
    #     File Excel berformat .xlsx atau .xlsm.
    #
    # Output:
    #     File Excel .xlsx dan informasi status pemrosesan.
    #
    # Catatan:
    #     Penggabungan header, normalisasi tipe data, serta
    #     penanganan kolom dilakukan pada modul pemrosesan.
    #     Antarmuka ini hanya menerima input dan menampilkan
    #     hasil yang dikembalikan oleh process_summary().

    with gr.Tab("Compile Summary"):

        gr.Markdown(
            """
            ### Cleaning Data Compile Summary

            Upload file Excel yang memiliki dua baris header.

            Sistem akan:
            - Menggabungkan dan menyingkat nama header.
            - Menghapus kolom `No`, `No.`, dan `-` yang berdiri sendiri.
            - Menormalisasi tipe data berdasarkan nama kolom.
            - Mengubah `No Data` menjadi `0` pada kolom numerik.
            - Mempertahankan seluruh baris data tanpa filtering.

            Output berupa file Excel `.xlsx`, tanpa schema JSON.
            """
        )

        with gr.Row():

            # Area input file dan tombol pemrosesan.
            with gr.Column():

                summary_input = gr.File(
                    label="Upload File Compile Summary",
                    file_types=[".xlsx", ".xlsm"],
                    type="filepath",
                )

                summary_button = gr.Button(
                    "⚙️ Proses Compile Summary",
                    variant="primary",
                )

            # Area informasi status dan unduhan hasil.
            with gr.Column():

                summary_status = gr.Textbox(
                    label="Status Compile Summary",
                    placeholder="Hasil proses akan muncul di sini...",
                    lines=9,
                    interactive=False,
                )

                summary_download = gr.File(
                    label="Download Excel Hasil Cleaning",
                    file_types=[".xlsx"],
                )

        # Menghubungkan tombol dengan fungsi pemrosesan.
        #
        # inputs:
        #     summary_input berisi path file yang diunggah.
        #
        # outputs:
        #     summary_download menerima file hasil.
        #     summary_status menerima informasi proses.
        #
        # Urutan output harus sesuai dengan nilai yang
        # dikembalikan oleh process_summary().

        summary_button.click(
            fn=process_summary,
            inputs=[summary_input],
            outputs=[summary_download, summary_status],
        )

    # =====================================================
    # FOOTER APLIKASI
    # =====================================================
    # Menampilkan identitas instansi pada bagian bawah
    # antarmuka. Elemen ini hanya bersifat informatif.

    gr.HTML(
        """
        <div class="footer">
            Data Cleaning • Balai Monitor Spektrum Frekuensi Radio & Infrastruktur Digital Kelas I Banda Aceh
        </div>
        """
    )


# =========================================================
# TITIK MASUK PROGRAM
# =========================================================
# Blok ini memastikan aplikasi hanya diluncurkan ketika
# file app.py dijalankan secara langsung.
#
# Aplikasi menggunakan CSS yang didefinisikan di atas.
# Konfigurasi port dan host tetap mengikuti pengaturan
# default Gradio karena tidak ditentukan secara eksplisit.
#
# Contoh menjalankan aplikasi:
#     python app.py

if __name__ == "__main__":
    demo.launch(css=CSS)
