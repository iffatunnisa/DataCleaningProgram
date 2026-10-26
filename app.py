
"""Antarmuka utama Data Cleaning - Balmon Aceh."""

import gradio as gr

from speed_testing import process_file
from compile_summary import process_summary


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


with gr.Blocks(
    title="Data Cleaning - Balmon Aceh",
) as demo:

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

        speed_button.click(
            fn=process_file,
            inputs=[speed_input],
            outputs=[speed_download, speed_status],
        )

    # =====================================================
    # TAB 2: COMPILE SUMMARY
    # =====================================================
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

        summary_button.click(
            fn=process_summary,
            inputs=[summary_input],
            outputs=[summary_download, summary_status],
        )

    # =====================================================
    # FOOTER
    # =====================================================
    gr.HTML(
        """
        <div class="footer">
            Data Cleaning • Balai Monitor Spektrum Frekuensi Radio & Infrastruktur Digital Kelas I Banda Aceh
        </div>
        """
    )


if __name__ == "__main__":
    demo.launch(css=CSS)
