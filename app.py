import streamlit as st
import pandas as pd
import io
import zipfile
from openpyxl import load_workbook

st.set_page_config(page_title="Generator Template SPD Honorium", page_icon="📄")

st.title("📄 Generator Template SPD Honorium")

# Upload database dan template
db_file = st.file_uploader("Upload Database (Excel/CSV)", type=["xlsx", "csv"])
template_file = st.file_uploader("Upload Template SPD", type=["xlsx"])

if db_file and template_file:
    # Baca database
    if db_file.name.endswith(".csv"):
        df = pd.read_csv(db_file)
    else:
        df = pd.read_excel(db_file)

    # Kolom wajib
    required_cols = [
        "Nama",
        "Honorarium Persiapan UKOMNAS",
        "Honorarium Pemantauan Briefing UKOMNAS",
        "Honorarium Pelaksanaan UKOMNAS"
    ]

    if not all(col in df.columns for col in required_cols):
        st.error(f"Kolom database harus mengandung: {', '.join(required_cols)}")
    else:
        st.subheader("Preview Database")
        st.dataframe(df.head())

        # -----------------------------
        # DOWNLOAD PER NAMA
        # -----------------------------
        nama_terpilih = st.selectbox("Pilih Nama", df["Nama"].unique())

        if st.button("🔄 Generate Template (Satu Nama)"):
            row = df[df["Nama"] == nama_terpilih].iloc[0]
            wb = load_workbook(template_file)
            ws = wb.active

            # Isi data honor
            ws["D26"] = row["Nama"]
            ws["C11"] = row["Honorarium Persiapan UKOMNAS"]
            ws["C12"] = row["Honorarium Pemantauan Briefing UKOMNAS"]
            ws["C13"] = row["Honorarium Pelaksanaan UKOMNAS"]

            # Hitung total honor
            total_honor = sum([
                row["Honorarium Persiapan UKOMNAS"],
                row["Honorarium Pemantauan Briefing UKOMNAS"],
                row["Honorarium Pelaksanaan UKOMNAS"]
            ])
            ws["C14"] = total_honor

            # Ambil PPH21 dari C15 di template (jika kosong = 0)
            pph21 = ws["C16"].value or 0

            # Hitung jumlah akhir
            ws["C18"] = total_honor - pph21

            buffer = io.BytesIO()
            wb.save(buffer)
            buffer.seek(0)

            st.success(f"Template untuk {nama_terpilih} berhasil dibuat!")
            st.download_button(
                label="📥 Download File",
                data=buffer,
                file_name=f"Template_{nama_terpilih}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # -----------------------------
        # DOWNLOAD SEMUA NAMA (ZIP)
        # -----------------------------
        if st.button("📦 Generate Semua Template (ZIP)"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for _, row in df.iterrows():
                    wb = load_workbook(template_file)
                    ws = wb.active

                    ws["D26"] = row["Nama"]
                    ws["C11"] = row["Honorarium Persiapan UKOMNAS"]
                    ws["C12"] = row["Honorarium Pemantauan Briefing UKOMNAS"]
                    ws["C13"] = row["Honorarium Pelaksanaan UKOMNAS"]

                    total_honor = sum([
                        row["Honorarium Persiapan UKOMNAS"],
                        row["Honorarium Pemantauan Briefing UKOMNAS"],
                        row["Honorarium Pelaksanaan UKOMNAS"]
                    ])
                    ws["C14"] = total_honor

                    pph21 = ws["C16"].value or 0
                    ws["C18"] = total_honor - pph21

                    file_buffer = io.BytesIO()
                    wb.save(file_buffer)
                    file_buffer.seek(0)

                    zf.writestr(f"Template_{row['Nama']}.xlsx", file_buffer.read())

            zip_buffer.seek(0)
            st.success("Semua template berhasil dibuat!")
            st.download_button(
                label="📥 Download ZIP Semua Template",
                data=zip_buffer,
                file_name="Semua_Template_SPD.zip",
                mime="application/zip"
            )
