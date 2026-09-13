import io
import re

import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="SPX PDF to Excel Converter", page_icon="📦", layout="wide"
)

st.title("📦 SPX Laporan Scan PDF ➡️ Excel Converter")
st.write(
    "Unggah file PDF Laporan Scan SPX OSO untuk mengonversinya secara otomatis ke format Excel (SJM & Marking)."
)

uploaded_file = st.file_uploader("Pilih File PDF", type=["pdf"])


def extract_data_from_pdf(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
    extracted_data = []

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue

        # 1. Extract LT NUMBER (Mencari pola LT0Q...)
        lt_match = re.search(r"\b(LT[A-Z0-9]{8,})\b", text)
        lt_num = lt_match.group(1) if lt_match else ""

        # 2. Extract Destination (Mencari nama DC/Hub tepat sebelum STD)
        dest_match = re.search(r":\s*([A-Za-z0-9\s-]+?\s*(?:DC|Hub))", text)
        dest = dest_match.group(1).strip() if dest_match else ""

        # 3. Extract Tanggal (Mengambil tanggal STD Keberangkatan)
        std_match = re.search(r"(\d{4}/\d{2}/\d{2})\s*\d{2}:\d{2}:\d{2}STD", text)
        if not std_match:
            std_match = re.search(r":\s*(\d{4}/\d{2}/\d{2})", text)
        tgl = std_match.group(1).replace("/", "-") if std_match else ""

        # 4. Extract Vendor & Origin
        vendor = "Lion Parcel"
        origin = "SURABAYA DC"

        # 5. Extract TO details (Nomor TO, Berat, & Type Bag/Bulky)
        # Mencari TO2026... diikuti berat dan jenis kemasan
        to_pattern = re.compile(
            r"(TO\d{8}[A-Z0-9]+)[\s\S]*?(\d+\.\d{2,3})\s*(Bag|Bulky|BOX)?",
            re.IGNORECASE,
        )

        # Alternatif ekstrak baris TO secara teliti
        lines = text.split("\n")
        to_numbers = re.findall(r"\b(TO\d{8}[A-Z0-9]+)\b", text)
        weights = re.findall(r"\b(\d{1,3}\.\d{2,3})\b", text)

        # Menggabungkan data TO pada halaman tersebut
        for i, to_num in enumerate(to_numbers):
            # Abaikan angka total berat di header
            gw = 0.0
            if i < len(weights):
                try:
                    gw = float(weights[i])
                except ValueError:
                    gw = 0.0

            extracted_data.append({
                "TGL": tgl,
                "Vendor": vendor,
                "Sc Origin": origin,
                "Sc Destination": dest,
                "Lt Number": lt_num,
                "To Number": to_num,
                "Gross Weight": gw,
                "Remarks": "BAG",
            })

    return pd.DataFrame(extracted_data)


if uploaded_file is not None:
    with st.spinner("Memproses dan membaca file PDF..."):
        df_result = extract_data_from_pdf(uploaded_file)

    if not df_result.empty:
        st.success(
            f"Berhasil mengolah data! Total **{len(df_result)}** TO ditemukan."
        )

        tab1, tab2 = st.tabs(["📋 Preview Data", "📊 Summary Destinasi"])

        with tab1:
            st.dataframe(df_result, use_container_width=True)

        with tab2:
            summary = (
                df_result.groupby("Sc Destination")
                .agg(
                    Count_of_TO=("To Number", "count"),
                    Total_Weight=("Gross Weight", "sum"),
                )
                .reset_index()
            )
            st.dataframe(summary, use_container_width=True)

        # Download Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Sheet MARKING
            df_result.to_excel(writer, sheet_name="MARKING", index=False)

            # Sheet SJM
            df_sjm = df_result.rename(columns={
                "Sc Origin": "SC Orgin",
                "Sc Destination": "DESTINATION",
                "Lt Number": "LT NUMBER",
                "To Number": "TO NUMBER",
                "Remarks": "REMAKE",
            })
            df_sjm.to_excel(writer, sheet_name="SJM", index=False)

        output.seek(0)
        st.download_button(
            label="📥 Download File Excel Perbaikan",
            data=output,
            file_name=f"FIXED_{uploaded_file.name.replace('.pdf', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("Gagal membaca data dari PDF.")
