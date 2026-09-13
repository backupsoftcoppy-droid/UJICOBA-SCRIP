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

        # Extract Header Metadata
        origin_match = re.search(r"Origin\s*:\s*(.*)", text)
        dest_match = re.search(r"Destination\s*:\s*(.*)", text)
        lt_match = re.search(r"Surat Jalan Line Haul\s*\n\s*([A-Z0-9]+)", text)
        std_match = re.search(r"STD\s*\([^)]*\)\s*:\s*(\d{4}/\d{2}/\d{2})", text)
        driver_match = re.search(r"PIC Gudang Origin\s*:\s*(.*)", text)

        origin = origin_match.group(1).strip() if origin_match else "Surabaya DC"
        dest = dest_match.group(1).strip() if dest_match else ""
        lt_num = lt_match.group(1).strip() if lt_match else ""
        tgl = std_match.group(1).replace("/", "-") if std_match else ""
        vendor = (
            driver_match.group(1).strip()
            if driver_match
            else "LINEHAUL LION PARCEL"
        )
        if "LION PARCEL" in vendor.upper():
            vendor = "Lion Parcel"

        # Extract TO Details from tables/text blocks
        # Pola Pencarian Nomor TO (Format: TO followed by digits/letters)
        to_matches = re.findall(
            r"(TO\d{8}[A-Z0-9]+)\s+([\d\.]+)\s+(BAG|BULKY|BOX)?", text
        )

        for match in to_matches:
            to_num = match[0]
            gross_weight = float(match[1]) if match[1] else 0.0
            remake = match[2] if match[2] else "BAG"

            extracted_data.append({
                "TGL": tgl,
                "Vendor": vendor,
                "Sc Origin": origin,
                "Sc Destination": dest,
                "Lt Number": lt_num,
                "To Number": to_num,
                "Gross Weight": gross_weight,
                "Remarks": remake,
            })

    return pd.DataFrame(extracted_data)


if uploaded_file is not None:
    with st.spinner("Memproses dan membaca file PDF..."):
        df_result = extract_data_from_pdf(uploaded_file)

    if not df_result.empty:
        st.success(
            f"Berhasil mengolah data! Total **{len(df_result)}** TO ditemukan."
        )

        # Preview Data Tab
        tab1, tab2 = st.tabs(["📋 Preview Data SJM / Marking", "📊 Summary"])

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

        # Export to Excel Stream
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_result.to_excel(writer, sheet_name="MARKING", index=False)

            # Buat Sheet SJM (Format Kolom SJM)
            df_sjm = df_result.rename(columns={
                "Sc Origin": "SC Orgin",
                "Sc Destination": "DESTINATION",
                "Lt Number": "LT NUMBER",
                "To Number": "TO NUMBER",
                "Remarks": "REMAKE",
            })
            df_sjm.to_excel(writer, sheet_name="SJM", index=False)

            # Summary Sheet
            summary.to_excel(writer, sheet_name="SUMMARY", index=False)

        output.seek(0)

        st.download_button(
            label="📥 Download File Excel",
            data=output,
            file_name=f"FIXED_SCRIPT_SJ_MANUAL_{uploaded_file.name.replace('.pdf', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error(
            "Tidak ada data TO yang berhasil diekstrak. Pastikan format PDF sesuai."
        )
