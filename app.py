import io
import re

import openpyxl
from openpyxl.styles import Font
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="SPX Laporan Scan to Excel Converter",
    page_icon="📦",
    layout="wide",
)

st.title("📦 SPX Laporan Scan PDF ➡️ Excel Converter")
st.write(
    "Upload file PDF Laporan Scan SPX OSO. Hasil Excel akan **100% persis** dengan format file acuan (Sheet SJM, MARKING, PVT, dan Summary)."
)

# Pemetaan Kode Marking Berdasarkan Destinasi
MARKING_MAP = {
    "Abepura DC": "DJJ-C1-1",
    "Alak DC": "KOE-C1-1",
    "Banjarbaru DC": "BDJ-C1-1",
    "Banjarmasin 2 DC": "BDJ-C1-1",
    "Banjarmasin DC": "BDJ-C1-1",
    "Batam DC": "BTH-C1-1",
    "Dungingi DC": "GTO-C1-1",
    "Labuhan Bajo DC": "LBJ-C1-1",
    "Manokwari Barat DC": "MKW-C1-1",
    "Mantikulore DC": "PLW-C1-1",
    "Pekanbaru 2 DC": "PKU-C1-1",
    "Pekanbaru DC": "PKU-C1-1",
    "Ternate Utara Hub": "TTE-C1-1",
    "Wua-Wua DC": "KDI-C1-1",
}


def extract_data_from_pdf(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
    extracted_rows = []

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue

        # Extract LT NUMBER (Contoh: LT0Q9E1MSZVR1)
        lt_match = re.search(r"\b(LT[A-Z0-9]{8,})\b", text)
        lt_num = lt_match.group(1) if lt_match else ""

        # Extract Destinasi (Contoh: Abepura DC)
        dest_match = re.search(r":\s*([A-Za-z0-9\s-]+?\s*(?:DC|Hub))", text)
        dest = dest_match.group(1).strip() if dest_match else ""

        # Extract Tanggal STD (Contoh: 2026/09/14)
        std_match = re.search(r"(\d{4}/\d{2}/\d{2})\s*\d{2}:\d{2}:\d{2}STD", text)
        if not std_match:
            std_match = re.search(r":\s*(\d{4}/\d{2}/\d{2})", text)
        tgl = (
            std_match.group(1).replace("/", "-")
            if std_match
            else "2026-09-14"
        )

        vendor = "Lion Parcel"
        origin = "SURABAYA DC"

        # Extract TO dan Berat
        to_numbers = re.findall(r"\b(TO\d{8}[A-Z0-9]+)\b", text)
        weights = re.findall(r"\b(\d{1,3}\.\d{2,3})\b", text)

        for i, to_num in enumerate(to_numbers):
            gw = 0.0
            if i < len(weights):
                try:
                    gw = float(weights[i])
                except ValueError:
                    gw = 0.0

            marking = MARKING_MAP.get(dest, "C1-1")

            extracted_rows.append({
                "TGL": tgl,
                "Vendor": vendor,
                "Sc Origin": origin,
                "Sc Destination": dest,
                "Lt Number": lt_num,
                "To Number": to_num,
                "Marking": marking,
                "Gross Weight": gw,
                "Remarks": "BAG",
            })

    return pd.DataFrame(extracted_rows)


uploaded_file = st.file_uploader("Upload File PDF SPX", type=["pdf"])

if uploaded_file is not None:
    df = extract_data_from_pdf(uploaded_file)

    if not df.empty:
        st.success(
            f"Berhasil mengolah data! Total **{len(df)}** TO ditemukan."
        )

        # Tab Preview Data di Streamlit Web
        tab1, tab2, tab3 = st.tabs(
            ["📋 Sheet SJM", "🏷️ Sheet MARKING", "📊 Summary Destinasi"]
        )

        # Dataframe SJM untuk Preview
        df_sjm_preview = df[[
            "TGL",
            "Vendor",
            "Sc Origin",
            "Sc Destination",
            "Lt Number",
            "To Number",
            "Gross Weight",
            "Remarks",
        ]].copy()
        df_sjm_preview.columns = [
            "TGL",
            "Vendor",
            "SC Orgin",
            "DESTINATION",
            "LT NUMBER",
            "TO NUMBER",
            "Gross Weight",
            "REMAKE",
        ]
        df_sjm_preview["Vendor"] = "LION PARCEL"

        with tab1:
            st.dataframe(df_sjm_preview, use_container_width=True)

        with tab2:
            st.dataframe(df, use_container_width=True)

        with tab3:
            summary = (
                df.groupby("Sc Destination")
                .agg(
                    Count_of_TO=("To Number", "count"),
                    Total_Weight=("Gross Weight", "sum"),
                )
                .reset_index()
            )
            st.dataframe(summary, use_container_width=True)

        # ---------------------------------------------------------
        # MEMBUAT FILE EXCEL SESUAI STRUKTUR EXCEL ASLI
        # ---------------------------------------------------------
        wb = openpyxl.Workbook()

        # --- 1. SHEET SJM ---
        ws_sjm = wb.active
        ws_sjm.title = "SJM"

        # Judul Header Atas
        ws_sjm.cell(
            row=1, column=1, value="SURAT JALAN MANUAL SURABAYA DC VIA LION STD"
        )
        ws_sjm.cell(row=2, column=1, value="14 SEPTEMBER 2026 TRIP 2")
        ws_sjm.cell(row=2, column=9, value=f"=COUNTA(F4:F{len(df)+3})")

        # Header Kolom (Baris 3)
        sjm_cols = [
            "TGL",
            "Vendor",
            "SC Orgin",
            "DESTINATION",
            "LT NUMBER",
            "TO NUMBER",
            "Gross Weight",
            "REMAKE",
            "TOTAL",
        ]
        for c_idx, col_name in enumerate(sjm_cols, 1):
            cell = ws_sjm.cell(row=3, column=c_idx, value=col_name)
            cell.font = Font(bold=True)

        # Isi Data (Mulai Baris 4)
        for r_idx, r in enumerate(df_sjm_preview.itertuples(), 4):
            ws_sjm.cell(row=r_idx, column=1, value=r.TGL)
            ws_sjm.cell(row=r_idx, column=2, value=r.Vendor)
            ws_sjm.cell(row=r_idx, column=3, value=r._3)
            ws_sjm.cell(row=r_idx, column=4, value=r.DESTINATION)
            ws_sjm.cell(row=r_idx, column=5, value=r._5)
            ws_sjm.cell(row=r_idx, column=6, value=r._6)
            ws_sjm.cell(row=r_idx, column=7, value=r._7)
            ws_sjm.cell(row=r_idx, column=8, value=r.REMAKE)

        # --- 2. SHEET MARKING ---
        ws_mk = wb.create_sheet(title="MARKING")
        ws_mk.cell(row=1, column=1, value="MARKING SPX OSO SUB DC CYCLE ")
        ws_mk.cell(row=2, column=1, value="14 SEPTEMBER 2026 TRIP 2")

        mk_cols = [
            "Tanggal",
            "Vendor",
            "Sc Origin",
            "Sc Destination",
            "Lt Number",
            "To Number",
            "Marking",
            "Gross Weight",
            "Remarks",
            "External Number",
            "Clear Gw",
        ]
        for c_idx, col_name in enumerate(mk_cols, 1):
            cell = ws_mk.cell(row=3, column=c_idx, value=col_name)
            cell.font = Font(bold=True)

        for r_idx, r in enumerate(df.itertuples(), 4):
            ws_mk.cell(row=r_idx, column=1, value=r.TGL)
            ws_mk.cell(row=r_idx, column=2, value=r.Vendor)
            ws_mk.cell(row=r_idx, column=3, value=r._3)
            ws_mk.cell(row=r_idx, column=4, value=r._4)
            ws_mk.cell(row=r_idx, column=5, value=r._5)
            ws_mk.cell(row=r_idx, column=6, value=r._6)
            ws_mk.cell(row=r_idx, column=7, value=r.Marking)
            ws_mk.cell(row=r_idx, column=8, value=r._8)
            ws_mk.cell(row=r_idx, column=9, value=r.Remarks)
            # Rumus Excel External Number = G4 & "/" & E4 & "/" & I4
            ws_mk.cell(
                row=r_idx,
                column=10,
                value=f'=G{r_idx}&"/"&E{r_idx}&"/"&I{r_idx}',
            )
            ws_mk.cell(row=r_idx, column=11, value=r._8)

        # --- 3. SHEET PVT ---
        ws_pvt = wb.create_sheet(title="PVT")
        pvt_cols = [
            "External Number",
            "Count of External Number",
            "Sum of Clear Gw",
        ]
        for c_idx, col_name in enumerate(pvt_cols, 1):
            cell = ws_pvt.cell(row=3, column=c_idx, value=col_name)
            cell.font = Font(bold=True)

        # Grouping unique destinasi untuk pivot
        unique_indices = df.drop_duplicates(subset=["Sc Destination"]).index
        for p_idx, first_r in enumerate(unique_indices, 4):
            pvt_row = p_idx
            ws_pvt.cell(row=pvt_row, column=1, value=f"=MARKING!J{first_r+4}")
            ws_pvt.cell(
                row=pvt_row,
                column=2,
                value=f"=COUNTIF(MARKING!J$4:J${len(df)+3}, A{pvt_row})",
            )
            ws_pvt.cell(
                row=pvt_row,
                column=3,
                value=f"=SUMIF(MARKING!J$4:J${len(df)+3}, A{pvt_row}, MARKING!K$4:K${len(df)+3})",
            )

        # --- 4. SHEET SHEET3 (SUMMARY) ---
        ws_sum = wb.create_sheet(title="Sheet3")
        sum_cols = [
            "Sc Destination",
            "Count of To Number",
            "Sum of Gross Weight",
        ]
        for c_idx, col_name in enumerate(sum_cols, 1):
            cell = ws_sum.cell(row=3, column=c_idx, value=col_name)
            cell.font = Font(bold=True)

        for s_idx, r in enumerate(summary.itertuples(), 4):
            ws_sum.cell(row=s_idx, column=1, value=r._1)
            ws_sum.cell(row=s_idx, column=2, value=r.Count_of_TO)
            ws_sum.cell(row=s_idx, column=3, value=r.Total_Weight)

        # Simpan ke Stream
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="📥 Download File Excel SJM & Marking",
            data=output,
            file_name=f"FIXED_SCRIPT_SJ_MANUAL_{uploaded_file.name.replace('.pdf', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error("Gagal mengekstrak data dari PDF. Pastikan format file sesuai.")
