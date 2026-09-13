import io
import re

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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

# Peta Kode Marking Berdasarkan Destinasi
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


def extract_data(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
    data_list = []

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue

        # Extract Header
        lt_match = re.search(r"\b(LT[A-Z0-9]{8,})\b", text)
        lt_num = lt_match.group(1) if lt_match else ""

        dest_match = re.search(r":\s*([A-Za-z0-9\s-]+?\s*(?:DC|Hub))", text)
        dest = dest_match.group(1).strip() if dest_match else ""

        std_match = re.search(r"(\d{4}/\d{2}/\d{2})\s*\d{2}:\d{2}:\d{2}STD", text)
        tgl = (
            std_match.group(1).replace("/", "-")
            if std_match
            else "2026-09-14"
        )

        vendor_match = re.search(r"Nama Vendor\s*:\s*(.*)", text)
        vendor = (
            vendor_match.group(1).strip() if vendor_match else "Lion Parcel"
        )
        if "LION" in vendor.upper():
            vendor = "Lion Parcel"

        origin = "SURABAYA DC"

        # Extract TO & Weights
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

            data_list.append({
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

    return pd.DataFrame(data_list)


uploaded_file = st.file_uploader("Upload File PDF SPX", type=["pdf"])

if uploaded_file is not None:
    df = extract_data(uploaded_file)

    if not df.empty:
        st.success(
            f"Berhasil membaca PDF! Total **{len(df)}** data TO ditemukan."
        )

        # Build Workbook via OpenPyXL to match EXACT structure & formulas
        wb = openpyxl.Workbook()

        # ---------------------------------------------------------
        # 1. SHEET: SJM
        # ---------------------------------------------------------
        ws_sjm = wb.active
        ws_sjm.title = "SJM"

        ws_sjm["A1"] = "SURAT JALAN MANUAL SURABAYA DC VIA LION STD"
        ws_sjm["A2"] = "14 SEPTEMBER 2026 TRIP 2"
        ws_sjm["I2"] = "=COUNTA(F4:F1000)"

        sjm_headers = [
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
        ws_sjm.append([])  # Row 3
        ws_sjm.row_dimensions[3].height = 20
        for col_idx, header in enumerate(sjm_headers, 1):
            cell = ws_sjm.cell(row=3, column=col_idx, value=header)
            cell.font = Font(bold=True)

        for row_idx, r in enumerate(df.itertuples(), 4):
            ws_sjm.append([
                r.TGL,
                "LION PARCEL",
                r._3,  # Sc Origin
                r._4,  # Sc Destination
                r._5,  # Lt Number
                r._6,  # To Number
                r._8,  # Gross Weight
                r.Remarks,
                "",
            ])

        # ---------------------------------------------------------
        # 2. SHEET: MARKING
        # ---------------------------------------------------------
        ws_mk = wb.create_sheet(title="MARKING")
        ws_mk["A1"] = "MARKING SPX OSO SUB DC CYCLE "
        ws_mk["A2"] = "14 SEPTEMBER 2026 TRIP 2"

        mk_headers = [
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
        ws_mk.append([])  # Row 3
        for col_idx, header in enumerate(mk_headers, 1):
            cell = ws_mk.cell(row=3, column=col_idx, value=header)
            cell.font = Font(bold=True)

        for row_idx, r in enumerate(df.itertuples(), 4):
            # Formula External Number = G4 & "/" & E4 & "/" & I4
            ext_formula = f'=G{row_idx}&"/"&E{row_idx}&"/"&I{row_idx}'
            clear_gw_formula = r._8

            ws_mk.append([
                r.TGL,
                r.Vendor,
                r._3,
                r._4,
                r._5,
                r._6,
                r.Marking,
                r._8,
                r.Remarks,
                ext_formula,
                clear_gw_formula,
            ])

        # ---------------------------------------------------------
        # 3. SHEET: PVT
        # ---------------------------------------------------------
        ws_pvt = wb.create_sheet(title="PVT")
        pvt_headers = [
            "External Number",
            "Count of External Number",
            "Sum of Clear Gw",
        ]
        ws_pvt.append([])  # Row 1
        ws_pvt.append([])  # Row 2
        ws_pvt.append(pvt_headers)  # Row 3

        # Group by External Number indexes
        unique_dests = df.drop_duplicates(subset=["Sc Destination"]).index
        for idx, first_row in enumerate(unique_dests, 4):
            pvt_row = idx + 3
            ext_ref = f"=MARKING!J{first_row}"
            count_formula = (
                f"=COUNTIF(MARKING!J$4:J${len(df)+3}, A{pvt_row-3})"
            )
            sum_formula = f"=SUMIF(MARKING!J$4:J${len(df)+3}, A{pvt_row-3}, MARKING!K$4:K${len(df)+3})"
            ws_pvt.append([ext_ref, count_formula, sum_formula])

        # ---------------------------------------------------------
        # 4. SHEET: Sheet3 (Summary)
        # ---------------------------------------------------------
        ws_sum = wb.create_sheet(title="Sheet3")
        ws_sum.append([])
        ws_sum.append([])
        ws_sum.append(
            ["Sc Destination", "Count of To Number", "Sum of Gross Weight"]
        )

        summary_df = (
            df.groupby("Sc Destination")
            .agg(
                Count_of_TO=("To Number", "count"),
                Total_Weight=("Gross Weight", "sum"),
            )
            .reset_index()
        )
        for r in summary_df.itertuples():
            ws_sum.append([r._1, r.Count_of_TO, r.Total_Weight])

        # Save to memory stream
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.dataframe(df, use_container_width=True)

        st.download_button(
            label="📥 Download File Excel SJM & Marking (Format Asli)",
            data=output,
            file_name=f"FIXED_SCRIPT_SJ_MANUAL_{uploaded_file.name.replace('.pdf', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
