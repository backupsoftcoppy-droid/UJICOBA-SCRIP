import io
import re

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="SPX Laporan Scan Converter", page_icon="📦", layout="wide"
)

st.title("📦 SPX Laporan Scan PDF ➡️ Excel Converter")

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
    "Sorong Utara DC": "SOQ-C1-1",
    "Tarakan Barat 4 Hub": "TRK-C1-1",
    "Tarakan Barat Hub": "TRK-C1-1",
    "Tarakan Timur Hub": "TRK-C1-1",
    "Tarakan Utara Hub": "TRK-C1-1",
    "Teluk Mutiara Hub": "ARD-C1-1",
    "Ternate Utara Hub": "TTE-C1-1",
    "Wua-Wua DC": "KDI-C1-1",
}


def apply_table_formatting(ws, start_row, max_col):
    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    for col in range(1, max_col + 1):
        cell = ws.cell(row=start_row, column=col)
        cell.font = Font(bold=True, name="Calibri")
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    for r in range(start_row, ws.max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = thin_border
            if r > start_row:
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        if col[0].column > max_col:
            continue
        for cell in col:
            if cell.row < start_row:
                continue
            if cell.value:
                val_str = str(cell.value)
                if len(val_str) > max_len:
                    max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 16)


def extract_data_from_pdf(pdf_file):
    reader = pypdf.PdfReader(pdf_file)
    extracted_rows = []
    valid_destinations = list(MARKING_MAP.keys())

    for page in reader.pages:
        text = page.extract_text()
        if not text:
            continue

        lt_match = re.search(r"\b(LT[A-Z0-9]{8,})\b", text)
        lt_num = lt_match.group(1) if lt_match else ""

        dest = ""
        for valid_dest in valid_destinations:
            if valid_dest in text and valid_dest != "SURABAYA DC":
                dest = valid_dest
                break

        if not dest:
            all_dcs = re.findall(
                r"\b([A-Za-z0-9\s-]+?\s*(?:DC|Hub))\b", text, re.IGNORECASE
            )
            for d in all_dcs:
                d_clean = d.strip()
                if "SURABAYA" not in d_clean.upper():
                    dest = d_clean
                    break

        std_match = re.search(r"(\d{4}/\d{2}/\d{2})\s*\d{2}:\d{2}:\d{2}STD", text)
        if not std_match:
            std_match = re.search(r":\s*(\d{4}/\d{2}/\d{2})", text)
        tgl = (
            std_match.group(1).replace("/", "-")
            if std_match
            else "2026-09-14"
        )

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
                "Vendor": "Lion Parcel",
                "Sc Origin": "SURABAYA DC",
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
        st.success(f"Berhasil! Total **{len(df)}** TO ditemukan.")

        df_sjm = pd.DataFrame({
            "TGL": df["TGL"],
            "Vendor": "LION PARCEL",
            "SC Orgin": df["Sc Origin"],
            "DESTINATION": df["Sc Destination"],
            "LT NUMBER": df["Lt Number"],
            "TO NUMBER": df["To Number"],
            "Gross Weight": df["Gross Weight"],
            "REMAKE": df["Remarks"],
            "TOTAL": "",
        })

        df_marking = pd.DataFrame({
            "Tanggal": df["TGL"],
            "Vendor": df["Vendor"],
            "Sc Origin": df["Sc Origin"],
            "Sc Destination": df["Sc Destination"],
            "Lt Number": df["Lt Number"],
            "To Number": df["To Number"],
            "Marking": df["Marking"],
            "Gross Weight": df["Gross Weight"],
            "Remarks": df["Remarks"],
            "External Number": df["Marking"]
            + "/"
            + df["Lt Number"]
            + "/"
            + df["Remarks"],
            "Clear Gw": df["Gross Weight"],
        })

        df_pvt = (
            df_marking.groupby("External Number")
            .agg(
                Count_of_External_Number=("To Number", "count"),
                Sum_of_Clear_Gw=("Clear Gw", "sum"),
            )
            .reset_index()
        )
        df_pvt.columns = [
            "External Number",
            "Count of External Number",
            "Sum of Clear Gw",
        ]

        df_sheet3 = (
            df.groupby("Sc Destination")
            .agg(
                Count_of_To_Number=("To Number", "count"),
                Sum_of_Gross_Weight=("Gross Weight", "sum"),
            )
            .reset_index()
        )
        df_sheet3.columns = [
            "Sc Destination",
            "Count of To Number",
            "Sum of Gross Weight",
        ]

        t1, t2, t3, t4 = st.tabs(
            ["📋 Sheet SJM", "🏷️ Sheet MARKING", "📑 Sheet PVT", "📊 Sheet3"]
        )

        with t1:
            st.markdown(
                "**SURAT JALAN MANUAL SURABAYA DC VIA LION STD | 14 SEPTEMBER 2026 TRIP 2**"
            )
            st.dataframe(df_sjm, use_container_width=True, hide_index=True)

        with t2:
            st.markdown("**MARKING SPX OSO SUB DC CYCLE | 14 SEPTEMBER 2026 TRIP 2**")
            st.dataframe(df_marking, use_container_width=True, hide_index=True)

        with t3:
            st.dataframe(df_pvt, use_container_width=True, hide_index=True)

        with t4:
            st.dataframe(df_sheet3, use_container_width=True, hide_index=True)

        wb = openpyxl.Workbook()
        red_fill = PatternFill(
            start_color="FF0000", end_color="FF0000", fill_type="solid"
        )
        yellow_fill = PatternFill(
            start_color="FFFF00", end_color="FFFF00", fill_type="solid"
        )
        grey_fill = PatternFill(
            start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"
        )
        thin_border = Border(
            left=Side(style="thin", color="000000"),
            right=Side(style="thin", color="000000"),
            top=Side(style="thin", color="000000"),
            bottom=Side(style="thin", color="000000"),
        )

        ws_sjm = wb.active
        ws_sjm.title = "SJM"
        ws_sjm.merge_cells("A1:H1")
        cell_r1 = ws_sjm.cell(
            row=1, column=1, value="SURAT JALAN MANUAL SURABAYA DC VIA LION STD"
        )
        cell_r1.font = Font(bold=True, size=11, name="Calibri")
        cell_r1.alignment = Alignment(horizontal="center", vertical="center")

        ws_sjm.merge_cells("A2:H2")
        cell_r2 = ws_sjm.cell(row=2, column=1, value="14 SEPTEMBER 2026 TRIP 2")
        cell_r2.font = Font(bold=True, size=11, name="Calibri")
        cell_r2.alignment = Alignment(horizontal="center", vertical="center")

        total_cell = ws_sjm.cell(row=2, column=9, value=len(df))
        total_cell.font = Font(bold=True, size=11, name="Calibri")
        total_cell.alignment = Alignment(horizontal="center", vertical="center")

        for c_idx, col_name in enumerate(df_sjm.columns, 1):
            ws_sjm.cell(row=3, column=c_idx, value=col_name)

        for r_idx, row_val in enumerate(df_sjm.itertuples(index=False), 4):
            for c_idx, val in enumerate(row_val, 1):
                ws_sjm.cell(row=r_idx, column=c_idx, value=val)

        apply_table_formatting(ws_sjm, start_row=3, max_col=len(df_sjm.columns))
        for r in range(1, 3):
            for c in range(1, 10):
                ws_sjm.cell(row=r, column=c).border = thin_border

        ws_mk = wb.create_sheet(title="MARKING")
        ws_mk.merge_cells("A1:K1")
        cell_mk1 = ws_mk.cell(
            row=1, column=1, value="MARKING SPX OSO SUB DC CYCLE"
        )
        cell_mk1.font = Font(bold=True, color="FFFF00", name="Calibri", size=11)
        cell_mk1.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, 12):
            cell = ws_mk.cell(row=1, column=c)
            cell.fill = red_fill
            cell.border = thin_border

        ws_mk.merge_cells("A2:K2")
        cell_mk2 = ws_mk.cell(row=2, column=1, value="14 SEPTEMBER 2026 TRIP 2")
        cell_mk2.font = Font(bold=True, color="FF0000", name="Calibri", size=10)
        cell_mk2.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, 12):
            cell = ws_mk.cell(row=2, column=c)
            cell.fill = yellow_fill
            cell.border = thin_border

        for c_idx, col_name in enumerate(df_marking.columns, 1):
            ws_mk.cell(row=3, column=c_idx, value=col_name)

        for r_idx, r in enumerate(df.itertuples(index=False), 4):
            ws_mk.cell(row=r_idx, column=1, value=r.TGL)
            ws_mk.cell(row=r_idx, column=2, value=r.Vendor)
            ws_mk.cell(row=r_idx, column=3, value=r._2)
            ws_mk.cell(row=r_idx, column=4, value=r._3)
            ws_mk.cell(row=r_idx, column=5, value=r._4)
            ws_mk.cell(row=r_idx, column=6, value=r._5)
            ws_mk.cell(row=r_idx, column=7, value=r.Marking)
            ws_mk.cell(row=r_idx, column=8, value=r._7)
            ws_mk.cell(row=r_idx, column=9, value=r.Remarks)
            ws_mk.cell(
                row=r_idx,
                column=10,
                value=f'=G{r_idx}&"/"&E{r_idx}&"/"&I{r_idx}',
            )
            ws_mk.cell(row=r_idx, column=11, value=r._7)

        footer_row = len(df) + 4
        for c in range(1, 12):
            cell = ws_mk.cell(row=footer_row, column=c)
            cell.fill = red_fill
            cell.border = thin_border

        apply_table_formatting(
            ws_mk, start_row=3, max_col=len(df_marking.columns)
        )

        ws_pvt = wb.create_sheet(title="PVT")
        for c_idx, col_name in enumerate(df_pvt.columns, 1):
            cell = ws_pvt.cell(row=1, column=c_idx, value=col_name)
            cell.fill = grey_fill

        unique_indices = df.drop_duplicates(subset=["Sc Destination"]).index
        last_pvt_row = 1
        for p_idx, first_r in enumerate(unique_indices, 2):
            ws_pvt.cell(row=p_idx, column=1, value=f"=MARKING!J{first_r+4}")
            ws_pvt.cell(
                row=p_idx,
                column=2,
                value=f"=COUNTIF(MARKING!J$4:J${len(df)+3}, A{p_idx})",
            )
            ws_pvt.cell(
                row=p_idx,
                column=3,
                value=f"=SUMIF(MARKING!J$4:J${len(df)+3}, A{p_idx}, MARKING!K$4:K${len(df)+3})",
            )
            last_pvt_row = p_idx

        gt_row = last_pvt_row + 1
        gt_cell1 = ws_pvt.cell(row=gt_row, column=1, value="Grand Total")
        gt_cell2 = ws_pvt.cell(
            row=gt_row, column=2, value=f"=SUM(B2:B{last_pvt_row})"
        )
        gt_cell3 = ws_pvt.cell(
            row=gt_row, column=3, value=f"=SUM(C2:C{last_pvt_row})"
        )

        gt_cell1.font = Font(bold=True, name="Calibri")
        gt_cell2.font = Font(bold=True, name="Calibri")
        gt_cell3.font = Font(bold=True, name="Calibri")

        gt_cell1.fill = grey_fill
        gt_cell2.fill = grey_fill
        gt_cell3.fill = grey_fill

        apply_table_formatting(
            ws_pvt, start_row=1, max_col=len(df_pvt.columns)
        )

        ws_sum = wb.create_sheet(title="Sheet3")

        for c_idx, col_name in enumerate(df_sheet3.columns, 1):
            cell = ws_sum.cell(row=1, column=c_idx, value=col_name)
            cell.fill = grey_fill

        unique_dests = df["Sc Destination"].unique()
        last_s3_row = 1
        for s_idx, dest_name in enumerate(unique_dests, 2):
            ws_sum.cell(row=s_idx, column=1, value=dest_name)
            ws_sum.cell(
                row=s_idx,
                column=2,
                value=f"=COUNTIF(MARKING!D$4:D${len(df)+3}, A{s_idx})",
            )
            ws_sum.cell(
                row=s_idx,
                column=3,
                value=f"=SUMIF(MARKING!D$4:D${len(df)+3}, A{s_idx}, MARKING!H$4:H${len(df)+3})",
            )
            last_s3_row = s_idx

        s3_gt_row = last_s3_row + 1
        s3_gt1 = ws_sum.cell(row=s3_gt_row, column=1, value="Grand Total")
        s3_gt2 = ws_sum.cell(
            row=s3_gt_row, column=2, value=f"=SUM(B2:B{last_s3_row})"
        )
        s3_gt3 = ws_sum.cell(
            row=s3_gt_row, column=3, value=f"=SUM(C2:C{last_s3_row})"
        )

        s3_gt1.font = Font(bold=True, name="Calibri")
        s3_gt2.font = Font(bold=True, name="Calibri")
        s3_gt3.font = Font(bold=True, name="Calibri")

        s3_gt1.fill = grey_fill
        s3_gt2.fill = grey_fill
        s3_gt3.fill = grey_fill

        apply_table_formatting(
            ws_sum, start_row=1, max_col=len(df_sheet3.columns)
        )

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        st.download_button(
            label="📥 Download File Excel SJM & Marking",
            data=output,
            file_name=f"FIXED_SCRIPT_SJ_MANUAL_{uploaded_file.name.replace('.pdf', '')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        
