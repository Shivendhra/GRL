# streamlit_app.py

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="File Registration",
    layout="wide"
)

st.title("📂 File Registration Generator")

# =========================================================
# SESSION STATE
# =========================================================

if "records" not in st.session_state:
    st.session_state.records = []

# =========================================================
# INPUT FORM
# =========================================================

with st.form("registration_form"):

    st.subheader("File Details")

    col1, col2 = st.columns(2)

    with col1:

        date = st.date_input("Date")

        file_ref = st.text_input("File Reference")

    with col2:

        property_name = st.text_area("Property")

    # =====================================================
    # CLIENT PARTICULARS
    # =====================================================

    st.subheader("Client(s) Particulars")

    col3, col4 = st.columns(2)

    with col3:

        st.markdown("### Vendors")

        vendor1 = st.text_input("Vendor 1")
        vendor2 = st.text_input("Vendor 2")
        vendor3 = st.text_input("Vendor 3")

    with col4:

        st.markdown("### Purchasers")

        purchaser1 = st.text_input("Purchaser 1")
        purchaser2 = st.text_input("Purchaser 2")
        purchaser3 = st.text_input("Purchaser 3")

    # =====================================================
    # SOLICITOR / FINANCIER
    # =====================================================

    st.subheader("Solicitor / Financier")

    col5, col6 = st.columns(2)

    with col5:

        v_solicitor = st.text_input("V Solicitor")

        v_financier = st.text_input("V Financier")

    with col6:

        p_solicitor = st.text_input("P Solicitor")

        p_financier = st.text_input("P Financier")

        b_solicitor = st.text_input("B Solicitor")

    submitted = st.form_submit_button("Add Record")

    # =====================================================
    # SAVE DATA
    # =====================================================

    if submitted:

        vendors = "\n".join([
            v for v in [
                vendor1,
                vendor2,
                vendor3
            ] if v.strip()
        ])

        purchasers = "\n".join([
            p for p in [
                purchaser1,
                purchaser2,
                purchaser3
            ] if p.strip()
        ])

        st.session_state.records.append({

            "Date": pd.to_datetime(date),

            "File Reference": file_ref,

            "Vendors": vendors,

            "Purchasers": purchasers,

            "Property": property_name,

            "V Solicitor": v_solicitor,

            "V Financier": v_financier,

            "P Solicitor": p_solicitor,

            "P Financier": p_financier,

            "B Solicitor": b_solicitor
        })

        st.success("✅ Record Added")

# =========================================================
# DISPLAY TABLE
# =========================================================

if st.session_state.records:

    df = pd.DataFrame(st.session_state.records)

    # SORT BY DATE
    df = df.sort_values(by="Date").reset_index(drop=True)

    # AUTO NUMBERING
    df.insert(0, "No", range(1, len(df) + 1))

    st.subheader("Preview")

    st.dataframe(df, use_container_width=True)

    # =====================================================
    # EXCEL GENERATOR
    # =====================================================

    def generate_excel(dataframe):

        wb = Workbook()

        ws = wb.active

        ws.title = "File Registration"

        # =================================================
        # TITLE
        # =================================================

        ws.merge_cells("A1:I1")

        ws["A1"] = "File Registration"

        ws["A1"].font = Font(
            bold=True,
            size=12
        )

        # =================================================
        # HEADERS
        # =================================================

        headers = [
            "No",
            "Date",
            "File Reference",
            "",
            "Client(s) Particulars",
            "Property",
            "R",
            "",
            ""
        ]

        for col_num, header in enumerate(headers, 1):

            cell = ws.cell(
                row=2,
                column=col_num
            )

            cell.value = header

            cell.font = Font(bold=True)

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center"
            )

        # =================================================
        # COLUMN WIDTHS
        # =================================================

        widths = {

            1: 6,     # No
            2: 12,    # Date
            3: 28,    # File Reference
            4: 5,     # V/P
            5: 40,    # Client Particulars
            6: 40,    # Property
            7: 8,     # R
            8: 20,    # Labels
            9: 25     # Values
        }

        for col, width in widths.items():

            ws.column_dimensions[
                get_column_letter(col)
            ].width = width

        # =================================================
        # DATA ROWS
        # =================================================

        start_row = 3

        for index, row in dataframe.iterrows():

            top_row = start_row
            bottom_row = start_row + 1

            # =============================================
            # MERGED CELLS
            # =============================================

            merge_columns = [
                1,  # No
                2,  # Date
                3,  # File Reference
                6,  # Property
                7,  # R
                8,  # Labels
                9   # Values
            ]

            for col in merge_columns:

                ws.merge_cells(
                    start_row=top_row,
                    start_column=col,
                    end_row=bottom_row,
                    end_column=col
                )

            # =============================================
            # MAIN INFO
            # =============================================

            ws.cell(top_row, 1).value = row["No"]

            ws.cell(top_row, 2).value = row["Date"].strftime("%d/%m/%y")

            ws.cell(top_row, 3).value = row["File Reference"]

            # =============================================
            # V / P COLUMN
            # =============================================

            ws.cell(top_row, 4).value = "V"

            ws.cell(bottom_row, 4).value = "P"

            # =============================================
            # CLIENT PARTICULARS
            # =============================================

            ws.cell(top_row, 5).value = row["Vendors"]

            ws.cell(bottom_row, 5).value = row["Purchasers"]

            # =============================================
            # PROPERTY COLUMN
            # =============================================

            ws.cell(top_row, 6).value = row["Property"]

            # =============================================
            # R COLUMN (EMPTY)
            # =============================================

            ws.cell(top_row, 7).value = ""

            # =============================================
            # LABEL COLUMN
            # =============================================

            label_text = (
                "V Solicitor\n"
                "V Financier\n"
                "P Solicitor\n"
                "P Financier\n"
                "B Solicitor"
            )

            ws.cell(top_row, 8).value = label_text

            # =============================================
            # VALUE COLUMN
            # =============================================

            value_text = (
                f'{row["V Solicitor"]}\n'
                f'{row["V Financier"]}\n'
                f'{row["P Solicitor"]}\n'
                f'{row["P Financier"]}\n'
                f'{row["B Solicitor"]}'
            )

            ws.cell(top_row, 9).value = value_text

            # =============================================
            # ALIGNMENT
            # =============================================

            for r in [top_row, bottom_row]:

                for c in range(1, 10):

                    cell = ws.cell(r, c)

                    cell.alignment = Alignment(
                        horizontal="left",
                        vertical="top",
                        wrap_text=True
                    )

            # =============================================
            # ROW HEIGHT
            # =============================================

            ws.row_dimensions[top_row].height = 45
            ws.row_dimensions[bottom_row].height = 45

            start_row += 2

        # =================================================
        # BORDERS
        # =================================================

        thin = Side(style="thin")

        for row in ws.iter_rows(
            min_row=1,
            max_row=ws.max_row,
            min_col=1,
            max_col=9
        ):

            for cell in row:

                cell.border = Border(
                    left=thin,
                    right=thin,
                    top=thin,
                    bottom=thin
                )

        # =================================================
        # FONT STYLING
        # =================================================

        for row in ws.iter_rows():

            for cell in row:

                cell.font = Font(size=10)

        # Header fonts
        for cell in ws[2]:

            cell.font = Font(
                bold=True,
                size=10
            )

        # =================================================
        # SAVE FILE
        # =================================================

        output = BytesIO()

        wb.save(output)

        output.seek(0)

        return output

    # =====================================================
    # DOWNLOAD BUTTON
    # =====================================================

    excel_file = generate_excel(df)

    st.download_button(
        label="📥 Download Excel File",
        data=excel_file,
        file_name="file_registration.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    