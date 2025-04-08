import streamlit as st
import pandas as pd
from io import BytesIO
import os

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="CSV to Branch-wise Excel", layout="wide")
st.title("📑 CSV to Branch-wise Excel Generator")

uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])

if uploaded_file:
    uploaded_filename = os.path.splitext(uploaded_file.name)[0]

    try:
        df = pd.read_csv(uploaded_file)

        if "Select Branch" not in df.columns:
            st.error("❌ Column 'Select Branch' not found in the uploaded CSV.")
        else:
            # Clean and identify branches
            df["Select Branch"] = df["Select Branch"].astype(str).str.strip()
            branches = df["Select Branch"].unique()

            # Create branch count table
            branch_counts = df["Select Branch"].value_counts().reset_index()
            branch_counts.columns = ["Branch", "Count"]
            total = branch_counts["Count"].sum()
            total_row = pd.DataFrame([["Total", total]], columns=["Branch", "Count"])
            branch_counts = pd.concat([branch_counts, total_row], ignore_index=True)

            # Overview Data
            test_title = uploaded_filename
            subject_name = "<Enter subject here>"

            # Generate initial Excel output
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                pd.DataFrame().to_excel(writer, sheet_name="Overview")  # Empty sheet to start
                df.to_excel(writer, sheet_name="All Data", index=False)

                for branch in branches:
                    branch_df = df[df["Select Branch"] == branch]
                    sheet_name = branch[:31]  # Excel sheet name limit
                    branch_df.to_excel(writer, sheet_name=sheet_name, index=False)

            output.seek(0)

            # Load workbook and edit Overview sheet
            workbook = load_workbook(output)
            overview_ws = workbook["Overview"]

            # Header: Overview
            overview_ws["A1"] = "Overview"
            overview_ws["A1"].font = Font(bold=True, size=16)  # Bold + Large

            # Row 2 left blank for spacing

            # Test Title
            overview_ws["A3"] = "Test Title:"
            overview_ws["B3"] = test_title

            # Row 4 left blank

            # Subject Name
            overview_ws["A5"] = "Subject Name:"
            overview_ws["B5"] = subject_name

            # Row 6 left blank

            # Branch-wise Count title
            overview_ws["A7"] = "Branch-wise Count"
            overview_ws["A7"].font = Font(bold=True)

            # Headers
            overview_ws.cell(row=8, column=1, value="Branch").font = Font(bold=True)
            overview_ws.cell(row=8, column=2, value="Count").font = Font(bold=True)

            # Data rows
            for idx, row in branch_counts.iterrows():
                r = 9 + idx
                overview_ws.cell(row=r, column=1, value=row["Branch"])
                cell = overview_ws.cell(row=r, column=2, value=row["Count"])
                if row["Branch"] == "Total":
                    overview_ws.cell(row=r, column=1).font = Font(bold=True)
                    cell.font = Font(bold=True)

            # Adjust column widths
            for col in range(1, 3):
                col_letter = get_column_letter(col)
                overview_ws.column_dimensions[col_letter].width = 20

            # Save final output
            final_output = BytesIO()
            workbook.save(final_output)
            final_output.seek(0)

            output_filename = f"BranchWise_{uploaded_filename}.xlsx"

            st.success("✅ File processed with Overview successfully!")
            st.download_button(
                label="📥 Download Excel File with Overview",
                data=final_output,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ Error reading CSV file: {e}")
