import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import datetime
import re

# =====================
# Page config
# =====================
st.set_page_config(page_title="EPA Comparison Tool (Terminal Mode)", layout="wide")

YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# =====================
# Helpers
# =====================
def clean(col):
    if pd.isna(col):
        return ""
    return re.sub(r"\s+", " ", str(col).replace("\n", " ").strip())

def normalize(col):
    return clean(col).lower().replace(" ", "").replace(".", "").replace("#", "")

def find_header_row(ws, max_rows=30):
    best, count = 1, 0
    for r in range(1, min(ws.max_row + 1, max_rows)):
        non_empty = sum(
            1 for c in range(1, ws.max_column + 1)
            if ws.cell(r, c).value not in [None, ""]
        )
        if non_empty > count:
            best, count = r, non_empty
    return best

def load_excel(file):
    file.seek(0)
    wb = openpyxl.load_workbook(BytesIO(file.read()), data_only=True)
    dfs = []
    for name in wb.sheetnames:
        ws = wb[name]
        header = find_header_row(ws)
        cols = [clean(ws.cell(header, c).value) or f"Unnamed_{c}" for c in range(1, ws.max_column + 1)]
        data = [
            [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            for r in range(header + 1, ws.max_row + 1)
        ]
        df = pd.DataFrame(data, columns=cols)
        dfs.append(df)
    wb.close()
    return pd.concat(dfs, ignore_index=True)

def detect_key_columns(columns):
    number = date = name = None
    for c in columns:
        n = normalize(c)
        if n in ["number", "no", "num"] and not number:
            number = c
        if n == "date" and not date:
            date = c
        if "applicant" in n and "name" in n and not name:
            name = c
    return number, date, name

def diff(a, b):
    if pd.isna(a) and pd.isna(b):
        return False
    return str(a).strip() != str(b).strip()

# =====================
# UI
# =====================
st.title("EPA Comparison Tool (Terminal-Equivalent)")

col1, col2 = st.columns(2)
with col1:
    raw_file = st.file_uploader("Raw / Previous Excel", type=["xlsx"])
with col2:
    latest_file = st.file_uploader("Latest Excel", type=["xlsx"])

st.divider()

col3, col4 = st.columns(2)
with col3:
    prev_date = st.text_input("Previous Date (e.g. 7/3)", value="")
with col4:
    latest_date = st.text_input("Latest Date (e.g. 8/29)", value="")

st.divider()

generate = st.button("🚀 Generate Comparison", type="primary")

# =====================
# Core logic
# =====================
if generate:
    if not raw_file or not latest_file:
        st.error("Please upload both files.")
        st.stop()

    if not prev_date or not latest_date:
        st.error("Please input both dates manually (terminal behavior).")
        st.stop()

    prev_df = load_excel(raw_file)
    latest_df = load_excel(latest_file)

    template_cols = list(prev_df.columns)

    num_col, date_col, name_col = detect_key_columns(template_cols)

    if not num_col:
        st.error("Number column not found.")
        st.stop()

    # normalize
    for c in template_cols:
        if c not in latest_df.columns:
            latest_df[c] = "-"

    prev_df = prev_df[template_cols]
    latest_df = latest_df[template_cols]

    prev_df[date_col] = prev_date
    latest_df[date_col] = latest_date

    output_df = pd.concat([prev_df, latest_df], ignore_index=True)

    # =====================
    # Excel output
    # =====================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        sheet = f"EPA {prev_date.replace('/','.')}_{latest_date.replace('/','.')}"
        output_df.to_excel(writer, index=False, sheet_name=sheet)

        ws = writer.sheets[sheet]

        prev_map = {
            str(r[num_col]): r
            for _, r in prev_df.iterrows()
            if pd.notna(r[num_col])
        }

        start_latest = len(prev_df) + 2

        for i, r in latest_df.iterrows():
            excel_row = start_latest + i
            key = str(r[num_col])
            is_new = key not in prev_map

            changed = False
            for j, col in enumerate(template_cols, start=1):
                if col in [num_col, date_col]:
                    continue

                cell = ws.cell(excel_row, j)

                if is_new:
                    cell.fill = YELLOW_FILL
                    changed = True
                else:
                    if diff(r[col], prev_map[key][col]):
                        cell.fill = YELLOW_FILL
                        changed = True

            if changed and name_col:
                ws.cell(excel_row, template_cols.index(name_col) + 1).fill = YELLOW_FILL

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col in ws.columns:
            width = min(max(len(str(c.value)) for c in col if c.value) + 2, 45)
            ws.column_dimensions[get_column_letter(col[0].column)].width = width

    output.seek(0)

    st.success("✅ Comparison generated (terminal-equivalent)")
    st.download_button(
        "📥 Download Excel",
        output,
        file_name=f"EPA {prev_date}_{latest_date}_Comparison.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
