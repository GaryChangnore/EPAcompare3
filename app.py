import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from io import BytesIO
from datetime import datetime
import re

# =========================
# Page config
# =========================
st.set_page_config(
    page_title="EPA Comparison Tool (Terminal-Equivalent)",
    layout="wide"
)

YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# =========================
# Helpers
# =========================
def clean(s):
    if s is None:
        return ""
    return str(s).strip().replace("\n", " ").replace("\r", " ")

def normalize(s):
    return re.sub(r"[^a-z0-9]", "", clean(s).lower())

def parse_date_input(d):
    if not d:
        return None
    if isinstance(d, datetime):
        return d
    try:
        return datetime.strptime(str(d), "%Y-%m-%d")
    except:
        return None

def format_short_date(dt):
    return f"{dt.month}/{dt.day}"

# =========================
# Excel loading (terminal-like)
# =========================
def load_excel(file):
    file.seek(0)
    wb = openpyxl.load_workbook(file, data_only=True)
    sheets = {}

    for sheet in wb.sheetnames:
        ws = wb[sheet]

        # find header row = row with most non-empty cells
        best_row = 1
        max_count = 0
        for r in range(1, min(30, ws.max_row) + 1):
            count = sum(
                1 for c in range(1, ws.max_column + 1)
                if ws.cell(r, c).value not in (None, "")
            )
            if count > max_count:
                best_row = r
                max_count = count

        headers = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(best_row, c).value
            headers.append(clean(v) if v else f"Unnamed_{c}")

        rows = []
        for r in range(best_row + 1, ws.max_row + 1):
            rows.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])

        df = pd.DataFrame(rows, columns=headers)
        sheets[sheet] = df

    return sheets

# =========================
# Key column detection (TERMINAL EQUIVALENT)
# =========================
def detect_key_columns(columns):
    number = date = name = None

    for c in columns:
        raw = clean(c).lower()
        n = normalize(c)

        # ---- Number (very permissive, terminal-style)
        if number is None:
            if (
                n in ["number", "no", "num", "numb"]
                or raw.startswith("no")
                or raw == "#"
                or ("permit" in raw and "no" in raw)
                or n.endswith("no")
            ):
                number = c
                continue

        # ---- Date
        if date is None and "date" in raw:
            date = c
            continue

        # ---- Applicant name
        if name is None and "applicant" in raw and "name" in raw:
            name = c
            continue

    # fallback = terminal behavior
    if number is None and len(columns) > 0:
        number = columns[0]

    return number, date, name

# =========================
# Comparison logic
# =========================
def compare_and_generate(prev_df, latest_df, number_col, name_col, prev_date, latest_date):
    prev_map = {}
    for _, r in prev_df.iterrows():
        k = str(r[number_col]).strip()
        if k:
            prev_map[k] = r

    output = pd.concat([prev_df, latest_df], ignore_index=True)

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        sheet_name = f"EPA {format_short_date(prev_date)} {format_short_date(latest_date)}"
        output.to_excel(writer, sheet_name=sheet_name, index=False)

        ws = writer.sheets[sheet_name]

        prev_len = len(prev_df)
        latest_start = prev_len + 2

        for i, r in latest_df.iterrows():
            excel_row = latest_start + i
            key = str(r[number_col]).strip()

            is_new = key not in prev_map
            changed = False

            for j, col in enumerate(output.columns, start=1):
                if col == number_col:
                    continue

                cell = ws.cell(excel_row, j)

                if is_new:
                    cell.fill = YELLOW_FILL
                    changed = True
                else:
                    old = prev_map[key].get(col)
                    new = r.get(col)
                    if str(old).strip() != str(new).strip():
                        cell.fill = YELLOW_FILL
                        changed = True

            if changed and name_col:
                idx = output.columns.get_loc(name_col) + 1
                ws.cell(excel_row, idx).fill = YELLOW_FILL

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col in ws.columns:
            width = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(width + 2, 50)

    bio.seek(0)
    return bio

# =========================
# UI
# =========================
st.title("EPA Comparison Tool (Terminal-Equivalent)")

col1, col2 = st.columns(2)

with col1:
    prev_file = st.file_uploader("Raw / Previous Excel", type=["xlsx"])
with col2:
    latest_file = st.file_uploader("Latest Excel", type=["xlsx"])

d1, d2 = st.columns(2)
with d1:
    prev_date_input = st.date_input("Previous Date (e.g. 7/3)")
with d2:
    latest_date_input = st.date_input("Latest Date (e.g. 8/29)")

if prev_file and latest_file:
    st.success("Files loaded")

    prev_date = parse_date_input(prev_date_input)
    latest_date = parse_date_input(latest_date_input)

    if not prev_date or not latest_date:
        st.error("Please input both dates.")
        st.stop()

    prev_sheets = load_excel(prev_file)
    latest_sheets = load_excel(latest_file)

    # terminal-style: use first sheet
    prev_df = list(prev_sheets.values())[0]
    latest_df = list(latest_sheets.values())[0]

    number_col, date_col, name_col = detect_key_columns(prev_df.columns)

    st.write("### Detected Columns")
    st.write(f"- Number: `{number_col}`")
    st.write(f"- Applicant Name: `{name_col}`")

    if st.button("🚀 Generate Comparison"):
        output = compare_and_generate(
            prev_df,
            latest_df,
            number_col,
            name_col,
            prev_date,
            latest_date
        )

        filename = f"EPA_{format_short_date(prev_date).replace('/','.')}_{format_short_date(latest_date).replace('/','.')}_Comparison.xlsx"

        st.download_button(
            "📥 Download Excel",
            data=output,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
