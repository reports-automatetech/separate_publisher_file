"""
Split Excel by Publisher Name — Streamlit App
Run with: streamlit run split_by_publisher_app.py
"""

import io
import os
import zipfile
import warnings

import numpy as np
import pandas as pd
import openpyxl
import streamlit as st

warnings.filterwarnings("ignore")

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Split by Publisher",
    page_icon="✂️",
    layout="centered",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'DM Mono', monospace;
    }

    .stApp {
        background: #0d0d0d;
        color: #e8e6df;
    }

    h1, h2, h3 {
        font-family: 'Syne', sans-serif !important;
    }

    .hero-title {
        font-family: 'Syne', sans-serif;
        font-size: 2.8rem;
        font-weight: 800;
        letter-spacing: -1px;
        color: #f0ede6;
        line-height: 1.1;
        margin-bottom: 0.2rem;
    }

    .hero-sub {
        font-family: 'DM Mono', monospace;
        font-size: 0.82rem;
        color: #7a7670;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin-bottom: 2.5rem;
    }

    .accent {
        color: #c8f04a;
    }

    .card {
        background: #161616;
        border: 1px solid #2a2a2a;
        border-radius: 10px;
        padding: 1.4rem 1.6rem;
        margin-bottom: 1rem;
    }

    .publisher-chip {
        display: inline-block;
        background: #1e1e1e;
        border: 1px solid #333;
        border-radius: 4px;
        padding: 2px 10px;
        margin: 3px 4px 3px 0;
        font-size: 0.78rem;
        color: #c8f04a;
        font-family: 'DM Mono', monospace;
    }

    .log-box {
        background: #0a0a0a;
        border: 1px solid #222;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        font-size: 0.78rem;
        color: #888;
        line-height: 1.8;
        max-height: 220px;
        overflow-y: auto;
        font-family: 'DM Mono', monospace;
    }

    .stat-row {
        display: flex;
        gap: 1rem;
        margin: 0.8rem 0;
    }

    .stat-box {
        flex: 1;
        background: #1a1a1a;
        border: 1px solid #2a2a2a;
        border-radius: 8px;
        padding: 0.9rem 1rem;
        text-align: center;
    }

    .stat-num {
        font-family: 'Syne', sans-serif;
        font-size: 2rem;
        font-weight: 800;
        color: #c8f04a;
        line-height: 1;
    }

    .stat-label {
        font-size: 0.7rem;
        color: #555;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        margin-top: 4px;
    }

    div.stButton > button {
        background: #c8f04a;
        color: #0d0d0d;
        font-family: 'Syne', sans-serif;
        font-weight: 800;
        font-size: 0.9rem;
        letter-spacing: 0.02em;
        border: none;
        border-radius: 6px;
        padding: 0.6rem 2rem;
        width: 100%;
        cursor: pointer;
        transition: opacity 0.15s;
    }

    div.stButton > button:hover {
        opacity: 0.85;
    }

    .stFileUploader > div {
        border: 1.5px dashed #333 !important;
        border-radius: 10px !important;
        background: #111 !important;
    }

    .stSelectbox > div > div {
        background: #161616 !important;
        border-color: #333 !important;
        color: #e8e6df !important;
    }

    hr {
        border-color: #222;
        margin: 1.5rem 0;
    }

    .stProgress > div > div {
        background: #c8f04a !important;
    }

    .success-banner {
        background: #162012;
        border: 1px solid #3a5c20;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        color: #c8f04a;
        font-family: 'Syne', sans-serif;
        font-weight: 700;
        font-size: 1.1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

PUBLISHER_COL = "publishername"
DATETIME_FMT  = "yyyy-mm-dd hh:mm:ss"


# ── Helpers ───────────────────────────────────────────────────────────────────

def scan_sheets(file_bytes: bytes, filename: str):
    """Return (valid_sheets, skipped_sheets, log_lines)."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    all_sheets = wb.sheetnames
    wb.close()

    valid, skipped, logs = [], [], []
    for sheet in all_sheets:
        try:
            header_df = pd.read_excel(
                io.BytesIO(file_bytes), sheet_name=sheet, nrows=0, engine="openpyxl"
            )
            cols_lower = [str(c).strip().lower() for c in header_df.columns]
            if PUBLISHER_COL in cols_lower:
                valid.append(sheet)
                logs.append(f"✅  \"{sheet}\" — publishername column found")
            else:
                skipped.append(sheet)
                logs.append(f"⏭️  \"{sheet}\" — no publishername column, skipped")
        except Exception as exc:
            skipped.append(sheet)
            logs.append(f"⚠️  \"{sheet}\" — error: {exc}")
    return valid, skipped, logs


def load_sheet_data(file_bytes: bytes, valid_sheets):
    """Return sheet_data dict and log lines."""
    sheet_data, logs = {}, []
    for sheet in valid_sheets:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, engine="openpyxl")
        df.columns = df.columns.str.strip()

        col_map = {c: c.strip().lower() for c in df.columns}
        pub_col = next(orig for orig, low in col_map.items() if low == PUBLISHER_COL)

        df[pub_col] = df[pub_col].astype(str).str.strip()
        df = df[df[pub_col].notna() & ~df[pub_col].isin(["", "nan", "None"])]

        dt_cols = [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]
        if dt_cols:
            logs.append(f"📅  \"{sheet}\" datetime cols: {dt_cols}")

        sheet_data[sheet] = {"df": df, "pub_col": pub_col, "dt_cols": dt_cols}
    return sheet_data, logs


def build_publisher_excel(publisher: str, sheet_data: dict) -> bytes | None:
    """Build an in-memory Excel for a single publisher. Returns None if no data."""
    output = io.BytesIO()
    tabs_written = 0

    with pd.ExcelWriter(
        output, engine="xlsxwriter", datetime_format=DATETIME_FMT, date_format="yyyy-mm-dd"
    ) as writer:
        workbook   = writer.book
        header_fmt = workbook.add_format(
            {"bold": True, "text_wrap": True, "valign": "top", "border": 1}
        )
        dt_fmt      = workbook.add_format({"num_format": DATETIME_FMT, "align": "left"})
        default_fmt = workbook.add_format({"align": "left"})

        for sheet, meta in sheet_data.items():
            df      = meta["df"]
            pub_col = meta["pub_col"]
            dt_cols = meta["dt_cols"]

            pub_df = df[df[pub_col] == publisher].copy().reset_index(drop=True)
            if pub_df.empty:
                continue

            for col in dt_cols:
                if pub_df[col].dt.tz is not None:
                    pub_df[col] = pub_df[col].dt.tz_localize(None)

            pub_df.to_excel(writer, sheet_name=sheet, index=False, startrow=0)
            worksheet = writer.sheets[sheet]

            for col_num, col_name in enumerate(pub_df.columns):
                worksheet.write(0, col_num, col_name, header_fmt)
                col_series  = pub_df[col_name]
                sample_strs = col_series.dropna().astype(str)
                max_data_w  = sample_strs.map(len).max() if not sample_strs.empty else 0
                col_width   = min(max(len(str(col_name)), max_data_w) + 2, 50)

                if col_name in dt_cols:
                    worksheet.set_column(col_num, col_num, max(col_width, 20), dt_fmt)
                else:
                    worksheet.set_column(col_num, col_num, col_width, default_fmt)

            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, pub_df.shape[0], pub_df.shape[1] - 1)
            tabs_written += 1

    if tabs_written == 0:
        return None

    output.seek(0)
    return output.read()


def build_zip(publisher_excels: dict) -> bytes:
    """Zip all publisher Excel bytes. publisher_excels: {filename: bytes}"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, data in publisher_excels.items():
            zf.writestr(fname, data)
    buf.seek(0)
    return buf.read()


# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown('<div class="hero-title">Split by <span class="accent">Publisher</span></div>', unsafe_allow_html=True)
st.markdown('<div class="hero-sub">Excel splitter · one file per publisher</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader(
    "Drop your Excel file here",
    type=["xlsx", "xls"],
    label_visibility="collapsed",
)

if not uploaded_file:
    st.markdown(
        '<div class="card" style="text-align:center;color:#555;font-size:0.82rem;">'
        '⬆ Upload an <code>.xlsx</code> file containing a <code>publishername</code> column'
        "</div>",
        unsafe_allow_html=True,
    )
    st.stop()

# ── Process ───────────────────────────────────────────────────────────────────
file_bytes = uploaded_file.read()
base_name  = os.path.splitext(uploaded_file.name)[0]

with st.spinner("Scanning sheets…"):
    valid_sheets, skipped_sheets, scan_logs = scan_sheets(file_bytes, uploaded_file.name)

# Log box
st.markdown('<div class="log-box">' + "<br>".join(scan_logs) + "</div>", unsafe_allow_html=True)

if not valid_sheets:
    st.error("No sheets with a `publishername` column were found. Please check your file.")
    st.stop()

with st.spinner("Loading sheet data…"):
    sheet_data, load_logs = load_sheet_data(file_bytes, valid_sheets)

if load_logs:
    st.markdown('<div class="log-box">' + "<br>".join(load_logs) + "</div>", unsafe_allow_html=True)

# Collect publishers
all_publishers = sorted(
    {val for meta in sheet_data.values() for val in meta["df"][meta["pub_col"]].unique()}
)

# Stats
total_rows = sum(meta["df"].shape[0] for meta in sheet_data.values())
st.markdown(
    f"""
    <div class="stat-row">
        <div class="stat-box"><div class="stat-num">{len(valid_sheets)}</div><div class="stat-label">Sheets</div></div>
        <div class="stat-box"><div class="stat-num">{len(all_publishers)}</div><div class="stat-label">Publishers</div></div>
        <div class="stat-box"><div class="stat-num">{total_rows:,}</div><div class="stat-label">Total Rows</div></div>
    </div>
    """,
    unsafe_allow_html=True,
)

# Publisher chips
chips_html = "".join(f'<span class="publisher-chip">{p}</span>' for p in all_publishers)
st.markdown(
    f'<div class="card"><div style="font-size:0.72rem;color:#555;text-transform:uppercase;'
    f'letter-spacing:0.1em;margin-bottom:0.5rem;">Publishers detected</div>{chips_html}</div>',
    unsafe_allow_html=True,
)

# ── Split Button ──────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
if st.button("✂️  Split & Download ZIP"):
    progress_bar = st.progress(0)
    status_text  = st.empty()
    publisher_excels = {}

    for i, publisher in enumerate(all_publishers):
        status_text.markdown(
            f'<span style="font-size:0.8rem;color:#666;">Processing: <b style="color:#c8f04a">{publisher}</b></span>',
            unsafe_allow_html=True,
        )
        excel_bytes = build_publisher_excel(publisher, sheet_data)
        if excel_bytes:
            safe_pub  = publisher.replace("/", "_").replace("\\", "_").replace(":", "_")
            fname     = f"{safe_pub}_{base_name}.xlsx"
            publisher_excels[fname] = excel_bytes
        progress_bar.progress((i + 1) / len(all_publishers))

    status_text.empty()
    progress_bar.empty()

    if not publisher_excels:
        st.warning("No data found for any publisher after filtering.")
        st.stop()

    zip_bytes = build_zip(publisher_excels)
    zip_name  = f"publisher_splits_{base_name}.zip"
    size_mb   = len(zip_bytes) / 1024 / 1024

    st.markdown(
        f'<div class="success-banner">✅ {len(publisher_excels)} file(s) ready — {size_mb:.2f} MB</div>',
        unsafe_allow_html=True,
    )

    st.download_button(
        label="⬇️  Download ZIP",
        data=zip_bytes,
        file_name=zip_name,
        mime="application/zip",
    )
