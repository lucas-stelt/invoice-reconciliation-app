import pandas as pd
import streamlit as st
from PIL import Image

# --- Page Setup ---
st.set_page_config(page_title="Invoice Reconciliation", page_icon="💼", layout="wide")

# --- Header with Logo ---
logo = Image.open("USAF Logo.png")
col1, col2 = st.columns([2, 10])
with col1:
    st.image(logo, width=100)
with col2:
    st.markdown("### Reconciliation Tool")
    st.markdown("*Developed by FinData Team*")

# --- Sidebar ---
st.sidebar.image(logo, width=200)
st.sidebar.title("💼 Finance Tools")
st.sidebar.markdown("Upload two invoice files (CSV or Excel). Select the matching columns and compare them.")
st.sidebar.markdown("*Supports pre-formatted Excel tables and detects header rows.*")

# --- Detect first usable header row (for Excel files with pre-header rows) ---
def detect_header_row(file, selected_sheet):
    try:
        preview = pd.read_excel(file, sheet_name=selected_sheet, header=None, nrows=20)
        for i, row in preview.iterrows():
            if row.notnull().sum() >= 2 and all(isinstance(val, str) or pd.isna(val) for val in row):
                return i
        return 0
    except Exception:
        return 0

# --- File Reading ---
def read_file(file, selected_sheet=None):
    try:
        if file.name.endswith('.xlsx'):
            xls = pd.ExcelFile(file)
            if selected_sheet is None:
                return {"sheet_names": xls.sheet_names}
            header_row = detect_header_row(file, selected_sheet)
            return pd.read_excel(file, sheet_name=selected_sheet, skiprows=header_row)
        else:
            return pd.read_csv(file)
    except Exception as e:
        st.error(f"❌ Failed to read file: {file.name}.\nError: {e}")
        return None

# --- Reconciliation Logic ---
def reconcile(df1, df2, key1, val1, key2, val2):
    df1 = df1[[key1, val1]].rename(columns={key1: 'InvoiceID', val1: 'Value1'})
    df2 = df2[[key2, val2]].rename(columns={key2: 'InvoiceID', val2: 'Value2'})

    merged = pd.merge(df1, df2, on='InvoiceID', how='outer')
    merged['Difference'] = merged['Value1'] - merged['Value2']
    merged['Status'] = merged.apply(
        lambda row: 'Match' if pd.isna(row['Difference']) or abs(row['Difference']) < 1e-8 else 'Mismatch',
        axis=1
    )
    return merged

# --- Mismatch Highlighter ---
def highlight_mismatches(df):
    return df.style.apply(
        lambda row: ['background-color: #ffcccc' if row['Status'] == 'Mismatch' else '' for _ in row],
        axis=1
    )

# --- Main App ---
st.divider()
file1 = st.file_uploader("📄 Upload First File", type=['csv', 'xlsx'], key="file1")
file2 = st.file_uploader("📄 Upload Second File", type=['csv', 'xlsx'], key="file2")

if file1 and file2:
    df1_raw = read_file(file1)
    if isinstance(df1_raw, dict) and "sheet_names" in df1_raw:
        sheet1 = st.selectbox("🗂 Select sheet for First File", df1_raw["sheet_names"], key="sheet1")
        df1 = read_file(file1, selected_sheet=sheet1)
    else:
        df1 = df1_raw

    df2_raw = read_file(file2)
    if isinstance(df2_raw, dict) and "sheet_names" in df2_raw:
        sheet2 = st.selectbox("🗂 Select sheet for Second File", df2_raw["sheet_names"], key="sheet2")
        df2 = read_file(file2, selected_sheet=sheet2)
    else:
        df2 = df2_raw

    if df1 is not None and df2 is not None:
        st.subheader("📈 Preview Uploaded Data")
        with st.expander("🔎 First File Preview"):
            st.dataframe(df1.head())
        with st.expander("🔎 Second File Preview"):
            st.dataframe(df2.head())

        st.subheader("🔑 Column Selection")
        key_col1 = st.selectbox("Select Key Column from First File", df1.columns)
        key_col2 = st.selectbox("Select Key Column from Second File", df2.columns)
        val_col1 = st.selectbox("Select Value Column from First File", df1.columns)
        val_col2 = st.selectbox("Select Value Column from Second File", df2.columns)

        if st.button("🔁 Run Reconciliation"):
            with st.spinner("Reconciling invoices..."):
                try:
                    result = reconcile(df1, df2, key_col1, val_col1, key_col2, val_col2)
                    st.session_state["recon_result"] = result
                    st.success("✅ Reconciliation Complete!")
                except Exception as e:
                    st.error(f"⚠️ Reconciliation Error: {e}")
                    st.stop()

# --- Show Results if Available ---
if "recon_result" in st.session_state:
    result = st.session_state["recon_result"]

    st.subheader("📄 Results Summary")
    match_count = (result['Status'] == 'Match').sum()
    mismatch_count = (result['Status'] == 'Mismatch').sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("✅ Matches", match_count)
    col2.metric("❌ Mismatches", mismatch_count)
    col3.metric("📊 Total Invoices", len(result))

    st.subheader("🧶 View Results")
    show_mismatches_only = st.checkbox("🔍 Show only mismatches")
    display_df = result[result['Status'] == 'Mismatch'] if show_mismatches_only else result

    st.dataframe(display_df, use_container_width=True)

    with st.expander("🔴 Styled Preview (Mismatches Highlighted)"):
        st.dataframe(highlight_mismatches(display_df), use_container_width=True)

    csv = result.to_csv(index=False).encode('utf-8')
    st.download_button("⬇️ Download Reconciliation Report", csv, "reconciliation.csv", "text/csv")
