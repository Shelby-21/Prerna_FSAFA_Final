import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Financial Dashboard", layout="wide")
st.title("📊 Financial Performance Dashboard")

# --------------------------------------------------
# STRICT DATA LOADER (EXCEL-AWARE, NO ASSUMPTIONS)
# --------------------------------------------------
@st.cache_data
def load_excel_sheet(file, sheet_name, header_keyword):
    raw = pd.read_excel(file, sheet_name=sheet_name, header=None)

    header_row = None
    for i in range(len(raw)):
        if header_keyword in raw.iloc[i].astype(str).values:
            header_row = i
            break

    if header_row is None:
        st.error(f"Header '{header_keyword}' not found in {sheet_name}")
        return None

    df = pd.read_excel(file, sheet_name=sheet_name, header=header_row)

    metric_col = df.columns[0]
    date_cols = pd.to_datetime(df.columns[1:], errors="coerce")

    df.columns = [metric_col] + date_cols.tolist()

    valid_cols = [metric_col] + [c for c in df.columns[1:] if isinstance(c, pd.Timestamp)]
    df = df[valid_cols]

    for c in df.columns[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.set_index(metric_col).T
    df.index.name = "Date"
    df = df.sort_index()

    return df


# --------------------------------------------------
# FILE UPLOADER (REQUIRED)
# --------------------------------------------------
uploaded_file = st.file_uploader("Upload Screener Excel File", type=["xlsx"])
if uploaded_file is None:
    st.stop()

# --------------------------------------------------
# LOAD EXACT SHEETS
# --------------------------------------------------
df_main = load_excel_sheet(uploaded_file, "Data Sheet", "Report Date")
df_ratios = load_excel_sheet(uploaded_file, "Ratio Analysis", "Years")
df_cashflow = load_excel_sheet(uploaded_file, "Cash Flow", "Narration")

df_mscore = pd.read_excel(uploaded_file, sheet_name="M-score")
df_fscore = pd.read_excel(uploaded_file, sheet_name="F-score")

# --------------------------------------------------
# DASHBOARD
# --------------------------------------------------
st.header("1️⃣ Key Financial Highlights")

latest = df_main.index[-1]
previous = df_main.index[-2]

metrics = ["Sales", "Net Profit", "Operating Profit"]
cols = st.columns(len(metrics))

for i, m in enumerate(metrics):
    if m in df_main.columns:
        cols[i].metric(
            m,
            f"₹{df_main.loc[latest, m]:,.0f} Cr",
            f"{df_main.loc[latest, m] - df_main.loc[previous, m]:,.0f} Cr"
        )

# --------------------------------------------------
st.divider()
st.subheader("📈 Financial Trends")

selected = st.multiselect(
    "Select metrics",
    df_main.columns.tolist(),
    default=["Sales", "Net Profit"]
)

if selected:
    fig = px.line(df_main, x=df_main.index, y=selected, markers=True)
    fig.update_xaxes(tickformat="%Y")
    st.plotly_chart(fig, use_container_width=True)

# --------------------------------------------------
st.divider()
st.subheader("📉 Ratio Analysis")

ratio_sel = st.multiselect(
    "Select ratios",
    df_ratios.columns.tolist()
)

if ratio_sel:
    fig_r = px.line(df_ratios, x=df_ratios.index, y=ratio_sel, markers=True)
    fig_r.update_xaxes(tickformat="%Y")
    st.plotly_chart(fig_r, use_container_width=True)

# --------------------------------------------------
st.divider()
st.subheader("💰 Cash Flow")

cf_cols = [
    "Cash from Operating Activity",
    "Cash from Investing Activity",
    "Cash from Financing Activity"
]

existing = [c for c in cf_cols if c in df_cashflow.columns]

if existing:
    fig_cf = px.bar(df_cashflow, x=df_cashflow.index, y=existing, barmode="group")
    fig_cf.update_xaxes(tickformat="%Y")
    st.plotly_chart(fig_cf, use_container_width=True)

# --------------------------------------------------
st.divider()
st.subheader("📊 M-Score (Beneish)")

st.dataframe(df_mscore, use_container_width=True)
st.info("M-score is used to detect earnings manipulation risk. Lower than -1.78 is generally considered safe.")

# --------------------------------------------------
st.subheader("📊 F-Score (Piotroski)")

st.dataframe(df_fscore, use_container_width=True)
st.info("F-score ranges from 0–9. Higher score indicates stronger financial health.")

# --------------------------------------------------
with st.expander("📂 Cleaned Financial Data (Data Sheet)"):
    st.dataframe(df_main)
