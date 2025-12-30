import streamlit as st
import pandas as pd
import plotly.express as px

# --- CONFIGURATION ---
st.set_page_config(layout="wide", page_title="HUL Financial Dashboard")
FILE_NAME = "Hind_Unilever_cleaned.xlsx"

# --- HELPER FUNCTIONS ---
def clean_financial_sheet(df, anchor_metric):
    """
    Cleans Screener.in style sheets. Finds the metric column, 
    identifies headers, transposes data, and cleans numeric strings.
    """
    if df is None or df.empty: return pd.DataFrame()
    
    # 1. Find the row and column where the data starts
    metric_row, metric_col = None, None
    for r in range(len(df)):
        for c in range(min(df.shape[1], 3)): # Check first 3 columns
            if str(df.iloc[r, c]).strip().lower() == anchor_metric.lower():
                metric_row, metric_col = r, c
                break
        if metric_row is not None: break
    
    if metric_row is None: return pd.DataFrame()

    # 2. Extract headers (Dates are usually 2 rows above or in a specific row)
    # We'll use the column names for values and index for metrics
    data = df.iloc[metric_row:, metric_col:].copy()
    data.columns = [f"Col_{i}" for i in range(data.shape[1])]
    data = data.rename(columns={"Col_0": "Metric"})
    
    # 3. Filter out "Narration" or empty metric rows
    data = data[data['Metric'].notna()]
    data = data[~data['Metric'].astype(str).str.contains('Narration', case=False)]
    
    # 4. Set Index and Transpose
    data = data.set_index('Metric').T
    
    # 5. Clean Numbers (Remove commas, convert to float)
    for col in data.columns:
        data[col] = data[col].astype(str).str.replace(',', '', regex=False)
        data[col] = pd.to_numeric(data[col], errors='coerce')
    
    # 6. Add a cleaner 'Period' index
    data.index = [f"Period {i+1}" for i in range(len(data))]
    return data

@st.cache_data
def load_data():
    try:
        xl = pd.read_excel(FILE_NAME, sheet_name=None, header=None)
        return xl
    except Exception as e:
        st.error(f"Error loading {FILE_NAME}: {e}")
        return None

# --- MAIN APP ---
st.title("📊 Hindustan Unilever Limited: Financial Analysis")
st.markdown(f"Analyzed directly from `{FILE_NAME}`")

sheets = load_data()

if sheets:
    # Processing Sheets
    df_pnl = clean_financial_sheet(sheets.get('Profit & Loss'), 'Sales')
    df_bs = clean_financial_sheet(sheets.get('Balance Sheet'), 'Equity Share Capital')
    df_cf = clean_financial_sheet(sheets.get('Cash Flow'), 'Cash from Operating Activity')
    df_q = clean_financial_sheet(sheets.get('Quarters'), 'Sales')
    df_ratios = clean_financial_sheet(sheets.get('Ratio Analysis'), 'Sales Growth')

    # --- TOP KPI METRICS ---
    # Using 'Data Sheet' for latest price/mcap
    ds = sheets.get('Data Sheet')
    try:
        mcap = ds.iloc[6, 2] # Based on your file structure
        price = ds.iloc[5, 2]
    except:
        mcap, price = 0, 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Market Cap", f"₹ {mcap:,.0f} Cr")
    c2.metric("Current Price", f"₹ {price:,.2f}")
    if not df_pnl.empty:
        c3.metric("Latest Annual Sales", f"₹ {df_pnl['Sales'].iloc[-1]:,.0f} Cr")

    st.divider()

    # --- TABS FOR NAVIGATION ---
    t1, t2, t3, t4, t5 = st.tabs(["Profit & Loss", "Balance Sheet", "Cash Flow", "Ratios", "Scores"])

    with t1:
        st.subheader("P&L Growth Trends")
        if not df_pnl.empty:
            sel = st.multiselect("Select Metrics", df_pnl.columns.tolist(), default=['Sales', 'Net profit'])
            fig = px.line(df_pnl, x=df_pnl.index, y=sel, markers=True, title="Annual Income Trends")
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(df_pnl)

    with t2:
        st.subheader("Balance Sheet Composition")
        if not df_bs.empty:
            sel_bs = st.multiselect("Select Items", df_bs.columns.tolist(), default=['Reserves', 'Borrowings'])
            fig_bs = px.bar(df_bs, x=df_bs.index, y=sel_bs, barmode='group')
            st.plotly_chart(fig_bs, use_container_width=True)

    with t3:
        st.subheader("Cash Flow Analysis")
        if not df_cf.empty:
            st.area_chart(df_cf[['Net Cash Flow']])
            st.dataframe(df_cf)

    with t4:
        st.subheader("Efficiency Ratios")
        if not df_ratios.empty:
            ratio = st.selectbox("Choose Ratio", df_ratios.columns.tolist())
            st.line_chart(df_ratios[ratio])

    with t5:
        col_f, col_m = st.columns(2)
        with col_f:
            st.subheader("Piotroski F-Score")
            st.table(sheets.get('F-score').dropna(how='all').iloc[:, 1:5])
        with col_m:
            st.subheader("Beneish M-Score")
            st.table(sheets.get('M-score').dropna(how='all').iloc[:, 1:5])

else:
    st.warning("Please ensure the Excel file is in the same folder as the script.")
