import streamlit as st
import pandas as pd
import plotly.express as px

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Financial Dashboard", layout="wide")
st.title("📊 Financial Performance Dashboard")

# --- DATA LOADING & CLEANING FUNCTION ---
@st.cache_data
def load_and_clean_data(file_path, sheet_name, header_keyword):
    """
    Loads data from Screener.in excel, finding the correct header row 
    and fixing the 'Year' column problem.
    """
    try:
        # 1. Read first 20 rows to find where the actual data starts
        df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=20)
        
        # Find the row index containing the specific keyword (e.g., 'Report Date' or 'Years')
        # We look for the keyword in the first few columns
        header_row = None
        for i, row in df_raw.iterrows():
            if header_keyword in row.values:
                header_row = i
                break
        
        if header_row is None:
            st.error(f"Could not find header '{header_keyword}' in sheet '{sheet_name}'")
            return None

        # 2. Load the full dataset using the discovered header row
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

        # 3. Clean the 'Year' Headers (The "Year Problem")
        # The first column is usually the Metric Name (e.g., "Sales"), the rest are dates
        metric_col_name = df.columns[0]
        date_columns = df.columns[1:]
        
        # Convert column headers to datetime objects to ensure correct sorting
        # Errors='coerce' will turn non-dates (like "Trailing") into NaT (Not a Time)
        new_columns = [metric_col_name] + pd.to_datetime(date_columns, errors='coerce').tolist()
        df.columns = new_columns

        # 4. Filter out columns that aren't dates (like 'Trailing', 'Best Case') for clean graphing
        # We keep the first column (Metric) and any column that is a valid timestamp
        valid_cols = [df.columns[0]] + [col for col in df.columns[1:] if isinstance(col, pd.Timestamp)]
        df = df[valid_cols]

        # 5. Clean Data Values (The "Revenue Problem")
        # Ensure all data is numeric (remove commas, handle non-numeric text)
        for col in df.columns[1:]:
            df[col] = pd.to_numeric(df[col], errors='coerce')

        # 6. Transpose for easier plotting (Years as rows, Metrics as columns)
        df_transposed = df.set_index(metric_col_name).T
        df_transposed.index.name = "Date"
        df_transposed = df_transposed.sort_index() # Sort chronologically
        
        return df_transposed

    except Exception as e:
        st.error(f"Error loading {sheet_name}: {e}")
        return None

# --- LOAD DATA ---
file_path = "Hind. Unilever (1).xlsx" # Ensure this matches your uploaded filename exactly

# Load key sheets
df_main = load_and_clean_data(file_path, "Data Sheet", "Report Date")
df_ratios = load_and_clean_data(file_path, "Ratio Analysis", "Years")
df_cashflow = load_and_clean_data(file_path, "Cash Flow", "Narration") # 'Narration' is usually the header in Cash Flow

if df_main is not None:
    # --- DASHBOARD LAYOUT ---

    # 1. KEY METRICS SNAPSHOT (Most recent data)
    st.header("1. Key Financial Highlights")
    latest_date = df_main.index[-1]
    last_year_date = df_main.index[-2]
    
    # metrics to display
    metrics = ['Sales', 'Net Profit', 'Operating Profit']
    
    cols = st.columns(len(metrics))
    for i, metric in enumerate(metrics):
        if metric in df_main.columns:
            val_now = df_main.loc[latest_date, metric]
            val_prev = df_main.loc[last_year_date, metric]
            delta = float(val_now) - float(val_prev)
            
            cols[i].metric(
                label=f"{metric} ({latest_date.strftime('%b %Y')})", 
                value=f"₹{val_now:,.0f} Cr", 
                delta=f"{delta:,.0f} Cr"
            )

    # 2. INTERACTIVE TREND CHARTS
    st.divider()
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📈 Revenue & Profit Trend")
        # User selects what to plot
        options = st.multiselect("Select Metrics to Compare:", df_main.columns.tolist(), default=['Sales', 'Net Profit'])
        
        if options:
            fig = px.line(df_main, x=df_main.index, y=options, markers=True)
            # Fix X-Axis to show nice dates like "Mar-18"
            fig.update_xaxes(tickformat="%b-%y", title_text="Year")
            st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("📉 Ratio Analysis")
        if df_ratios is not None:
            ratio_options = st.multiselect("Select Ratios:", df_ratios.columns.tolist(), default=['ROCE %', 'ROE %'])
            if ratio_options:
                fig_ratio = px.line(df_ratios, x=df_ratios.index, y=ratio_options, markers=True)
                fig_ratio.update_xaxes(tickformat="%b-%y", title_text="Year")
                st.plotly_chart(fig_ratio, use_container_width=True)

    # 3. CASH FLOW ANALYSIS
    st.divider()
    st.subheader("💰 Cash Flow Analysis")
    if df_cashflow is not None:
        cf_metrics = ['Cash from Operating Activity', 'Cash from Investing Activity', 'Cash from Financing Activity']
        # Clean up names if they exist
        existing_cf = [m for m in cf_metrics if m in df_cashflow.columns]
        
        if existing_cf:
            fig_cf = px.bar(df_cashflow, x=df_cashflow.index, y=existing_cf, barmode='group')
            fig_cf.update_xaxes(tickformat="%b-%y")
            st.plotly_chart(fig_cf, use_container_width=True)

    # 4. VIEW RAW DATA
    with st.expander("📂 View Raw Data (Cleaned)"):
        st.dataframe(df_main.style.format("{:,.2f}"))

else:
    st.warning("Please upload the Excel file to the repository or ensure the path is correct.")
