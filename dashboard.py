import streamlit as st
import pandas as pd
import xlwings as xw
import time
import os

# --- Configuration ---
EXCEL_FILE = 'Live_DAP.xlsx'
SHEET_NAME = 'DAP_Main' # Sheet containing the main market data
CSV_FALLBACK = 'Live_DAP.xlsx - DAP_Main.csv' # Fallback file path

@st.cache_data(ttl=2) # Cache the data for 2 seconds to simulate a near-live refresh rate
def get_live_data():
    """
    Connects to the open Excel workbook using xlwings and reads the data.
    If the connection fails, it falls back to reading the static CSV file.
    """
    try:
        # --- XLWINGS LIVE CONNECTION ---
        # NOTE: The Excel workbook must be open for xlwings to capture live updates
        book = xw.Book(EXCEL_FILE)
        sheet = book.sheets[SHEET_NAME]
        
        # Read the data, assuming the market table starts at A1 and has a header
        # The 'expand='table'' option is a robust way to read the whole block.
        data = sheet.range('A1').options(pd.DataFrame, header=1, index=False, expand='table').value
        
        st.success("Successfully fetched data from live Excel sheet via xlwings.")

    except Exception as e:
        # --- FALLBACK TO STATIC CSV ---
        st.error(f"Failed to connect to Excel with xlwings: {e}")
        st.info(f"Falling back to reading the static CSV snapshot: {CSV_FALLBACK}")
        
        try:
            # Read the CSV snapshot provided by the user
            if os.path.exists(CSV_FALLBACK):
                 # Skip the first row as the CSV header seems to be duplicated
                data = pd.read_csv(CSV_FALLBACK, header=1) 
            else:
                st.error(f"Fallback CSV file not found: {CSV_FALLBACK}")
                return pd.DataFrame()
        except Exception as csv_e:
            st.error(f"Failed to read fallback CSV: {csv_e}")
            return pd.DataFrame()

    # --- Post-Processing ---
    display_columns = ['Outrights', 'Last', 'Settle', 'BidQty', 'Bid', 'Ask', 'AskQty', 'VWAP']
    
    # Clean up and select columns
    data = data.dropna(subset=['Outrights'])
    
    return data[display_columns]

def main():
    st.set_page_config(layout="wide")
    st.title("📈 Live Market Prices Dashboard")
    st.markdown("---")
    
    # -------------------------------------------------------------------------
    # Manual Refresh Section
    # -------------------------------------------------------------------------
    st.header("Manual Data Refresh")
    col1, col2 = st.columns([1, 4])
    
    # Button to force data reload, clearing the cache
    if col1.button("Force Refresh from Excel"):
        # The short 'ttl=2' on the cache means a button click will almost always fetch new data.
        st.balloons()
        
    col2.info("Press the button to force a new read from the live Excel sheet. This is necessary to view updated prices if your Excel sheet is actively calculating.")
    
    st.markdown("---")

    # -------------------------------------------------------------------------
    # Live Data Display Section
    # -------------------------------------------------------------------------
    st.header("DAP Main Market Prices - Live View")
    
    # Fetch the data
    df = get_live_data()
    
    if not df.empty:
        # Display the DataFrame as a highly formatted table
        st.dataframe(
            df, 
            hide_index=True,
            use_container_width=True,
            # Configure columns for better display of market data
            column_config={
                "Last": st.column_config.NumberColumn("Last", format="%.3f"),
                "Settle": st.column_config.NumberColumn("Settle", format="%.3f"),
                "Bid": st.column_config.NumberColumn("Bid", format="%.3f"),
                "Ask": st.column_config.NumberColumn("Ask", format="%.3f"),
                "VWAP": st.column_config.NumberColumn("VWAP", format="%.4f"),
            }
        )
        st.caption(f"Last updated: **{time.strftime('%H:%M:%S')}**")
    else:
        st.error("No data to display. Please check file paths and ensure your Excel file is open if using xlwings.")


if __name__ == "__main__":
    main()
