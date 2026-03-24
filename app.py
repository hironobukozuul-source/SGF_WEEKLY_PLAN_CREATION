import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Production Schedule Viewer", layout="wide")

st.title("📊 Production & Case Schedule Dashboard")

# Define file paths based on your uploads
CASE_FILE = "Weekly_12,13,14,15,16,17.xlsx - Case.csv"
WEEKLY_FILES = {
    "Week 12": "Weekly_12,13,14,15,16,17.xlsx - Wk12.csv",
    "Week 13": "Weekly_12,13,14,15,16,17.xlsx - Wk13.csv",
    "Week 14": "Weekly_12,13,14,15,16,17.xlsx - Wk14.csv",
    "Week 15": "Weekly_12,13,14,15,16,17.xlsx - Wk15.csv",
    "Week 16": "Weekly_12,13,14,15,16,17.xlsx - Wk16.csv",
    "Week 17": "Weekly_12,13,14,15,16,17.xlsx - Wk17.csv",
}

tab1, tab2 = st.tabs(["📋 Case Summary", "🗓️ Weekly Schedules"])

with tab1:
    st.header("Case Summary Analysis")
    try:
        # The Case file has 4-5 rows of junk/headers. 
        # We read it and try to reconstruct the columns for each production line.
        df_case = pd.read_csv(CASE_FILE, header=None)
        
        # Line names are in row 4, column labels in row 5
        line_names = df_case.iloc[4].fillna(method='ffill')
        col_labels = df_case.iloc[5]
        
        # Clean data (everything from row 6 onwards)
        data = df_case.iloc[6:].reset_index(drop=True)
        
        # Sidebar selection for Line
        unique_lines = [l for l in line_names.unique() if isinstance(l, str) and l != 'wk']
        selected_line = st.selectbox("Select Production Line", unique_lines)
        
        # Filter columns for the selected line
        line_mask = line_names == selected_line
        line_df = data.loc[:, line_mask]
        line_df.columns = col_labels[line_mask]
        
        # Remove empty rows
        line_df = line_df.dropna(subset=['Material', 'date'], how='all')
        
        if not line_df.empty:
            st.subheader(f"Schedule for {selected_line}")
            st.dataframe(line_df, use_container_width=True)
            
            # Simple metric
            if 'ton' in line_df.columns:
                total_ton = pd.to_numeric(line_df['ton'], errors='coerce').sum()
                st.metric("Total Tonnage", f"{total_ton:,.2f} T")
        else:
            st.warning("No data found for this line.")
            
    except Exception as e:
        st.error(f"Error loading Case file: {e}")

with tab2:
    st.header("Weekly Timeline")
    selected_wk = st.selectbox("Select Week", list(WEEKLY_FILES.keys()))
    
    try:
        # Weekly sheets are very wide. We skip the first few rows of meta-info.
        df_wk = pd.read_csv(WEEKLY_FILES[selected_wk], header=None)
        
        # Row 7 usually contains the hour markers (3, 7, 11, 15...)
        # Row 8+ contains the actual production blocks
        st.write(f"Displaying raw data for {selected_wk} (Rows 6-30):")
        
        # Basic cleaning: remove completely empty rows/cols
        display_df = df_wk.iloc[5:50].dropna(axis=1, how='all').dropna(axis=0, how='all')
        
        st.dataframe(display_df, use_container_width=True)
        
        st.info("💡 Note: The weekly view is formatted as a visual grid. Columns represent 4-hour time blocks.")
        
    except Exception as e:
        st.error(f"Error loading {selected_wk}: {e}")
