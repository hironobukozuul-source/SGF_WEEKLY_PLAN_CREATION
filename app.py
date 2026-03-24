import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Production Schedule Viewer", layout="wide")

st.title("📊 Dynamic Production Dashboard")
st.info("Upload your 'Case' and 'Weekly' CSV or Excel files below to begin.")

# 1. File Uploader (Accepts multiple files)
uploaded_files = st.sidebar.file_uploader(
    "Upload Production Files", 
    type=["csv", "xlsx", "xlsm"], 
    accept_multiple_files=True
)

if not uploaded_files:
    st.warning("Please upload files in the sidebar to see the data.")
    st.stop()

# Organize files into categories
case_file = None
weekly_files = {}

for f in uploaded_files:
    fname = f.name.lower()
    if "case" in fname:
        case_file = f
    elif "wk" in fname or "week" in fname:
        # Extract week number if possible, or just use filename
        weekly_files[f.name] = f

# --- TAB 1: CASE SUMMARY ---
tab1, tab2 = st.tabs(["📋 Case Summary", "🗓️ Weekly Schedules"])

with tab1:
    if case_file:
        st.header(f"Analysis: {case_file.name}")
        try:
            # Flexible reading: CSV vs Excel
            if case_file.name.endswith('.csv'):
                df_case = pd.read_csv(case_file, header=None)
            else:
                df_case = pd.read_excel(case_file, header=None)

            # Logic to find the data:
            # We look for the row that contains 'Material' to set as header
            header_idx = df_case[df_case.stack().str.contains('Material', na=False).any(level=0)].index[0]
            
            # Line names are usually 1 row above 'Material'
            line_names = df_case.iloc[header_idx - 1].fillna(method='ffill')
            col_labels = df_case.iloc[header_idx]
            
            data = df_case.iloc[header_idx + 1:].reset_index(drop=True)
            
            unique_lines = [l for l in line_names.unique() if isinstance(l, str) and l.lower() != 'wk']
            selected_line = st.selectbox("Select Production Line", unique_lines)
            
            line_mask = line_names == selected_line
            line_df = data.loc[:, line_mask]
            line_df.columns = col_labels[line_mask]
            
            # Clean up empty rows
            line_df = line_df.dropna(subset=[col for col in line_df.columns if 'Material' in str(col)], how='all')
            
            st.dataframe(line_df, use_container_width=True)
        except Exception as e:
            st.error(f"Could not parse Case file: {e}")
    else:
        st.info("Upload a file with 'Case' in the name to see Case Summary.")

# --- TAB 2: WEEKLY SCHEDULES ---
with tab2:
    if weekly_files:
        selected_wk_name = st.selectbox("Select Weekly Sheet", list(weekly_files.keys()))
        target_file = weekly_files[selected_wk_name]
        
        try:
            if target_file.name.endswith('.csv'):
                df_wk = pd.read_csv(target_file, header=None)
            else:
                df_wk = pd.read_excel(target_file, header=None)
            
            # Find the row with time markers (3, 7, 11...) or Day names (MON, TUE)
            # For your specific files, row 5 or 6 usually starts the visual grid
            start_row = st.slider("Adjust starting row (to skip headers)", 0, 20, 5)
            
            display_df = df_wk.iloc[start_row:].dropna(axis=1, how='all').dropna(axis=0, how='all')
            st.dataframe(display_df, use_container_width=True)
            
        except Exception as e:
            st.error(f"Error loading weekly sheet: {e}")
    else:
        st.info("Upload files with 'Wk' or 'Week' in the name to see Weekly Timeline.")
