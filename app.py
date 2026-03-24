import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font
from datetime import datetime, time, timedelta
from io import BytesIO

class ProductionGanttSystem:
    def __init__(self):
        self.START_HOUR = 3  # Day starts at 03:00
        self.BASE_COL = 27   # Column AA
        self.COL_PER_DAY = 96
        self.DAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]

    def build_and_fill(self, fill_df, start_date):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Wk_{start_date.isocalendar()[1]}"
        
        # 1. SETUP THE TEMPLATE (Styling & Headers)
        thin = Side(border_style="thin")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        
        # Draw Time Scale for each day
        for d_idx, day_name in enumerate(self.DAYS):
            col_start = self.BASE_COL + (d_idx * self.COL_PER_DAY)
            day_dt = start_date + timedelta(days=d_idx)
            
            # Day Header
            ws.merge_cells(start_row=8, start_column=col_start, end_row=8, end_column=col_start+95)
            ws.cell(row=8, column=col_start, value=f"{day_name} {day_dt.date()}").alignment = Alignment(horizontal="center")
            
            # Hour Labels (3, 7, 11, 15, 19, 23 cycle)
            for h in range(24):
                hr_val = (self.START_HOUR + h) % 24
                c = col_start + (h * 4)
                ws.merge_cells(start_row=9, start_column=c, end_row=9, end_column=c+3)
                cell = ws.cell(row=9, column=c, value=f"{hr_val}:00")
                cell.alignment = Alignment(horizontal="center")
                cell.border = border

        # 2. POPULATE PRODUCTION DATA
        line_rows = {"Pump": 11, "Ref-1": 14, "Ref-2": 17, "Ref-3": 20} # Expand as needed
        
        for idx, row in fill_df.iterrows():
            if idx < 3: continue # Skip headers
            
            date_val = row[0]
            prod_name = str(row[2])
            start_t = row[4]
            end_t = row[5]

            if not isinstance(start_t, time) or not isinstance(end_t, time): continue

            # Calculate Day Offset (0=Mon, 1=Tue, etc.)
            try:
                day_offset = (date_val.date() - start_date.date()).days
            except: day_offset = 0 
            
            # Color Mapping
            color = "D3D3D3" # Default Setup (Gray)
            if "DVB" in prod_name: color = "FFCC00" # Yellow
            elif "LSL" in prod_name: color = "99CCFF" # Blue
            
            # Calculate Columns
            sc = self.get_col(start_t, day_offset)
            ec = self.get_col(end_t, day_offset)
            
            # Paint the cells
            target_row = line_rows.get("Pump", 11) # Logic to detect line from prod_name goes here
            for c in range(sc, ec):
                ws.cell(row=target_row, column=c).fill = PatternFill("solid", fgColor=color)
                if c == sc: ws.cell(row=target_row, column=c).value = prod_name

        output = BytesIO()
        wb.save(output)
        return output.getvalue()

    def get_col(self, t, day_off):
        rel_h = t.hour - self.START_HOUR
        if t.hour < self.START_HOUR: rel_h += 24
        return self.BASE_COL + (day_off * 96) + (rel_h * 4) + (t.minute // 15)

# --- STREAMLIT UI ---
st.title("🚀 Auto-Gantt Engine")
uploaded_file = st.file_uploader("Upload Fill_XX.xlsm", type=["xlsm"])
start_date = st.date_with_week = st.date_input("Select Monday Start Date", datetime(2026, 3, 30))

if uploaded_file:
    if st.button("Build From Scratch"):
        df = pd.read_excel(uploaded_file, sheet_name='Fill', header=None)
        engine = ProductionGanttSystem()
        final_excel = engine.build_and_fill(df, start_date)
        
        st.download_button("📥 Download Generated Schedule", final_excel, "Weekly_Output.xlsx")
