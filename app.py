import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, PatternFill, Border, Side, Font
from datetime import datetime, time, timedelta, date
from io import BytesIO

class ProductionGanttSystem:
    def __init__(self):
        self.START_HOUR = 3  # Day starts at 03:00 AM
        self.BASE_COL = 27   # Column AA
        self.COL_PER_DAY = 96
        self.DAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
        
        # Mapping lines to specific rows in the generated sheet
        self.LINE_ROWS = {
            "Pump": 11,
            "Ref-1": 14,
            "Ref-2": 17,
            "Ref-3": 20,
            "Ref-4": 23,
            "Mini-1": 26
        }

    def build_and_fill(self, fill_df, start_date):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"Wk_{start_date.isocalendar()[1]}"
        
        # Styling
        thin = Side(border_style="thin")
        border = Border(top=thin, left=thin, right=thin, bottom=thin)
        center_align = Alignment(horizontal="center", vertical="center")
        
        # 1. SETUP THE TEMPLATE GRID
        for d_idx, day_name in enumerate(self.DAYS):
            col_start = self.BASE_COL + (d_idx * self.COL_PER_DAY)
            # Ensure we are working with date objects for math
            current_date = start_date + timedelta(days=d_idx)
            
            # Day Header (Merged across 96 columns)
            ws.merge_cells(start_row=8, start_column=col_start, end_row=8, end_column=col_start + 95)
            day_cell = ws.cell(row=8, column=col_start, value=f"{day_name} {current_date}")
            day_cell.alignment = center_align
            day_cell.font = Font(bold=True)
            
            # Hour Labels (3:00 to 2:00 next day)
            for h in range(24):
                hr_val = (self.START_HOUR + h) % 24
                c = col_start + (h * 4)
                ws.merge_cells(start_row=9, start_column=c, end_row=9, end_column=c + 3)
                hr_cell = ws.cell(row=9, column=c, value=f"{hr_val}:00")
                hr_cell.alignment = center_align
                hr_cell.border = border

        # 2. POPULATE PRODUCTION DATA
        for idx, row in fill_df.iterrows():
            if idx < 3: continue # Skip headers in the XLSM
            
            # Extract data from Fill sheet columns
            task_date = row[0]   # Column A
            line_name = str(row[1]) # Column B (e.g. "Pump")
            prod_name = str(row[2]) # Column C
            start_t = row[4]    # Column E
            end_t = row[5]      # Column F

            if not isinstance(start_t, time) or not isinstance(end_t, time):
                continue

            # Calculate Day Offset relative to the selected start_date
            try:
                # Convert task_date to date object if it's a timestamp
                if hasattr(task_date, 'date'):
                    actual_date = task_date.date()
                else:
                    actual_date = task_date
                day_offset = (actual_date - start_date).days
            except:
                continue # Skip if date is invalid or out of range
            
            if day_offset < 0 or day_offset > 6: continue

            # Color Logic
            color = "D3D3D3" # Default Gray for setups
            if "DVB" in prod_name: color = "FFCC00" # Yellow
            elif "LSL" in prod_name: color = "99CCFF" # Blue
            
            # Determine Column indices
            sc = self.get_col_index(start_t, day_offset)
            ec = self.get_col_index(end_t, day_offset)
            
            # Find the row for this line
            target_row = self.LINE_ROWS.get(line_name, 11) # Default to row 11 if unknown
            
            # Paint the grid
            fill = PatternFill("solid", fgColor=color)
            for col_idx in range(sc, ec):
                ws.cell(row=target_row, column=col_idx).fill = fill
                if col_idx == sc:
                    ws.cell(row=target_row, column=col_idx).value = prod_name
                    ws.cell(row=target_row, column=col_idx).font = Font(size=8)

        # Final Formatting for Line Names
        for name, r in self.LINE_ROWS.items():
            ws.cell(row=r, column=2, value=name).font = Font(bold=True)

        output = BytesIO()
        wb.save(output)
        return output.getvalue()

    def get_col_index(self, t, day_off):
        rel_h = t.hour - self.START_HOUR
        if t.hour < self.START_HOUR: rel_h += 24
        return self.BASE_COL + (day_off * self.COL_PER_DAY) + (rel_h * 4) + (t.minute // 15)

# --- Streamlit Frontend ---
st.set_page_config(page_title="Gantt Automator", layout="wide")
st.title("🏭 SGF Weekly Plan Creator")

with st.sidebar:
    st.header("Settings")
    # Streamlit returns a datetime.date object here
    monday_start = st.date_input("Select Monday Start Date", date(2026, 3, 30))
    uploaded_file = st.file_uploader("Upload Fill_XX.xlsm", type=["xlsm"])

if uploaded_file:
    if st.button("Generate Weekly Schedule"):
        try:
            # Read the XLSM
            df = pd.read_excel(uploaded_file, sheet_name='Fill', header=None)
            
            engine = ProductionGanttSystem()
            excel_data = engine.build_and_fill(df, monday_start)
            
            st.success("✅ Schedule Generated Successfully!")
            st.download_button(
                label="📥 Download Excel Result",
                data=excel_data,
                file_name=f"Weekly_Plan_{monday_start}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"An error occurred: {e}")
