import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from io import BytesIO

# --- GANTT ENGINE ---
class GanttEngine:
    def __init__(self):
        self.START_HOUR = 3  # Day starts at 03:00
        self.BASE_COL = 27   # Based on Wk14.csv, 03:00 is roughly Col AA (27)
        self.COL_PER_HOUR = 4 # 15 min increments

    def get_column_index(self, time_obj):
        if time_obj is None: return None
        h, m = time_obj.hour, time_obj.minute
        # Adjust for 3 AM start
        rel_h = h + (24 - self.START_HOUR) if h < self.START_HOUR else h - self.START_HOUR
        return self.BASE_COL + (rel_h * self.COL_PER_HOUR) + (m // 15)

    def process(self, fill_file, template_file, week_name):
        # Read Input
        df = pd.read_excel(fill_file, sheet_name='Fill')
        wb = openpyxl.load_workbook(template_file)
        
        if week_name not in wb.sheetnames:
            return None, f"Sheet {week_name} not found in template!"
        
        ws = wb[week_name]
        
        # Simple Line-to-Row Mapping (Customize this for your lines)
        line_map = {"Pump": 11, "Ref1": 15, "Ref2": 19} 

        for _, row in df.iterrows():
            # Adjust these indices based on your 'Fill' sheet columns
            line_name = str(row.iloc[2]) # Col C
            start_t = row.iloc[4]        # Col E
            finish_t = row.iloc[5]       # Col F
            prod_code = str(row.iloc[2])

            if pd.isna(start_t) or pd.isna(finish_t): continue

            start_col = self.get_column_index(start_t)
            end_col = self.get_column_index(finish_t)
            
            # Row mapping logic
            target_row = 11 # Default row
            for key, val in line_map.items():
                if key in line_name: target_row = val

            # Color Logic
            color = "FFCC00" if "DVB" in prod_code else "99CCFF"
            fill = PatternFill(start_color=color, end_color=color, fill_type='solid')

            for c in range(start_col, end_col):
                ws.cell(row=target_row, column=c).fill = fill
        
        # Save to memory buffer
        out_buf = BytesIO()
        wb.save(out_buf)
        return out_buf, None

# --- STREAMLIT UI ---
st.set_page_config(page_title="Production Gantt Maker", layout="wide")
st.title("📊 Production Schedule Automator")
st.markdown("Upload your **Fill** file and **Weekly Template** to generate the schedule.")

with st.sidebar:
    st.header("Settings")
    week_selection = st.selectbox("Select Week Sheet", ["Wk14", "Wk15", "Wk16", "Wk17"])

col1, col2 = st.columns(2)
with col1:
    fill_input = st.file_uploader("Upload 'Fill' File (xlsm)", type=['xlsm', 'xlsx'])
with col2:
    temp_input = st.file_uploader("Upload 'Weekly' Template (xlsx)", type=['xlsx'])

if fill_input and temp_input:
    if st.button("Generate Gantt Chart"):
        engine = GanttEngine()
        result, error = engine.process(fill_input, temp_input, week_selection)
        
        if error:
            st.error(error)
        else:
            st.success("Gantt Chart Generated!")
            st.download_button(
                label="📥 Download Result",
                data=result.getvalue(),
                file_name=f"Generated_Weekly_{week_selection}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
