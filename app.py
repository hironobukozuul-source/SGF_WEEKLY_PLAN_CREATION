import pandas as pd
import datetime
from datetime import timedelta
import openpyxl
from openpyxl.styles import PatternFill

class ProductionGanttSystem:
    def __init__(self):
        # Index 26 = Column AA (where 03:00 usually starts)
        self.BASE_COL_INDEX = 26 
        self.START_HOUR = 3

    def get_monday_start(self, df):
        """Finds the first valid date in Column A and returns the Monday of that week."""
        for val in df.iloc[:, 0]:
            if pd.isna(val) or isinstance(val, str):
                # Try to parse strings, skip if they aren't dates (like "Apr" or "月")
                try:
                    ts = pd.to_datetime(val, errors='coerce')
                    if pd.isna(ts): continue
                    d = ts.date()
                except:
                    continue
            elif isinstance(val, (datetime.datetime, datetime.date)):
                d = val if isinstance(val, datetime.date) else val.date()
            else:
                continue
            
            # If we found a valid date, return the Monday
            return d - timedelta(days=d.weekday())
        return None

    def get_column_for_time(self, time_val):
        """Maps a time object to the specific column index in the Weekly sheet."""
        if pd.isna(time_val): return None
        
        # Ensure we have a time or datetime object
        if isinstance(time_val, str):
            try:
                time_val = pd.to_datetime(time_val).time()
            except:
                return None
        
        h, m = time_val.hour, time_val.minute
        
        # Handle the 3:00 AM start offset
        # 3:00 -> 0, 4:00 -> 1 ... 2:00 (next day) -> 23
        rel_hour = h - self.START_HOUR if h >= self.START_HOUR else h + (24 - self.START_HOUR)
        
        # Each hour has 4 columns (15-min increments)
        return self.BASE_COL_INDEX + (rel_hour * 4) + (m // 15)

    def build_and_fill(self, fill_df):
        start_date = self.get_monday_start(fill_df)
        if not start_date:
            raise ValueError("Could not find a valid start date in Column A of the Fill sheet.")

        # Dictionary to store data per week/sheet
        # Format: { "Wk14": { row_index: { col_index: color_hex } } }
        output_updates = {}

        for idx, row in fill_df.iterrows():
            # 1. Validate Date (Column A)
            raw_date = row[0]
            try:
                task_ts = pd.to_datetime(raw_date, errors='coerce')
                # CRITICAL: Skip if not a valid date (handles 'NaTType' error)
                if pd.isna(task_ts) or task_ts is pd.NaT:
                    continue
                task_date = task_ts.date()
            except:
                continue

            # 2. Calculate which week and day this belongs to
            # (task_date and start_date are both datetime.date now)
            delta_days = (task_date - start_date).days
            wk_num = (delta_days // 7) + 12 # Starting from Week 12
            sheet_name = f"Wk{int(wk_num)}"
            day_offset = delta_days % 7

            # 3. Get Production Times (Column E=Start, F=End)
            start_t = row[4]
            end_t = row[5]
            product_name = str(row[2]) if pd.notnull(row[2]) else "Unknown"

            start_col = self.get_column_for_time(start_t)
            end_col = self.get_column_for_time(end_t)

            if start_col and end_col:
                # Logic to determine Row in Weekly sheet:
                # Based on the file, 'Pump' starts around row 11? 
                # You may need to adjust this row mapping based on Line Name (Col C)
                line_name = str(row[1]).strip()
                base_row = 11 # Default for Pump
                if "Awa" in line_name: base_row = 8
                
                # Each day is usually offset by a certain number of rows?
                # If the sheet is one long vertical list, calculate target_row:
                target_row = base_row # Adjust this if Monday/Tuesday are separated vertically

                if sheet_name not in output_updates:
                    output_updates[sheet_name] = []
                
                output_updates[sheet_name].append({
                    "row": target_row,
                    "start_col": start_col,
                    "end_col": end_col,
                    "label": product_name
                })

        return output_updates, start_date

# --- Usage with openpyxl ---
def apply_to_excel(template_path, updates):
    wb = openpyxl.load_workbook(template_path, keep_vba=True)
    fill_color = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
    
    for sheet_name, tasks in updates.items():
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for t in tasks:
                for c in range(t['start_col'], t['end_col'] + 1):
                    cell = ws.cell(row=t['row'], column=c + 1) # openpyxl is 1-based
                    cell.fill = fill_color
                    if c == t['start_col']:
                        cell.value = t['label']
    
    output_file = "Gantt_Updated.xlsm"
    wb.save(output_file)
    return output_file
