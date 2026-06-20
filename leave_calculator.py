import sys
import os
import glob
from datetime import date, timedelta
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                              QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                              QTextEdit, QTabWidget, QTableWidget, QTableWidgetItem,
                              QHeaderView, QMessageBox, QStatusBar, QProgressBar)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor
import sys, os

if getattr(sys, 'frozen', False):
    # Running as a PyInstaller bundle
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ATTENDANCE_FOLDER = os.path.join(BASE_DIR, "attendance_2026")
EMPLOYEES_FILE    = os.path.join(BASE_DIR, "employees.xlsx")
BALANCE_2025_FILE = os.path.join(BASE_DIR, "balance_2025.xlsx")
HOLIDAYS_FILE     = os.path.join(BASE_DIR, "holidays.xlsx")
NOT_ELIGIBLE_FILE = os.path.join(BASE_DIR, "not_eligible_shifts.xlsx")
QUEUE_NATURE_FILE = os.path.join(BASE_DIR, "Queue_Nature.xlsx")
# ========== FILE PATHS ==========
ATTENDANCE_FOLDER = "./attendance_2026"
EMPLOYEES_FILE    = "employees.xlsx"
BALANCE_2025_FILE = "balance_2025.xlsx"
HOLIDAYS_FILE     = "holidays.xlsx"
NOT_ELIGIBLE_FILE = "not_eligible_shifts.xlsx"
QUEUE_NATURE_FILE = "Queue_Nature.xlsx"

CUTOFF_DATE = date.today()
PER_MIN_DAYS = 52

# Leave types that deduct from annual balance with their weights
LEAVE_TYPE_WEIGHTS = {
    "annual": 1.0,
    "half annual": 0.5,
    "casual": 1.0,
    "annual exception": 1.0
}

class DataLoader(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    
    def run(self):
        try:
            data = {}
            
            # --- 1. Queue Nature (FTE / Per Minute) ---
            self.progress.emit("Loading Queue Nature...")
            if os.path.exists(QUEUE_NATURE_FILE):
                qn = pd.read_excel(QUEUE_NATURE_FILE)
                qn.columns = [c.strip() for c in qn.columns]
                nature_col = next((c for c in qn.columns if c.lower() == 'nature'), None)
                queue_col = next((c for c in qn.columns if c.lower() == 'queue'), None)
                if nature_col and queue_col:
                    qn = qn.rename(columns={nature_col: 'Nature', queue_col: 'Queue'})
                    fte_queues = qn[qn['Nature'].str.upper() == 'FTE']['Queue'].tolist()
                    per_min_queues = qn[qn['Nature'].str.upper() != 'FTE']['Queue'].tolist()
                else:
                    fte_queues, per_min_queues = [], []
            else:
                fte_queues, per_min_queues = [], []
            
            # --- 2. Employees ---
            self.progress.emit("Loading employees...")
            employees = pd.read_excel(EMPLOYEES_FILE)
            # Rename known columns
            col_map = {
                'UAE ID': 'uae_id',
                'ACD ID': 'acd_id',
                'Hiring Date': 'hiring_date',
                'LWD': 'lwd',
                'Date of Certification': 'certification_date'
            }
            employees.rename(columns={k:v for k,v in col_map.items() if k in employees.columns}, inplace=True)
            # Convert date columns
            for dc in ['hiring_date', 'lwd', 'certification_date']:
                if dc in employees.columns:
                    employees[dc] = pd.to_datetime(employees[dc], errors='coerce', dayfirst=True)
            
            # Unify agent name column
            if 'Agent Name' in employees.columns:
                employees.rename(columns={'Agent Name': 'Name'}, inplace=True)
            if 'Name' not in employees.columns:
                employees['Name'] = 'Unknown'
            
            # Capture Title and Queue (Q)
            if 'Title' in employees.columns:
                pass  # already exists
            else:
                employees['Title'] = ''
            
            # For Queue (maybe column 'Q' or 'Queue')
            if 'Q' in employees.columns:
                employees.rename(columns={'Q': 'Queue'}, inplace=True)
            elif 'Queue' in employees.columns:
                pass
            else:
                employees['Queue'] = ''
            
            # --- 3. 2025 Balance ---
            self.progress.emit("Loading 2025 balance...")
            balance_2025 = pd.read_excel(BALANCE_2025_FILE)
            if 'Annual_Used' in balance_2025.columns:
                balance_2025.rename(columns={'UAE ID':'uae_id', 'Annual_Used':'annual_used_2025'}, inplace=True)
            elif 'Annual' in balance_2025.columns:
                balance_2025.rename(columns={'UAE ID':'uae_id', 'Annual':'annual_used_2025'}, inplace=True)
            else:
                balance_2025.rename(columns={'UAE ID':'uae_id'}, inplace=True)
                balance_2025['annual_used_2025'] = 0
            
            # --- 4. Holidays (flexible) ---
            self.progress.emit("Loading holidays...")
            holidays_raw = pd.read_excel(HOLIDAYS_FILE)
            holidays_raw.columns = ['Holiday_Date', 'Holiday_Name']
            holidays_raw['Holiday_Date'] = pd.to_datetime(holidays_raw['Holiday_Date'], errors='coerce')
            holidays = holidays_raw
            
            # --- 5. Not eligible shifts ---
            not_eligible_list = []
            if os.path.exists(NOT_ELIGIBLE_FILE):
                try:
                    ne_df = pd.read_excel(NOT_ELIGIBLE_FILE)
                except:
                    ne_df = pd.read_csv(NOT_ELIGIBLE_FILE)
                if not ne_df.empty:
                    not_eligible_list = ne_df.iloc[:,0].astype(str).str.strip().str.lower().tolist()
            
            # --- 6. Attendance ---
            self.progress.emit("Loading attendance files...")
            all_files = sorted(glob.glob(f"{ATTENDANCE_FOLDER}/*.csv"))
            if not all_files:
                raise Exception("No attendance files found in the specified folder")
            att_list = []
            for i, f in enumerate(all_files):
                self.progress.emit(f"Reading file {i+1}/{len(all_files)}...")
                df = pd.read_csv(f, dayfirst=True)
                rename_att = {
                    'ACD ID': 'acd_id',
                    'Date of Join': 'date_of_join',
                    'Final Status': 'final_status',
                    'Queue': 'queue'
                }
                df.rename(columns={k:v for k,v in rename_att.items() if k in df.columns}, inplace=True)
                for dc in ['Date', 'date_of_join']:
                    if dc in df.columns:
                        df[dc] = pd.to_datetime(df[dc], errors='coerce', dayfirst=True)
                att_list.append(df)
            attendance = pd.concat(att_list, ignore_index=True)
            
            # --- 7. Determine queue from first attendance (overwrite if missing) ---
            if 'queue' in attendance.columns and 'acd_id' in attendance.columns:
                first_att = attendance.sort_values('Date').groupby('acd_id').first().reset_index()
                first_att = first_att[['acd_id', 'queue']]
                employees['acd_str'] = employees['acd_id'].astype(str).str.strip()
                first_att['acd_str'] = first_att['acd_id'].astype(str).str.strip()
                # Only fill Queue from attendance if employee doesn't already have one
                employees = employees.merge(first_att[['acd_str', 'queue']], on='acd_str', how='left', suffixes=('', '_att'))
                employees['Queue'] = employees['Queue'].replace('', pd.NA).fillna(employees['queue']).fillna('')
                employees.drop(columns=['queue'], inplace=True)
            
            # --- 8. Go Live Date ---
            self.progress.emit("Calculating Go Live...")
            def calc_go_live(row):
                cert = row.get('certification_date')
                if pd.isna(cert):
                    return pd.NaT
                q = str(row.get('Queue', '')).strip()
                if q in fte_queues:
                    return cert + timedelta(days=1)
                elif q in per_min_queues:
                    return cert - timedelta(days=PER_MIN_DAYS)
                else:
                    return cert + timedelta(days=1)  # default FTE
            employees['go_live'] = employees.apply(calc_go_live, axis=1)
            
            # --- 9. Link UAE ID to attendance ---
            self.progress.emit("Linking UAE ID...")
            attendance['acd_str'] = attendance['acd_id'].astype(str).str.strip()
            valid_dates = attendance['date_of_join'].notna()
            attendance.loc[valid_dates, 'join_key'] = (
                attendance.loc[valid_dates, 'acd_str'] + '_' + 
                attendance.loc[valid_dates, 'date_of_join'].dt.strftime('%Y%m%d')
            )
            valid_hiring = employees['hiring_date'].notna()
            employees.loc[valid_hiring, 'join_key'] = (
                employees.loc[valid_hiring, 'acd_str'] + '_' + 
                employees.loc[valid_hiring, 'hiring_date'].dt.strftime('%Y%m%d')
            )
            attendance = attendance.merge(
                employees[['join_key', 'uae_id']].dropna(subset=['join_key']), 
                on='join_key', how='left'
            )
            attendance = attendance.dropna(subset=['uae_id'])
            attendance['final_status_clean'] = attendance['final_status'].str.strip().str.lower()
            
            # --- 10. Used annual leave (2026) ---
            self.progress.emit("Calculating used leaves...")
            leave_mask = attendance['final_status_clean'].isin(LEAVE_TYPE_WEIGHTS.keys())
            leave_days = attendance[leave_mask].copy()
            leave_days['day_weight'] = leave_days['final_status_clean'].map(LEAVE_TYPE_WEIGHTS)
            leave_days = leave_days.merge(holidays, left_on='Date', right_on='Holiday_Date', how='left')
            leave_days = leave_days.merge(employees[['uae_id', 'go_live']], on='uae_id', how='left')
            
            leave_days['exclude'] = (
                leave_days['Holiday_Date'].notna() &
                (leave_days['Date'] >= leave_days['go_live']) &
                (~leave_days['final_status_clean'].isin(not_eligible_list))
            )
            leave_days['exclusion_weight'] = leave_days['day_weight'].where(leave_days['exclude'], 0.0)
            summary_2026 = leave_days.groupby('uae_id').agg(
                gross_taken=('day_weight', 'sum'),
                excluded_weight=('exclusion_weight', 'sum')
            ).reset_index()
            summary_2026['net_used_2026'] = summary_2026['gross_taken'] - summary_2026['excluded_weight']
            
            # --- 11. Public holidays earned & used ---
            self.progress.emit("Calculating public holiday balance...")
            all_hol = attendance.merge(holidays, left_on='Date', right_on='Holiday_Date', how='inner')
            all_hol = all_hol.merge(employees[['uae_id', 'go_live']], on='uae_id', how='left')
            public_earned = all_hol[
                (all_hol['Date'] >= all_hol['go_live']) &
                (all_hol['final_status_clean'].isin(['available'])) &
                (~all_hol['final_status_clean'].isin(not_eligible_list))
            ].groupby('uae_id').size().reset_index(name='public_holidays_earned')
            
            comp_days = attendance[attendance['final_status_clean'] == 'comp'].copy()
            comp_days = comp_days.merge(employees[['uae_id', 'go_live']], on='uae_id', how='left')
            comp_used = comp_days[comp_days['Date'] >= comp_days['go_live']].groupby('uae_id').size().reset_index(name='comp_days_used')
            
            # --- 12. Annual leave entitlement ---
            self.progress.emit("Calculating entitlements...")
            emp_master = employees[['uae_id', 'hiring_date', 'lwd', 'go_live', 'certification_date', 'Queue', 'acd_id', 'Name', 'Title']].copy()
            emp_master['calc_date'] = emp_master['lwd'].fillna(pd.Timestamp(CUTOFF_DATE))
            
            def calc_entitlement(row):
                h = row['hiring_date']
                c = row['calc_date']
                if pd.isna(h) or pd.isna(c): return 0.0
                days = (c - h).days
                years = days / 365.25
                if years < 1:
                    return round((days / 365.25) * 15, 2)
                else:
                    full = int(years)
                    rem = days - (full * 365.25)
                    return round(15 + (full - 1) * 21 + (rem / 365.25) * 21, 2)
            
            emp_master['total_earned'] = emp_master.apply(calc_entitlement, axis=1)
            
            # --- 13. Final merge ---
            self.progress.emit("Assembling results...")
            final = emp_master.merge(balance_2025[['uae_id', 'annual_used_2025']], on='uae_id', how='left')
            final = final.merge(summary_2026[['uae_id', 'net_used_2026']], on='uae_id', how='left')
            final = final.merge(public_earned, on='uae_id', how='left')
            final = final.merge(comp_used, on='uae_id', how='left')
            
            for c in ['annual_used_2025', 'net_used_2026']:
                final[c] = final[c].fillna(0.0)
            for c in ['public_holidays_earned', 'comp_days_used']:
                final[c] = final[c].fillna(0).astype(int)
            
            final['total_used_annual'] = final['annual_used_2025'] + final['net_used_2026']
            final['remaining_annual'] = final['total_earned'] - final['total_used_annual']
            final['public_holidays_remaining'] = final['public_holidays_earned'] - final['comp_days_used']
            
            for col, na_val in [('lwd','Active'), ('go_live','N/A'), ('certification_date','N/A')]:
                if col in final.columns:
                    final[f'{col}_display'] = final[col].dt.strftime('%Y-%m-%d').fillna(na_val)
            final['calc_date_display'] = final['calc_date'].dt.strftime('%Y-%m-%d')
            final['hiring_date_display'] = final['hiring_date'].dt.strftime('%Y-%m-%d')
            final['service_years'] = ((final['calc_date'] - final['hiring_date']).dt.days / 365.25).round(2)
            
            data['final_data'] = final
            data['attendance'] = attendance
            data['employees'] = employees
            data['holidays'] = holidays
            data['not_eligible_list'] = not_eligible_list
            
            self.progress.emit("✅ Ready!")
            self.finished.emit(data)
        except Exception as e:
            import traceback
            self.error.emit(f"{str(e)}\n{traceback.format_exc()}")

class LeaveCalculator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.final_data = None
        self.attendance = None
        self.employees = None
        self.holidays = None
        self.not_eligible_list = []
        self.initUI()
        self.load_data_thread()
    
    def initUI(self):
        self.setWindowTitle("Leave Balance Calculator")
        self.setGeometry(100, 100, 950, 680)
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        
        title = QLabel("Leave Balance Calculator - Egyptian Labor Law No.14 of 2025")
        title.setFont(QFont("Arial", 12, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        input_frame = QWidget()
        input_frame.setStyleSheet("background-color: #ecf0f1; border-radius: 10px; padding: 15px;")
        il = QHBoxLayout(input_frame)
        il.addWidget(QLabel("🔍 UAE ID:"))
        self.uae_input = QLineEdit()
        self.uae_input.setPlaceholderText("Enter UAE ID...")
        self.uae_input.setMinimumHeight(35)
        self.uae_input.setEnabled(False)
        il.addWidget(self.uae_input)
        
        self.calc_btn = QPushButton("Calculate Balance")
        self.calc_btn.clicked.connect(self.calculate_balance)
        self.calc_btn.setMinimumHeight(35)
        self.calc_btn.setEnabled(False)
        self.calc_btn.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold;")
        il.addWidget(self.calc_btn)
        
        self.detail_btn = QPushButton("📅 Holiday Details")
        self.detail_btn.clicked.connect(self.show_holiday_details)
        self.detail_btn.setMinimumHeight(35)
        self.detail_btn.setEnabled(False)
        self.detail_btn.setStyleSheet("background-color: #2980b9; color: white;")
        il.addWidget(self.detail_btn)
        layout.addWidget(input_frame)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(0)
        self.progress_label = QLabel("⏳ Loading data...")
        self.progress_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.progress_label)
        layout.addWidget(self.progress_bar)
        
        self.tabs = QTabWidget()
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setFont(QFont("Courier New", 11))
        self.result_text.setStyleSheet("background-color: #2c3e50; color: #ecf0f1; padding: 15px;")
        self.tabs.addTab(self.result_text, "📋 Results")
        
        self.holiday_table = QTableWidget()
        self.tabs.addTab(self.holiday_table, "📅 Holidays")
        layout.addWidget(self.tabs)
        
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
    
    def load_data_thread(self):
        self.loader = DataLoader()
        self.loader.progress.connect(lambda m: (self.progress_label.setText(f"⏳ {m}"), self.status_bar.showMessage(m)))
        self.loader.finished.connect(self.on_data_loaded)
        self.loader.error.connect(lambda m: QMessageBox.critical(self, "Error", m))
        self.loader.start()
    
    def on_data_loaded(self, data):
        self.final_data = data['final_data']
        self.attendance = data['attendance']
        self.employees = data['employees']
        self.holidays = data['holidays']
        self.not_eligible_list = data['not_eligible_list']
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(100)
        self.progress_label.setText("✅ Ready!")
        self.uae_input.setEnabled(True)
        self.calc_btn.setEnabled(True)
        self.detail_btn.setEnabled(True)
        self.uae_input.setFocus()
        self.status_bar.showMessage(f"✅ {len(self.final_data)} employees loaded")
    
    def find_employee(self, uid):
        emp = self.final_data[self.final_data['uae_id'].astype(str).str.strip() == uid]
        if len(emp) == 0:
            emp = self.final_data[self.final_data['uae_id'].astype(str).str.strip() == "UAE" + uid]
        return emp.iloc[0] if len(emp) > 0 else None
    
    def calculate_balance(self):
        uid = self.uae_input.text().strip()
        if not uid: return
        emp = self.find_employee(uid)
        if emp is None:
            QMessageBox.information(self, "Not Found", f"UAE: {uid}")
            return
        result = f"""
╔══════════════════════════════════════╗
║   Leave Balance Report - Law 14/2025  ║
╚══════════════════════════════════════╝
👤 UAE ID      : {emp['uae_id']}
🆔 ACD ID      : {emp.get('acd_id','?')}
📛 Name        : {emp.get('Name','?')}
🏷️  Title       : {emp.get('Title','?')}
📡 Queue (Q)   : {emp.get('Queue','?')}
──────────────────────────────────────
📅 Hiring Date : {emp.get('hiring_date_display','?')}
📜 Certified   : {emp.get('certification_date_display','?')}
🚀 Go Live     : {emp.get('go_live_display','?')}
🏁 Last Working Day: {emp.get('lwd_display','?')}
⏳ Service     : {emp['service_years']} years
──────────────────────────────────────
📋 Annual Leave:
  Entitled     : {emp['total_earned']} days
  Used 2025    : {emp['annual_used_2025']} days
  Used 2026    : {emp['net_used_2026']} days
  ➤ Remaining  : {emp['remaining_annual']} days
──────────────────────────────────────
🎉 Public Holidays:
  Earned       : {emp['public_holidays_earned']} days
  Used (Comp)  : {emp['comp_days_used']} days
  ➤ Remaining  : {emp['public_holidays_remaining']} days
"""
        self.result_text.setText(result)
        self.tabs.setCurrentIndex(0)
    
    def show_holiday_details(self):
        uid = self.uae_input.text().strip()
        if not uid: return
        emp = self.find_employee(uid)
        if emp is None: return
        
        go_live = emp.get('go_live')
        details = []
        for _, hol in self.holidays.iterrows():
            hd = hol['Holiday_Date']
            ed = self.attendance[(self.attendance['uae_id'] == emp['uae_id']) & (self.attendance['Date'] == hd)]
            is_earned = False
            if pd.isna(go_live) or hd < go_live:
                reason, status = "Before Go Live", "-"
            elif len(ed) == 0:
                reason, status = "No record", "Absent"
            else:
                status = ed.iloc[0]['final_status_clean']
                if status in self.not_eligible_list:
                    reason = f"Not eligible: {status}"
                elif status == 'available':
                    is_earned = True
                    reason = "Counted"
                else:
                    reason = "Not Available"
            details.append((hd.strftime('%Y-%m-%d'), hol.get('Holiday_Name',''), status, "✅" if is_earned else "❌", reason))
        
        self.holiday_table.setRowCount(len(details))
        self.holiday_table.setColumnCount(5)
        self.holiday_table.setHorizontalHeaderLabels(['Date', 'Holiday', 'Status', 'Earned', 'Reason'])
        self.holiday_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        for i, d in enumerate(details):
            for j, v in enumerate(d):
                item = QTableWidgetItem(v)
                if j == 3:
                    item.setBackground(QColor('#d4edda' if v == '✅' else '#f8d7da'))
                self.holiday_table.setItem(i, j, item)
        self.tabs.setCurrentIndex(1)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = LeaveCalculator()
    window.show()
    sys.exit(app.exec_())