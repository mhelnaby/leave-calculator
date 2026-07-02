cat > "/home/mostafa/Documents/New APP/leave_calculator.py" << 'ENDOFFILE'
import sys, os, glob, threading, time
import warnings
warnings.filterwarnings('ignore')

from datetime import date, timedelta
import pandas as pd
from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import uvicorn

# ========== Paths ==========
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

ATTENDANCE_FOLDER = os.path.join(BASE_DIR, "attendance_2026")
EMPLOYEES_FILE    = os.path.join(BASE_DIR, "employees.xlsx")
BALANCE_2025_FILE = os.path.join(BASE_DIR, "balance_2025.xlsx")
HOLIDAYS_FILE     = os.path.join(BASE_DIR, "holidays.xlsx")
NOT_ELIGIBLE_FILE = os.path.join(BASE_DIR, "not_eligible_shifts.xlsx")
QUEUE_NATURE_FILE = os.path.join(BASE_DIR, "Queue_Nature.xlsx")

CUTOFF_DATE = date.today()
PER_MIN_DAYS = 52

LEAVE_TYPE_WEIGHTS = {
    "annual": 1.0, "half annual": 0.5, "casual": 1.0, "annual exception": 1.0
}

# ========== Data loading ==========
def load_all_data():
    fte_queues, per_min_queues = [], []
    if os.path.exists(QUEUE_NATURE_FILE):
        qn = pd.read_excel(QUEUE_NATURE_FILE)
        qn.columns = [c.strip() for c in qn.columns]
        nature_col = next((c for c in qn.columns if c.lower() == 'nature'), None)
        queue_col = next((c for c in qn.columns if c.lower() == 'queue'), None)
        if nature_col and queue_col:
            qn = qn.rename(columns={nature_col: 'Nature', queue_col: 'Queue'})
            fte_queues = qn[qn['Nature'].str.upper() == 'FTE']['Queue'].tolist()
            per_min_queues = qn[qn['Nature'].str.upper() != 'FTE']['Queue'].tolist()

    employees = pd.read_excel(EMPLOYEES_FILE)
    col_map = {
        'UAE ID': 'uae_id', 'ACD ID': 'acd_id',
        'Hiring Date': 'hiring_date', 'LWD': 'lwd',
        'Date of Certification': 'certification_date'
    }
    employees.rename(columns={k: v for k, v in col_map.items() if k in employees.columns}, inplace=True)
    for dc in ['hiring_date', 'lwd', 'certification_date']:
        if dc in employees.columns:
            employees[dc] = pd.to_datetime(employees[dc], errors='coerce', dayfirst=True)

    if 'Agent Name' in employees.columns:
        employees.rename(columns={'Agent Name': 'Name'}, inplace=True)
    if 'Name' not in employees.columns:
        employees['Name'] = 'Unknown'
    if 'Title' not in employees.columns:
        employees['Title'] = ''
    for col in ['LOB', 'S-LOB']:
        if col not in employees.columns:
            employees[col] = ''

    balance_2025 = pd.read_excel(BALANCE_2025_FILE)
    if 'Annual_Used' in balance_2025.columns:
        balance_2025.rename(columns={'UAE ID': 'uae_id', 'Annual_Used': 'annual_used_2025'}, inplace=True)
    elif 'Annual' in balance_2025.columns:
        balance_2025.rename(columns={'UAE ID': 'uae_id', 'Annual': 'annual_used_2025'}, inplace=True)
    else:
        balance_2025.rename(columns={'UAE ID': 'uae_id'}, inplace=True)
        balance_2025['annual_used_2025'] = 0

    holidays_raw = pd.read_excel(HOLIDAYS_FILE)
    holidays_raw.columns = ['Holiday_Date', 'Holiday_Name']
    holidays_raw['Holiday_Date'] = pd.to_datetime(holidays_raw['Holiday_Date'], errors='coerce')
    holidays = holidays_raw

    not_eligible_list = []
    if os.path.exists(NOT_ELIGIBLE_FILE):
        try:
            ne_df = pd.read_excel(NOT_ELIGIBLE_FILE)
        except:
            ne_df = pd.read_csv(NOT_ELIGIBLE_FILE)
        if not ne_df.empty:
            not_eligible_list = ne_df.iloc[:, 0].astype(str).str.strip().str.lower().tolist()

    all_files = sorted(glob.glob(f"{ATTENDANCE_FOLDER}/*.csv"))
    if not all_files:
        raise Exception("No attendance CSV files found")
    att_list = []
    for f in all_files:
        df = pd.read_csv(f, dayfirst=True)
        rename_att = {
            'ACD ID': 'acd_id', 'Date of Join': 'date_of_join',
            'Final Status': 'final_status', 'Queue': 'queue'
        }
        df.rename(columns={k: v for k, v in rename_att.items() if k in df.columns}, inplace=True)
        for dc in ['Date', 'date_of_join']:
            if dc in df.columns:
                df[dc] = pd.to_datetime(df[dc], errors='coerce', dayfirst=True)
        att_list.append(df)
    attendance = pd.concat(att_list, ignore_index=True)

    if 'queue' in attendance.columns and 'acd_id' in attendance.columns:
        first_att = attendance.sort_values('Date').groupby('acd_id').first().reset_index()
        first_att = first_att[['acd_id', 'queue']]
        employees['acd_str'] = employees['acd_id'].astype(str).str.strip()
        first_att['acd_str'] = first_att['acd_id'].astype(str).str.strip()
        employees = employees.merge(first_att[['acd_str', 'queue']], on='acd_str', how='left', suffixes=('', '_att'))
        employees['Queue'] = employees['Queue'].fillna(employees['queue']).fillna('')
        employees.drop(columns=['queue'], inplace=True)

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
            return cert + timedelta(days=1)
    employees['go_live'] = employees.apply(calc_go_live, axis=1)

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

    # Public holidays earned – using HIRING DATE
    all_hol = attendance.merge(holidays, left_on='Date', right_on='Holiday_Date', how='inner')
    all_hol = all_hol.merge(employees[['uae_id', 'hiring_date']], on='uae_id', how='left')
    public_earned = all_hol[
        (all_hol['Date'] >= all_hol['hiring_date']) &
        (all_hol['final_status_clean'].isin(['available'])) &
        (~all_hol['final_status_clean'].isin(not_eligible_list))
    ].groupby('uae_id').size().reset_index(name='public_holidays_earned')

    comp_days = attendance[attendance['final_status_clean'] == 'comp'].copy()
    comp_days = comp_days.merge(employees[['uae_id', 'hiring_date']], on='uae_id', how='left')
    comp_used = comp_days[comp_days['Date'] >= comp_days['hiring_date']].groupby('uae_id').size().reset_index(name='comp_days_used')

    emp_master = employees[['uae_id', 'hiring_date', 'lwd', 'go_live', 'certification_date',
                            'LOB', 'S-LOB', 'acd_id', 'Name', 'Title']].copy()
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

    for col, na_val in [('lwd', 'Active'), ('go_live', 'N/A'), ('certification_date', 'N/A')]:
        if col in final.columns:
            final[f'{col}_display'] = final[col].dt.strftime('%Y-%m-%d').fillna(na_val)
    final['calc_date_display'] = final['calc_date'].dt.strftime('%Y-%m-%d')
    final['hiring_date_display'] = final['hiring_date'].dt.strftime('%Y-%m-%d')
    final['service_years'] = ((final['calc_date'] - final['hiring_date']).dt.days / 365.25).round(2)

    return final, attendance, holidays, not_eligible_list


# ========== FastAPI App ==========
app = FastAPI(title="Leave Balance Calculator", version="2.4")

app.mount("/static", StaticFiles(directory=BASE_DIR), name="static")

final_df, attendance_df, holidays_df, not_eligible_shifts = load_all_data()

def find_employee(uid: str):
    uid_clean = uid.strip().lower()
    emp = final_df[final_df['uae_id'].astype(str).str.strip().str.lower() == uid_clean]
    if len(emp) == 0 and not uid_clean.startswith('uae'):
        emp = final_df[final_df['uae_id'].astype(str).str.strip().str.lower() == 'uae' + uid_clean]
    return emp.iloc[0] if len(emp) > 0 else None

def make_serializable(val):
    if pd.isna(val): return None
    if isinstance(val, (pd.Timestamp, date)): return val.isoformat()
    if hasattr(val, 'item'): return val.item()
    return val

# ========== Redesigned Main Page ==========
@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    html = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
    <meta charset="UTF-8">
    <title>Leave Balance Calculator v2.4</title>
    <link rel="icon" href="/static/creativity.png" type="image/png">
    <style>
      * { margin: 0; padding: 0; box-sizing: border-box; }
      body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        min-height: 100vh;
        display: flex;
        justify-content: center;
        align-items: center;
        padding: 20px;
      }
      .toast {
        position: fixed;
        top: 20px;
        right: 20px;
        background: #2c3e50;
        color: white;
        padding: 16px 24px;
        border-radius: 12px;
        box-shadow: 0 10px 25px rgba(0,0,0,0.2);
        z-index: 9999;
        animation: slideIn 0.5s ease, fadeOut 0.5s 7s ease forwards;
        font-weight: 500;
      }
      @keyframes slideIn {
        from { transform: translateX(100%); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
      }
      @keyframes fadeOut {
        to { opacity: 0; visibility: hidden; }
      }
      .card {
        max-width: 950px;
        width: 100%;
        background: white;
        border-radius: 20px;
        box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        padding: 40px;
        margin: 20px;
      }
      .header {
        display: flex;
        align-items: center;
        gap: 20px;
        margin-bottom: 30px;
      }
      .header img {
        width: 64px;
        height: 64px;
        object-fit: contain;
      }
      .header h1 {
        font-size: 28px;
        color: #2c3e50;
      }
      .search-area {
        display: flex;
        gap: 15px;
        flex-wrap: wrap;
        margin-bottom: 20px;
      }
      input[type="text"] {
        flex: 1;
        min-width: 250px;
        padding: 14px 18px;
        font-size: 16px;
        border: 2px solid #e0e6ed;
        border-radius: 12px;
        transition: border-color 0.3s;
        outline: none;
      }
      input[type="text"]:focus {
        border-color: #3498db;
      }
      .btn {
        padding: 14px 28px;
        font-size: 16px;
        font-weight: 600;
        border: none;
        border-radius: 12px;
        cursor: pointer;
        transition: transform 0.2s, background 0.3s;
      }
      .btn:hover { transform: translateY(-2px); }
      .btn-primary { background: #27ae60; color: white; }
      .btn-primary:hover { background: #219a52; }
      .btn-info { background: #2980b9; color: white; }
      .btn-info:hover { background: #1c6ea4; }
      .btn-danger { background: #e74c3c; color: white; }
      .btn-danger:hover { background: #c0392b; }
      .report {
        background: #2c3e50;
        color: #ecf0f1;
        padding: 25px;
        margin-top: 25px;
        border-radius: 16px;
        font-family: 'Courier New', monospace;
        white-space: pre-wrap;
        line-height: 1.6;
      }
      .hidden { display: none; }
      .error { color: #e74c3c; margin-top: 15px; font-weight: 600; }
      table {
        width: 100%;
        border-collapse: collapse;
        margin-top: 25px;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 5px 15px rgba(0,0,0,0.05);
      }
      th {
        background: #2980b9;
        color: white;
        padding: 14px 16px;
        text-align: left;
        font-weight: 600;
      }
      td {
        padding: 12px 16px;
        border-bottom: 1px solid #ecf0f1;
        background: white;
      }
      .earned { background-color: #d4edda; }
      .not-earned { background-color: #f8d7da; }
      .footer {
        text-align: center;
        margin-top: 30px;
        color: #7f8c8d;
        font-size: 14px;
      }
      .button-group { display: flex; gap: 15px; flex-wrap: wrap; margin-top: 25px; }
    </style>
    </head>
    <body>

    <div id="toast" class="toast hidden"></div>

    <div class="card">
      <div class="header">
        <img src="/static/creativity.png" alt="Logo">
        <h1>Leave Balance Calculator</h1>
      </div>
      <div class="search-area">
        <input type="text" id="uae_input" placeholder="Enter UAE ID (e.g., UAE00058)">
        <button class="btn btn-primary" onclick="calculate()">Calculate</button>
      </div>
      <div id="error_msg" class="error"></div>
      <div id="report" class="report hidden"></div>
      <div id="holiday_btn_container" class="hidden button-group">
        <button class="btn btn-info" onclick="loadHolidays()">📅 Holiday Details</button>
      </div>
      <div id="holiday_section" class="hidden">
        <h2 style="margin-top: 30px; color: #2c3e50;">📅 Holiday Details</h2>
        <table id="holiday_table">
          <thead><tr><th>Date</th><th>Holiday</th><th>Status</th><th>Earned</th><th>Reason</th></tr></thead>
          <tbody></tbody>
        </table>
      </div>
      <div class="button-group">
        <button class="btn btn-danger" onclick="shutdown()">⏻ Shutdown Server</button>
      </div>
      <div class="footer">
        Created and developed by: <strong>Mostafa Hasab ElNaby</strong> | All rights reserved 2026<br>
        Version 2.4
      </div>
    </div>

    <script>
    function showToast(message) {
      const toast = document.getElementById("toast");
      toast.textContent = message;
      toast.classList.remove("hidden");
      void toast.offsetWidth;
      toast.style.animation = "slideIn 0.5s ease, fadeOut 0.5s 7s ease forwards";
    }

    window.addEventListener("DOMContentLoaded", () => {
      showToast("👋 Welcome to the Leave Balance Calculator!");
    });

    let currentUid = '';

    async function calculate() {
      const uid = document.getElementById("uae_input").value.trim();
      if (!uid) return;
      document.getElementById("error_msg").textContent = "";
      document.getElementById("report").classList.add("hidden");
      document.getElementById("holiday_btn_container").classList.add("hidden");
      document.getElementById("holiday_section").classList.add("hidden");

      let resp = await fetch(`/balance/${uid}`);
      if (resp.status === 404) {
        document.getElementById("error_msg").textContent = "Employee not found";
        return;
      }
      let data = await resp.json();
      currentUid = data.uae_id;

      let report = `╔══════════════════════════════════════╗
║   Leave Balance Report - Law 14/2025  ║
╚══════════════════════════════════════╝
👤 UAE ID      : ${data.uae_id}
🆔 ACD ID      : ${data.acd_id ?? '?'}
📛 Name        : ${data.Name ?? '?'}
🏷️  Title       : ${data.Title ?? ''}
📡 LOB         : ${data.LOB ?? ''}
📡 S-LOB       : ${data['S-LOB'] ?? ''}
──────────────────────────────────────
📅 Hiring Date : ${data.hiring_date_display ?? '?'}
📜 Certified   : ${data.certification_date_display ?? '?'}
🚀 Go Live     : ${data.go_live_display ?? '?'}
🏁 Last Working Day: ${data.lwd_display ?? '?'}
⏳ Service     : ${data.service_years} years
──────────────────────────────────────
📋 Annual Leave:
  Entitled     : ${data.total_earned} days
  Used 2025    : ${data.annual_used_2025} days
  Used 2026    : ${data.net_used_2026} days
  ➤ Remaining  : ${data.remaining_annual} days
──────────────────────────────────────
🎉 Public Holidays:
  Earned       : ${data.public_holidays_earned} days
  Used (Comp)  : ${data.comp_days_used} days
  ➤ Remaining  : ${data.public_holidays_remaining} days`;
      document.getElementById("report").textContent = report;
      document.getElementById("report").classList.remove("hidden");
      document.getElementById("holiday_btn_container").classList.remove("hidden");
    }

    async function loadHolidays() {
      if (!currentUid) return;
      document.getElementById("holiday_section").classList.add("hidden");
      let resp = await fetch(`/holiday_details/${currentUid}`);
      let holidays = await resp.json();
      let tbody = document.querySelector("#holiday_table tbody");
      tbody.innerHTML = "";
      holidays.forEach(h => {
        let row = tbody.insertRow();
        row.insertCell().textContent = h.date;
        row.insertCell().textContent = h.name;
        row.insertCell().textContent = h.status;
        let earnedCell = row.insertCell();
        earnedCell.textContent = h.earned ? "✅" : "❌";
        earnedCell.className = h.earned ? "earned" : "not-earned";
        row.insertCell().textContent = h.reason;
      });
      document.getElementById("holiday_section").classList.remove("hidden");
    }

    async function shutdown() {
      if (confirm("Shutdown the server? A thank-you page will appear and the server will stop.")) {
        document.body.innerHTML = `
          <div style="display:flex; justify-content:center; align-items:center; height:100vh; background:linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);">
            <div style="text-align:center; background:white; padding:60px; border-radius:20px; box-shadow:0 20px 40px rgba(0,0,0,0.1);">
              <h1 style="color:#2c3e50;">Thank you for using the Leave Balance Calculator</h1>
              <p style="color:#7f8c8d;">The server is shutting down. You can close this tab.</p>
              <p style="color:#7f8c8d; font-size:14px;">Created and developed by: <strong>Mostafa Hasab ElNaby</strong> | All rights reserved 2026<br>Version 2.4</p>
            </div>
          </div>
        `;
        fetch('/shutdown');
        setTimeout(() => {
          window.close();
        }, 27000);
      }
    }
    </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html)


@app.get("/balance/{uae_id}")
async def get_balance(uae_id: str):
    emp = find_employee(uae_id)
    if emp is None:
        return JSONResponse(status_code=404, content={"error": "Employee not found"})
    result = {k: make_serializable(v) for k, v in emp.to_dict().items()}
    return result


@app.get("/holiday_details/{uae_id}")
async def holiday_details(uae_id: str):
    emp = find_employee(uae_id)
    if emp is None:
        return JSONResponse(status_code=404, content={"error": "Employee not found"})
    hiring_date = emp['hiring_date']
    details = []
    for _, hol in holidays_df.iterrows():
        hd = hol['Holiday_Date']
        ed = attendance_df[(attendance_df['uae_id'] == emp['uae_id']) & (attendance_df['Date'] == hd)]
        is_earned = False
        if pd.isna(hiring_date) or hd < hiring_date:
            reason, status = "Before Hiring Date", "-"
        elif len(ed) == 0:
            reason, status = "No record", "Absent"
        else:
            status = ed.iloc[0]['final_status_clean']
            if status in not_eligible_shifts:
                reason = f"Not eligible: {status}"
            elif status == 'available':
                is_earned = True
                reason = "Counted"
            else:
                reason = "Not Available"
        details.append({
            "date": hd.strftime('%Y-%m-%d'),
            "name": hol.get('Holiday_Name', ''),
            "status": status,
            "earned": is_earned,
            "reason": reason
        })
    return details


@app.get("/shutdown")
async def shutdown():
    import os as _os
    threading.Thread(target=lambda: (time.sleep(1), _os._exit(0))).start()
    return HTMLResponse(content="""
    <html><body style="margin:0; padding:0;">
    <div style="display:flex; justify-content:center; align-items:center; height:100vh; background:linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);">
      <div style="text-align:center; background:white; padding:60px; border-radius:20px; box-shadow:0 20px 40px rgba(0,0,0,0.1);">
        <h1 style="color:#2c3e50;">Thank you for using the Leave Balance Calculator</h1>
        <p style="color:#7f8c8d;">The server is shutting down. You can close this tab.</p>
        <p style="color:#7f8c8d; font-size:14px;">Created and developed by: <strong>Mostafa Hasab ElNaby</strong> | All rights reserved 2026<br>Version 2.4</p>
      </div>
    </div>
    </body></html>
    """)


if __name__ == "__main__":
    uvicorn.run("leave_calculator:app", host="0.0.0.0", port=8000, reload=False)
ENDOFFILE
