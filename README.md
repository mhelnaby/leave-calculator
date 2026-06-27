# 🏖️ Leave Balance Calculator

**Version 2.4**  
A powerful, self‑contained web application that calculates employees' annual leave balances and earned public holidays according to Egyptian Labor Law No. 14 of 2025.

---

## 🚀 Features

- 🔍 **Search by UAE ID** – instantly retrieve detailed leave information for any employee.
- 📊 **Annual Leave Calculation** – automatically computes:
  - Total entitled days (15 days first year, 21 days thereafter).
  - Used days in 2025.
  - Used days in 2026 (with correct handling of half‑day and casual leaves).
  - Remaining balance.
- 🎉 **Public Holidays** – calculates how many official holidays the employee has earned (only if they were available on those days), and how many have been used as comp off.
- 📅 **Holiday Details Table** – shows a day‑by‑day breakdown of all public holidays and explains why each was/wasn't earned.
- 🎨 **Modern UI** – clean, responsive design with a custom icon and smooth animations.
- 🔄 **Queue & Go‑Live Date** – respects the nature of each queue (FTE / Per‑Minute) to determine the correct Go‑Live date.
- 🛡️ **Robust Data Linking** – uses ACD ID + Hiring Date to uniquely identify employees, preventing mix‑ups from reused ACD IDs.
- 📦 **All‑in‑one** – all calculations happen on the server; users only need a web browser.
- ✍️ **Developer Credit** – proudly displays the creator's name and copyright.

---

## 📋 How It Works

1. The app loads employee data (`employees.xlsx`), 2025 used leaves (`balance_2025.xlsx`), public holidays (`holidays.xlsx`), shift types that don't earn holidays (`not_eligible_shifts.xlsx`), and queue types (`Queue_Nature.xlsx`).
2. It reads daily attendance records from CSV files inside the `attendance_2026/` folder.
3. For each employee, it computes:
   - **Annual leave entitlement** based on hiring date and length of service.
   - **Net annual leave used in 2026**, subtracting any official holidays that fell on a leave day (only after the employee's Go‑Live date).
   - **Public holidays earned** – each official holiday when the employee was marked as `available` and after their Go‑Live date.
   - **Comp off days used** – recorded as `comp` in attendance.
4. The final balances are displayed in a clean web interface.

---

## 🖥️ How to Run (Development)

### Prerequisites
- Python 3.9 or higher
- Git (optional)

### Setup
1. Clone the repository:
   ```bash
   git clone https://github.com/mhelnaby/leave-calculator.git
   cd leave-calculator
