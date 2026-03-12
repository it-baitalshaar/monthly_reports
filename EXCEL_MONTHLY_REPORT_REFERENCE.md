# Monthly Report Excel Sheet – Layout & Formulas Reference

This document describes the Excel template used for monthly attendance reports and all formulas/logic used by the generators.

---

## 1. Template file

- **Name:** `monthly_attendence_report_tempelate.xlsx`
- **Location:** same folder as `app.py` / `monthly_attendance_report_generator.py`
- **Sheet name:** `Sheet1` (one sheet per employee is copied from this template)

---

## 2. Cell map – where data is written

| Cell(s) | Content | Written by script |
|--------|---------|--------------------|
| **A1, A2, B1, B2, C1, C2** | Period text | `"Period: YYYY-MM-DD to YYYY-MM-DD"` (first empty cell used) |
| **B4** | Employee name | From DB `Employee.name` |
| **I2** | Salary | From DB `Employee.salary` |
| **I3, I4, I5, H3, H4** | Monthly hours | First writable cell gets `monthly_hours` (see formula below) |
| **A7, A8, …** | Date | One row per day; `A{row} = date` |
| **B7, B8, …** | Status code | See status codes below |
| **C7, C8, …** | Working hours | Total working hours for that day |
| **D7, D8, …** | Overtime (Present) | Overtime when status is Present |
| **E7, E8, …** | Overtime (Weekend) | Overtime when status is Weekend |
| **F7, F8, …** | Overtime (Holiday) | Overtime when status is Holiday-Work |
| **H7, H8, …** | Projects | Project name(s) and hours (e.g. `"Project A 4hrs + Project B 4hrs"`) |
| **J7, J8, …** | Notes | Attendance notes for that day |

- **Data starts at row 7** (one row per calendar day).
- **Editable columns** (unlocked): B, C, D, E, F, H, J (rows 7–50).

---

## 3. Formulas used in the system

### 3.1 Monthly hours (used by Python, not an Excel formula)

Used to compute the value written into the template (e.g. I3 / I4 / I5 / H3 / H4):

```
monthly_hours = number_of_days × 8
number_of_days = (to_date - from_date).days + 1   (inclusive)
```

**Examples:**

- 31 days → 248 hours  
- 30 days → 240 hours  
- 29 days → 232 hours  
- 28 days → 224 hours  

You can override this with a manual value in the web app or CLI.

### 3.2 Row total / validation (from trial salary script)

In `trials/trial_python_script_for_salary_report.py`, each data row uses:

- **Column G (per row):**  
  `=IF(B{r}="SL","SL","")`  
  – Shows "SL" when status is Sick Leave, else blank.

- **Column I (per row):**  
  `=IFERROR(IF(SUM(C{r}:G{r})>24,"إجمالي > 24 ساعة.",SUM(C{r}:G{r})),"")`  
  – Sum of columns C–G for that row; if &gt; 24 shows Arabic message “Total &gt; 24 hours”, otherwise shows the sum.

The **main** monthly report generator (`monthly_attendance_report_generator.py` and `app.py`) does **not** write these formulas; it only writes **values** into A, B, C, D, E, F, H, J. So if your template has totals or other formulas (e.g. in G, I or a summary section), they must be **in the template file itself** and will calculate from the values the script writes.

---

## 4. Status codes (column B)

| DB status | status_attendance | Written in Excel (B) |
|-----------|-------------------|----------------------|
| present | Present | **P** |
| present | Weekend | **W** |
| present | Holiday-Work | **H** |
| absent | Absence without excuse | **AWO** |
| absent | Sick Leave | **SL** |
| absent | Absence with excuse | **A** |
| vacation | (any) | **V** |

---

## 5. Working hours and overtime (columns C, D, E, F)

- **C (working hours):** Sum of `working_hours` from `Attendance_projects` for that day. If status is Weekend or Holiday-Work and total is 0, script writes **8**.
- **D:** Overtime hours when `status_attendance = 'Present'`.
- **E:** Overtime hours when `status_attendance = 'Weekend'`.
- **F:** Overtime hours when `status_attendance = 'Holiday-Work'`.

All of these are **values** written by the script, not Excel formulas.

---

## 6. Template expectations (from README)

- Sheet named **Sheet1**.
- **Row 7** = first row for daily data.
- **Columns:** A = Date, B = Status, C = Working hours, D/E/F = Overtime (Present/Weekend/Holiday), H = Projects, J = Notes.
- **B4** = employee name, **I2** = salary.
- Any **totals, deductions, or payroll formulas** (e.g. total hours, net pay) must be defined **in the template**; the script only fills the cells listed in the cell map above.

---

## 7. Generated output files

- **CLI:** `monthly_attendance_report_YYYY_MM_DD_to_YYYY_MM_DD.xlsx`
- **Web:** `monthly_attendance_report_{month_name}_{year}.xlsx` (e.g. `monthly_attendance_report_january_2026.xlsx`), optionally with `_construction` or `_maintenance` suffix when filtered.

If you need the exact formulas that are **inside** your current template (e.g. in summary rows or columns G/I), open `monthly_attendence_report_tempelate.xlsx` in Excel and check the formula bar for those cells; this reference documents how the generator fills the sheet and which formulas are used in code.
