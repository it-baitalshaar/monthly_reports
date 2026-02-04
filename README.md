# Monthly Attendance Report Generator

A Python script to generate monthly attendance reports from a PostgreSQL database using an Excel template.

## Features

- **Date Range Input**: Specify custom date ranges (from date to to date)
- **Automatic Monthly Hours Calculation**: Automatically calculates monthly hours based on number of days
  - 31 days = 248 hours (31 × 8)
  - 30 days = 240 hours (30 × 8)
  - 29 days = 232 hours (29 × 8)
  - 28 days = 224 hours (28 × 8)
- **Manual Override**: Option to manually enter monthly hours if needed
- **Template-Based**: Uses `monthly_attendence_report_tempelate.xlsx` as the base template
- **Multi-Employee Support**: Generates separate sheets for each employee
- **PDF Export**: Optional PDF export for each employee's attendance report
- **Cell Protection**: Automatically unlocks editable cells in the Excel template

## Installation

1. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the script:
```bash
python monthly_attendance_report_generator.py
```

2. Follow the prompts:
   - Enter the **FROM date** (YYYY-MM-DD format, e.g., 2026-01-01)
   - Enter the **TO date** (YYYY-MM-DD format, e.g., 2026-01-31)
   - Choose to use auto-calculated monthly hours or enter manually

3. The script will:
   - Connect to the database
   - Fetch attendance data for the specified date range
   - Generate an Excel file with separate sheets for each employee
   - Save the output file in the same directory as the script
   - Optionally export PDF files for each employee

## Output

### Excel File
The generated Excel file will be named:
```
monthly_attendance_report_YYYY_MM_DD_to_YYYY_MM_DD.xlsx
```

### PDF Files (Optional)
If you choose to export PDFs, they will be saved in a `pdf_reports` folder:
```
pdf_reports/
  ├── Employee_Name_1_YYYY_MM.pdf
  ├── Employee_Name_2_YYYY_MM.pdf
  └── ...
```

## Configuration

Update the database connection parameters in the script if needed:
- `host`
- `port`
- `dbname`
- `user`
- `password`

## Monthly Hours Cell Location

The script attempts to write monthly hours to cell `I3` by default. If your template uses a different cell for monthly hours, update line 147 in `monthly_attendance_report_generator.py`:

```python
employee_sheet['I3'] = monthly_hours  # Change 'I3' to your cell reference
```

## Template Requirements

The template file (`monthly_attendence_report_tempelate.xlsx`) should:
- Have a sheet named "Sheet1"
- Use row 7 as the starting row for attendance data
- Have columns for: Date (A), Status (B), Working Hours (C), Overtime (D, E, F), Projects (H), Notes (J)
- Employee name goes in cell B4
- Salary goes in cell I2
