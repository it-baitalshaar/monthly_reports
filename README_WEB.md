# Monthly Attendance Report Generator - Web Application

A Flask web application to generate monthly attendance reports from your PostgreSQL database.

## Features

- 🌐 **Web Interface**: Easy-to-use web interface - no command line needed!
- 📅 **Date Range Selection**: Select any date range via web form
- 🔢 **Automatic Filename Generation**: Files automatically named based on month/year (e.g., `monthly_attendance_report_january_2026.xlsx`)
- ⚡ **Auto Monthly Hours**: Automatically calculates monthly hours based on number of days
- 📄 **PDF Export**: Optional PDF export for each employee
- 📊 **Excel Generation**: Creates professional Excel reports with separate sheets per employee

## Installation

1. Install required packages:
```bash
pip install -r requirements.txt
```

## Running the Web Application

1. Start the Flask server:
```bash
python app.py
```

2. Open your web browser and go to:
```
http://localhost:5000
```

3. Fill in the form:
   - **From Date**: Start date of the report period
   - **To Date**: End date of the report period
   - **Monthly Hours**: (Optional) Leave empty for auto-calculation
   - **Export PDF**: Check if you want PDF files too

4. Click "Generate Report" and download your file!

## File Structure

```
app_for_monthly_reports/
├── app.py                          # Flask web application
├── new_script_for_employee.py      # Standalone script version
├── monthly_attendance_report_generator.py  # Enhanced script with date inputs
├── templates/
│   ├── index.html                  # Main form page
│   └── success.html                # Success/download page
├── generated_reports/               # Generated Excel files (created automatically)
│   └── pdfs/                       # Generated PDF files (if PDF export enabled)
├── monthly_attendence_report_tempelate.xlsx  # Excel template
└── requirements.txt
```

## Automatic Filename Generation

The system automatically generates filenames based on the date range:

- **January 2026** → `monthly_attendance_report_january_2026.xlsx`
- **February 2026** → `monthly_attendance_report_february_2026.xlsx`
- **March 2026** → `monthly_attendance_report_march_2026.xlsx`
- And so on...

## Monthly Hours Calculation

If you don't specify monthly hours, the system automatically calculates:
- **31 days** = 248 hours (31 × 8)
- **30 days** = 240 hours (30 × 8)
- **29 days** = 232 hours (29 × 8)
- **28 days** = 224 hours (28 × 8)

## Configuration

Update database connection in `app.py`:
```python
DB_CONFIG = {
    'host': 'your-host',
    'port': 6543,
    'dbname': 'postgres',
    'user': 'your-user',
    'password': 'your-password'
}
```

## Production Deployment

For production, consider:
1. Change `app.secret_key` to a secure random string
2. Use a production WSGI server (e.g., Gunicorn)
3. Set `debug=False` in `app.run()`
4. Use environment variables for database credentials
5. Add authentication/authorization

## Troubleshooting

- **Template not found**: Make sure `monthly_attendence_report_tempelate.xlsx` is in the same directory as `app.py`
- **Database connection error**: Check your database credentials in `app.py`
- **Port already in use**: Change the port in `app.py` (default: 5000)
