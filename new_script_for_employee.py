import pandas as pd
import psycopg2
from openpyxl import load_workbook
from datetime import datetime
import os
import calendar

# Database connection parameters
host = 'aws-0-ap-south-1.pooler.supabase.com'
port = 6543
dbname = 'postgres'
user = 'postgres.fhsvgeacwnnvqidyhnok'
password = 'Bait-Alshaar20'

def generate_filename_from_date_range(from_date, to_date):
    """
    Generate filename automatically based on date range
    Examples: 
    - Jan 2026 -> monthly_attendance_report_january_2026.xlsx
    - Feb 2026 -> monthly_attendance_report_february_2026.xlsx
    """
    # Get month name and year from the date range
    month_name = calendar.month_name[from_date.month].lower()
    year = from_date.year
    
    filename = f"monthly_attendance_report_{month_name}_{year}.xlsx"
    return filename

# Get the script directory and template path
script_dir = os.path.dirname(os.path.abspath(__file__))
template_path = os.path.join(script_dir, 'monthly_attendence_report_tempelate.xlsx')

# Check if template exists
if not os.path.exists(template_path):
    print(f"Error: Template file not found at: {template_path}")
    exit(1)

# Date range (you can modify these dates)
from_date = datetime(2026, 1, 1).date()
to_date = datetime(2026, 1, 31).date()

# Generate output filename automatically
output_filename = generate_filename_from_date_range(from_date, to_date)
file_path = os.path.join(script_dir, output_filename)

# Load the template file
workbook = load_workbook(template_path)
template_sheet = workbook['Sheet1']


# Connect to the database
try:
    connection = psycopg2.connect(
        host=host,
        port=port,
        dbname=dbname,
        user=user,
        password=password
    )
    cursor = connection.cursor()
    print("Connected to the database!")

    query = '''
        SELECT 
            e.employee_id,
            e.name,
            e.salary,
            a.date,
            a.notes,
            a.status,
            a.status_attendance,
            p.project_name,
            ap.working_hours,
            ap.overtime_hours
        FROM 
            "Employee" e
        LEFT JOIN 
            "Attendance" a ON e.employee_id = a.employee_id
        LEFT JOIN
            "Attendance_projects" ap ON a.id = ap.attendance_id
        LEFT JOIN
            projects p ON ap.project_id = p.project_id
        WHERE 
            a.date BETWEEN %s AND %s
        ORDER BY 
            e.employee_id, a.date;
    '''

    cursor.execute(query, (from_date, to_date))
    data = cursor.fetchall()

    df = pd.DataFrame(data, columns=[
        'employee_id', 'name', 'salary', 'date', 'notes', 'status',
        'status_attendance', 'project_name', 'working_hours', 'overtime_hours'
    ])

    grouped_data = df.groupby('employee_id')

    for employee_id, employee_data in grouped_data:
        employee_sheet = workbook.copy_worksheet(template_sheet)

        first_row = employee_data.iloc[0]
        employee_sheet['B4'] = first_row['name']
        employee_sheet.title = f"{first_row['name']}"
        employee_sheet.sheet_view.rightToLeft = template_sheet.sheet_view.rightToLeft
        employee_sheet['I2'] = first_row['salary']

        # Write date range and month name to the correct template cells
        employee_sheet['F2'] = from_date
        employee_sheet['F3'] = to_date
        employee_sheet['F4'] = calendar.month_name[from_date.month]

        row_date = 7  # Start from row 7

        # --- KEY FIX: Group by date so each day = exactly one row ---
        daily_grouped = employee_data.groupby('date')

        for date_value, day_data in daily_grouped:
            # These values are the same for all rows of the same day
            status = day_data.iloc[0]['status']
            status_attendance = day_data.iloc[0]['status_attendance']
            notes = day_data.iloc[0]['notes']

            # --- Write the date ---
            employee_sheet[f'A{row_date}'] = date_value

            # --- Write notes ---
            employee_sheet[f'J{row_date}'] = notes

            # --- Write attendance status code ---
            if status == 'present' and status_attendance == 'Present':
                employee_sheet[f'B{row_date}'] = 'P'
            elif status == 'present' and status_attendance == 'Weekend':
                employee_sheet[f'B{row_date}'] = 'W'
            elif status == 'present' and status_attendance == 'Holiday-Work':
                employee_sheet[f'B{row_date}'] = 'H'
            elif status == 'absent' and status_attendance == 'Absence without excuse':
                employee_sheet[f'B{row_date}'] = 'AWO'
            elif status == 'absent' and status_attendance == 'Sick Leave':
                employee_sheet[f'B{row_date}'] = 'SL'
            elif status == 'absent' and status_attendance == 'Absence with excuse':
                employee_sheet[f'B{row_date}'] = 'A'
            elif status == 'vacation':
                employee_sheet[f'B{row_date}'] = 'V'

            # --- Sum total working hours for the day ---
            total_working_hours = day_data['working_hours'].sum()

            # If working hours are all 0 but it's Weekend or Holiday-Work, default to 8
            if total_working_hours == 0:
                if status_attendance in ('Weekend', 'Holiday-Work'):
                    total_working_hours = 8

            if total_working_hours != 0:
                employee_sheet[f'C{row_date}'] = total_working_hours

            # --- Sum overtime hours per type (Present / Weekend / Holiday-Work) ---
            total_overtime = day_data['overtime_hours'].sum()

            if total_overtime != 0:
                if status_attendance == 'Present':
                    employee_sheet[f'D{row_date}'] = total_overtime
                elif status_attendance == 'Weekend':
                    employee_sheet[f'E{row_date}'] = total_overtime
                elif status_attendance == 'Holiday-Work':
                    employee_sheet[f'F{row_date}'] = total_overtime

            # --- Build the combined project string ---
            # Example output: "Home's Mr.Abdullah 3hrs + Saqia 20 5hrs"
            if len(day_data) == 1:
                # Single project — just write the project name
                project_text = day_data.iloc[0]['project_name']
            else:
                # Multiple projects — combine with hours
                project_parts = []
                for _, proj_row in day_data.iterrows():
                    proj_name = proj_row['project_name'] if proj_row['project_name'] else 'Unknown'
                    proj_hours = int(proj_row['working_hours']) if proj_row['working_hours'] is not None and proj_row['working_hours'] != 0 else 0
                    project_parts.append(f"{proj_name} {proj_hours}hrs")
                project_text = " + ".join(project_parts)

            employee_sheet[f'H{row_date}'] = project_text

            # Move to next row
            row_date += 1

        print(f"Data written for Employee ID: {employee_id} — {first_row['name']}")

    workbook.save(file_path)
    print(f"Excel file saved successfully: {output_filename}")
    print(f"Full path: {file_path}")

except Exception as error:
    print(f"Error: {error}")

finally:
    if connection:
        cursor.close()
        connection.close()
        print("Database connection closed.")