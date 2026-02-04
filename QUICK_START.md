# Quick Start Guide

## 🚀 Running the Web Application

### Step 1: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 2: Start the Server
```bash
python app.py
```

### Step 3: Open Browser
Go to: **http://localhost:5000**

### Step 4: Generate Report
1. Select **From Date** (e.g., 2026-01-01)
2. Select **To Date** (e.g., 2026-01-31)
3. (Optional) Enter monthly hours or leave empty for auto-calculation
4. (Optional) Check "Export PDF" if you want PDF files
5. Click **Generate Report**
6. Download your Excel file!

## 📝 Standalone Script Usage

If you prefer command-line:

```bash
python new_script_for_employee.py
```

**Note**: Edit the dates in the script (lines 25-26) before running.

## 📁 Output Files

- **Excel files**: Saved in `generated_reports/` folder
- **PDF files**: Saved in `generated_reports/pdfs/` folder (if PDF export enabled)

## 🔧 Configuration

Update database credentials in:
- `app.py` (for web app)
- `new_script_for_employee.py` (for standalone script)

## ✨ Features

✅ Automatic filename generation (january_2026, february_2026, etc.)  
✅ Auto-calculate monthly hours based on days  
✅ Web interface - no coding needed!  
✅ PDF export option  
✅ Multiple employee support  
