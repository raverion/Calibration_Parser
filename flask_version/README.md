# Measurement Data Analyzer - Web Application

A Flask-based web application for analyzing measurement data from CSV and TXT files, generating Excel reports with tolerance charts, and interactive HTML reports.

## Features

### Equipment-Specific Reports
- Process CSV (output data) and TXT (input data) measurement files
- Multi-channel tolerance chart visualization
- PASS/FAIL analysis (Mean & Mean±2σ checks)
- Excel reports with formatted tables and embedded charts
- Interactive HTML reports with Plotly charts

### Cross-Equipment Comparison (Coming Soon)
- Compare multiple equipment samples of the same type
- Statistical comparison across samples
- Batch quality assessment

## Installation

1. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Running the Application

### One-Click Launch (Recommended)

**Windows:** Double-click `launch.bat`

**macOS/Linux:** Double-click `launch.py` (or run `./launch.sh` in terminal)

The launcher will:
1. Check and install any missing dependencies
2. Start the Flask server
3. Automatically open your browser to http://localhost:5000

### Manual Launch

If you prefer to run manually:

```bash
pip install -r requirements.txt
python app.py
```

Then open http://localhost:5000 in your browser.

## Usage

1. **Select Report Type**: Choose "Equipment-Specific Report" on the home page
2. **Upload Files**: Drag and drop or select your CSV/TXT measurement files
3. **Configure**: 
   - Select measurement types for files with multiple types
   - Set reference values and tolerances for each test point
   - Optionally save/load configuration files
4. **Process**: Click "Process Files" to generate reports
5. **Results**: Download Excel report and/or view interactive HTML report

## File Naming Convention

Files should follow naming patterns like:
- `DeviceName_TestValue_RangeValue_CHx.csv` for CSV files
- `DeviceName_TestValue_RangeValue_Samples.txt` for TXT files

Examples:
- `VT2816A_m2V5_R10V_CH3.csv` (Voltage: -2.5V, Range: 10V, Channel: 3)
- `VIO2004_3mA_R10mA_CH1.txt` (Current: 3mA, Range: 10mA, Channel: 1)

## Project Structure

```
measurement_webapp/
├── app.py              # Flask application
├── parsers.py          # File parsing utilities
├── excel_charts.py     # Excel chart generation
├── html_report.py      # HTML report generation
├── utils.py            # Shared utilities
├── requirements.txt    # Python dependencies
├── static/
│   ├── css/
│   │   └── style.css   # Application styles
│   └── js/
│       ├── upload.js   # Upload page scripts
│       └── configure.js # Configuration page scripts
├── templates/
│   ├── base.html       # Base template
│   ├── index.html      # Home page
│   ├── upload.html     # File upload page
│   ├── configure.html  # Configuration page
│   ├── results.html    # Results page
│   └── coming_soon.html # Placeholder for future features
├── uploads/            # Temporary upload storage
└── outputs/            # Generated reports
```

## Configuration Files

You can save and load test configurations as JSON files. The configuration includes:
- Reference values for each test point
- Tolerance values
- Range settings

This allows you to reuse configurations across multiple analysis sessions.

## Notes

- Session data (uploaded files and outputs) are stored temporarily and cleaned up when starting a new session
- The application supports both single-channel and multi-channel measurement files
- Files are automatically categorized as Input (TXT) or Output (CSV) based on extension
