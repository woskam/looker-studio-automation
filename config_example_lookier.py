# Configuration Example
# Copy this to config_personal.py and customize with your settings

# ============================================
# LOOKER STUDIO CONFIGURATION
# ============================================

# Your Looker Studio dashboard URL
# Find this by opening your dashboard in Chrome and copying the full URL
LOOKER_URL = "https://lookerstudio.google.com/reporting/YOUR-REPORT-ID-HERE/page/YOUR-PAGE-ID-HERE"

# Example URLs:
# LOOKER_URL = "https://lookerstudio.google.com/reporting/03c24722-c1a6-4f81-9b7c-7dcdae1b2ee2/page/p_wnf5qms8bd"
# LOOKER_URL = "https://lookerstudio.google.com/u/0/reporting/1a2b3c4d-5e6f-7g8h-9i0j-1k2l3m4n5o6p/page/abcd"

# ============================================
# FILE PATHS
# ============================================

# Where to save downloaded CSV files and logs
# Use raw string (r"...") for Windows paths
OUTPUT_FOLDER = r"C:\Users\YourName\Documents\LookerData"

# Examples:
# OUTPUT_FOLDER = r"C:\Data\Weekly_Reports"
# OUTPUT_FOLDER = r"D:\Projects\Looker\Exports"

# ============================================
# DOWNLOAD SETTINGS
# ============================================

# File pattern for downloaded files from Looker Studio
# This is the pattern your Looker export uses
# Use wildcards (*) to match variable parts
DOWNLOAD_FILE_PATTERN = "*Table*.csv"

# Examples based on your Looker export naming:
# DOWNLOAD_FILE_PATTERN = "*Table*.csv"         # Matches "SEA Table (1).csv", "Data Table Export.csv"
# DOWNLOAD_FILE_PATTERN = "Export_*.csv"        # Matches "Export_20250101.csv"
# DOWNLOAD_FILE_PATTERN = "Looker_Data_*.csv"   # Matches "Looker_Data_Week45.csv"
# DOWNLOAD_FILE_PATTERN = "*.csv"               # Matches any CSV (not recommended if you download other CSVs)

# ============================================
# CONSOLIDATION SETTINGS
# ============================================

# Same as OUTPUT_FOLDER if you want consolidated file in same location
# Or use different folder for master files
WEEKLY_DATA_FOLDER = r"C:\Users\YourName\Documents\LookerData"

# Master file name (will be created in WEEKLY_DATA_FOLDER)
MASTER_FILE_NAME = "master_data_all_weeks.xlsx"

# ============================================
# ADVANCED SETTINGS (Optional)
# ============================================

# Maximum wait time for download completion (seconds)
DOWNLOAD_TIMEOUT = 60

# Wait time after login prompt (if not auto-logged in)
LOGIN_WAIT_TIME = 60

# Wait time for dashboard to fully load (seconds)
DASHBOARD_LOAD_TIME = 10

# Wait time after applying date filter (seconds)
DATE_FILTER_WAIT_TIME = 20

# ============================================
# TASK SCHEDULER SETTINGS (Windows)
# ============================================

# These are just reference values for setting up Task Scheduler
# Day of week to run (0 = Monday, 1 = Tuesday, etc.)
SCHEDULE_DAY = 0  # Monday

# Time to run (24-hour format)
SCHEDULE_TIME = "09:00"

# ============================================
# USAGE INSTRUCTIONS
# ============================================

"""
1. Copy this file to 'config_personal.py'
2. Update LOOKER_URL with your dashboard URL
3. Update OUTPUT_FOLDER with your desired save location
4. Update DOWNLOAD_FILE_PATTERN to match your Looker export filename
5. Save the file
6. The main script will import from config_personal.py if it exists

To use this configuration in your script:
    
    try:
        from config_personal import *
    except ImportError:
        # Use default values from looker_download.py
        pass
"""
