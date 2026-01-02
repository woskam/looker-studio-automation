"""
Script to combine all weekly CSV files into one master file
With week number column
"""

import os
import pandas as pd
import glob
from datetime import datetime
import logging

# ============================================
# CONFIGURATION
# ============================================
WEEKLY_DATA_FOLDER = r"C:\path\to\your\weekly\data\folder"
MASTER_FILE = os.path.join(WEEKLY_DATA_FOLDER, "master_data_all_weeks.xlsx")
LOG_FILE = os.path.join(WEEKLY_DATA_FOLDER, "consolidate_log.txt")
# ============================================

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def consolidate_weekly_data():
    """Combine all weekly CSV files into one Excel file"""
    
    try:
        logging.info("=== Start Consolidation of Weekly Data ===")
        
        # Find all data_weekXX_YYYY.csv files
        pattern = os.path.join(WEEKLY_DATA_FOLDER, "data_week*.csv")
        csv_files = glob.glob(pattern)
        
        if not csv_files:
            logging.warning("No weekly CSV files found!")
            return False
        
        logging.info(f"Found {len(csv_files)} weekly CSV files")
        
        # List to store all dataframes
        all_data = []
        
        # Read each CSV file and add week number
        for csv_file in sorted(csv_files):
            try:
                # Extract week number and year from filename
                # Format: data_week45_2025.csv
                filename = os.path.basename(csv_file)
                parts = filename.replace('.csv', '').split('_')
                
                week_str = parts[1]  # "week45"
                year_str = parts[2]  # "2025"
                
                week_number = int(week_str.replace('week', ''))
                year = int(year_str)
                
                logging.info(f"Processing: {filename} (Week {week_number}, {year})")
                
                # Read CSV file
                df = pd.read_csv(csv_file)
                
                # Add week number and year columns (at the beginning)
                df.insert(0, 'Year', year)
                df.insert(1, 'Week', week_number)
                
                all_data.append(df)
                
            except Exception as e:
                logging.error(f"Error processing {csv_file}: {e}")
                continue
        
        if not all_data:
            logging.error("No data successfully processed")
            return False
        
        # Combine all dataframes
        logging.info("Combining all data...")
        master_df = pd.concat(all_data, ignore_index=True)
        
        # Sort by Year and Week (newest at the bottom)
        master_df = master_df.sort_values(['Year', 'Week'])
        
        logging.info(f"Total number of rows in master file: {len(master_df)}")
        logging.info(f"Columns: {', '.join(master_df.columns.tolist())}")
        
        # Save as Excel file
        logging.info(f"Saving to: {MASTER_FILE}")
        master_df.to_excel(MASTER_FILE, index=False, engine='openpyxl')
        
        # Also create a backup with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(WEEKLY_DATA_FOLDER, f"master_data_backup_{timestamp}.xlsx")
        master_df.to_excel(backup_file, index=False, engine='openpyxl')
        logging.info(f"Backup saved: {backup_file}")
        
        logging.info("=== Consolidation Successfully Completed ===")
        print(f"\n✓ Master file created: {MASTER_FILE}")
        print(f"  Total {len(master_df)} rows")
        print(f"  Weeks: {master_df['Week'].min()} - {master_df['Week'].max()}")
        
        return True
        
    except Exception as e:
        logging.error(f"ERROR during consolidation: {str(e)}", exc_info=True)
        print(f"\n✗ Consolidation failed: {e}")
        return False

if __name__ == "__main__":
    success = consolidate_weekly_data()
    
    if success:
        exit(0)
    else:
        exit(1)
