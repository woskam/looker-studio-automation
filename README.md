# Looker Studio Automated Data Export

Automate weekly data exports from Looker Studio dashboards using Python and Selenium. This tool automatically selects date ranges, exports data to CSV, and consolidates weekly reports into a master Excel file.

## Features

- 🤖 **Automated Chrome Control**: Uses Selenium to navigate Looker Studio dashboards
- 📅 **Smart Date Selection**: Automatically selects previous week (Sunday-Saturday)
- 💾 **Auto-Download**: Exports data as CSV with proper formatting
- 📊 **Data Consolidation**: Combines all weekly files into master Excel
- 🔄 **ChromeDriver Management**: Automatically downloads matching ChromeDriver version
- 🔐 **Persistent Login**: Cookie copying utility to stay logged in
- 📝 **Comprehensive Logging**: Detailed logs with screenshots for debugging

## Requirements

- **Windows OS** (required for Chrome automation)
- **Python 3.7+**
- **Google Chrome** browser installed
- **Looker Studio** dashboard access

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/looker-studio-automation.git
cd looker-studio-automation
```

2. Install required packages:
```bash
pip install -r requirements.txt
```

3. Configure your settings (see Configuration section)

## Configuration

Edit the configuration section in `looker_download.py`:

```python
# Your Looker Studio dashboard URL
LOOKER_URL = "https://lookerstudio.google.com/reporting/YOUR-REPORT-ID/page/YOUR-PAGE-ID"

# Where to save downloaded files
OUTPUT_FOLDER = r"C:\path\to\your\output\folder"

# File pattern for your Looker export (adjust to match your filename)
DOWNLOAD_FILE_PATTERN = "*Table*.csv"
```

Also update `consolidate_weekly_data.py`:
```python
WEEKLY_DATA_FOLDER = r"C:\path\to\your\weekly\data\folder"
```

## Usage

### First-Time Setup

1. **Copy Chrome cookies** (to stay logged in):
```bash
python copy_chrome_cookies.py
```
- Close all Chrome windows first
- This copies your login cookies to the automation profile

### Running the Download

Run the main script:
```bash
python looker_download.py
```

The script will:
1. Open Chrome with automation profile
2. Navigate to your Looker Studio dashboard
3. Select previous week's date range (Sunday-Saturday)
4. Export the table data as CSV
5. Save with filename format: `data_week{XX}_{YYYY}.csv`
6. Automatically run consolidation (optional)

### Consolidating Data

To manually combine all weekly CSV files:
```bash
python consolidate_weekly_data.py
```

Creates:
- `master_data_all_weeks.xlsx` - All weeks combined
- `master_data_backup_{timestamp}.xlsx` - Timestamped backup

## Scheduling with Task Scheduler

To run automatically every Monday:

1. Open **Task Scheduler** (Windows)
2. Create Basic Task → Name it "Looker Studio Weekly Download"
3. Trigger: **Weekly** on **Monday** at your preferred time
4. Action: **Start a Program**
   - Program: `pythonw.exe` (silent mode) or `python.exe`
   - Arguments: `"C:\path\to\looker_download.py"`
   - Start in: `C:\path\to\project\folder`

## How It Works

### Date Range Selection

The script calculates the previous week's Sunday-Saturday range:
- Today is any day → finds last Sunday → goes back 7 more days
- Handles month/year boundaries automatically
- Uses ISO week numbers for file naming

### Table Export Process

1. **Find Table**: Locates Looker Studio table component (`ng2-canvas-component`)
2. **Hover Action**: Moves mouse to top-right corner to reveal menu button
3. **Click Export**: Finds and clicks the 3-dot menu → Export data
4. **Format Options**: Selects CSV with "Keep value formatting"
5. **Download**: Waits for file completion, renames with week number

### ChromeDriver Management

- Detects your Chrome version from Windows registry
- Downloads matching ChromeDriver from official Chrome for Testing API
- Stores in `~/.chromedriver` for reuse
- Auto-updates if older than 7 days

## Troubleshooting

### "Not logged in" error
```bash
python copy_chrome_cookies.py
```
Run this to copy your login cookies.

### ChromeDriver version mismatch
The script auto-downloads the correct version. If issues persist:
- Update Chrome to the latest version
- Delete `~/.chromedriver` folder to force re-download

### Table not found
- Check the `LOOKER_URL` is correct
- Look at debug screenshots in output folder
- Your Looker table might use different classes (check logs)

### Export button not appearing
- The script takes screenshots at each step
- Check `after_hover.png` to see if menu appeared
- Hover position might need adjustment for your dashboard

### Date selection fails
- Screenshots show exactly which calendar elements were found
- Week calculations are logged - verify they're correct
- Calendar structure might differ (check `date_menu_after_click.png`)

## File Structure

```
looker-studio-automation/
│
├── looker_download.py          # Main download script
├── consolidate_weekly_data.py  # Combine weekly CSVs
├── copy_chrome_cookies.py      # Cookie copy utility
├── requirements.txt            # Python dependencies
└── README.md                   # This file

Output folder structure:
├── data_week45_2025.csv        # Weekly downloads
├── data_week46_2025.csv
├── master_data_all_weeks.xlsx  # Consolidated data
├── download_log.txt            # Detailed logs
├── consolidate_log.txt
└── *.png                       # Debug screenshots
```

## Debug Screenshots

The script automatically saves screenshots:
- `debug_initial.png` - Dashboard loaded
- `date_menu_after_click.png` - Date picker opened
- `after_date_selection.png` - Dates selected
- `table_before_hover.png` - Table found
- `after_hover.png` - Menu revealed
- `menu_button{N}.png` - Each menu attempt

Use these to diagnose issues!

## Advanced Configuration

### Custom Date Ranges

Modify the date calculation in `looker_download.py`:
```python
# For last 30 days instead of previous week:
last_sunday = today - timedelta(days=30)
last_saturday = today - timedelta(days=1)
```

### Different File Naming

Change the pattern in `consolidate_weekly_data.py`:
```python
pattern = os.path.join(WEEKLY_DATA_FOLDER, "your_custom_pattern*.csv")
```

### Multiple Dashboards

Create separate config files or use command-line arguments:
```python
if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        LOOKER_URL = sys.argv[1]
    success = download_looker_data()
```

## Known Limitations

- **Windows only** (Chrome automation uses Windows-specific paths)
- **Single dashboard** per run (can be extended for multiple)
- **Requires GUI** (cannot run headless due to Looker's anti-bot measures)
- **Manual first login** (automated logins may be blocked by Google)

## Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable
5. Submit a pull request

## License

MIT License - feel free to modify and use for your purposes.

## Acknowledgments

Built with:
- [Selenium](https://www.selenium.dev/) - Browser automation
- [Pandas](https://pandas.pydata.org/) - Data processing
- [OpenPyXL](https://openpyxl.readthedocs.io/) - Excel file creation

## Support

If you encounter issues:
1. Check the log files (`download_log.txt`, `consolidate_log.txt`)
2. Review debug screenshots in output folder
3. Enable verbose logging by setting `logging.DEBUG`
4. Open an issue on GitHub with logs and screenshots

---

**Note**: This tool automates interaction with Looker Studio. Make sure your usage complies with your organization's policies and Google's Terms of Service.
