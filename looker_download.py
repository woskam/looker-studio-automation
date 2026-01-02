"""
Looker Studio Automated Weekly Download Script
Downloads data every Monday for the last 7 days
"""

import os
import time
import glob
from datetime import datetime, timedelta
from pathlib import Path
import logging

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import requests
import zipfile
import io

# ============================================
# CONFIGURATION
# ============================================
LOOKER_URL = "https://lookerstudio.google.com/reporting/YOUR-REPORT-ID/page/YOUR-PAGE-ID"
OUTPUT_FOLDER = r"C:\path\to\your\output\folder"
DOWNLOAD_FOLDER = os.path.join(os.path.expanduser("~"), "Downloads")
LOG_FILE = os.path.join(OUTPUT_FOLDER, "download_log.txt")

# File pattern for downloaded files
DOWNLOAD_FILE_PATTERN = "*Table*.csv"  # Adjust to match your Looker export filename

# ============================================

# Setup logging with UTF-8 encoding to handle special characters
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

def get_week_number():
    """Calculate week number and year for the previous week (Sunday to Saturday)"""
    today = datetime.now()
    
    # How many days ago was the most recent Sunday?
    days_since_sunday = (today.weekday() + 1) % 7
    if days_since_sunday == 0:
        days_since_sunday = 7
    
    # Sunday of previous week
    last_sunday = today - timedelta(days=days_since_sunday + 7)
    
    # Use ISO week number of the Sunday
    week_number = last_sunday.isocalendar()[1]
    year = last_sunday.year
    
    return week_number, year

def get_chrome_version():
    """Get Chrome version from registry (Windows only)"""
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon")
        version, _ = winreg.QueryValueEx(key, "version")
        winreg.CloseKey(key)
        return version
    except:
        return None

def download_chromedriver():
    """Download correct chromedriver version for win64"""
    driver_dir = os.path.join(os.path.expanduser("~"), ".chromedriver")
    os.makedirs(driver_dir, exist_ok=True)
    
    driver_path = os.path.join(driver_dir, "chromedriver.exe")
    
    # If driver already exists and is recent (< 7 days), use it
    if os.path.exists(driver_path):
        file_age = time.time() - os.path.getmtime(driver_path)
        if file_age < 7 * 24 * 3600:  # 7 days
            logging.info(f"Using existing chromedriver: {driver_path}")
            return driver_path
    
    # Get Chrome version
    chrome_version = get_chrome_version()
    if not chrome_version:
        logging.error("Cannot determine Chrome version")
        return None
    
    # Get major version (e.g., 142 from 142.0.7444.162)
    major_version = chrome_version.split('.')[0]
    
    logging.info(f"Chrome version: {chrome_version}")
    
    # Download correct chromedriver for win64
    try:
        # Get latest version for this Chrome version
        api_url = f"https://googlechromelabs.github.io/chrome-for-testing/latest-versions-per-milestone-with-downloads.json"
        response = requests.get(api_url)
        data = response.json()
        
        if major_version not in data['milestones']:
            logging.error(f"No chromedriver for Chrome {major_version}")
            return None
        
        # Find win64 download URL
        downloads = data['milestones'][major_version]['downloads']['chromedriver']
        win64_url = None
        for download in downloads:
            if download['platform'] == 'win64':
                win64_url = download['url']
                break
        
        if not win64_url:
            logging.error("No win64 chromedriver found")
            return None
        
        logging.info(f"Downloading chromedriver from: {win64_url}")
        
        # Download and unzip
        response = requests.get(win64_url)
        with zipfile.ZipFile(io.BytesIO(response.content)) as zip_file:
            # Find chromedriver.exe in the zip
            for file_name in zip_file.namelist():
                if file_name.endswith('chromedriver.exe'):
                    # Extract to driver_dir
                    with zip_file.open(file_name) as source:
                        with open(driver_path, 'wb') as target:
                            target.write(source.read())
                    break
        
        logging.info(f"Chromedriver installed: {driver_path}")
        return driver_path
        
    except Exception as e:
        logging.error(f"Error downloading chromedriver: {e}")
        return None

def setup_chrome_driver():
    """Setup Chrome driver with automation profile"""
    chrome_options = Options()
    
    # Use a dedicated automation profile (not the default profile)
    automation_dir = os.path.join(os.path.expanduser("~"), "chrome_automation_data")
    chrome_options.add_argument(f"--user-data-dir={automation_dir}")
    
    # Extra options for stability
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--start-maximized")
    
    # Download settings
    prefs = {
        "download.default_directory": DOWNLOAD_FOLDER,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.notifications": 2  # Block notifications
    }
    chrome_options.add_experimental_option("prefs", prefs)
    
    # Disable automation flags
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    
    # Download correct chromedriver
    driver_path = download_chromedriver()
    if not driver_path:
        raise Exception("Cannot download chromedriver")
    
    service = Service(driver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    
    return driver

def wait_for_download(filename_pattern, timeout=60):
    """Wait until download is complete"""
    logging.info(f"Waiting for download: {filename_pattern}")
    
    end_time = time.time() + timeout
    while time.time() < end_time:
        # Search for files matching the pattern
        files = glob.glob(os.path.join(DOWNLOAD_FOLDER, filename_pattern))
        
        if files:
            # Check if there's no .crdownload file (Chrome download in progress)
            downloading = glob.glob(os.path.join(DOWNLOAD_FOLDER, "*.crdownload"))
            if not downloading:
                # Take the most recent file
                latest_file = max(files, key=os.path.getctime)
                logging.info(f"Download complete: {latest_file}")
                return latest_file
        
        time.sleep(1)
    
    raise TimeoutError(f"Download timeout after {timeout} seconds")

def close_chrome_processes():
    """Close all Chrome processes"""
    import subprocess
    try:
        # Try to close Chrome processes
        subprocess.run(['taskkill', '/F', '/IM', 'chrome.exe'], 
                      capture_output=True, 
                      timeout=5)
        time.sleep(2)  # Wait a moment
        logging.info("Chrome processes closed")
    except:
        pass  # No problem if Chrome is not running

def download_looker_data():
    """Main function to download Looker Studio data"""
    driver = None
    
    try:
        logging.info("=== Start Looker Studio Download ===")
        
        # Check if output folder exists
        os.makedirs(OUTPUT_FOLDER, exist_ok=True)
        
        # Setup driver
        logging.info("Starting Chrome browser...")
        driver = setup_chrome_driver()
        
        # Go to Looker Studio dashboard
        logging.info(f"Navigating to dashboard...")
        driver.get(LOOKER_URL)
        
        # Check if we're logged in - if not, wait for manual login
        time.sleep(5)
        if "accounts.google.com" in driver.current_url or "signin" in driver.current_url.lower():
            logging.info("=" * 60)
            logging.info("NOT LOGGED IN - Please log in manually in the Chrome window")
            logging.info("Waiting 60 seconds for login...")
            logging.info("=" * 60)
            time.sleep(60)  # Give user time to log in
        
        # Wait until page is loaded
        logging.info("Waiting for dashboard to load...")
        time.sleep(10)  # Extra wait time for Looker Studio
        
        # Screenshot to see HTML structure
        debug_path = os.path.join(OUTPUT_FOLDER, "debug_initial.png")
        driver.save_screenshot(debug_path)
        logging.info(f"Debug screenshot: {debug_path}")
        
        # Print all visible divs with 'table' or 'chart' in class
        logging.info("Searching for table/chart elements...")
        possible_tables = driver.find_elements(By.XPATH, "//*[contains(@class, 'table') or contains(@class, 'chart') or contains(@class, 'grid') or contains(@class, 'data')]")
        logging.info(f"Found {len(possible_tables)} possible table/chart elements")
        
        for idx, elem in enumerate(possible_tables[:10]):  # Max 10
            try:
                if elem.is_displayed():
                    class_name = elem.get_attribute('class')
                    tag_name = elem.tag_name
                    logging.info(f"  Element {idx}: <{tag_name}> class='{class_name[:100]}'")
            except:
                pass
        
        # Select dates: Sunday to Saturday of previous week
        logging.info("Selecting dates (Sunday to Saturday of previous week)...")
        
        # Calculate the dates
        today = datetime.now()
        
        # Debug logging
        logging.info(f"Today: {today.strftime('%A %d %B %Y')}, weekday={today.weekday()}")
        
        # weekday(): 0=Monday, 1=Tuesday, 2=Wednesday, 3=Thursday, 4=Friday, 5=Saturday, 6=Sunday
        # We want: Sunday of PREVIOUS week (not this week!)
        
        # Step 1: Calculate how many days ago the most recent Sunday was
        # Monday = 1 day ago, Tuesday = 2 days, etc.
        # Sunday itself = 0 days (but we want previous Sunday then)
        days_since_sunday = (today.weekday() + 1) % 7
        
        if days_since_sunday == 0:
            # Today is Sunday - the most recent Sunday is TODAY
            # But we want Sunday of PREVIOUS week
            days_to_last_week_sunday = 7
        else:
            # Otherwise: go back to most recent Sunday, then 7 more days back
            days_to_last_week_sunday = days_since_sunday + 7
        
        logging.info(f"Days back to Sunday of previous week: {days_to_last_week_sunday}")
        
        # Sunday of previous week
        last_sunday = today - timedelta(days=days_to_last_week_sunday)
        logging.info(f"Sunday of previous week: {last_sunday.strftime('%A %d %B %Y')}")
        
        # Verify it's actually a Sunday
        if last_sunday.weekday() != 6:
            logging.error(f"ERROR: Calculated date is not a Sunday! weekday={last_sunday.weekday()}")
        
        # Saturday of previous week = Sunday + 6 days
        last_saturday = last_sunday + timedelta(days=6)
        logging.info(f"Saturday of previous week: {last_saturday.strftime('%A %d %B %Y')}")
        
        # Verify it's actually a Saturday
        if last_saturday.weekday() != 5:
            logging.error(f"ERROR: Calculated date is not a Saturday! weekday={last_saturday.weekday()}")
        
        # Extract day numbers for clicking in calendar
        start_day = last_sunday.day
        end_day = last_saturday.day
        start_month = last_sunday.strftime("%B")  # Full month name
        end_month = last_saturday.strftime("%B")
        
        # Format dates (Windows compatible - no %-d)
        start_date_str = last_sunday.strftime("%b %d, %Y").replace(" 0", " ")  # "Nov 10, 2025"
        end_date_str = last_saturday.strftime("%b %d, %Y").replace(" 0", " ")  # "Nov 16, 2025"
        
        logging.info(f"Period: {start_date_str} - {end_date_str}")
        logging.info(f"Start day number: {start_day}, End day number: {end_day}")
        
        try:
            # Click on date range selector to open calendars
            date_selector = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'date') or contains(@aria-label, 'date') or contains(@class, 'date-range')]"))
            )
            date_selector.click()
            time.sleep(3)  # Wait for calendars to appear
            
            # Screenshot AFTER clicking on date selector
            date_menu_screenshot = os.path.join(OUTPUT_FOLDER, "date_menu_after_click.png")
            driver.save_screenshot(date_menu_screenshot)
            logging.info(f"Screenshot of date calendars: {date_menu_screenshot}")
            
            # The calendars are now visible - no "Custom" option needed!
            # Click directly on the day numbers in the calendars
            
            # STEP 1: Click on start date (Sunday = day 10 in left calendar)
            logging.info(f"Trying to click start date: day {start_day}")
            start_clicked = False
            
            # Try multiple ways to find the start day
            start_selectors = [
                # Method 1: Via aria-label with full date
                f"//button[@aria-label='{start_month} {start_day}, {last_sunday.year}']",
                f"//button[@aria-label='{start_month[:3]} {start_day}, {last_sunday.year}']",
                # Method 2: Via day number in calendar cell
                f"//div[contains(@class, 'mat-calendar-body-cell-content') and text()='{start_day}']",
                f"//button[contains(@class, 'mat-calendar-body-cell') and .//div[text()='{start_day}']]",
                # Method 3: Simple - just button with day number
                f"//button[contains(@class, 'mat-calendar') and contains(., '{start_day}')]",
            ]
            
            for selector in start_selectors:
                try:
                    start_day_button = driver.find_element(By.XPATH, selector)
                    if start_day_button.is_displayed():
                        start_day_button.click()
                        time.sleep(1)
                        logging.info(f"✓ Start date clicked with selector: {selector}")
                        start_clicked = True
                        break
                except Exception as e:
                    logging.debug(f"Start selector '{selector}' doesn't work: {e}")
                    continue
            
            if not start_clicked:
                logging.error(f"Could not click start date (day {start_day}) with any selector")
            
            # STEP 2: Click on end date (Saturday = day 15 in RIGHT calendar)
            logging.info(f"Trying to click end date: day {end_day} in RIGHT calendar")
            end_clicked = False
            
            # First: find all calendars on the page
            try:
                all_calendars = driver.find_elements(By.XPATH, "//div[contains(@class, 'mat-calendar-content')]")
                logging.info(f"Found {len(all_calendars)} calendars")
                
                if len(all_calendars) >= 2:
                    # There are 2 calendars - use the SECOND (right) for end date
                    right_calendar = all_calendars[1]
                    logging.info("Searching in right calendar for end date...")
                    
                    # Search day-button WITHIN the right calendar
                    end_selectors_in_calendar = [
                        f".//button[@aria-label='{end_month} {end_day}, {last_saturday.year}']",
                        f".//button[@aria-label='{end_month[:3]} {end_day}, {last_saturday.year}']",
                        f".//div[contains(@class, 'mat-calendar-body-cell-content') and text()='{end_day}']",
                        f".//button[contains(@class, 'mat-calendar-body-cell') and .//div[text()='{end_day}']]",
                        f".//button[contains(., '{end_day}') and not(contains(@class, 'mat-calendar-body-disabled'))]",
                    ]
                    
                    for selector in end_selectors_in_calendar:
                        try:
                            end_day_button = right_calendar.find_element(By.XPATH, selector)
                            if end_day_button.is_displayed():
                                end_day_button.click()
                                time.sleep(1)
                                logging.info(f"✓ End date clicked in RIGHT calendar with selector: {selector}")
                                end_clicked = True
                                break
                        except Exception as e:
                            logging.debug(f"End selector '{selector}' doesn't work in right calendar: {e}")
                            continue
                else:
                    logging.warning(f"Expected 2 calendars but found {len(all_calendars)}")
                    
            except Exception as e:
                logging.warning(f"Could not find calendars: {e}")
            
            # Fallback: try global selectors (if right calendar approach doesn't work)
            if not end_clicked:
                logging.info("Right calendar method didn't work, try global selectors...")
                end_selectors = [
                    f"//button[@aria-label='{end_month} {end_day}, {last_saturday.year}']",
                    f"//button[@aria-label='{end_month[:3]} {end_day}, {last_saturday.year}']",
                ]
                
                for selector in end_selectors:
                    try:
                        # Find ALL buttons with this day
                        all_day_buttons = driver.find_elements(By.XPATH, selector)
                        logging.info(f"Found {len(all_day_buttons)} buttons for day {end_day}")
                        
                        # Click on the SECOND (right calendar)
                        if len(all_day_buttons) >= 2:
                            all_day_buttons[1].click()  # Index 1 = second calendar
                            time.sleep(1)
                            logging.info(f"✓ End date clicked (second button) with selector: {selector}")
                            end_clicked = True
                            break
                        elif len(all_day_buttons) == 1:
                            all_day_buttons[0].click()
                            time.sleep(1)
                            logging.info(f"✓ End date clicked (only button) with selector: {selector}")
                            end_clicked = True
                            break
                    except Exception as e:
                        logging.debug(f"Global selector '{selector}' doesn't work: {e}")
                        continue
            
            if not end_clicked:
                logging.error(f"Could not click end date (day {end_day}) with any method")
            
            # Screenshot AFTER date selection
            after_date_screenshot = os.path.join(OUTPUT_FOLDER, "after_date_selection.png")
            driver.save_screenshot(after_date_screenshot)
            logging.info(f"Screenshot after date selection: {after_date_screenshot}")
            
            # STEP 3: Click Apply button
            try:
                apply_button = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Apply') or contains(., 'Apply')]"))
                )
                apply_button.click()
                logging.info(f"✓ Apply button clicked, dates applied: {start_date_str} - {end_date_str}")
                
                # IMPORTANT: Wait long enough for the table to reload with new data
                logging.info("Waiting for table reload with new data (20 seconds)...")
                time.sleep(20)
                
                # Try to wait until loading indicator disappears
                try:
                    logging.info("Checking if loading indicator is present...")
                    # Wait until any loading spinners are gone
                    WebDriverWait(driver, 30).until_not(
                        EC.presence_of_element_located((By.XPATH, "//*[contains(@class, 'loading') or contains(@class, 'spinner') or contains(@class, 'progress')]"))
                    )
                    logging.info("✓ Loading indicator disappeared")
                except:
                    logging.info("No loading indicator found (or already gone)")
                
                # Extra wait time for safety
                time.sleep(5)
                
                # Screenshot AFTER data reload
                after_reload_screenshot = os.path.join(OUTPUT_FOLDER, "after_data_reload.png")
                driver.save_screenshot(after_reload_screenshot)
                logging.info(f"Screenshot after data reload: {after_reload_screenshot}")
                
            except Exception as e:
                logging.warning(f"Could not find/click Apply button: {e}")
                # Try pressing ESC to close calendar
                try:
                    from selenium.webdriver.common.keys import Keys
                    driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
                    time.sleep(2)
                    logging.info("ESC pressed to close calendar")
                except:
                    pass
                
        except Exception as e:
            logging.warning(f"Could not select dates (dashboard might already use the right period): {e}")
        
        # Search for the table/chart container first
        logging.info("Searching for table...")
        time.sleep(3)
        
        # Close any overlay/dialogs first
        try:
            close_buttons = driver.find_elements(By.XPATH, "//button[contains(text(), 'Cancel') or contains(text(), 'Close') or contains(@aria-label, 'Close')]")
            for btn in close_buttons:
                if btn.is_displayed():
                    btn.click()
                    time.sleep(1)
                    logging.info("Overlay closed")
                    break
        except:
            pass
        
        # Click on backdrop if it exists
        try:
            backdrop = driver.find_element(By.CLASS_NAME, "cdk-overlay-backdrop")
            if backdrop.is_displayed():
                backdrop.click()
                time.sleep(1)
                logging.info("Backdrop clicked")
        except:
            pass
        
        # Find the table (Looker Studio uses ng2-canvas-component)
        table = None
        try:
            table = driver.find_element(By.XPATH, "//ng2-canvas-component[contains(@class, 'simple-table')]")
            logging.info("Looker Studio table component found!")
        except:
            try:
                table = driver.find_element(By.XPATH, "//ng2-component-header[contains(@class, 'simple-table')]")
                logging.info("Looker Studio header component found!")
            except:
                logging.error("No Looker Studio table component found")
                raise Exception("Table component not found")
        
        # Scroll to the table AND wait until it's fully visible
        logging.info("Scrolling to table and waiting until fully loaded...")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});", table)
        time.sleep(3)
        
        # Check once more if there are no loading indicators
        try:
            loading_elements = driver.find_elements(By.XPATH, "//*[contains(@class, 'loading') or contains(@class, 'spinner')]")
            visible_loaders = [elem for elem in loading_elements if elem.is_displayed()]
            if visible_loaders:
                logging.info(f"Still {len(visible_loaders)} loading indicators visible, waiting...")
                time.sleep(10)
        except:
            pass
        
        # Screenshot of the table BEFORE hovering
        table_before_hover = os.path.join(OUTPUT_FOLDER, "table_before_hover.png")
        driver.save_screenshot(table_before_hover)
        logging.info(f"Screenshot of table before hover: {table_before_hover}")
        
        # Hover strategy: top-right IN the table (where 3-dots are)
        logging.info("Hovering top-right IN the table (where 3-dots are)...")
        actions = ActionChains(driver)
        
        # Hover over the table
        actions.move_to_element(table).perform()
        time.sleep(1)
        
        # Get the size and position of the table
        try:
            size = table.size
            width = size['width']
            height = size['height']
            
            logging.info(f"Table size: {width}x{height}")
            
            # Hover top-right IN the table - where the 3-dots appear
            # Try different positions top-right
            for x_offset in [width - 80, width - 100, width - 120, width - 150]:
                for y_offset in [10, 20, 30, 40]:
                    try:
                        actions.move_to_element_with_offset(table, x_offset, y_offset).perform()
                        time.sleep(0.5)
                        logging.info(f"  Hover position: x={x_offset}, y={y_offset}")
                    except Exception as e:
                        logging.warning(f"  Hover failed at {x_offset},{y_offset}: {e}")
            
            time.sleep(2)
            logging.info("Hover over table top-right executed")
            
        except Exception as e:
            logging.warning(f"Could not hover with offset: {e}")
        
        # Search for buttons in the ng2-component-header (where the 3-dots are)
        logging.info("Searching for buttons in ng2-component-header...")
        
        export_data_found = False
        
        # Find the header component
        try:
            header = driver.find_element(By.XPATH, "//ng2-component-header[contains(@class, 'simple-table')]")
        except Exception as e:
            logging.error(f"Could not find ng2-component-header: {e}")
            raise Exception("Header not found")
        
        # Try to force hover state with JavaScript
        logging.info("Forcing hover state with JavaScript...")
        driver.execute_script("""
            arguments[0].classList.add('hover');
            arguments[0].classList.add('mat-hover');
            
            // Trigger mouse events
            var event = new MouseEvent('mouseenter', {bubbles: true, cancelable: true});
            arguments[0].dispatchEvent(event);
            
            var event2 = new MouseEvent('mouseover', {bubbles: true, cancelable: true});
            arguments[0].dispatchEvent(event2);
        """, header)
        
        time.sleep(2)
        
        # Now search for buttons (including hidden but in DOM)
        all_buttons_in_header = header.find_elements(By.XPATH, ".//button")
        logging.info(f"Found {len(all_buttons_in_header)} buttons in header (including hidden)")
        
        # Log all buttons
        header_buttons = []
        for idx, btn in enumerate(all_buttons_in_header):
            aria_label = btn.get_attribute('aria-label') or ''
            class_name = btn.get_attribute('class') or ''
            is_displayed = btn.is_displayed()
            logging.info(f"  Button {idx}: visible={is_displayed}, aria='{aria_label[:50]}', class='{class_name[:50]}'")
            
            # Add (even if not visible - we'll try them anyway)
            header_buttons.append(btn)
        
        # Screenshot AFTER hover
        screenshot_path = os.path.join(OUTPUT_FOLDER, "after_hover.png")
        driver.save_screenshot(screenshot_path)
        logging.info(f"Screenshot after hover: {screenshot_path}")
        
        if not header_buttons:
            logging.error("No buttons found in header DOM at all")
            raise Exception("No buttons in header DOM")
        
        # Try each button
        logging.info(f"Trying {len(header_buttons)} buttons...")
        for idx, button in enumerate(header_buttons):
            try:
                aria_label = button.get_attribute('aria-label') or ''
                logging.info(f"Trying button #{idx+1} (aria: '{aria_label[:50]}')")
                
                # Skip filter button
                if 'filter' in aria_label.lower():
                    logging.info(f"  Skipping - this is the filter button")
                    continue
                
                # Try click
                try:
                    # Extra hover for this specific button
                    actions = ActionChains(driver)
                    actions.move_to_element(button).perform()
                    time.sleep(1)
                    
                    button.click()
                    logging.info(f"  Button #{idx+1} clicked")
                    time.sleep(3)  # Extra wait time for menu
                    
                    # Search for Export data with multiple variants
                    export_found = False
                    export_selectors = [
                        "//*[text()='Export data']",
                        "//*[contains(text(), 'Export')]",
                        "//button[contains(., 'Export')]",
                        "//div[contains(., 'Export data')]",
                        "//*[@aria-label='Export data']"
                    ]
                    
                    for selector in export_selectors:
                        try:
                            export_data = driver.find_element(By.XPATH, selector)
                            if export_data.is_displayed():
                                logging.info(f"Export option found with selector: {selector}")
                                export_found = True
                                export_data_found = True
                                
                                # Click on Export
                                logging.info("Clicking on 'Export'...")
                                export_data.click()
                                time.sleep(2)
                                break
                        except:
                            continue
                    
                    if export_found:
                        break
                    else:
                        logging.info(f"  Button #{idx+1} has no Export option")
                        # Screenshot of menu
                        menu_screenshot = os.path.join(OUTPUT_FOLDER, f"menu_button{idx+1}.png")
                        driver.save_screenshot(menu_screenshot)
                        logging.info(f"  Menu screenshot: {menu_screenshot}")
                        
                        # Close menu
                        try:
                            driver.find_element(By.TAG_NAME, "body").click()
                            time.sleep(1)
                        except:
                            pass
                        
                except Exception as e:
                    logging.warning(f"  Error with button #{idx+1}: {e}")
                    continue
                    
            except Exception as e:
                logging.warning(f"  Error with button #{idx+1}: {e}")
                continue
        
        if not export_data_found:
            screenshot_path = os.path.join(OUTPUT_FOLDER, "debug_no_export_data.png")
            driver.save_screenshot(screenshot_path)
            logging.error(f"'Export data' not found in any menu. Screenshot: {screenshot_path}")
            raise Exception("'Export data' not found")
        
        # CSV is default already selected, check "Keep value formatting"
        logging.info("Checking 'Keep value formatting'...")
        
        try:
            # Find the checkbox for "Keep value formatting"
            keep_formatting_selectors = [
                "//input[@type='checkbox' and following-sibling::*[contains(text(), 'Keep value formatting')]]",
                "//mat-checkbox[contains(., 'Keep value formatting')]",
                "//*[contains(text(), 'Keep value formatting')]",
            ]
            
            for selector in keep_formatting_selectors:
                try:
                    keep_formatting = driver.find_element(By.XPATH, selector)
                    keep_formatting.click()
                    time.sleep(1)
                    logging.info(f"'Keep value formatting' checked with: {selector}")
                    break
                except:
                    continue
        except Exception as e:
            logging.warning(f"Could not check 'Keep value formatting': {e}")
        
        # Click on "Export" button
        logging.info("Clicking on 'Export' button in dialog...")
        export_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Export') or contains(., 'Export')]"))
        )
        export_button.click()
        
        # Wait for download
        logging.info("Download started, waiting for completion...")
        downloaded_file = wait_for_download(DOWNLOAD_FILE_PATTERN, timeout=60)
        
        # Rename and move file
        week_num, year = get_week_number()
        new_filename = f"data_week{week_num:02d}_{year}.csv"
        new_filepath = os.path.join(OUTPUT_FOLDER, new_filename)
        
        # FIX: Check if destination file already exists and remove it first
        if os.path.exists(new_filepath):
            logging.info(f"Existing file found, removing: {new_filepath}")
            os.remove(new_filepath)
        
        # Move and rename
        os.rename(downloaded_file, new_filepath)
        logging.info(f"File saved as: {new_filepath}")
        
        logging.info("=== Download Successfully Completed ===")
        
        # Optional: Automatically combine weekly files into master file
        try:
            logging.info("Starting automatic file combination...")
            import consolidate_weekly_data
            consolidate_weekly_data.consolidate_weekly_data()
        except Exception as e:
            logging.warning(f"Could not run automatic combination (not critical): {e}")
        
        return True
        
    except Exception as e:
        logging.error(f"ERROR during download: {str(e)}", exc_info=True)
        return False
        
    finally:
        if driver:
            logging.info("Closing browser...")
            driver.quit()

if __name__ == "__main__":
    success = download_looker_data()
    
    if success:
        print("\n✓ Download successfully completed!")
        exit(0)
    else:
        print("\n✗ Download failed - check log file for details")
        exit(1)
