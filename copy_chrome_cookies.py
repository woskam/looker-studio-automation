"""
Helper script to copy cookies from your normal Chrome to automation profile
Run this once to stay logged in
"""

import os
import shutil
import sqlite3

def copy_chrome_cookies():
    """Copy cookies from normal Chrome to automation profile"""
    
    # Source: normal Chrome profile
    normal_chrome = os.path.join(os.path.expanduser("~"), "AppData", "Local", "Google", "Chrome", "User Data", "Default")
    
    # Destination: automation profile
    automation_profile = os.path.join(os.path.expanduser("~"), "chrome_automation_data", "Default")
    os.makedirs(automation_profile, exist_ok=True)
    
    # Important files to copy
    files_to_copy = [
        "Cookies",
        "Login Data",
        "Web Data"
    ]
    
    print("Copying Chrome authentication data...")
    print(f"From: {normal_chrome}")
    print(f"To: {automation_profile}")
    print()
    
    success_count = 0
    for filename in files_to_copy:
        source = os.path.join(normal_chrome, filename)
        dest = os.path.join(automation_profile, filename)
        
        if os.path.exists(source):
            try:
                shutil.copy2(source, dest)
                print(f"✓ {filename} copied")
                success_count += 1
            except Exception as e:
                print(f"✗ {filename} could not be copied: {e}")
        else:
            print(f"⚠ {filename} not found in normal Chrome")
    
    print()
    if success_count > 0:
        print(f"✓ {success_count} files successfully copied!")
        print("\nYou can now run looker_download.py.")
        print("You should be automatically logged in to Looker Studio.")
    else:
        print("✗ No files copied. Make sure Chrome is completely closed.")

if __name__ == "__main__":
    print("=== Chrome Cookie Copy Tool ===\n")
    print("IMPORTANT: Close ALL Chrome windows before continuing!\n")
    input("Press Enter when Chrome is completely closed...")
    
    copy_chrome_cookies()
