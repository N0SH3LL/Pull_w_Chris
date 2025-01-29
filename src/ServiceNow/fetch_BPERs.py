"""Needs updated BPER_sysids.json to find all the files. Goes python>powershell>firefox>servicenow to fetch BPERs by sys_id. Has to download, rename, and move one by one. Not ideal, need API access
"""

import os
import subprocess
import time
import logging
import shutil
import json
import sys
import glob
from typing import Optional, Dict, List, Any

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Base URL for the ServiceNow instance
BASE_URL = "https://wpshelpdesk.servicenowservices.com/"
DOWNLOAD_FILENAME = "sn_compliance_bulk_policy_exception.pdf"

def load_sysids(json_file: str) -> Dict[str, str]:
   """Load ServiceNow sys_ids from config file.

   Args:
       json_file: Path to sys_ids JSON config

   Returns:
       Dict mapping BPER IDs to sys_ids

   Raises:
       FileNotFoundError: If config file missing
       JSONDecodeError: If invalid JSON
   """
   
   try:
       with open(json_file, 'r') as f:
           return json.load(f)
   except FileNotFoundError:
       logging.error(f"sysids.json file not found at {json_file}")
       sys.exit(1)
   except json.JSONDecodeError:
       logging.error(f"Error decoding JSON from {json_file}") 
       sys.exit(1)

def get_downloads_folder() -> str:
   """Get Windows default Downloads folder path.

   Returns:
       str: Path to Downloads folder
   """
   return os.path.join(os.path.expanduser('~'), 'Downloads')

def generate_powershell_command(sys_id: str, download_dir: str, bper: str) -> str: # I'm sure there is a more efficient way
   """Generate PowerShell command for downloading BPER.

   Args:
       sys_id: ServiceNow sys_id
       download_dir: Download directory path
       bper: BPER ID

   Returns:
       PowerShell command string
   """
   return f"""
    $PdfUrl = "{BASE_URL}sn_compliance_bulk_policy_exception.do?sys_id={sys_id}&PDF&sysparm_view=&related_list_filter=&sysparm_domain="
    $process = Start-Process "firefox.exe" -ArgumentList $PdfUrl -PassThru

    # Wait for download to complete
    $maxWaitTime = 60  # Maximum wait time in seconds
    $startTime = Get-Date
    do {{
        Start-Sleep -Seconds 1
        $downloadedFile = Get-ChildItem -Path "{download_dir}" | Where-Object {{ $_.Name -like "*{DOWNLOAD_FILENAME}*" }} | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $elapsedTime = (Get-Date) - $startTime
    }} until ($downloadedFile -or $elapsedTime.TotalSeconds -ge $maxWaitTime)

    if ($downloadedFile) {{
        Write-Output "File downloaded: $($downloadedFile.FullName)"
    }} else {{
        Write-Output "Download failed or timed out"
    }}

    # Close the Firefox tab
    $process | Stop-Process -Force

    # Wait a bit before next download
    Start-Sleep -Seconds 2
    """

def download_and_rename_bper(bper, sys_id, downloads_dir, destination_dir): # Download and rename a single BPER
    powershell_command = generate_powershell_command(sys_id, downloads_dir, bper)
    try:
        result = subprocess.run(["powershell", "-Command", powershell_command], capture_output=True, text=True, check=True)
        logging.info(result.stdout)

        # Check if file was downloaded
        downloaded_files = glob.glob(os.path.join(downloads_dir, f"*{DOWNLOAD_FILENAME}*"))
        if downloaded_files:
            downloaded_file = max(downloaded_files, key=os.path.getctime)
            new_filename = f"{bper}.pdf"
            destination_path = os.path.join(destination_dir, new_filename)
            
            # Move and rename the file
            shutil.move(downloaded_file, destination_path)
            logging.info(f"Successfully downloaded, renamed, and moved: {destination_path}")
            return destination_path
        else:
            logging.error(f"Download failed for {bper}")
            return None
    except subprocess.CalledProcessError as e:
        logging.error(f"Error running PowerShell command for {bper}: {e}")
        return None

def main(bper_list, destination_dir): # Main function for multiple BPERs
    """Main function for multiple BPERs.

   Args:
       bper_list: List of BPER IDs to fetch
       destination_dir: Directory for storing downloads

   Raises:
       subprocess.CalledProcessError: If PowerShell commands fail
       OSError: If directory creation fails
   """
    
    os.makedirs(destination_dir, exist_ok=True)
    downloads_dir = get_downloads_folder()

    json_file = os.path.join(os.path.dirname(__file__), '..', '..', 'config', 'BPER_sysids.json')
    sysids = load_sysids(json_file)

    # Initial login
    login_url = f"{BASE_URL}now/nav/ui/classic/params/target/home.do"
    subprocess.run(["powershell", "-Command", f"Start-Process 'firefox.exe' -ArgumentList '{login_url}'"])
    input("Please log in to ServiceNow in the opened Firefox window. Press Enter when done...")

    for bper in bper_list:
        sys_id = sysids.get(bper)
        if sys_id:
            pdf_path = download_and_rename_bper(bper, sys_id, downloads_dir, destination_dir)
            if pdf_path:
                print(f"BPER PDF downloaded, renamed, and moved successfully: {pdf_path}")
            else:
                print(f"Failed to download, rename, or move PDF for {bper}.")
        else:
            logging.warning(f"No sys_id found for BPER number: {bper}")

        # Additional delay between downloads
        time.sleep(2)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python fetch_BPERs.py <destination_directory> <bper1> [bper2] [bper3] ...")
        sys.exit(1)
    
    destination_dir = sys.argv[1]
    bper_list = sys.argv[2:]
    main(bper_list, destination_dir)
