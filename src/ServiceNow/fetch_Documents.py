"""
Similar to fetch_BPERs, but for the supporting documents. Has to match, and depends on version, so having an updated doc_sysids.json is key. Downloads should be reviewed.
"""

import os
import subprocess
import time
import logging
import shutil
import json
import sys
import glob

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Base URL for the ServiceNow instance
BASE_URL = "https://wpshelpdesk.servicenowservices.com/"
DOWNLOAD_FILENAME = "sys_attachment.do"

def load_sysids(json_file):
    # Load the sysids from the JSON file - these have to be manually gathered by copying from the html and pulling; at least until I get API access
    try:
        with open(json_file, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        logging.error(f"doc_sysids.json file not found at {json_file}")
        sys.exit(1)
    except json.JSONDecodeError:
        logging.error(f"Error decoding JSON from {json_file}")
        sys.exit(1)

def get_downloads_folder():
    # Get the path to the Windows default Downloads folder
    return os.path.join(os.path.expanduser('~'), 'Downloads')

def generate_powershell_command(sys_id, download_dir, doc_name):
    # Generate the PowerShell command to download a single document
    # TODO move this to the powershell folder and call it 
    return f"""
    $DocUrl = "{BASE_URL}sys_attachment.do?sys_id={sys_id}&sysparm_this_url=dms_document_revision.do%3Fsys_id%3D3cd43fa71be1315411edeb57624bcbab%26sysparm_record_list%3Dstage%253Dpublished%255EORDERBYname%26sysparm_record_row%3D2%26sysparm_record_rows%3D1043%26sysparm_record_target%3Ddm"
    $process = Start-Process "firefox.exe" -ArgumentList $DocUrl -PassThru

    # Wait for download to complete
    $maxWaitTime = 60  # Maximum wait time in seconds
    $startTime = Get-Date
    do {{
        Start-Sleep -Seconds 1
        $downloadedFile = Get-ChildItem -Path "{download_dir}" | Where-Object {{ $_.Name -like "*.*" }} | Sort-Object LastWriteTime -Descending | Select-Object -First 1
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
    Start-Sleep -Seconds 1
    """

def download_and_move_document(doc_name, sys_id, downloads_dir, destination_dir):
    """Download and move document from ServiceNow.

   Args:
       doc_name: Name of document to download
       sys_id: ServiceNow sys_id
       downloads_dir: Download directory path
       destination_dir: Final destination path

   Returns:
       Path to downloaded file or None if failed

   Raises:
       subprocess.CalledProcessError: If PowerShell command fails 
       OSError: If file operations fail
   """
    powershell_command = generate_powershell_command(sys_id, downloads_dir, doc_name)
    try:
        logging.info(f"Attempting to download document: {doc_name}")
        result = subprocess.run(["powershell", "-Command", powershell_command], capture_output=True, text=True, check=True)
        logging.info(result.stdout)

        # Check if file was downloaded
        downloaded_files = glob.glob(os.path.join(downloads_dir, "*.*"))
        if downloaded_files:
            downloaded_file = max(downloaded_files, key=os.path.getctime)
            destination_path = os.path.join(destination_dir, os.path.basename(downloaded_file))
            
            # Move the file without renaming
            if os.path.exists(destination_path):
                logging.warning(f"File already exists at destination: {destination_path}")
                # Append timestamp to make the filename unique
                base, ext = os.path.splitext(destination_path)
                destination_path = f"{base}_{int(time.time())}{ext}"
            
            shutil.move(downloaded_file, destination_path)
            logging.info(f"Successfully downloaded and moved: {destination_path}")
            return destination_path
        else:
            logging.error(f"Download failed for {doc_name}")
            return None
    except subprocess.CalledProcessError as e:
        logging.error(f"Error running PowerShell command for {doc_name}: {e}")
        logging.error(f"PowerShell output: {e.output}")
        return None

def main(doc_list, destination_dir) -> None:
    """Orchestrate the document download process for multiple documents 
   
   Args:
       doc_list: List of document IDs to download
       destination_dir: Directory to store downloads

   Raises:
       FileNotFoundError: If sysids config missing
       subprocess.CalledProcessError: If PowerShell fails
       OSError: If directory creation fails
   """
    os.makedirs(destination_dir, exist_ok=True)
    downloads_dir = get_downloads_folder()

    json_file = os.path.join(os.path.dirname(__file__), '..', '..', 'config', 'doc_sysids.json')
    sysids = load_sysids(json_file)

    # Initial login
    # TODO Make one login work for both this and the SCC pull without multiple prompts
    login_url = f"{BASE_URL}now/nav/ui/classic/params/target/home.do"
    subprocess.run(["powershell", "-Command", f"Start-Process 'firefox.exe' -ArgumentList '{login_url}'"])
    input("Please log in to ServiceNow in the opened Firefox window. Press Enter when done...")

    for doc_name in doc_list:
        sys_id = sysids.get(doc_name)
        if sys_id:
            doc_path = download_and_move_document(doc_name, sys_id, downloads_dir, destination_dir)
            if doc_path:
                print(f"Document downloaded and moved successfully: {doc_path}")
            else:
                print(f"Failed to download or move document: {doc_name}")
        else:
            logging.warning(f"No sys_id found for document: {doc_name}")

        # Additional delay between downloads
        time.sleep(1)

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python fetch_Documents.py <destination_directory> <doc1> [doc2] [doc3] ...")
        sys.exit(1)
    
    destination_dir = sys.argv[1]
    doc_list = sys.argv[2:]
    main(doc_list, destination_dir)
