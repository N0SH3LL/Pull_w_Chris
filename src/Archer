"""
This file basically uses requests to headlessly hit up Archer, get the cookie and header info, and construct the url according to attestation number, then stream the downloads as html, before converting to http. May create empty files if the url path is missing or takes too long, so double check the size of everything downloaded. 
[!] FAIR WARNING [!] this will generate a file for each attestation you give it, even if that attestation's url doesn't exist. Check for emtpy files. 
"""

import os
import requests
from requests_negotiate_sspi import HttpNegotiateAuth
import logging
from urllib.parse import urljoin
import time
import concurrent.futures
import subprocess
from typing import Dict, List, Optional, Union

# Set up logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# Base URLs
BASE_URL = "https://wps-archapp-p01.corp.wpsic.com/RSAarcher/"
DEFAULT_URL = urljoin(BASE_URL, "Default.aspx")
ATTESTATION_URL_TEMPLATE = urljoin(BASE_URL, "Foundation/Print.aspx?exportSourceType=RecordView&levelId=171&contentId={}&castContentId=0&layoutId=460")

# Path to the CA certificate file
CA_CERT_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), '../..', 'config', 'cert.cer'))
POWERSHELL_SCRIPT_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), '../scripts/convert_to_pdf.ps1'))

def create_session() -> requests.Session:
    """Create authenticated session for Archer requests.

    Returns:
        requests.Session: Configured session with SSPI auth and custom SSL cert

    Raises:
        FileNotFoundError: If CA certificate file not found
    """
    session = requests.Session()
    session.auth = HttpNegotiateAuth()
    session.verify = CA_CERT_PATH
    return session

def get_initial_headers() -> Dict[str, str]:
    """Get headers required for initial Archer connection.

    Returns:
        Dict[str, str]: Dictionary of required HTTP headers
    """
    return {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/png,image/svg+xml,*/*;q=0.8',
        'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Accept-Language': 'en-US,en;q=0.5',
        'Connection': 'keep-alive',
        'Host': 'wps-archapp-p01.corp.wpsic.com',
        'Priority': 'u=0, i',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:129.0) Gecko/20100101 Firefox/129.0'
    }

def get_attestation_headers(cookies: Dict[str, str]) -> Dict[str, str]:
    """Get headers required for attestation requests.

    Args:
        cookies: Dictionary of cookies from authenticated session

    Returns:
        Dict[str, str]: Dictionary of required HTTP headers
    """
    return {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Encoding': 'gzip, deflate, br, zstd',
        'Accept-Language': 'en-US,en;q=0.9',
        'Connection': 'keep-alive',
        'Cookie': '; '.join([f"{name}={value}" for name, value in cookies.items()]),
        'Host': 'wps-archapp-p01.corp.wpsic.com',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate', 
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
        'sec-ch-ua': '"Chromium";v="128", "Not;A=Brand";v="24", "Google Chrome";v="128"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"'
    }

def fetch_initial_cookies(session: requests.Session) -> Optional[requests.cookies.RequestsCookieJar]:
    """Fetch initial authentication cookies by making a request to the Defaul.aspx page

    Args:
        session: Authenticated requests session

    Returns:
        Optional[RequestsCookieJar]: Session cookies if successful, None on failure

    Raises:
        requests.RequestException: For network or HTTP errors
    """
    try:
        headers = get_initial_headers()
        logging.debug(f"Sending request to {DEFAULT_URL}")
        response = session.get(DEFAULT_URL, headers=headers, timeout=60, allow_redirects=False)
        log_request_response("Fetch Initial Cookies", response.request, response)
        
        if response.status_code == 302:
            redirect_url = response.headers.get('Location')
            logging.debug(f"Following redirect to: {redirect_url}")
            response = session.get(urljoin(BASE_URL, redirect_url), headers=headers, timeout=60)
            log_request_response("Fetch Initial Cookies (After Redirect)", response.request, response)
        
        response.raise_for_status()
        logging.info("Successfully fetched initial cookies")
        return session.cookies
    except requests.RequestException as e:
        logging.error(f"Failed to fetch initial cookies: {str(e)}")
        return None

def fetch_attestation_html(session: requests.Session, attestation_id: int, cookies: Dict[str, str], max_retries: int = 5, wait_time: int = 10) -> Optional[str]:
    """Fetch HTML content for a specific attestation.

    Args:
        session: Authenticated requests session
        attestation_id: Archer attestation ID number
        cookies: Session cookies
        max_retries: Maximum number of retry attempts
        wait_time: Seconds to wait between retries

    Returns:
        Optional[str]: HTML content if successful, None on failure

    Raises:
        requests.RequestException: For network or HTTP errors
    """
    url = ATTESTATION_URL_TEMPLATE.format(attestation_id)
    headers = get_attestation_headers(cookies)
    
    for attempt in range(max_retries):
        try:
            logging.debug(f"Fetching attestation {attestation_id} (Attempt {attempt + 1})")
            response = session.get(url, headers=headers, timeout=60, allow_redirects=False)
            
            log_request_response(f"Fetch Attestation (Attempt {attempt + 1})", response.request, response)
            
            if response.status_code == 302:
                redirect_url = response.headers.get('Location')
                logging.debug(f"Following redirect to: {redirect_url}")
                response = session.get(urljoin(BASE_URL, redirect_url), headers=headers, timeout=60)
                log_request_response("Fetch Attestation (After Redirect)", response.request, response)
            
            response.raise_for_status()
            
            logging.info(f"Successfully fetched HTML for attestation {attestation_id}")
            return response.text
            
        except requests.RequestException as e:
            logging.warning(f"Attempt {attempt + 1} failed for attestation {attestation_id}: {str(e)}")
            if attempt == max_retries - 1:
                logging.error(f"Failed to fetch attestation {attestation_id} after all attempts")
                return None
            time.sleep(wait_time)
    
    return None

def log_request_response(step_name: str, request: requests.PreparedRequest, response: requests.Response) -> None:
    """Log details of HTTP request/response for debugging.

    Args:
        step_name: Name of the current processing step
        request: The HTTP request object
        response: The HTTP response object
    """
    logging.debug(f"--- {step_name} ---")
    logging.debug(f"Request URL: {request.url}")
    logging.debug(f"Request Method: {request.method}")
    logging.debug(f"Request Headers: {request.headers}")
    logging.debug(f"Request Body: {request.body}")
    
    logging.debug(f"Response Status Code: {response.status_code}")
    logging.debug(f"Response Headers: {response.headers}")
    logging.debug(f"Response Cookies: {response.cookies.get_dict()}")
    logging.debug(f"Response Content (first 500 chars): {response.text[:500]}")
    logging.debug("--- End of Request/Response ---\n")

def save_html(html_content: str, output_path: str) -> None:
    """Save HTML content (couldnt find a conversion library that works with 32 bit python; would just stream direct to weasyprint or pdfkit or smth)

    Args:
        html_content: HTML string to save
        output_path: Path to save HTML file

    Raises:
        IOError: If file cannot be written
    """
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        logging.info(f"HTML saved: {output_path}")
    except Exception as e:
        logging.error(f"Error saving HTML: {str(e)}")
        raise

# TODO move powershell to scripts folder
def convert_html_to_pdf(html_dir: str) -> None:
    """Convert HTML files to PDF using Powershell and Microsoft Word. (work around)

    Args:
        html_dir: Directory containing HTML files

    Raises:
        subprocess.CalledProcessError: If PowerShell conversion fails
    """
    powershell_script = f"""
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    Get-ChildItem -Path "{html_dir}" -Filter *.html | ForEach-Object {{
        $doc = $word.Documents.Open($_.FullName)
        $pdf_path = $_.FullName -replace '\.html$','.pdf'
        $doc.SaveAs([ref] $pdf_path, [ref] 17)
        $doc.Close()
        Write-Host "Converted $($_.Name) to PDF"
    }}
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
    """
    
    try:
        subprocess.run(["powershell", "-Command", powershell_script], check=True)
        logging.info("PDF conversion completed successfully")
    except subprocess.CalledProcessError as e:
        logging.error(f"Error during PDF conversion: {e}")
        raise

def fetch_attestations(attestation_ids: List[int], output_dir: str) -> None:
    """Main function to fetch attestations, save as HTML and convert to PDF.

    Args:
        attestation_ids: List of attestation IDs to fetch
        output_dir: Directory to save output files

    Raises:
        OSError: If output directory cannot be created
    """
    os.makedirs(output_dir, exist_ok=True)
    
    results = batch_fetch_attestations(attestation_ids, output_dir)
    
    successful_downloads = sum(results)
    logging.info(f"Successfully downloaded {successful_downloads} out of {len(attestation_ids)} attestations.")

    convert_html_to_pdf(output_dir)

def batch_fetch_attestations(attestation_ids: List[int], output_dir: str, max_workers: int = 5) -> List[bool]:
    """Download multiple attestations concurrently.

    Args:
        attestation_ids: List of attestation IDs to fetch
        output_dir: Directory to save output files
        max_workers: Maximum number of concurrent downloads

    Returns:
        List[bool]: Success/failure status for each attestation

    Raises:
        RuntimeError: If initial authentication fails
    """
    session = create_session()
    cookies = fetch_initial_cookies(session)
    
    if not cookies:
        logging.error("Failed to fetch initial cookies. Exiting.")
        return

    def download_single_attestation(attestation_id: int) -> bool:
        html_content = fetch_attestation_html(session, attestation_id, cookies)
        if html_content:
            html_path = os.path.join(output_dir, f"{attestation_id}.html")
            save_html(html_content, html_path)
            return True
        else:
            logging.warning(f"Failed to process attestation {attestation_id}")
            return False

    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor: # some leftover threading from my attempt to speed this up. kept breaking when I would mess with it, so here it stays
        futures = [executor.submit(download_single_attestation, att_id) for att_id in attestation_ids]
        concurrent.futures.wait(futures)
    
    return [future.result() for future in futures]

if __name__ == "__main__":
    attestation_ids = [316070, 392957]  # List of attestations, for KAIZEN script, this is passed in differently
    output_dir = "attestation_html"
    fetch_attestations(attestation_ids, output_dir)
