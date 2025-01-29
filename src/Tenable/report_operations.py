from .api_client import TenableSCClient
from typing import Dict, Any, List, Optional, Tuple
import os
import re

def create_report(client: TenableSCClient, report_data: Dict[str, Any]) -> Dict[str, Any]: # Create a new report using the provided report data.
    return client.post('report', data=report_data)

def get_report_status(client: TenableSCClient, report_id: str) -> Dict[str, Any]: # Get the status of a report by its ID.
    return client.get(f'report/{report_id}')

def download_report(client: TenableSCClient, report_id: str) -> bytes: # Download report by ID
    response = client.post(f'report/{report_id}/download', data={}, raw_response=True)
    return response.content

def send_scan_to_report_template(client: TenableSCClient, scan_id: str, template_id: str) -> Dict[str, Any]: # Send scan output to a report template.

    report_data = {
        "definition": {
            "name": f"Report for Scan {scan_id}",
            "type": "pdf",  # format
            "template": {"id": template_id},
            "chapters": [
                {
                    "type": "scan",
                    "scanID": scan_id
                }
            ]
        }
    }
    return create_report(client, report_data)

def fetch_generated_report(client: TenableSCClient, report_id: str, max_retries: int = 10, delay: int = 30) -> bytes: # Fetch a generated report, waiting for it to complete if necessary.
    import time

    for _ in range(max_retries):
        report_status = get_report_status(client, report_id)
        if report_status['status'] == 'Completed':
            return download_report(client, report_id)
        elif report_status['status'] in ['Error', 'Cancelled']:
            raise Exception(f"Report generation failed with status: {report_status['status']}")
        time.sleep(delay)
    
    raise TimeoutError("Report generation timed out")

def list_reports(client: TenableSCClient) -> Dict[str, Any]: # list all reports
    params = {
        'fields': 'id,name,owner,status,startTime,finishTime'
    }
    return client.get('report', params=params)

def filter_reports(reports_data: Dict[str, Any], key: str, value: str) -> List[Dict[str, Any]]: # much easier to filter reports on owner, otherwise it tries to look at everyone & everything and gets overwhelmed
    filtered = []
    for report in reports_data['response'].get('usable', []):
        if key == 'owner':
            if report.get('owner', {}).get('username') == value:
                filtered.append(report)
        elif str(report.get(key)) == value:
            filtered.append(report)
    return filtered

def find_report_by_name(reports_data: Dict[str, Any], name: str) -> Optional[Dict[str, Any]]: # Had to create a search function to grab reports by name, because iirc it uses mainly report ids
    
    for report in reports_data['response'].get('usable', []):
        if report['name'].lower() == name.lower():
            return report
    return None

def download_report(client: TenableSCClient, report_id: str) -> Tuple[bytes, str]:
    
    response = client.post(f'report/{report_id}/download', data={}, raw_response=True)
    content_type = response.headers.get('Content-Type', '')
    
    if 'pdf' in content_type: # Not sure why this has to be specified.
        extension = '.pdf'
    elif 'rtf' in content_type:
        extension = '.rtf'
    elif 'csv' in content_type:
        extension = '.csv'
    elif 'asr' in content_type:
        extension = '.asr'
    elif 'arf' in content_type:
        extension = '.arf'
    elif 'lasr' in content_type:
        extension = '.lasr'
    else:
        extension = ''  # Default to no extension if unknown

    return response.content, extension

def download_report_by_name(client: TenableSCClient, report_name: str) -> Tuple[Optional[bytes], str]: # Download individual reports by their names.
    reports_data = list_reports(client)
    report = find_report_by_name(reports_data, report_name)
    if report:
        return download_report(client, report['id'])
    return None, ''

def download_reports_for_owner(client: TenableSCClient, owner_id: str, project_dir: str) -> None: # Download all the reports for a specific owner. 
    reports_data = list_reports(client)
    owner_reports = filter_reports(reports_data, 'owner', owner_id)
    
    for report in owner_reports:
        report_name = report['name']
        report_id = report['id']
        
        # [!] This is how it handles scan/report names. Currently hardcoded for very specific values. If it's not grabbing stuff correctly, look here to derive the expected format
        # Strip TDL-PDF(Scan:) or TDL-CSV(Scan:) from the name
        stripped_name = re.sub(r'^TDL-(PDF|CSV) \(Scan[_:] TDL-', '', report_name)
        
        # Remove only one closing parenthesis from the end
        stripped_name = stripped_name[:-1] if stripped_name.endswith(')') else stripped_name
        
        # Separate SCC name from the report type and number
        match = re.match(r'(.*?)(-(?:Info|PassFail).*)', stripped_name)
        if match:
            scc_name, report_suffix = match.groups()
            scc_name = scc_name.strip()
        else:
            print(f"Unexpected report name format: {stripped_name}")
            continue
        
        report_content, file_extension = download_report(client, report_id)
        
        if report_content:
            # Determine the appropriate directory
            if 'Info' in report_suffix:
                dir_path = os.path.join(project_dir, scc_name, "Manual", "Automated Info")
            elif 'PassFail' in report_suffix:
                dir_path = os.path.join(project_dir, scc_name, "Automated")
            else:
                dir_path = os.path.join(project_dir, scc_name)
            
            # Create directory if it doesn't exist
            os.makedirs(dir_path, exist_ok=True)
            
            # Construct the filename
            filename = f"{scc_name}{report_suffix}{file_extension}"
            
            # Save the file
            file_path = os.path.join(dir_path, filename)
            with open(file_path, 'wb') as f:
                f.write(report_content)
            
            print(f"Downloaded report: {file_path}")
        else:
            print(f"Failed to download report: {stripped_name}")
