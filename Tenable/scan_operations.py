"""
Tenable scan operations module.

This module provides functions for managing and executing scans in Tenable SecurityCenter,
including creating, copying, editing, launching and deleting scans. It reads inventory files for each scan, and chunks into smaller batches for better performance and reliability.

Typical usage example:
    client = TenableSCClient()
    scans = list_scans(client)
    scan_id = find_scan_by_name(scans, "My Scan")
    result = launch_scan_by_name(client, "My Scan")

"""

from typing import Dict, Any, Optional, List
from .api_client import TenableSCClient
import re
import requests
import math
from datetime import datetime, timedelta

def edit_scan_ip_list(client: TenableSCClient, scan_id: str, ip_list: str) -> Dict[str, Any]:
    """Updates the IP target list for a specified scan.

    Args:
        client: Authenticated TenableSCClient instance
        scan_id: ID of the scan to update
        ip_list: Comma-separated list of IP addresses/ranges

    Returns:
        Dict containing API response with updated scan details
        Example: {'response': {'id': '123', 'name': 'scan1', 'ipList': '10.0.0.1,10.0.0.2'}}

    Raises:
        requests.exceptions.RequestException: If API request fails
    """
    data = {
        'ipList': ip_list
    }
    return client.patch(f'scan/{scan_id}', data=data)
def list_scans(client: TenableSCClient) -> Dict[str, Any]:
    """Retrieves list of all accessible scans.

    Args:
        client: Authenticated TenableSCClient instance

    Returns:
        Dict containing scan list in the format:
        {
            'response': {
                'usable': [
                    {
                        'id': str,
                        'name': str,
                        'status': str,
                        'owner': Dict,
                        'createdTime': str,
                        'schedule': Dict
                    },
                    ...
                ]
            }
        }
    """
    params = {
        'fields': 'id,name,status,owner,createdTime,schedule'
    }
    return client.get('scan', params=params)
def find_scan_by_name(scans_data: Dict[str, Any], name: str) -> Optional[Dict[str, Any]]:
    """Finds a scan by its name (case-insensitive).

    Args:
        scans_data: Dict containing scan list from list_scans()
        name: Name of scan to find

    Returns:
        Dict containing scan details if found, None otherwise
    """
    for scan in scans_data['response'].get('usable', []):
        if scan['name'].lower() == name.lower():
            return scan
    return None
def launch_scan(client: TenableSCClient, scan_id: str) -> Dict[str, Any]:
    """Launches a scan from the scan id value.

    Args:
        client: Authenticated TenableSCClient instance
        scan_id: ID of scan to launch

    Returns:
        Dict containing launch status and details
        Example: {'response': {'scanResult': {'id': '123', 'status': 'Running'}}}
    """
    return client.post(f'scan/{scan_id}/launch', data={})
def launch_scan_by_name(client: TenableSCClient, scan_name: str) -> Dict[str, Any]:
    """Launches a scan by its name.

    Args:
        client: Authenticated TenableSCClient instance
        scan_name: Name of scan to launch

    Returns:
        Dict containing launch status and details

    Raises:
        ValueError: If scan with specified name is not found
    """
    scans_data = list_scans(client)
    scan = find_scan_by_name(scans_data, scan_name)
    if scan:
        return launch_scan(client, scan['id'])
    raise ValueError(f"Scan with name '{scan_name}' not found.")
def copy_scan(client: TenableSCClient, scan_id: str, new_name: str) -> Dict[str, Any]:
    """Creates a copy of an existing scan with a new name.

    Args:
        client: Authenticated TenableSCClient instance
        scan_id: ID of scan to copy
        new_name: Name for the new scan copy

    Returns:
        Dict containing new scan details or error information
        Success format: {'response': {'scan': {'id': str, 'name': str, 'uuid': str}}}
        Error format: {'error': str}

    Notes:
        - Uses a hardcoded target user ID (156) - may need updating for different users
        - New scan inherits most settings from original but can be modified post-copy
    """
    base_url = client.base_url.rstrip('/')
    url = f"{base_url}/scan/{scan_id}/copy"
    
    data = {
        "name": new_name,
        "targetUser": {
            "id": "156" # My id number for tenable, =/= to user id at WPS
        }
    }
    
    try:
        return client.post(f'scan/{scan_id}/copy', data=data)
    except requests.exceptions.RequestException as e:
        return {"error": f"Request Error occurred: {str(e)}"}
def edit_scan(client: TenableSCClient, scan_id: str, schedule: Dict[str, Any] = None) -> Dict[str, Any]:
    """Updates scan settings, particularly scheduling.

    Args:
        client: Authenticated TenableSCClient instance
        scan_id: ID of scan to edit
        schedule: Optional dict containing schedule settings:
            For one-time schedule: {'type': 'ical', 'start': 'TZID=America/New_York:YYYYMMDDTHHmmss'}
            For dependent schedule: {'type': 'dependent', 'dependentID': 'scan_id'}

    Returns:
        Dict containing updated scan details or error information
        Success format: {'response': {'id': str, 'name': str, 'schedule': Dict}}
        Error format: {'error': str}
    """
    url = f"{client.base_url.rstrip('/')}/scan/{scan_id}"
    
    data = {}
    if schedule:
        data['schedule'] = schedule

    try:
        response = requests.patch(url, headers=client.headers, json=data, verify=client.ca_cert_path)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        return {"error": f"Request Error occurred: {str(e)}. URL: {url}"}
def delete_scan(client: TenableSCClient, scan_id: str) -> Dict[str, Any]:
    """Deletes a scan.

    Args:
        client: Authenticated TenableSCClient instance
        scan_id: ID of scan to delete

    Returns:
        Dict containing deletion confirmation
    """
    return client.delete(f'scan/{scan_id}')
def chunk_and_create_scans(client: TenableSCClient, base_scan_name: str, inventory_file: str, start_time: str, chunk_size: int = 6):
    """Creates multiple dependent scans by copying a base scan, each targeting a subset of devices.
    Splits device list into chunks, will prompt for chunk size. First scan is scheduled for the specified start time,
    and subsequent scans are configured to start after their previous finishes.

    Args:
        client: Authenticated TenableSCClient instance
        base_scan_name: Name of the template scan to copy
        inventory_file: Path to file containing target IP addresses (one per line) - comes from main_gui.py, which finds the inventory files in each scc directory according to '{scc_name}-Inventory.txt'
        start_time: Start time for first scan in format: YYYYMMDDTHHmmss (ex. 20250121T163000)
        chunk_size: Number of devices per scan chunk (default: 6)

    Notes:
        - Each scan is named: "{base_scan_name} (XofY)" where X is chunk number, Y is total chunks
        - File should contain one IP address or range per line
        - Invalid or empty lines in inventory file are ignored
        - If any step fails for a chunk, that chunk is skipped but processing continues
        - All scans copy settings from base_scan_name template. These base scans need to be set up in Tenable prior to execution. 
    """
    # Read the inventory file
    with open(inventory_file, 'r') as f:
        devices = [line.strip() for line in f if line.strip()]
    
    # Calculate the number of scans needed
    num_scans = math.ceil(len(devices) / chunk_size)
    
    # Get the original scan
    scans_data = list_scans(client)
    original_scan = find_scan_by_name(scans_data, base_scan_name)
    if not original_scan:
        print(f"Error: Base scan '{base_scan_name}' not found.")
        return

    previous_scan_id = None

    for i in range(num_scans):
        # Calculate the chunk of devices for this scan
        start = i * chunk_size
        end = min((i + 1) * chunk_size, len(devices))
        chunk = devices[start:end]
        
        # Create the new scan name
        new_scan_name = f"{base_scan_name} ({i+1}of{num_scans})"
        
        # Copy the scan
        copy_result = copy_scan(client, original_scan['id'], new_scan_name)
        if 'error' in copy_result:
            print(f"Error copying scan: {copy_result['error']}")
            continue
        
        new_scan_id = copy_result['response']['scan']['id']
        
        # Edit the IP list
        ip_list = ','.join(chunk)
        edit_result = edit_scan_ip_list(client, new_scan_id, ip_list)
        if 'error' in edit_result:
            print(f"Error editing scan IP list: {edit_result['error']}")
            continue
        
        # Set up the schedule
        if i == 0:
            # First scan is scheduled at the specified time
            schedule = {
                "type": "ical",
                "start": f"TZID=America/New_York:{start_time}",
            }
        else:
            # Subsequent scans depend on the previous scan
            schedule = {
                "type": "dependent",
                "dependentID": previous_scan_id
            }
        
        # Edit the scan schedule
        schedule_result = edit_scan(client, new_scan_id, schedule=schedule)
        if 'error' in schedule_result:
            print(f"Error setting scan schedule: {schedule_result['error']}")
            continue
        
        if i == 0:
            print(f"Scheduled first scan: {new_scan_name} for {start_time} America/New_York")
        else:
            print(f"Set scan {new_scan_name} to be dependent on scan with ID: {previous_scan_id}")
        
        print(f"Created scan: {new_scan_name}")
        print(f"Devices: {ip_list}")
        print("---")

        # Update the previous_scan_id for the next iteration
        previous_scan_id = new_scan_id

    print(f"Completed creating and setting dependencies for {num_scans} scans.")
