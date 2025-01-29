"""
Main module for Tenable operations. May actually be more in here than is used in through the gui, I was playing with functionality on command line initially. 

This module includes functions for:
- Filtering and managing scan data
- Editing scan configurations 
- Launching and monitoring scans
- Managing reports and report templates
- Handling IP lists and inventory files

"""

import argparse
import json
from api_client import TenableSCClient
from src import scan_operations, report_operations
from scan_operations import chunk_and_create_scans
from typing import Dict, Any
import os

def filter_scans(scans_data, key, value):
    """
    Filter scan data based on a key-value pair.
    
    Args:
        scans_data: Raw scan data from Tenable SC (dict or JSON string)
        key: Filter key (owner)
        value: Filter value to match
    
    Returns:
        list: Filtered list of scan data matching the criteria
    """
    if isinstance(scans_data, str):
        try:
            scans_data = json.loads(scans_data)
        except json.JSONDecodeError:
            print("Error: Unable to parse scans data as JSON.")
            return []

    if not isinstance(scans_data, dict) or 'response' not in scans_data:
        print(f"Unexpected data structure. Received: {type(scans_data)}")
        return []

    scans = scans_data['response'].get('usable', [])
    filtered = []
    for scan in scans:
        if key.startswith('owner.'):
            owner_key = key.split('.')[1]
            if str(scan.get('owner', {}).get(owner_key)) == value:
                filtered.append(scan)
        elif str(scan.get(key)) == value:
            filtered.append(scan)
    return filtered

def read_ip_list_from_file(file_path: str) -> str:
    """
    Read and parse IP addresses from inventory file, converting space, comma or newline separated to comma-separated format.
    
    Args:
        file_path: Path to file containing IP addresses
        
    Returns:
        str: Comma-separated list of IP addresses
    """
    with open(file_path, 'r') as f:
        content = f.read().strip()
    return ','.join(filter(None, [ip.strip() for ip in content.replace('\r', '\n').replace('\n', ',').replace(' ', ',').split(',')]))

def edit_scan_by_name(client: TenableSCClient, scan_name: str, ip_list_file: str) -> None:
    """
    Edit an existing scan's IP list.
    
    Args:
        client: TenableSCClient instance
        scan_name: Name of scan to edit
        ip_list_file: Path to file containing IP addresses
    """
    scans_data = scan_operations.list_scans(client)
    scan = scan_operations.find_scan_by_name(scans_data, scan_name)
    if not scan:
        print(f"Error: Scan with name '{scan_name}' not found.")
        return

    ip_list = read_ip_list_from_file(ip_list_file)
    result = scan_operations.edit_scan_ip_list(client, scan['id'], ip_list)
    
    updated_info = {}
    if 'response' in result:
        scan_data = result['response']
        updated_info = {
            "id": scan_data.get("id"),
            "name": scan_data.get("name"),
            "description": scan_data.get("description"),
            "ipList": scan_data.get("ipList")
        }
    
    if all(value is not None for value in updated_info.values()):
        print(f"Scan '{scan_name}' updated with new IP list.")
        print("Updated scan information:")
        print(json.dumps(updated_info, indent=2))
    else:
        print(f"Scan '{scan_name}' was updated, but couldn't retrieve all updated information.")
        print("Available updated information:")
        print(json.dumps({k: v for k, v in updated_info.items() if v is not None}, indent=2))

def launch_scan_by_name(client: TenableSCClient, scan_name: str) -> None:
    """
    Launch a scan by its name and monitor initialization status.
    
    Args:
        client: TenableSCClient instance
        scan_name: Name of scan to launch
        
    Prints status updates and error messages if launch fails
    """
    try:
        result = scan_operations.launch_scan_by_name(client, scan_name)
        
        if 'response' in result:
            scan_result = result['response'].get('scanResult', {})
            print(f"Scan '{scan_name}' launched successfully.")
            print(f"Scan ID: {scan_result.get('id')}")
            print(f"Job ID: {scan_result.get('jobID')}")
            print(f"Status: {scan_result.get('status')}")
        else:
            print(f"Scan '{scan_name}' launch initiated, but couldn't retrieve launch details.")
        
    except ValueError as e:
        print(f"Error: {str(e)}")
    except Exception as e:
        print(f"Error launching scan '{scan_name}': {str(e)}")

    # Print the full response for debugging
    #print("\nFull API response:")
    #print(json.dumps(result, indent=2))

def copy_scan_by_name(client: TenableSCClient, scan_name: str, new_name: str) -> None:
    """
    Create a copy of an existing scan with a new name.
    
    Args:
        client: TenableSCClient instance
        scan_name: Name of scan to copy
        new_name: Name for the copied scan
        
    Prints status of copy operation and new scan details
    """
    scans_data = scan_operations.list_scans(client)
    scan = scan_operations.find_scan_by_name(scans_data, scan_name)
    if not scan:
        print(f"Error: Scan with name '{scan_name}' not found.")
        return
    
    result = scan_operations.copy_scan(client, scan['id'], new_name)
    if 'error' in result:
        print(f"Error copying scan '{scan_name}': {result['error']}")
        if 'url' in result:
            print(f"Attempted URL: {result['url']}")
    elif 'response' in result and 'scan' in result['response']:
        new_scan = result['response']['scan']
        print(f"Scan '{scan_name}' copied successfully.")
        print(f"New scan name: '{new_scan['name']}'")
        print(f"New scan ID: {new_scan['id']}")
        print(f"New scan UUID: {new_scan['uuid']}")
    else:
        print(f"Unexpected response when copying scan '{scan_name}'")
        print(f"Full response: {result}")

def edit_scan_details(client: TenableSCClient, scan_name: str, new_name: str = None, schedule: Dict[str, Any] = None) -> None:
    """
    Edit scan details - name and scheduling.
    
    Args:
        client: TenableSCClient instance
        scan_name: Name of scan to edit
        new_name: Optional new name for the scan
        schedule: Optional dictionary containing schedule parameters
        
    Prints status of update operation
    """
    scans_data = scan_operations.list_scans(client)
    scan = scan_operations.find_scan_by_name(scans_data, scan_name)
    if not scan:
        print(f"Error: Scan with name '{scan_name}' not found.")
        return
    
    result = scan_operations.edit_scan(client, scan['id'], name=new_name, schedule=schedule)
    if 'response' in result:
        print(f"Scan '{scan_name}' updated successfully.")
        if new_name:
            print(f"New name: {result['response'].get('name')}")
        if schedule:
            print(f"Schedule updated: {result['response'].get('schedule')}")
    else:
        print(f"Error updating scan '{scan_name}'")

def delete_scan_by_name(client: TenableSCClient, scan_name: str) -> None:
    """
    Delete an existing scan by name.
    
    Args:
        client: TenableSCClient instance
        scan_name: Name of scan to delete
    """
    scans_data = scan_operations.list_scans(client)
    scan = scan_operations.find_scan_by_name(scans_data, scan_name)
    if not scan:
        print(f"Error: Scan with name '{scan_name}' not found.")
        return
    
    result = scan_operations.delete_scan(client, scan['id'])
    if 'response' in result:
        print(f"Scan '{scan_name}' deleted successfully.")
    else:
        print(f"Error deleting scan '{scan_name}'")

def main():
    """
    entry point for using CLI. messy if/elif statements that handle each use case. TODO improve this (make a gui for the api? seems counterintuitive)
    """
    parser = argparse.ArgumentParser(description="Tenable SC Automation")
    parser.add_argument('action', choices=['list_scans', 'edit_scan', 'launch_scan', 'create_report', 'fetch_report', 'list_reports', 'download_report', 'copy_scan', 'edit_scan_details', 'delete_scan', 'chunk_and_scan'])
    parser.add_argument('--filter', nargs=2, metavar=('KEY', 'VALUE'), help='Filter scans or reports by KEY VALUE pair')
    parser.add_argument('--scan-id', help='Scan ID for launching, editing a scan, or report creation')
    parser.add_argument('--scan-name', help='Scan name for launching or editing a scan')
    parser.add_argument('--new-scan-name', help='New name for the scan when copying or editing')
    parser.add_argument('--schedule', help='JSON string for scan schedule when editing')
    parser.add_argument('--report-name', help='Name of the report to download')
    parser.add_argument('--output-dir', default='.', help='Directory to save the downloaded report')
    parser.add_argument('--template-id', help='Report template ID')
    parser.add_argument('--report-id', help='Report ID for fetching')
    parser.add_argument('--ip-list-file', help='File containing IP list for editing scan')
    parser.add_argument('--inventory-file', help='Path to the inventory file')
    parser.add_argument('--chunk-size', type=int, default=6, help='Number of devices per scan')
    parser.add_argument('--start-time', help='Start time for the first scan (format: YYYYMMDDTHHMMSS)')
    parser.add_argument('--user', help='Download all reports for a specific user')
    parser.add_argument('--all', action='store_true', help='Download all reports for the specified user')

    args = parser.parse_args()

    client = TenableSCClient()

    if args.action == 'list_scans':
        scans_data = scan_operations.list_scans(client)
        if args.filter:
            key, value = args.filter
            filtered_scans = filter_scans(scans_data, key, value)
            print(json.dumps(filtered_scans, indent=2))
        else:
            print(json.dumps(scans_data, indent=2))

    elif args.action == 'edit_scan':
        if args.scan_name and args.ip_list_file:
            edit_scan_by_name(client, args.scan_name, args.ip_list_file)
        else:
            print("Error: Scan name and IP list file are required for editing a scan's IP list.")
            print("Use --scan-name <scan_name> --ip-list-file <file_path>")

    elif args.action == 'launch_scan':
        if args.scan_id:
            result = scan_operations.launch_scan(client, args.scan_id)
            print(json.dumps(result, indent=2))
        elif args.scan_name:
            launch_scan_by_name(client, args.scan_name)
        else:
            print("Error: Either Scan ID or Scan Name is required for launching a scan.")
            print("Use --scan-id <scan_id> or --scan-name <scan_name>")

    elif args.action == 'edit_scan_ip_list':
        if not args.scan_name or not args.ip_list_file:
            print("Error: Scan name and IP list file are required for editing a scan's IP list.")
            return
        edit_scan_by_name(client, args.scan_name, args.ip_list_file)

    elif args.action == 'launch_scan':
        if args.scan_id:
            result = scan_operations.launch_scan(client, args.scan_id)
            print(json.dumps(result, indent=2))
        elif args.scan_name:
            launch_scan_by_name(client, args.scan_name)
        else:
            print("Error: Either Scan ID or Scan Name is required for launching a scan. Use --scan-id <scan_id> or --scan-name <scan_name>")

    elif args.action == 'create_report':
        if not args.scan_id or not args.template_id:
            print("Error: Scan ID and template ID are required for creating a report.")
            return
        result = report_operations.send_scan_to_report_template(client, args.scan_id, args.template_id)
        print(json.dumps(result, indent=2))

    elif args.action == 'fetch_report':
        if not args.report_id:
            print("Error: Report ID is required for fetching a report.")
            return
        try:
            report_content = report_operations.fetch_generated_report(client, args.report_id)
            with open(f"report_{args.report_id}.pdf", "wb") as f:
                f.write(report_content)
            print(f"Report downloaded as report_{args.report_id}.pdf")
        except Exception as e:
            print(f"Error fetching report: {str(e)}")

    elif args.action == 'list_reports':
        reports_data = report_operations.list_reports(client)
        
        if args.filter:
            key, value = args.filter
            filtered_reports = report_operations.filter_reports(reports_data, key, value)
            for report in filtered_reports:
                print(f"Report ID: {report['id']}, Name: {report['name']}, "
                      f"Owner: {report['owner']['username']}, Status: {report['status']}, "
                      f"Start Time: {report.get('startTime', 'N/A')}, Finish Time: {report.get('finishTime', 'N/A')}")
            print(f"Total filtered reports: {len(filtered_reports)}")
        else:
            if 'response' in reports_data and 'usable' in reports_data['response']:
                for report in reports_data['response']['usable']:
                    print(f"Report ID: {report['id']}, Name: {report['name']}, "
                          f"Owner: {report['owner']['username']}, Status: {report['status']}, "
                          f"Start Time: {report.get('startTime', 'N/A')}, Finish Time: {report.get('finishTime', 'N/A')}")
                print(f"Total reports: {len(reports_data['response']['usable'])}")
            else:
                print("No reports found or error in response.")
                print(json.dumps(reports_data, indent=2))

    elif args.action == 'download_report':
        if args.all and args.user:
            reports_data = report_operations.list_reports(client)
            if 'response' in reports_data and 'usable' in reports_data['response']:
                user_reports = [report for report in reports_data['response']['usable'] 
                                if report['owner']['username'].lower() == args.user.lower()]
                if not user_reports:
                    print(f"No reports found for user: {args.user}")
                    return
                for report in user_reports:
                    report_name = report['name']
                    report_id = report['id']
                    report_content, file_extension = report_operations.download_report(client, report_id)
                    if report_content:
                        filename = os.path.join(args.output_dir, f"{report_name}{file_extension}")
                        with open(filename, 'wb') as f:
                            f.write(report_content)
                        print(f"Report downloaded successfully: {filename}")
                    else:
                        print(f"Failed to download report: {report_name}")
                print(f"Downloaded {len(user_reports)} reports for user: {args.user}")
            else:
                print("No reports found or error in response.")
        elif not args.report_name:
            print("Error: Report name is required for downloading a report, or use --all --user <username> to download all reports for a specific user.")
            return
        else:
            report_content, file_extension = report_operations.download_report_by_name(client, args.report_name)
            if report_content:
                report_name = args.report_name
                filename = os.path.join(args.output_dir, f"{report_name}{file_extension}")
                with open(filename, 'wb') as f:
                    f.write(report_content)
                print(f"Report downloaded successfully: {filename}")
            else:
                print(f"Report '{args.report_name}' not found.")

    elif args.action == 'copy_scan':
        if not args.scan_name or not args.new_scan_name:
            print("Error: Scan name and new scan name are required for copying a scan.")
            return
        copy_scan_by_name(client, args.scan_name, args.new_scan_name)

    elif args.action == 'edit_scan_details':
        if not args.scan_name:
            print("Error: Scan name is required for editing scan details.")
            return
        schedule = json.loads(args.schedule) if args.schedule else None
        edit_scan_details(client, args.scan_name, new_name=args.new_scan_name, schedule=schedule)

    elif args.action == 'chunk_and_scan':
        if not args.scan_name or not args.inventory_file or not args.start_time:
            print("Error: Base scan name, inventory file, and start time are required.")
            return
        chunk_and_create_scans(client, args.scan_name, args.inventory_file, args.start_time, args.chunk_size)

    elif args.action == 'delete_scan':
        if not args.scan_name:
            print("Error: Scan name is required for deleting a scan.")
            return
        delete_scan_by_name(client, args.scan_name)


if __name__ == "__main__":
    main()
