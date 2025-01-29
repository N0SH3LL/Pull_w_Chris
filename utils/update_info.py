"""
update_info.py

Handles updating document information for BPERs, Attestations, 
and Supporting Documents to the progress.json file.
"""

import os
import json
import re
import fitz
from src.SCC import scc_check
from src.SCC import scc_read
from src.utils import file_operations
from datetime import datetime
from difflib import SequenceMatcher

def update_bper_info(bper_dict, base_directories):
    """
    Grab info for BPERs and write to dict
    
    Args:
        bper_dict (dict): Dictionary containing BPER information
        base_directories (dict): Dictionary containing base directory paths for different document types
        
    Returns:
        None: Updates the bper_dict in place
    """
    for key, value_list in bper_dict.items(): # goes through every item in every bper entry
        for value in value_list:
            if value.get('false_positive', False):
                print(f"Skipping BPER '{key}' marked as false positive.") # skip if marked as false pos
                continue

            if 'manually_linked' in value:
                file_path = value['manually_linked'] # use manually_linked path if assigned
            else:
                source_directory = base_directories['bper']
                file_path = os.path.join(source_directory, f"{key}.pdf")

            if os.path.isfile(file_path):
                valid_to_date, approval_status, tla_present = file_operations.extract_BPER_info(file_path) # TODO: FIX Counterintuitively, the actual pulling of information comes from the file_operations file; just where it started, hasn't been fixed yet. 
                value['Valid to'] = valid_to_date # write these values to the dictionaries
                value['Approval Status'] = approval_status
                value['TLA'] = tla_present
                value['Updated from filename'] = os.path.basename(file_path) # stores file info is from and date it grabbed it
                value['Updated from timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                print(f"Updated BPER '{key}' - 'Valid to': {valid_to_date}, 'Approval Status': {approval_status}, 'TLA': {tla_present}") # print out everything it updated
            else:
                print(f"File not found for BPER: {key}") # if BPER with that name isn't present

def convert_datetime_to_string(obj):
    """
    File starts throwing errors unless the datetime objects it gets are converted to strings.    
   
     Args:
        obj: Object potentially containing datetime values (dict, list, datetime, or other)
        
    Returns:
        Object with all datetime values converted to strings
    """
    if isinstance(obj, dict):
        return {key: convert_datetime_to_string(value) for key, value in obj.items()}
    elif isinstance(obj, list):
        return [convert_datetime_to_string(item) for item in obj]
    elif isinstance(obj, datetime):
        return obj.isoformat()
    return obj

def update_attestation_info(attestation_dict, base_directories):
    """
    Updates attestation information 
    
    Args:
        attestation_dict (dict): Dictionary containing attestation information
        base_directories (dict): Dictionary containing base directory paths for different document types
        
    Returns:
        dict: Updated attestation dictionary
    """
    for key, value_list in attestation_dict.items():
        for value in value_list:
            if value.get('false_positive', False):
                print(f"Skipping Attestation '{key}' marked as false positive.")
                continue

            if 'manually_linked' in value:
                file_path = value['manually_linked'] # manually linked helps if it doesn't automatically find the doc
            else:
                source_directory = base_directories['attestation']
                file_path = os.path.join(source_directory, f"{key}.pdf")

            if os.path.isfile(file_path):
                try:
                    # Extract text content from PDF
                    with fitz.open(file_path) as doc:
                        text = ""
                        for page in doc:
                            text += page.get_text()
                    
                    # Process extracted text for attestation information
                    approval_status, valid_to_date, review_date, assessment_date, overall_status = file_operations.extract_attest_info(text)
                    
                    if approval_status != "Status: Error":
                        # Update attestation data
                        value.update({
                            'Approval Status': approval_status,
                            'Valid to': valid_to_date,
                            'Review Date': review_date,
                            'Assessment Date': assessment_date,
                            'Overall Status': overall_status,
                            'Updated from filename': os.path.basename(file_path),
                            'Updated from timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        })
                        print(f"Updated '{key}' with status: {approval_status}, valid to date: {valid_to_date}, "
                              f"review date: {review_date}, assessment date: {assessment_date}, "
                              f"overall status: {overall_status}")
                    else:
                        print(f"Error extracting information for Attestation: {key}")
                except Exception as e:
                    print(f"Error processing PDF file for Attestation {key}: {str(e)}")
            else:
                print(f"File not found for Attestation: {key}")

    return attestation_dict

def update_doc_info(doc_dict, base_directories):
    """
    Updates supporting document information in the master dict.
    
    Args:
        doc_dict (dict): Dictionary containing document information
        base_directories (dict): Dictionary containing base directory paths for different document types
        
    Returns:
        None: Updates the doc_dict in place
    """
    for doc_name, value_list in doc_dict.items(): # goes through every item in every document entry
        for value in value_list:
            if value.get('false_positive', False):
                print(f"Skipping Document '{value['Doc name']}' marked as false positive.") # skip if marked as false pos
                continue

            if 'manually_linked' in value:
                file_path = value['manually_linked'] # use manually_linked path if assigned
            else:
                # Find best matching document in directory
                source_directory = base_directories['doc']
                matching_files = [f for f in os.listdir(source_directory) if (f.endswith('.docx') or f.endswith('.doc') or f.endswith('.xlsx') or f.endswith('.xls') or f.endswith('.pdf'))] # has to handle additional file types
                if matching_files:
                    # Use sequence matcher to find closest filename match
                    best_match = max(matching_files, key=lambda x: SequenceMatcher(None, doc_name, x).ratio())
                    match_ratio = SequenceMatcher(None, doc_name, best_match).ratio()
                    if match_ratio >= 0.8:
                        file_path = os.path.join(source_directory, best_match) # matching for document names
                    else:
                        print(f"No close match found for Document: {doc_name}") # no matches better than the ratio
                        continue
                else:
                    print(f"No matching file found for Document: {doc_name}") # no matches at all
                    continue

            if os.path.isfile(file_path):
                most_recent_date = file_operations.extract_Doc_info(file_path) # TODO:fix Counterintuitively, the actual pulling of information comes from the file_operations file; just where it started, hasn't been fixed yet.
                if most_recent_date:
                    for entry in value_list: # write these values to dictionaries
                        entry['Last update'] = most_recent_date
                        entry['Version'] = extract_version(os.path.basename(file_path))
                        entry['Updated from filename'] = os.path.basename(file_path)
                        entry['Updated from timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    print(f"Updated '{doc_name}' with 'Last update': {most_recent_date}, 'Version': {extract_version(os.path.basename(file_path))}")
            else:
                print(f"File not found for Document: {doc_name}")

def extract_version(filename):
    """
    grabs last two digits after underscore in name.
    
    Args:
        filename (str): Name of the file to extract version from
        
    Returns:
        str: Version number if found, empty string otherwise
    """
    version_match = re.search(r'_(\d{2})(?=\.docx$|\.doc$)', filename)
    return version_match.group(1) if version_match else ''

def update_scc_info(scc_dict, scc_dir, progress_data):
    """
    Updates SCC information in the master dict.
    
    Args:
        scc_dict (dict): Dictionary containing SCC information
        scc_dir (str): Directory containing SCC files
        progress_data (dict): Full progress tracking dictionary
        
    Returns:
        dict: Updated SCC dictionary
    """
    print("Entering update_scc_info function")
    print(f"Number of SCCs to process: {len(scc_dict)}")

    for file_path, scc_info in scc_dict.items():
        print(f"Processing SCC: {scc_info['SCC']}")
        # Extract the SCC name from the file path
        scc_name = os.path.splitext(os.path.basename(file_path))[0]
        scc_name = re.sub(r'_\d+$', '', scc_name)  # Remove the version number from the SCC name

        # Search for files with a similar name pattern in the specified directory
        matching_files = [f for f in os.listdir(scc_dir) if f.startswith(scc_name) and f.endswith('.xlsx')]

        if matching_files:
            # Process most recent version of SCC file
            latest_file = max(matching_files, key=lambda x: os.path.getmtime(os.path.join(scc_dir, x)))
            latest_file_path = os.path.join(scc_dir, latest_file)

            # Process the SCC file using scc_read
            bper_dict, doc_dict, attestation_dict, method_dict = scc_read.process_excel_file(latest_file_path)

            # Update SCC info using scc_check
            updated_scc_info = scc_check.process_scc_file(latest_file_path)
            scc_info.update(updated_scc_info)

            # Update checks information
            for stig_id, details in method_dict.items():
                progress_data["Checks"][stig_id] = {
                    'SCC': scc_info['SCC'],
                    'Evidence method': details['Evidence Method']
                }

            # Collect evidence methods for this SCC
            scc_methods = set(details['Evidence Method'].lower() for details in method_dict.values())
            scc_info['Evidence Methods'] = list(scc_methods)

            # Update other dictionaries in progress_data
            for key, value in bper_dict.items():
                if key not in progress_data['BPERs']:
                    progress_data['BPERs'][key] = []
                progress_data['BPERs'][key].append(value)

            for key, value in doc_dict.items():
                if key not in progress_data['Documents']:
                    progress_data['Documents'][key] = []
                progress_data['Documents'][key].append(value)

            for key, value in attestation_dict.items():
                if key not in progress_data['Attestations']:
                    progress_data['Attestations'][key] = []
                progress_data['Attestations'][key].append(value)

            print(f"Updated SCC {scc_name} with {len(scc_methods)} evidence methods: {scc_info['Evidence Methods']}")
            print(f"Number of checks for this SCC: {len(method_dict)}")
        else:
            print(f"No matching SCC files found for: {scc_info['SCC']}")

    print("Exiting update_scc_info function")
    return scc_dict

def update_progress_info(progress_file, base_directories=None, scc_dir=None):
    """
    Writes all the dictionaries to progress.json.
    
    Args:
        progress_file (str): Path to progress.json file
        base_directories (dict, optional): Dictionary of base directory paths
        scc_dir (str, optional): Directory containing SCC files
        
    Returns:
        None: Updates progress.json file directly
    """
    print("Entering update_progress_info function")
    # Load current progress data
    with open(progress_file, 'r') as file:
        progress_data = json.load(file)
    
    # Extract component dictionaries
    scc_dict = progress_data.get('SCC', {})
    bper_dict = progress_data.get('BPERs', {})
    attestation_dict = progress_data.get('Attestations', {})
    doc_dict = progress_data.get('Documents', {})
    
    # Update document information if directories provided
    if base_directories:
        update_bper_info(bper_dict, base_directories)
        update_attestation_info(attestation_dict, base_directories)
        update_doc_info(doc_dict, base_directories)
    
    # Update SCC information if directory provided
    if scc_dir:
        updated_scc_dict = update_scc_info(scc_dict, scc_dir, progress_data)
        progress_data['SCC'] = updated_scc_dict  # Ensure we're saving the updated SCC dictionary
    
    # Update main progress data
    progress_data.update({
        'BPERs': bper_dict,
        'Attestations': attestation_dict,
        'Documents': doc_dict,
        'SCC': scc_dict
    })

    # Update timestamp
    program_settings = progress_data.get('Program Settings', {})
    program_settings['Pull Info Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    progress_data['Program Settings'] = program_settings
    
    print("Saving updated progress data")
    # Save updated progress data
    with open(progress_file, 'w') as file:
        json.dump(convert_datetime_to_string(progress_data), file, indent=4) # datetime objects > strings
    
    print("Progress information updated successfully.")
