import os
import shutil
import re
import docx2txt
import argparse
import json
import fitz
import logging
from typing import Tuple, List
from difflib import SequenceMatcher
from datetime import datetime

def update_dictionaries_and_copy_files(bper_dict, doc_dict, attestation_dict, base_directories, master_directory): # goes through lists, and sets up for copy_and_update
    for key, value_list in bper_dict.items(): # BPER list
        value = value_list[0]  
        if value.get('false_positive', False):
            print(f"Skipping BPER '{key}' marked as false positive.") # skip false pos
            continue

        if 'manually_linked' in value:
            source_file_path = value['manually_linked'] # use the path stored in manually_linked if present
        else:
            source_directory = base_directories['bper']
            source_file_path = os.path.join(source_directory, f"{key}.pdf") # otherwise creates the path, expects only pdf, BPER names should match exactly

        if os.path.isfile(source_file_path):
            bper_dict = copy_and_update(value, source_file_path, master_directory, bper_dict, use_margin_for_error=False) # if file is present, copy it over
        else:
            value['Gathered'] = False
            print(f"File not found for BPER: {key}") # not found, print outcome

    for key, value_list in doc_dict.items(): # Supporting Docs list
        value = value_list[0]  
        if value.get('false_positive', False):
            print(f"Skipping Document '{value['Doc name']}' marked as false positive.") # skip false pos
            continue

        if 'manually_linked' in value:
            source_file_path = value['manually_linked'] # use the path stored in manually_linked if present
        else:
            source_directory = base_directories['doc']
            doc_name = value['Doc name']
            # Needs to match because of doc names are all over the place
            matching_files = [f for f in os.listdir(source_directory) if (f.lower().endswith('.docx') or f.lower().endswith('.doc') or f.lower().endswith('.xlsx') or f.lower().endswith('.xls') or f.lower().endswith('.pdf'))]
            if matching_files:
                best_match = max(matching_files, key=lambda x: SequenceMatcher(None, doc_name.lower(), x.lower()).ratio())
                match_ratio = SequenceMatcher(None, doc_name.lower(), best_match.lower()).ratio()
                if match_ratio >= 0.8:  # threshold feeds into SequenceMatcher, basically closeness of match
                    source_file_path = os.path.join(source_directory, best_match)
                else:
                    value['Gathered'] = False
                    print(f"No close match found for Document: {doc_name}") # not found, ratio is below the match_ratio
                    continue
            else:
                value['Gathered'] = False
                print(f"No matching file found for Document: {doc_name}") # not found at all
                continue

        if os.path.isfile(source_file_path): # when appropriate match is found, copy it over, update the dictionary
            doc_dict = copy_and_update(value, source_file_path, master_directory, doc_dict, use_margin_for_error=True)
        else:
            value['Gathered'] = False
            print(f"File not found for Document: {value['Doc name']}") # not found at all

    for key, value_list in attestation_dict.items(): # Attestation list
        value = value_list[0]  
        if value.get('false_positive', False):
            print(f"Skipping Attestation '{key}' marked as false positive.") # skip false pos
            continue

        if 'manually_linked' in value:
            source_file_path = value['manually_linked'] # use manually_linked if present
        else:
            source_directory = base_directories['attestation']
            source_file_path = os.path.join(source_directory, f"{key}.pdf") # otherwise, use the path, only expects pdf

        if os.path.isfile(source_file_path):
            attestation_dict = copy_and_update(value, source_file_path, master_directory, attestation_dict, use_margin_for_error=False) # if present, copy over and update dict
        else:
            value['Gathered'] = False
            print(f"File not found for Attestation: {key}") # not found, print outcome

    return bper_dict, doc_dict, attestation_dict # output dicts for writing to progress.json

def copy_and_update(entry, source_file_path, master_directory, doc_dict, use_margin_for_error=False):
    scc_name = entry['SCC']
    source_file_name = os.path.basename(source_file_path)
    item_name = entry.get('Doc name') or entry.get('BPER name') or entry.get('Attestation num')

    dest_subdir = 'Attestations' if entry.get('Attestation num') else 'Exceptions and Deviations' if entry.get('BPER name') else 'Supporting Documents'

    for value in doc_dict.get(item_name, []):
        dest_directory = os.path.join(master_directory, value['SCC'], dest_subdir)

        if not os.path.exists(dest_directory):
            os.makedirs(dest_directory)

        dest_file_path = os.path.join(dest_directory, source_file_name)

        try:
            shutil.copy2(source_file_path, dest_file_path) # copy file
            print(f"Copied {source_file_name} to {dest_file_path}")

            value['Gathered'] = True
            value['Gathered file'] = source_file_name
            value['Gathered timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S') # update dictionary values
        except Exception as e:
            print(f"Error copying {source_file_name}: {e}") # error handling
            value['Gathered'] = False

    return doc_dict

import re

def extract_attest_info(text):
    try:
        # Preprocess the text
        text = re.sub(r'\s+', ' ', text).strip()

        # Extract the approval status
        approval_status_match = re.search(r"reviewer status:?\s*(\w+(?:\s+\w+)*)", text, re.IGNORECASE)
        approval_status = approval_status_match.group(1) if approval_status_match else "Status: Not Found"

        # Extract the valid to date (estimated close date)
        valid_to_match = re.search(r"estimated close\s+date:?\s*(\d{1,2}/\d{1,2}/\d{4})", text, re.IGNORECASE)
        valid_to_date = valid_to_match.group(1) if valid_to_match else "N/A"

        # Extract the review date
        review_date_match = re.search(r"review date:?\s*(\d{1,2}/\d{1,2}/\d{4})", text, re.IGNORECASE)
        review_date = review_date_match.group(1) if review_date_match else "N/A"

        # Extract the assessment date
        assessment_date_match = re.search(r"assessment date:?\s*(\d{1,2}/\d{1,2}/\d{4})", text, re.IGNORECASE)
        assessment_date = assessment_date_match.group(1) if assessment_date_match else "N/A"

        # Extract the overall status
        overall_status_match = re.search(r"overall status:?\s*(\w+)", text, re.IGNORECASE)
        overall_status = overall_status_match.group(1) if overall_status_match else "N/A"

        print(f"Extracted approval status: {approval_status}")
        print(f"Extracted valid to date: {valid_to_date}")
        print(f"Extracted review date: {review_date}")
        print(f"Extracted assessment date: {assessment_date}")
        print(f"Extracted overall status: {overall_status}")

        return approval_status, valid_to_date, review_date, assessment_date, overall_status
    except Exception as e:
        print(f"Error extracting information: {e}")
        return "Status: Error", "N/A", "N/A", "N/A", "N/A"
        return "Status: Error", "N/A", "N/A", "N/A", "N/A"

def extract_BPER_info(pdf_path): #grabbing info from BPERs
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc:
            text += page.get_text()

        # Extract date
        date_match = re.search(r"Valid To:\s*(\d{4}-\d{2}-\d{2}) \d{2}:\d{2}:\d{2}", text) 
        valid_to_date = date_match.group(1) if date_match else "N/A"

        # Extract text between "State:" and "CMS:"
        status_match = re.search(r"State:\s*(\S+)", text)
        approval_status = status_match.group(1).strip() if status_match else "Status: Not Found"

        # Check for "Technical Limitation"
        tla_match = re.search(r"Technical Limitation[^.]*", text)
        tla_present = bool(tla_match)

        return valid_to_date, approval_status, tla_present
    except Exception as e:
        print(f"Error processing {pdf_path}: {e}"'\n') # error message
        return "N/A", "Status: Error", False
    finally:
        doc.close() # close doc 

def extract_Doc_info(filepath): # grab info from Supporting Docs
    try:
        text = docx2txt.process(filepath)
        last_50_words = " ".join(text.split()[-50:]) #searches only the last 50 words (faster)

        date_pattern = r"\b\d{1,2}/\d{1,2}/\d{4}\b" # date pattern, grabs all dates in last 50 words
        dates = re.findall(date_pattern, last_50_words)

        date_objects = []
        for date in dates: # checks to make sure dates are in reasonable time frame
            try:
                date_object = datetime.strptime(date, "%m/%d/%Y")
                date_objects.append(date_object)
            except ValueError:
                pass

        valid_dates = [date for date in date_objects if 2000 <= date.year <= 2050]

        if valid_dates:
            return max(valid_dates).strftime("%Y-%m-%d") # returns the most recent date (assumes most recent would = last updated)
        return None
    except Exception as e:
        print(f"Error processing file {filepath}: {e}") # error handling
        return None
    
def load_dictionary(file_path): # vestigial from scripting
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return {}
    with open(file_path, 'r') as file:
        return json.load(file)   
    
def main(): # for use from command line
        parser = argparse.ArgumentParser(description='Update dictionaries and copy files.')
        parser.add_argument('bper_dict_file', type=str, help='Path to the BPER dictionary JSON file')
        parser.add_argument('doc_dict_file', type=str, help='Path to the Doc dictionary JSON file')
        parser.add_argument('attestation_dict_file', type=str, help='Path to the Attestation dictionary JSON file')
        parser.add_argument('master_directory', type=str, help='Where subdirectories will be created')
        args = parser.parse_args()

        bper_dict = load_dictionary(args.bper_dict_file)
        doc_dict = load_dictionary(args.doc_dict_file)
        attestation_dict = load_dictionary(args.attestation_dict_file)

        update_dictionaries_and_copy_files(bper_dict, doc_dict, attestation_dict, args.master_directory)

        print("\nUpdated BPER Dictionary:")
        print(bper_dict)

        print("\nUpdated Doc Dictionary:")
        print(doc_dict)

        print("\nUpdated Attestation Dictionary:")
        print(attestation_dict)

if __name__ == "__main__": # for use from command line
    print(f"Current working directory: {os.getcwd()}")
    main()
