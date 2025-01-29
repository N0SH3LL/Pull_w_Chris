
import os
import re
import fitz 
import argparse
import sys

# DEPRECATED !!This file is from when BPERs had to be bulk downloaded, and should not be used anymore!!


def check_already_processed(input_pdf, bper_input_directory): # check to skip processing if BPERs have already been sliced

    if "BPER" not in os.path.basename(input_pdf).upper(): # skips anything without BPER in the name
        print(f"Skipping {input_pdf} (does not contain 'BPER')")
        return False
    
    pdf_document = fitz.open(input_pdf)
    print(f"Checking {input_pdf}")  # Display check status
    first_three_bpers = set()
    for page in pdf_document:
        if "TDL Control:" in page.get_text():
            bper_name = extract_bper_text(page.get_text()).upper()
            first_three_bpers.add(bper_name)
            if len(first_three_bpers) == 3:
                break
    pdf_document.close()

    for bper in first_three_bpers: # check BPER directory for those filenames
        if not os.path.exists(os.path.join(bper_input_directory, f'{bper}.pdf')):
            return False
    return True

def extract_bper_text(page_text): # Search for 'BPER' XXXXXXX from pdf
    match = re.search(r'BPER\d+', page_text)
    return match.group(0) if match else 'rename_me' # file output as rename_me if it can't find a BPER name

def split_pdf(input_pdf): # Split a PDF into separate files based on 'TDL Control:' markers
    if check_already_processed(input_pdf, os.path.dirname(input_pdf)):
        print(f"{os.path.basename(input_pdf)}: already split") #skip if already split
        return
    else:
        print(f"Splitting {os.path.basename(input_pdf)}")

    pdf_document = fitz.open(input_pdf)
    current_output = None
    output_name = None

    for page in pdf_document: 
        page_text = page.get_text() # pull out text from every page until >

        if "TDL Control:" in page_text: # best phrase I could find for separation
            if current_output:
                current_output.save(os.path.join(os.path.dirname(input_pdf), f'{output_name.upper()}.pdf')) # save this file and start a new one
            current_output = fitz.open()  
            output_name = extract_bper_text(page_text) # name from extract_bper_text

        if current_output is not None:
            current_output.insert_pdf(pdf_document, from_page=page.number, to_page=page.number) # adds pages without TDL control to current_output, building new file

    if current_output:
        current_output.save(os.path.join(os.path.dirname(input_pdf), f'{output_name.upper()}.pdf')) # makes sure last file is saved, bc no separator

    pdf_document.close()

def process_directory(directory_path): # Process all PDFs in a directory for splitting and renaming
    print(f"Processing directory: {directory_path}")
    for filename in os.listdir(directory_path):
        if filename.lower().endswith(".pdf") and not re.match(r"BPER\d+.pdf", filename, re.IGNORECASE):
            split_pdf(os.path.join(directory_path, filename))
    
    for filename in os.listdir(directory_path):
        name, extension = os.path.splitext(filename)
        uppercase_filename = name.upper() + extension.lower()  # Rename file to uppercase
        if filename != uppercase_filename:
            os.rename(os.path.join(directory_path, filename), os.path.join(directory_path, uppercase_filename))

if __name__ == "__main__": # Command-line interface to specify directory for PDF processing
    parser = argparse.ArgumentParser(description='Process all PDFs in a directory.')
    parser.add_argument('DirectoryPath', metavar='directory_path', type=str, help='The path to the directory containing PDF files')
    args = parser.parse_args()

    process_directory(args.DirectoryPath)
