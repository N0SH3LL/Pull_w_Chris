
"""
I'm a visual person, and have deviated from PEP8 guidelines to organize the code below in the way that makes the most sense to me. I've tried to spatially organize everything according to how it presents within the GUI. Within the comments, the leading information serves to orient in this fashion '{GUI Screen supported} - {part of that screen supported} - {what it does} - {explanation}. folding (Ctrl+K, Ctrl+0) and Ctrl+F are your friends.  
"""
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import messagebox
from tkinter import *
from tkinter import ttk 
from ttkthemes import ThemedTk
import os
import glob
import json
import re
import subprocess
from datetime import datetime
from openpyxl import Workbook
from difflib import SequenceMatcher
import KAIZEN
from src.utils import split_bper
from src.utils import json_to_excel
from src.utils import doc_validation
from src.utils import file_operations
from src.utils import update_info 
import src.SCC.scc_check
import src.SCC.scc_read
import src.SCC.scc_tables
from src.Archer.fetch_attestations import fetch_attestations, HttpNegotiateAuth
import src.ServiceNow.fetch_Documents
import src.ServiceNow.fetch_BPERs
import src.Tenable.report_operations
import src.Tenable.api_client
import src.Tenable.scan_operations

# directory variables
existing_project_dir = None 
supporting_docs_dir = None
attestation_dir = None
bpers_dir = None
progress_file = None
scc_dir = None
project_dir = None
template_dir = None

### GUI Functions ###
## Welcome Screen ##
def select_directory(prompt): # pop up for selecting dirs
    directory = filedialog.askdirectory(title=prompt)
    return directory
def show_welcome(): # Welcome - Screen - show the first screen
    clear_frames() # hide other frames
    welcome_screen.pack(fill="both", expand=True) # show welcome screen
def ask_scc_gathered(): # Welcome - Button - Start New Project - ask if SCC's have been gather & skip if so 
    return messagebox.askyesno("SCC Gathering", "Have the SCCs already been gathered?")

def select_sccs_to_gather(): # Weclome- Button - Start New Project - Deselect SCC's not needed
    selection_window = tk.Toplevel()
    selection_window.title("Select SCCs to Gather") # kept here bc temp screen
    selection_window.geometry("400x600")

    # Load SCC list from doc_sysids
    config_dir = os.path.join(os.path.dirname(__file__), 'config')
    doc_sysids_path = os.path.join(config_dir, 'doc_sysids.json')
    
    try:
        with open(doc_sysids_path, 'r') as f:
            doc_sysids = json.load(f)
    except Exception as e:
        messagebox.showerror("Error", f"Could not load SCC list: {str(e)}")
        return None

    # Get list of SCCs
    scc_list = [doc for doc in doc_sysids.keys() if "SCC" in doc.upper()]
    selected_sccs = {scc: tk.BooleanVar(value=True) for scc in scc_list}

    # Main container
    main_frame = ttk.Frame(selection_window)
    main_frame.pack(fill="both", expand=True)

    # List container
    list_frame = ttk.Frame(main_frame)
    list_frame.pack(fill="both", expand=True, padx=10, pady=10)

    # Create canvas and scrollbar
    canvas = tk.Canvas(list_frame)
    scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
    
    # Create frame for checkbuttons
    scrollable_frame = ttk.Frame(canvas)
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    # Add mousewheel scrolling
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    # Unbind mousewheel when mouse leaves the window
    def _unbind_mousewheel(e):
        canvas.unbind_all("<MouseWheel>")
    
    def _bind_mousewheel(e):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    canvas.bind('<Enter>', _bind_mousewheel)
    canvas.bind('<Leave>', _unbind_mousewheel)

    # Create window inside canvas
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=canvas.winfo_width())

    # Add checkbuttons for each SCC
    for scc in scc_list:
        cb = ttk.Checkbutton(scrollable_frame, text=scc, variable=selected_sccs[scc])
        cb.pack(anchor="w", pady=2, padx=5, fill="x")

    # Configure canvas expansion
    def configure_canvas(event):
        canvas.itemconfig(1, width=event.width)  # 1 is the ID of the first (and only) window
    canvas.bind('<Configure>', configure_canvas)

    # Pack canvas and scrollbar
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Configure canvas scrolling
    canvas.configure(yscrollcommand=scrollbar.set)

    # Button frame with themed styling
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(fill="x", padx=10, pady=(0, 10))

    # Store results
    result = {}
    
    def confirm_selection():
        canvas.unbind_all("<MouseWheel>")  # Cleanup bindings
        result['sccs'] = [scc for scc, var in selected_sccs.items() if var.get()]
        selection_window.destroy()

    def select_all():
        for var in selected_sccs.values():
            var.set(True)

    def deselect_all():
        for var in selected_sccs.values():
            var.set(False)

    # Create button container for equal spacing
    button_container = ttk.Frame(button_frame)
    button_container.pack(fill="x", expand=True)

    # Add all buttons with equal spacing
    ttk.Button(button_container, text="Select All", command=select_all).pack(
        side="left", expand=True, padx=5)
    ttk.Button(button_container, text="Deselect All", command=deselect_all).pack(
        side="left", expand=True, padx=5)
    ttk.Button(button_container, text="Confirm", command=confirm_selection).pack(
        side="left", expand=True, padx=5)

    # Center the window on screen
    selection_window.update_idletasks()
    width = selection_window.winfo_width()
    height = selection_window.winfo_height()
    x = (selection_window.winfo_screenwidth() // 2) - (width // 2)
    y = (selection_window.winfo_screenheight() // 2) - (height // 2)
    selection_window.geometry(f'{width}x{height}+{x}+{y}')
    
    # Make window modal
    selection_window.transient(selection_window.master)
    selection_window.grab_set()
    
    # Wait for window to close
    selection_window.wait_window()
    
    return result.get('sccs', [])
def start_new_project(): # Welcome - Button - Start New Project 
    global project_dir, progress_file, scc_dir, bpers_dir, attestation_dir, supporting_docs_dir
    
    # Ask user to select the project directory
    parent_dir = filedialog.askdirectory(title="Select location for new project")
    if not parent_dir:
        error_label.config(text="No directory selected. Project creation cancelled.")
        return

    # Create new project directory with current date
    project_name = f"TDL {datetime.now().strftime('%Y-%m-%d')}"
    project_dir = os.path.join(parent_dir, project_name)
    
    try:
        # Create main project directory and subdirectories
        os.makedirs(project_dir)
        scc_dir = os.path.join(project_dir, "SCCs")
        bpers_dir = os.path.join(project_dir, "BPERs")
        attestation_dir = os.path.join(project_dir, "Attestations")
        supporting_docs_dir = os.path.join(project_dir, "Documents")
        
        for dir_path in [scc_dir, bpers_dir, attestation_dir, supporting_docs_dir]:
            os.makedirs(dir_path)

        # Create and populate progress.json
        progress_file = os.path.join(project_dir, "progress.json")
        initial_progress_data = {
            "Program Settings": {
                "Project Directory": project_dir,
                "SCC Directory": scc_dir,
                "BPERs Directory": bpers_dir,
                "Attestation Directory": attestation_dir,
                "Supporting Documents Directory": supporting_docs_dir,
                "Directories Built": True,
                "Templates Built": False,
                "Gather and Sort Date": "",
                "Pull Info Date": "",
                "Checklists generated": ""
            },
            "SCC": {},
            "BPERs": {},
            "Attestations": {},
            "Documents": {},
            "Checks": {}
        }
        
        # Ask if SCCs have already been gathered
        sccs_gathered = ask_scc_gathered()
        
        if not sccs_gathered:
            # Show SCC selection window and get selected SCCs
            selected_sccs = select_sccs_to_gather()
            if selected_sccs is None:  # Error loading SCC list
                return
            
            # Fetch selected SCC documents from SNow only if they haven't been gathered
            fetch_scc_documents(project_dir, scc_dir, selected_sccs)
        else:
            messagebox.showinfo("SCC Processing", "Please copy your SCC files to the SCCs directory.")
            # Wait for user to copy files
            messagebox.showinfo("SCC Processing", "Click OK once you've copied the SCC files.")

        # Process SCC files
        for scc_file in os.listdir(scc_dir):
            if scc_file.endswith('.xlsx') or scc_file.endswith('.xls'):
                file_path = os.path.join(scc_dir, scc_file)
                try:
                    # Run scc_check on the file
                    scc_info = src.SCC.scc_check.process_scc_file(file_path)
                    
                    # Update the progress data with SCC information
                    initial_progress_data["SCC"][file_path] = scc_info
                except Exception as e:
                    # If an error occurs, add the file name with blank values
                    print(f"Error processing {scc_file}: {str(e)}")
                    initial_progress_data["SCC"][file_path] = {
                        "SCC": os.path.splitext(scc_file)[0],  # Use filename without extension as SCC name
                        "Version": "",
                        "SCM Name": "",
                        "Last Review Date": "",
                        "SCC Guidance source presence": False,
                        "SCC Policy and Procedure presence": False,
                        "Exception column presence": False,
                        "Deviation column presence": False,
                        "TLA column presence": False,
                        "Compliance method column presence": False,
                        "WPS config sup doc presence": False,
                        "Reviewed within 180 days": False,
                        "SCC System Scope Presence": False,
                        "Evidence Methods": []
                    }

        # Write the populated progress data to progress.json
        with open(progress_file, 'w') as file:
            json.dump(initial_progress_data, file, indent=4)

        # Update GUI and go to second screen
        update_directory_labels()
        show_options()
        error_label.config(text=f"New project created successfully at {project_dir}")

    except Exception as e:
        error_label.config(text=f"Error creating project: {str(e)}")

def fetch_scc_documents(project_dir, scc_dir, selected_sccs=None): # Welcome - Support - New Project - supports start new project by getting the sccs

    if not os.path.exists(scc_dir):
        error_label.config(text=f"Error: SCC directory not found at {scc_dir}")
        return

    config_dir = os.path.join(os.path.dirname(__file__), 'config')
    doc_sysids_path = os.path.join(config_dir, 'doc_sysids.json')
    doc_sysids = load_doc_sysids(doc_sysids_path)
    
    if doc_sysids is None:
        return
    
    # Filter SCCs based on selection if provided
    if selected_sccs is not None:
        scc_docs = selected_sccs
    else:
        scc_docs = [doc for doc in doc_sysids.keys() if "SCC" in doc.upper()]
    
    try:
        error_label.config(text=f"Fetching {len(scc_docs)} SCC documents...")
        root.update()
        src.ServiceNow.fetch_Documents.main(scc_docs, scc_dir)
        error_label.config(text=f"Successfully fetched {len(scc_docs)} SCC documents.")
    except Exception as e:
        error_label.config(text=f"Error fetching SCC documents: {str(e)}")


def update_existing_project(): # Welcome - Button - Update Existing - Actions on update existing button click
    global progress_file, project_dir
    progress_file = filedialog.askopenfilename(title="Select the progress.json file", filetypes=[("JSON Files", "*.json")])
    if progress_file:
        project_dir = os.path.dirname(progress_file) 
        load_project_settings() 
        update_directory_labels() 
        show_options() # show options screen
    else:
        error_label.config(text="Please select a valid progress.json file.")
def load_project_settings(): # Welcome - Button - Update Existing - Load the progress file
    if progress_file:
        with open(progress_file, 'r') as file: # get below info from progress.json
            progress_data = json.load(file)
            program_settings = progress_data.get('Program Settings', {})
            global scc_dir, bpers_dir, attestation_dir, supporting_docs_dir, template_dir
            scc_dir = program_settings.get('SCC Directory', '')
            bpers_dir = program_settings.get('BPERs Directory', '')
            attestation_dir = program_settings.get('Attestation Directory', '')
            supporting_docs_dir = program_settings.get('Supporting Documents Directory', '')
            template_dir = program_settings.get('Template Directory', '')
            update_directory_labels() 
            update_status_labels(program_settings) 
            
            # Update SCC list
            scc_list = list(progress_data.get('SCC', {}).keys())
            scc_listbox.delete(0, tk.END) # clear current list
            for scc_path in scc_list:
                scc_data = progress_data.get('SCC', {}).get(scc_path, {})
                scc_name = scc_data.get('SCC')
                if scc_name:
                    scc_listbox.insert(tk.END, scc_name) # add SCC names to listbox
            
            # Update "Items Not Gathered" lists
            not_gathered_attestations = []
            not_gathered_bpers = []
            not_gathered_documents = []
            for item_type, item_dict in [('Attestations', progress_data.get('Attestations', {})),
                                        ('BPERs', progress_data.get('BPERs', {})),
                                        ('Documents', progress_data.get('Documents', {}))]:
                for item_id, item_data_list in item_dict.items():
                    for item_data in item_data_list:
                        if not item_data.get('Gathered', True) and not item_data.get('false_positive', False):
                            item_name = item_data.get('BPER name') or item_data.get('Attestation num') or item_data.get('Doc name')
                            scc = item_data.get('SCC')
                            if item_type == 'Attestations':
                                not_gathered_attestations.append(f"{item_name} - {scc}")
                            elif item_type == 'BPERs':
                                not_gathered_bpers.append(f"{item_name} - {scc}")
                            else:
                                not_gathered_documents.append(f"{item_name} - {scc}")
            
            not_gathered_attestations_listbox.delete(0, tk.END) # clear current list
            for item in not_gathered_attestations:
                not_gathered_attestations_listbox.insert(tk.END, item) # add items to listbox
            
            not_gathered_bpers_listbox.delete(0, tk.END) # clear current list
            for item in not_gathered_bpers:
                not_gathered_bpers_listbox.insert(tk.END, item) # add items to listbox
            
            not_gathered_documents_listbox.delete(0, tk.END) # clear current list
            for item in not_gathered_documents:
                not_gathered_documents_listbox.insert(tk.END, item) # add items to listbox
            
            # Update date labels
            last_info_pull_date = program_settings.get('Pull Info Date', 'N/A')
            last_doc_pull_date = program_settings.get('Gather and Sort Date', 'N/A')
            last_checklist_generated_date = program_settings.get('Checklists generated', 'N/A')
            last_info_pull_label.config(text=f"Last Info Pull: {last_info_pull_date}")
            last_doc_pull_label.config(text=f"Last Doc Pull: {last_doc_pull_date}")
            last_checklist_generated_label.config(text=f"Last Checklist Generated: {last_checklist_generated_date}")

## Options Screen ##
def show_options(): # Options - Screen - show the second screen (options/main screen)
    clear_frames() # hide other frames
    options_screen.pack(fill="both", expand=True) # show options screen
    dashboard_screen.pack_forget() # hide dashboard screen
    scans_screen.pack_forget() # hide scans screen
# Options - First Row 
def pull_information(): # Options - Button - Pull Information - Parses SCC's and any documents in the BPER/Attestation/SupDoc folders
    if progress_file:
        base_directories = {
            'bper': bpers_dir,
            'attestation': attestation_dir,
            'doc': supporting_docs_dir
        }
        update_info.update_progress_info(progress_file, base_directories, scc_dir)
        
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        program_settings = progress_data.get('Program Settings', {})
        program_settings['Pull Info Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        progress_data['Program Settings'] = program_settings
        progress_data = remove_duplicates_from_progress(progress_data) # remove duplicates

        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)
        
        update_status_labels(program_settings)
        print("Pull information completed successfully.")
    else:
        error_label.config(text="Please select a valid progress.json file.")

def remove_duplicates_from_progress(progress_data): # exactly what it sounds like, was having this problem bc multiple pull infos would double and triple up doc names 
    for category in ['BPERs', 'Documents', 'Attestations']:
        for key, entries in progress_data[category].items():
            seen_sccs = set()
            unique_entries = []
            for entry in entries:
                scc = entry['SCC']
                if scc not in seen_sccs:
                    seen_sccs.add(scc)
                    unique_entries.append(entry)
            progress_data[category][key] = unique_entries
    return progress_data

def build_dirs(): # Options - Button - Build - Builds out the TDL directories
    if progress_file and project_dir:
        KAIZEN.create_directories(project_dir)  # create directories
        
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        program_settings = progress_data.get('Program Settings', {})
        program_settings['Directories Built'] = True
        progress_data['Program Settings'] = program_settings
        
        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)
        
        build_dirs_status.config(text="Done")  # update status label
        error_label.config(text="Directories built.") # TODO : this is broken, it displays built after an info pull, but before dirs are present. Not sure if the problem is this file or not
    else:
        error_label.config(text="Please select a valid progress.json file and project directory.")
def build_templates(): # Options - Button - Build - Creates the templates
    if progress_file and project_dir and template_dir:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
            method_dict = progress_data.get('Checks', {})
            KAIZEN.build_templates(method_dict, project_dir, template_dir) # build templates

        program_settings = progress_data.get('Program Settings', {})
        program_settings['Templates Built'] = True
        progress_data['Program Settings'] = program_settings

        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)

        build_templates_status.config(text="Done") # update status label
    else:
        error_label.config(text="Please select a valid progress.json file, project directory, and template directory.")
def gather_docs(): # Options - Button - Gather - Starts the doc gathering process
    if progress_file and bpers_dir and attestation_dir and supporting_docs_dir:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
            bper_dict = progress_data.get('BPERs', {})
            doc_dict = progress_data.get('Documents', {})
            attestation_dict = progress_data.get('Attestations', {})

        # Load document sysids
        config_dir = os.path.join(os.path.dirname(__file__), 'config')
        doc_sysids_path = os.path.join(config_dir, 'doc_sysids.json')
        doc_sysids = load_doc_sysids(doc_sysids_path)
        
        if doc_sysids is None:
            return

        # Identify documents to fetch
        docs_to_fetch = []
        for doc_name, doc_info_list in doc_dict.items():
            for doc_info in doc_info_list:
                if not doc_info.get('Gathered', False) and not doc_info.get('false_positive', False):
                    matched_name, sysid = match_document_name(doc_name, doc_sysids)
                    if matched_name:
                        docs_to_fetch.append(matched_name)
                    else:
                        print(f"Warning: No good match found for document {doc_name}")

        # Fetch documents
        if docs_to_fetch:
            gather_docs_status.config(text="Fetching documents...")
            root.update()
            
            if fetch_documents(docs_to_fetch, supporting_docs_dir):
                # Update progress data for fetched documents
                for doc_name, doc_info_list in doc_dict.items():
                    matched_name, _ = match_document_name(doc_name, doc_sysids)
                    if matched_name in docs_to_fetch:
                        for doc_info in doc_info_list:
                            doc_info['Gathered'] = True
                            doc_info['Gathered timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                
                gather_docs_status.config(text="Documents fetched")
            else:
                gather_docs_status.config(text="Error fetching documents")
                return

        # Fetch attestations
        gather_docs_status.config(text="Fetching attestations...")
        root.update()

        try:
            attestation_ids = [int(att_id) for att_id in attestation_dict.keys()]
            
            fetch_attestations(attestation_ids, attestation_dir)
            
            # Update progress data
            for att_id in attestation_ids:
                if str(att_id) in attestation_dict:
                    html_path = os.path.join(attestation_dir, f"{att_id}.html")
                    pdf_path = os.path.join(attestation_dir, f"{att_id}.pdf")
                    attestation_dict[str(att_id)][0]['Gathered'] = os.path.exists(html_path) or os.path.exists(pdf_path)
            
            progress_data['Attestations'] = attestation_dict
            
            messagebox.showinfo("Success", "Attestations fetched successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to fetch attestations: {str(e)}")

        # gather_docs_status.config(text="Attestations fetched")

        # Load BPER sysids
        bper_sysids_path = os.path.join(config_dir, 'BPER_sysids.json')
        bper_sysids = load_bper_sysids(bper_sysids_path)
        
        if bper_sysids is None:
            return

        # Identify BPERs to fetch
        bpers_to_fetch = []
        for bper_name, bper_info_list in bper_dict.items():
            for bper_info in bper_info_list:
                if not bper_info.get('Gathered', False) and not bper_info.get('false_positive', False):
                    if bper_name in bper_sysids:
                        bpers_to_fetch.append(bper_name)
                    else:
                        print(f"Warning: BPER {bper_name} not found in BPER_sysids.json")

        # Fetch BPERs
        if bpers_to_fetch:
            gather_docs_status.config(text="Fetching BPERs...")
            root.update()
            
            if fetch_bpers(bpers_to_fetch, bpers_dir):
                # Update progress data for fetched BPERs
                for bper_name in bpers_to_fetch:
                    for bper_info in bper_dict[bper_name]:
                        bper_info['Gathered'] = True
                        bper_info['Gathered timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                
                gather_docs_status.config(text="BPERs fetched")
            else:
                gather_docs_status.config(text="Error fetching BPERs")
                return

        # Process other documents as before
        base_directories = {'bper': bpers_dir, 'doc': supporting_docs_dir, 'attestation': attestation_dir}
        
        updated_bper_dict, updated_doc_dict, updated_attestation_dict = file_operations.update_dictionaries_and_copy_files(
            bper_dict, doc_dict, attestation_dict, base_directories, project_dir
        )
        
        progress_data['BPERs'] = updated_bper_dict
        progress_data['Documents'] = updated_doc_dict
        progress_data['Attestations'] = updated_attestation_dict

        program_settings = progress_data.get('Program Settings', {})
        program_settings['Gather and Sort Date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        progress_data['Program Settings'] = program_settings
        
        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)
        
        gather_docs_status.config(text=program_settings['Gather and Sort Date'])
        messagebox.showinfo("Success", "Documents gathered and sorted successfully!")
    else:
        error_label.config(text="Please select a valid progress.json file and all required directories.")
def load_bper_sysids(file_path): # Options - Support - Gather - loads sysids from reference file
    
    try:
        with open(file_path, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        messagebox.showerror("Error", f"BPER_sysids.json file not found at {file_path}")
        return None
    except json.JSONDecodeError:
        messagebox.showerror("Error", "Invalid JSON in BPER_sysids.json")
        return None
def fetch_bpers(bper_list, destination_dir): # Options - Support - Gather - pulls bpers from SNow
    
    try:
        src.ServiceNow.fetch_BPERs.main(bper_list, destination_dir)
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to fetch BPERs: {str(e)}")
        return False
def fetch_documents(doc_list, destination_dir): # Options - Support - Gather - pulls supDocs from SNow
    
    try:
        src.ServiceNow.fetch_Documents.main(doc_list, destination_dir)
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Failed to fetch documents: {str(e)}")
        return False
def prepare_doc_name(name): # Options - Support - Gather - supports supDoc name handling from sysid reference
    # Clean up document name by removing file extensions and '...'
    name = name.lower()  # Convert to lowercase for case-insensitive matching
    name = re.sub(r'\.[^.]+$', '', name)  # Remove file extension
    name = name.rstrip('.')  # Remove trailing dots
    name = name.rstrip('…')  # Remove trailing ellipsis (…)
    return name.strip()  # Remove leading/trailing whitespace
def match_document_name(progress_name, sysids_dict, threshold=0.5): # Options - Support - Supports the gather docs button, maps SNow sys_ids to doc names for BPER/supdocs
    
    progress_name_clean = prepare_doc_name(progress_name)
    best_match = None
    best_ratio = 0

    for sysid_name, sysid in sysids_dict.items():
        sysid_name_clean = prepare_doc_name(sysid_name)
        ratio = SequenceMatcher(None, progress_name_clean, sysid_name_clean).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_match = (sysid_name, sysid)

    if best_ratio >= threshold:
        print(f"Matched: {progress_name} -> {best_match[0]} (Accuracy: {best_ratio:.2f})")
        return best_match
    else:
        print(f"No good match found for: {progress_name} (Best accuracy: {best_ratio:.2f})")
        return None, None
def load_doc_sysids(file_path): # Options - Support - Supports the gather docs button, maps SNow sys_ids to doc names for BPER/ supdocs
    
    try:
        with open(file_path, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        messagebox.showerror("Error", f"doc_sysids.json file not found at {file_path}")
        return None
    except json.JSONDecodeError:
        messagebox.showerror("Error", "Invalid JSON in doc_sysids.json")
        return None
def generate_md_files(): # Options - Button - Generate MD Files - generate markdown files for the SCC checklists
    if progress_file and project_dir:
        src.SCC.scc_tables.generate_scc_info_docs(progress_file) 
        
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        program_settings = progress_data.get('Program Settings', {})
        program_settings['Checklists generated'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        progress_data['Program Settings'] = program_settings
        
        # Store the generated file paths in the respective "SCC" dictionary entry
        for scc_file, scc_data in progress_data.get('SCC', {}).items():
            scc_name = scc_data.get('SCC')
            if scc_name:
                md_file_path = os.path.join(project_dir, f"{scc_name}_info.md")
                if os.path.exists(md_file_path):
                    scc_data['Info Doc Path'] = md_file_path
        
        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)
        
        generate_md_status.config(text=program_settings['Checklists generated']) # update status label
    else:
        error_label.config(text="Please select a valid progress.json file and project directory.")
def update_document_validation(): # Options - Button - Update - Updates the doc validation tracker in the templates folder
    if progress_file and template_dir:
        template_path = os.path.join(template_dir, "Document Validation.xlsx")
        if os.path.exists(template_path):
            try:
                doc_validation.update_document_validation(progress_file, template_path)
                error_label.config(text="Document Validation.xlsx updated successfully.")
            except Exception as e:
                error_label.config(text=f"Error updating Document Validation.xlsx: {str(e)}")
        else:
            error_label.config(text="Document Validation.xlsx not found in the template directory.")
    else:
        error_label.config(text="Please select a valid progress.json file and template directory.")
def update_status_labels(program_settings): # Options - Presentation - updates status indicators underneath top row buttons 
    # Update status labels with program settings
    directories_built = program_settings.get('Directories Built', False)
    templates_built = program_settings.get('Templates Built', False)
    gather_sort_date = program_settings.get('Gather and Sort Date', 'Not Done')
    checklist_generated = program_settings.get('Checklists generated', 'Not Done')
    pull_info_date = program_settings.get('Pull Info Date', 'Not Done')

    build_dirs_status.config(text="Done" if directories_built else "Not Done")
    build_templates_status.config(text="Done" if templates_built else "Not Done")
    gather_docs_status.config(text=gather_sort_date)
    generate_md_status.config(text=checklist_generated)
    pull_info_status.config(text=pull_info_date)
# Options - Second Row
def output_progress(): # Options - Button - Output Progress - puts progress.json info into an excel file
    if progress_file and project_dir:
        try:
            # Load the JSON data from the progress file
            with open(progress_file, 'r') as file:
                data = json.load(file)

            # Create a new workbook
            wb = Workbook()

            # Create a sheet for each high-level dictionary
            for key in data:
                json_to_excel.create_sheet(wb, key, data[key])

            # Remove the default sheet created by openpyxl
            default_sheet = wb['Sheet']
            wb.remove(default_sheet)

            # Save the workbook in the project directory
            output_file = os.path.join(project_dir, 'progress.xlsx')
            wb.save(output_file)

            error_label.config(text="Progress exported successfully!")
        except Exception as e:
            error_label.config(text=f"Error exporting progress: {str(e)}")
    else:
        error_label.config(text="Please select a valid progress.json file and project directory.")
def add_or_redo_scc(): # Options - Button - Add or redo an SCC
    # Open file explorer dialog to select an Excel file
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    
    if file_path:
        # Extract SCC name from the file path and remove extension and trailing "_**"
        scc_name = os.path.splitext(os.path.basename(file_path))[0]
        scc_name = re.sub(r'_\d{2}$', '', scc_name).strip()
        
        # Load progress data from progress.json
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        # Check if the SCC exists in progress data
        if scc_name in progress_data['SCC']:
            # Remove all entries with the matching SCC name
            for key in ['BPERs', 'Attestations', 'Documents', 'SCC', 'Checks']:
                progress_data[key] = {k: v for k, v in progress_data[key].items() if v.get('SCC') != scc_name}
        
        # Process the selected Excel file
        bper_dict, doc_dict, attestation_dict, method_dict = src.SCC.scc_read.process_excel_file(file_path)
        
        # Update progress data with the new information
        for key, value in bper_dict.items():
            progress_data['BPERs'][key] = value
        for key, value in doc_dict.items():
            progress_data['Documents'][key] = value
        for key, value in attestation_dict.items():
            progress_data['Attestations'][key] = value
        for stig_id, details in method_dict.items():
            progress_data['Checks'][stig_id] = {
                'SCC': scc_name,
                'Evidence method': details['Evidence Method']
            }
        
        # Save the updated progress data to progress.json
        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)
        
        error_label.config(text=f"SCC '{scc_name}' added or updated successfully.") # update status message
def remove_scc(): # Options - Button - Remove an SCC - 
    remove_scc_window = tk.Toplevel(root) # create new window for SCC removal
    remove_scc_window.title("Remove an SCC")
    remove_scc_window.geometry("400x300")

    scc_list_frame = tk.Frame(remove_scc_window)
    scc_list_frame.pack(fill="both", expand=True, padx=10, pady=10)

    scc_list_label = tk.Label(scc_list_frame, text="Select an SCC to remove:")
    scc_list_label.pack()

    scc_list_listbox = tk.Listbox(scc_list_frame, font=("Arial", 10))
    scc_list_listbox.pack(fill="both", expand=True)

    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)

        for scc_path in progress_data.get('SCC', {}):
            scc_data = progress_data['SCC'][scc_path]
            scc_name = scc_data.get('SCC')
            if scc_name:
                scc_list_listbox.insert(tk.END, scc_name) # add SCC names to listbox

    delete_button = tk.Button(remove_scc_window, text="Delete", command=lambda: delete_scc(scc_list_listbox.get(scc_list_listbox.curselection())))
    delete_button.pack(pady=10)
def delete_scc(scc_name): # Options - Support - Supports remove an scc
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)

        for key in ['BPERs', 'Attestations', 'Documents']:
            updated_dict = {}
            for item_key, item_list in progress_data[key].items():
                updated_list = [item for item in item_list if item.get('SCC') != scc_name]
                if updated_list:
                    updated_dict[item_key] = updated_list
            progress_data[key] = updated_dict

        progress_data['SCC'] = {k: v for k, v in progress_data['SCC'].items() if v.get('SCC') != scc_name}
        progress_data['Checks'] = {k: v for k, v in progress_data['Checks'].items() if v.get('SCC') != scc_name}

        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)

        error_label.config(text=f"SCC '{scc_name}' removed successfully.")
        load_project_settings()  # Refresh the dashboard after removing an SCC
    else:
        error_label.config(text="Please select a valid progress.json file.")
        error_label.config(text="Please select a valid progress.json file.")
def sync_button_click(): # Options - Button - Sync - Syncs the progress file with the md files
    if progress_file:
        src.SCC.scc_tables.sync_progress_info(progress_file) # sync progress info
    else:
        error_label.config(text="Please select a valid progress.json file.")
# Options - Selected Dirs Section
def update_directory_labels(): # Options - Area - Selected Directories - Update directory labels on the options screen with the selected paths
    if progress_file:
        progress_file_label.config(text=f"Progress File: {progress_file}")
    else:
        progress_file_label.config(text="Progress File: Not selected")

    if bpers_dir:
        bpers_dir_label.config(text=f"BPERs Directory: {bpers_dir}")
    else:
        bpers_dir_label.config(text="BPERs Directory: Not selected")

    if attestation_dir:
        attestation_dir_label.config(text=f"Attestation Directory: {attestation_dir}")
    else:
        attestation_dir_label.config(text="Attestation Directory: Not selected")

    if supporting_docs_dir:
        supporting_docs_dir_label.config(text=f"Supporting Documents Directory: {supporting_docs_dir}")
    else:
        supporting_docs_dir_label.config(text="Supporting Documents Directory: Not selected")

    if project_dir:
        project_dir_label.config(text=f"Project Directory: {project_dir}")
    else:
        project_dir_label.config(text="Project Directory: Not selected")

    if scc_dir:
        scc_dir_label.config(text=f"SCC Directory: {scc_dir}")
    else:
        scc_dir_label.config(text="SCC Directory: Not selected")
def select_bpers_directory(): # Options - Button - Select BPER Directory
    global bpers_dir
    bpers_dir = select_directory("Please select the BPERs directory") # select BPERs directory
    if bpers_dir:
        bpers_dir_label.config(text=f"BPERs Directory: {bpers_dir}") # update label
        save_project_settings() # save settings
def select_attestation_directory(): # Options - Button - Select Attestation Directory
    global attestation_dir
    attestation_dir = select_directory("Please select the Attestation directory") # select attestation directory
    if attestation_dir:
        attestation_dir_label.config(text=f"Attestation Directory: {attestation_dir}") # update label
        save_project_settings() # save settings
def select_supporting_docs_directory(): # Options - Button - Select Documents Directory
    global supporting_docs_dir
    supporting_docs_dir = select_directory("Please select the Supporting Documents directory") # select supporting docs directory
    if supporting_docs_dir:
        supporting_docs_dir_label.config(text=f"Supporting Documents Directory: {supporting_docs_dir}") # update label
        save_project_settings() # save settings
def select_scc_directory(): # Options - Button - Select SCC Directory
    global scc_dir
    scc_dir = select_directory("Please select the SCC directory") # select SCC directory
    if scc_dir:
        scc_dir_label.config(text=f"SCC Directory: {scc_dir}") # update label
        save_project_settings() # save settings
def select_progress_file(): # Options - Button - Select progress file
    global progress_file
    progress_file = filedialog.askopenfilename(title="Select the progress.json file", filetypes=[("JSON Files", "*.json")])
    if progress_file:
        progress_file_label.config(text=f"Progress File: {progress_file}") # update label
        load_project_settings() # load settings
        update_directory_labels() # update directory labels
def select_project_directory(): # Options - Button - Select Project Directory
    global project_dir
    project_dir = select_directory("Select the project directory") # select project directory
    if project_dir:
        project_dir_label.config(text=f"Project Directory: {project_dir}") # update label
        save_project_settings() # save settings
def select_template_directory(): # Options - Button - Select Template Directory
    global template_dir
    template_dir = select_directory("Please select the Template directory") # select template directory
    if template_dir:
        template_dir_label.config(text=f"Template Directory: {template_dir}") # update label
        save_project_settings() # save settings
def save_project_settings(): # Options - Support - Supports the buttons that select directories
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        program_settings = progress_data.get('Program Settings', {})
        program_settings['SCC Directory'] = scc_dir
        program_settings['BPERs Directory'] = bpers_dir
        program_settings['Attestation Directory'] = attestation_dir
        program_settings['Supporting Documents Directory'] = supporting_docs_dir
        program_settings['Template Directory'] = template_dir
        progress_data['Program Settings'] = program_settings

        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)

## Dashboard Screen ##
def show_dashboard(): # Dashboard - Screen - show the third screen (Doc Dashboard)
    clear_frames() # hide other frames
    dashboard_screen.pack(fill="both", expand=True) # show dashboard screen
    refresh_dashboard() # update dashboard data
def refresh_dashboard(): # Dashboard - Presentation -  handles the dashboard screen
    scc_listbox.delete(0, tk.END)  # clear current list
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        for scc_path, scc_data in progress_data['SCC'].items():
            scc_name = scc_data.get('SCC')
            if scc_name:
                # Check if any linked items are not gathered
                has_ungathered_items = False
                for category in ['Attestations', 'BPERs', 'Documents']:
                    for item_data_list in progress_data.get(category, {}).values():
                        for item_data in item_data_list:
                            if item_data.get('SCC') == scc_name and not item_data.get('Gathered', False):
                                has_ungathered_items = True
                                break
                        if has_ungathered_items:
                            break
                    if has_ungathered_items:
                        break
                
                # Insert SCC name with appropriate background color
                if has_ungathered_items:
                    scc_listbox.insert(tk.END, scc_name)
                    scc_listbox.itemconfig(tk.END, {'bg': '#FFB6C1'})  # Light red
                else:
                    scc_listbox.insert(tk.END, scc_name)
                    scc_listbox.itemconfig(tk.END, {'bg': '#90EE90'})  # Light green

        categories = ['Documents', 'Attestations', 'BPERs']
        gathered_counts = []
        total_counts = []

        for category in categories:
            items = progress_data.get(category, {})
            gathered_count = sum(1 for item_data_list in items.values() for item_data in item_data_list if item_data.get('Gathered', False) and not item_data.get('false_positive', False))
            total_count = sum(1 for item_data_list in items.values() for item_data in item_data_list if not item_data.get('false_positive', False))
            gathered_counts.append(gathered_count)
            total_counts.append(total_count)

        pie_chart_text = ""
        for category, gathered_count, total_count in zip(categories, gathered_counts, total_counts):
            percentage = (gathered_count / total_count) * 100 if total_count > 0 else 0
            pie_chart_text += f"{category}: {gathered_count}/{total_count} ({percentage:.2f}%)\n"

        pie_chart_label.config(text=pie_chart_text)
    else:
        pie_chart_label.config(text="No progress file selected.")
    
    # Update "Items Not Gathered" lists
    not_gathered_attestations = []
    not_gathered_bpers = []
    not_gathered_documents = []
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        for item_type, item_dict in [('Attestations', progress_data.get('Attestations', {})),
                                     ('BPERs', progress_data.get('BPERs', {})),
                                     ('Documents', progress_data.get('Documents', {}))]:
            for item_id, item_data_list in item_dict.items():
                for item_data in item_data_list:
                    if not item_data.get('Gathered', True) and not item_data.get('false_positive', False):
                        item_name = item_data.get('BPER name') or item_data.get('Attestation num') or item_data.get('Doc name')
                        scc = item_data.get('SCC')
                        if item_type == 'Attestations':
                            not_gathered_attestations.append(f"{item_name} - {scc}")
                        elif item_type == 'BPERs':
                            not_gathered_bpers.append(f"{item_name} - {scc}")
                        else:
                            not_gathered_documents.append(f"{item_name} - {scc}")
    
    not_gathered_attestations_listbox.delete(0, tk.END)  # clear current list
    for item in not_gathered_attestations:
        not_gathered_attestations_listbox.insert(tk.END, item)  # add items to listbox
    
    not_gathered_bpers_listbox.delete(0, tk.END)  # clear current list
    for item in not_gathered_bpers:
        not_gathered_bpers_listbox.insert(tk.END, item)  # add items to listbox
    
    not_gathered_documents_listbox.delete(0, tk.END)  # clear current list
    for item in not_gathered_documents:
        not_gathered_documents_listbox.insert(tk.END, item)  # add items to listbox
    
    # Update date labels
    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        program_settings = progress_data.get('Program Settings', {})
        last_info_pull_date = program_settings.get('Pull Info Date', 'N/A')
        last_doc_pull_date = program_settings.get('Gather and Sort Date', 'N/A')
        last_checklist_generated_date = program_settings.get('Checklists generated', 'N/A')
    else:
        last_info_pull_date = 'N/A'
        last_doc_pull_date = 'N/A'
        last_checklist_generated_date = 'N/A'
    
    last_info_pull_label.config(text=f"Last Info Pull: {last_info_pull_date}")
    last_doc_pull_label.config(text=f"Last Doc Pull: {last_doc_pull_date}")
    last_checklist_generated_label.config(text=f"Last Checklist Generated: {last_checklist_generated_date}")
# Dashboard - SCC List Section
def open_scc_markdown_file(event): # Dashboard - Button - opens relevant checklist if you click on an SCC name
    selected_indices = scc_listbox.curselection()
    if selected_indices:
        selected_scc = scc_listbox.get(selected_indices[0])
        if progress_file:
            with open(progress_file, 'r') as file:
                progress_data = json.load(file)
                for scc_path, scc_data in progress_data['SCC'].items():
                    if scc_data.get('SCC') == selected_scc:
                        md_file_path = scc_data.get('Info Doc Path')
                        if md_file_path and os.path.exists(md_file_path):
                            os.startfile(md_file_path)
                            return
                error_label.config(text=f"Markdown file not found for SCC: {selected_scc}")
    else:
        error_label.config(text="Please select a valid progress.json file.")
# Dashboard - Items Not Gathered Section
def mark_as_false_positive(item_type): #  Dashboard - Button - Mark as false positive - designates selected documents as false pos in progress.json
    # Mark selected items as false positive based on item type
    if item_type == "BPERs":
        selected_items = [not_gathered_bpers_listbox.get(idx) for idx in not_gathered_bpers_listbox.curselection()]
        listbox = not_gathered_bpers_listbox
        dict_key = "BPERs"
    elif item_type == "Attestations":
        selected_items = [not_gathered_attestations_listbox.get(idx) for idx in not_gathered_attestations_listbox.curselection()]
        listbox = not_gathered_attestations_listbox
        dict_key = "Attestations"
    elif item_type == "Documents":
        selected_items = [not_gathered_documents_listbox.get(idx) for idx in not_gathered_documents_listbox.curselection()]
        listbox = not_gathered_documents_listbox
        dict_key = "Documents"
    
    with open(progress_file, 'r') as file:
        progress_data = json.load(file)
    
    for selected_item in selected_items:
        item_name, scc = selected_item.split(" - ")
        
        for item_id, item_data_list in progress_data[dict_key].items():
            for item_data in item_data_list:
                if item_data.get("BPER name") == item_name or item_data.get("Attestation num") == item_name or item_data.get("Doc name") == item_name:
                    item_data["false_positive"] = True
                    break
    
    with open(progress_file, 'w') as file:
        json.dump(progress_data, file, indent=4)
    
    for idx in reversed(listbox.curselection()): # remove marked items from listbox
        listbox.delete(idx)
    
    item_name, scc = selected_item.split(" - ")
    
    with open(progress_file, 'r') as file:
        progress_data = json.load(file)
    
    for item_id, item_data_list in progress_data[dict_key].items():
        for item_data in item_data_list:
            if item_data.get("BPER name") == item_name or item_data.get("Attestation num") == item_name or item_data.get("Doc name") == item_name:
                item_data["false_positive"] = True
                break
    
    with open(progress_file, 'w') as file:
        json.dump(progress_data, file, indent=4)
    
    listbox.delete(listbox.curselection())
def manually_link_files(item_type): # Dashboard - Button - Assign Match - Manually link selected files to items based on item type
    if item_type == "BPERs":
        selected_items = [not_gathered_bpers_listbox.get(idx) for idx in not_gathered_bpers_listbox.curselection()]
        listbox = not_gathered_bpers_listbox
        dict_key = "BPERs"
        directory = bpers_dir
    elif item_type == "Attestations":
        selected_items = [not_gathered_attestations_listbox.get(idx) for idx in not_gathered_attestations_listbox.curselection()]
        listbox = not_gathered_attestations_listbox
        dict_key = "Attestations"
        directory = attestation_dir
    elif item_type == "Documents":
        selected_items = [not_gathered_documents_listbox.get(idx) for idx in not_gathered_documents_listbox.curselection()]
        listbox = not_gathered_documents_listbox
        dict_key = "Documents"
        directory = supporting_docs_dir
    
    with open(progress_file, 'r') as file:
        progress_data = json.load(file)
    
    for selected_item in selected_items:
        item_name, scc = selected_item.split(" - ")
        
        file_path = filedialog.askopenfilename(initialdir=directory, title=f"Select file for {item_name}")
        
        if file_path:
            for item_id, item_data_list in progress_data[dict_key].items():
                if len(item_data_list) > 1:  # Check if there are multiple sub-values
                    for item_data in item_data_list:
                        if item_data.get("Doc name") == item_name:
                            for sub_item_data in item_data_list:
                                sub_item_data["manually_linked"] = file_path
                            break
                else:
                    for item_data in item_data_list:
                        if item_data.get("BPER name") == item_name or item_data.get("Attestation num") == item_name or item_data.get("Doc name") == item_name:
                            item_data["manually_linked"] = file_path
                            break
    
    with open(progress_file, 'w') as file:
        json.dump(progress_data, file, indent=4)

## Scans Screen ## 
def show_scans(): # Scans - Screen - show the fourth screen (Scans)
    clear_frames()
    scans_screen.pack(fill="both", expand=True)
    update_inventory_display()
    update_scan_status_display()
    refresh_report_status()
# Scans - Inventories Section
def update_inventory_display(): # Scans - Area - Inventories pane updates/ populates the inventory pane
    for widget in inventory_frame.winfo_children():
        widget.destroy()

    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        scc_dict = progress_data.get('SCC', {})
        
        for scc_path, scc_info in scc_dict.items():
            scc_name = scc_info['SCC']
            inventory_file = scc_info.get('Inventory File', '')
            evidence_methods = [method.lower() for method in scc_info.get('Evidence Methods', [])]
            
            # Check if inventory is required (automated or manual-auto info)
            inventory_required = 'automated' in evidence_methods or 'manual-auto info' in evidence_methods
            
            if inventory_required:
                if inventory_file:
                    status = "Present"
                    color = "#90EE90"  # Light green
                else:
                    status = "Not Present"
                    color = "#FFB6C1"  # Light red
            else:
                status = "Not Required"
                color = "#FFFFFF"  # White
            
            scc_frame = tk.Frame(inventory_frame, bg=color)
            scc_frame.pack(fill="x", padx=5, pady=2)
            
            scc_label = tk.Label(scc_frame, text=f"{scc_name}: {status}", bg=color, anchor="w")
            scc_label.pack(fill="x")
    else:
        placeholder_label = tk.Label(inventory_frame, text="No progress file selected", bg="#FFFFFF")
        placeholder_label.pack(pady=10)

    inventory_canvas.configure(scrollregion=inventory_canvas.bbox("all"))
def check_inventories(): # Scans - Button - Check Inventories - checks for inventory files in built out SCC dirs
    if not progress_file or not project_dir:
        error_label.config(text="Please select a valid progress.json file and project directory.")
        return

    print("Checking inventories...")
    inventory_status = []

    try:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)

        for scc_path, scc_info in progress_data['SCC'].items():
            scc_name = scc_info['SCC']
            inventory_file_name = f"{scc_name}-Inventory.txt"
            inventory_file_path = os.path.join(project_dir, scc_name, inventory_file_name)

            if os.path.exists(inventory_file_path):
                progress_data['SCC'][scc_path]['Inventory File'] = inventory_file_path
                inventory_status.append(f"Found inventory for {scc_name}")
            else:
                progress_data['SCC'][scc_path]['Inventory File'] = ""
                inventory_status.append(f"No inventory found for {scc_name}")

        # Save the updated progress data
        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)

        # Update the last inventory check timestamp
        program_settings = progress_data.get('Program Settings', {})
        program_settings['Last Inventory Check'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        progress_data['Program Settings'] = program_settings

        with open(progress_file, 'w') as file:
            json.dump(progress_data, file, indent=4)

        print("Inventory check completed and progress.json updated.")
        
        # Update the inventory display
        update_inventory_display()
        update_scan_status_display()
    except Exception as e:
        error_message = f"Error checking inventories: {str(e)}"
        print(error_message)
        error_label.config(text=error_message)
# Scans - Scan Status Section
def update_scan_status_display(): # Scans - Area - Scan Status Pane - updates/ populates scan status pane
    for widget in scan_status_frame.winfo_children():
        widget.destroy()

    if progress_file:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        scc_dict = progress_data.get('SCC', {})
        
        for scc_path, scc_info in scc_dict.items():
            scc_name = scc_info['SCC']
            inventory_file = scc_info.get('Inventory File', '')
            evidence_methods = [method.lower() for method in scc_info.get('Evidence Methods', [])]
            
            if 'automated' in evidence_methods:
                passfail_status = scc_info.get('PassFail_Status', 'Ready' if inventory_file else 'Not Ready')
                if not inventory_file:
                    color = "#FFB6C1"  # Light red
                elif passfail_status == 'Queued':
                    color = "#90EE90"  # Light green
                else:
                    color = "#FFFFFF"  # White
                passfail_frame = tk.Frame(scan_status_frame, bg=color)
                passfail_frame.pack(fill="x", padx=5, pady=2)
                passfail_label = tk.Label(passfail_frame, text=f"{scc_name} -PassFail: {passfail_status}", bg=color, anchor="w")
                passfail_label.pack(fill="x")
            
            if 'manual-auto info' in evidence_methods:
                info_status = scc_info.get('Info_Status', 'Ready' if inventory_file else 'Not Ready')
                if not inventory_file:
                    color = "#FFB6C1"  # Light red
                elif info_status == 'Queued':
                    color = "#90EE90"  # Light green
                else:
                    color = "#FFFFFF"  # White
                info_frame = tk.Frame(scan_status_frame, bg=color)
                info_frame.pack(fill="x", padx=5, pady=2)
                info_label = tk.Label(info_frame, text=f"{scc_name} -Info: {info_status}", bg=color, anchor="w")
                info_label.pack(fill="x")
            
            if not evidence_methods:
                no_scan_frame = tk.Frame(scan_status_frame, bg="#FFFFFF")
                no_scan_frame.pack(fill="x", padx=5, pady=2)
                no_scan_label = tk.Label(no_scan_frame, text=f"{scc_name}: No scans required", bg="#FFFFFF", anchor="w")
                no_scan_label.pack(fill="x")
    else:
        placeholder_label = tk.Label(scan_status_frame, text="No progress file selected", bg="#FFFFFF")
        placeholder_label.pack(pady=10)

    scan_status_canvas.configure(scrollregion=scan_status_canvas.bbox("all"))
def initiate_scans(): # Scans - Button - Initiate Scans - Launches the scans for SCC's that are ready
    if not progress_file:
        error_label.config(text="Please select a valid progress.json file.")
        return

    # Ask user if they want to launch all scans
    launch_all = messagebox.askyesno("Launch Scans", "Do you want to launch all scans?", parent=root)
    
    with open(progress_file, 'r') as file:
        progress_data = json.load(file)

    # Get inputs
    chunk_size = simpledialog.askinteger("Input", "Enter chunk size:", minvalue=1, maxvalue=100, parent=root)
    if chunk_size is None:
        return  # User cancelled

    start_time = simpledialog.askstring("Input", "Enter start time (format: YYYYMMDDTHHMMSS):", parent=root)
    if not start_time or not re.match(r'\d{8}T\d{6}', start_time):
        error_label.config(text="Invalid start time format. Please use YYYYMMDDTHHMMSS.", parent=root)
        return

    access_key = simpledialog.askstring("Tenable Credentials", "Enter Tenable Access Key:", parent=root) # Tenable API access key
    if not access_key:
        return
        
    secret_key = simpledialog.askstring("Tenable Credentials", "Enter Tenable Secret Key:", parent=root, show="*") # Tenable API secret key
    if not secret_key:
        return

    client = src.Tenable.api_client.TenableSCClient(access_key, secret_key)

    if launch_all:
        # Original - launch all scans
        for scc_path, scc_info in progress_data['SCC'].items():
            scc_name = scc_info['SCC']
            inventory_file = scc_info.get('Inventory File')
            
            if not inventory_file:
                print(f"Skipping {scc_name}: No inventory file found.")
                continue

            evidence_methods = [method.lower() for method in scc_info.get('Evidence Methods', [])]
            
            if 'automated' in evidence_methods:
                progress_data['SCC'][scc_path]['PassFail_Status'] = 'Queued'
                passfail_scan_name = f"TDL-{scc_name}-PassFail"
                print(f"Initiating PassFail scan for {scc_name}")
                src.Tenable.scan_operations.chunk_and_create_scans(client, passfail_scan_name, inventory_file, start_time, chunk_size)
            
            if 'manual-auto info' in evidence_methods:
                progress_data['SCC'][scc_path]['Info_Status'] = 'Queued'
                info_scan_name = f"TDL-{scc_name}-Info"
                print(f"Initiating Info scan for {scc_name}")
                src.Tenable.scan_operations.chunk_and_create_scans(client, info_scan_name, inventory_file, start_time, chunk_size)
    else:
        # Show dialog for selecting individual SCC
        select_window = tk.Toplevel()
        select_window.title("Select SCC to Scan")
        select_window.geometry("400x300")

        # Create listbox for SCC selection
        scc_listbox = tk.Listbox(select_window, selectmode=tk.SINGLE)
        scc_listbox.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Populate listbox with eligible SCCs
        eligible_sccs = []
        for scc_path, scc_info in progress_data['SCC'].items():
            scc_name = scc_info['SCC']
            inventory_file = scc_info.get('Inventory File')
            evidence_methods = [method.lower() for method in scc_info.get('Evidence Methods', [])]
            
            if inventory_file and ('automated' in evidence_methods or 'manual-auto info' in evidence_methods):
                scc_listbox.insert(tk.END, scc_name)
                eligible_sccs.append((scc_path, scc_info))

        def launch_selected_scan():
            selection = scc_listbox.curselection()
            if not selection:
                messagebox.showerror("Error", "Please select an SCC")
                return
                
            selected_idx = selection[0]
            scc_path, scc_info = eligible_sccs[selected_idx]
            scc_name = scc_info['SCC']
            inventory_file = scc_info.get('Inventory File')
            evidence_methods = [method.lower() for method in scc_info.get('Evidence Methods', [])]

            # Launch scans for selected SCC
            if 'automated' in evidence_methods:
                progress_data['SCC'][scc_path]['PassFail_Status'] = 'Queued'
                passfail_scan_name = f"TDL-{scc_name}-PassFail"
                print(f"Initiating PassFail scan for {scc_name}")
                src.Tenable.scan_operations.chunk_and_create_scans(client, passfail_scan_name, inventory_file, start_time, chunk_size)
            
            if 'manual-auto info' in evidence_methods:
                progress_data['SCC'][scc_path]['Info_Status'] = 'Queued'
                info_scan_name = f"TDL-{scc_name}-Info"
                print(f"Initiating Info scan for {scc_name}")
                src.Tenable.scan_operations.chunk_and_create_scans(client, info_scan_name, inventory_file, start_time, chunk_size)
            
            select_window.destroy()

        # Add launch button
        launch_btn = tk.Button(select_window, text="Launch Scan", command=launch_selected_scan)
        launch_btn.pack(pady=10)

    # Save the updated progress data
    with open(progress_file, 'w') as file:
        json.dump(progress_data, file, indent=4)

    # Update the scan list display
    populate_scan_list()

    print("Scan initiation process completed.")
def populate_scan_list(): # Scans - Support - supports the initiate scans
    # Clear existing widgets in the scan frame
    for widget in scan_frame.winfo_children():
        widget.destroy()

    if not progress_file:
        error_label.config(text="Please select a valid progress.json file.")
        return

    with open(progress_file, 'r') as file:
        progress_data = json.load(file)

    for scc_path, scc_info in progress_data['SCC'].items():
        scc_name = scc_info['SCC']
        evidence_methods = [method.lower() for method in scc_info.get('Evidence Methods', [])]

        if 'automated' in evidence_methods:
            status = scc_info.get('PassFail_Status', 'Ready')
            color = "#90EE90" if status == "Queued" else "#FFFFFF"  # Light green if Queued, white if Ready
            scan_item_frame = tk.Frame(scan_frame, bg=color)
            scan_item_frame.pack(fill="x", padx=5, pady=2)
            scan_item_label = tk.Label(scan_item_frame, text=f"{scc_name}-PassFail: {status}", bg=color, anchor="w")
            scan_item_label.pack(fill="x")

        if 'manual-auto info' in evidence_methods:
            status = scc_info.get('Info_Status', 'Ready')
            color = "#90EE90" if status == "Queued" else "#FFFFFF"  # Light green if Queued, white if Ready
            scan_item_frame = tk.Frame(scan_frame, bg=color)
            scan_item_frame.pack(fill="x", padx=5, pady=2)
            scan_item_label = tk.Label(scan_item_frame, text=f"{scc_name}-Info: {status}", bg=color, anchor="w")
            scan_item_label.pack(fill="x")

    # Update the canvas scroll region
    scan_canvas.configure(scrollregion=scan_canvas.bbox("all"))
# Scans - Report Status Section
def refresh_report_status(): # Scans - Button - Refresh Report Status - updates / populates report status pane
    # Clear existing widgets in the report status frame
    for widget in report_status_frame.winfo_children():
        widget.destroy()

    if progress_file and project_dir:
        with open(progress_file, 'r') as file:
            progress_data = json.load(file)
        
        scc_dict = progress_data.get('SCC', {})
        
        for scc_path, scc_info in scc_dict.items():
            scc_name = scc_info['SCC']
            evidence_methods = [method.lower() for method in scc_info.get('Evidence Methods', [])]
            
            passfail_required = 'automated' in evidence_methods
            info_required = 'manual-auto info' in evidence_methods
            
            if not passfail_required and not info_required:
                status = "Reports: Not Required"
                color = "#90EE90"  # Light green
            else:
                passfail_collected = False
                info_collected = False
                
                if passfail_required:
                    automated_folder = os.path.join(project_dir, scc_name, "Automated")
                    passfail_collected = any(file.endswith('.csv') for file in os.listdir(automated_folder)) if os.path.exists(automated_folder) else False
                
                if info_required:
                    info_folder = os.path.join(project_dir, scc_name, "Manual", "Automated Info")
                    info_collected = any(file.endswith('.pdf') for file in os.listdir(info_folder)) if os.path.exists(info_folder) else False
                
                all_required_collected = (not passfail_required or passfail_collected) and (not info_required or info_collected)
                
                if all_required_collected:
                    status = "Reports: All Collected"
                    color = "#90EE90"  # Light green
                elif (passfail_required and passfail_collected) or (info_required and info_collected):
                    status = "Reports: Partially Collected"
                    color = "#FFFF00"  # Yellow
                else:
                    status = "Reports: Not Collected"
                    color = "#FFB6C1"  # Light red
            
            scc_frame = tk.Frame(report_status_frame, bg=color)
            scc_frame.pack(fill="x", padx=5, pady=2)
            
            scc_label = tk.Label(scc_frame, text=f"{scc_name}: {status}", bg=color, anchor="w")
            scc_label.pack(fill="x")
    else:
        placeholder_label = tk.Label(report_status_frame, text="No progress file selected", bg="#FFFFFF")
        placeholder_label.pack(pady=10)

    # Update the canvas scroll region
    report_status_canvas.configure(scrollregion=report_status_canvas.bbox("all"))
def gather_reports(): # Scans - Button - Gather Reports - pulls reports from SNow 
    if not progress_file or not project_dir:
        error_label.config(text="Please select a valid progress.json file and project directory.")
        return
    
    client = src.Tenable.api_client()
    owner_id = "0029702"  # my ID, hardcoded
    
    try:
        # Download reports
        src.report_operations.download_reports_for_owner(client, owner_id, project_dir)
        
        # Process reports
        KAIZEN.gather_and_process_reports(project_dir)
        
        # Run PowerShell script in each "Automated" folder
        run_powershell_script_in_automated_folders(project_dir)
        
        refresh_report_status()
        messagebox.showinfo("Success", "Reports gathered, processed, and parsed successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to gather, process, or parse reports: {str(e)}")
def run_powershell_script_in_automated_folders(project_dir): # Scans - Support - Gather Reports - checks for automated folder presence and runs script
    for root, dirs, files in os.walk(project_dir):
        if "Automated" in dirs:
            automated_folder = os.path.join(root, "Automated")
            run_powershell_script(automated_folder)
            organize_output_files(automated_folder)
def run_powershell_script(folder_path): # Scans - Support - Gather Reports - function to actually run the script
    script_path = os.path.join(os.path.dirname(__file__), "SCReport-Parse-Multiple 1.ps1")
    subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", script_path], cwd=folder_path)
def organize_output_files(folder_path): # Scans - Support - Gather Reports - function to organize/sort files
    metadata_folder = os.path.join(folder_path, "MetaData")
    failed_folder = os.path.join(folder_path, "Failed")
    
    os.makedirs(metadata_folder, exist_ok=True)
    os.makedirs(failed_folder, exist_ok=True)
    
    for file in os.listdir(folder_path):
        if file.endswith("_MetaData.csv"):
            os.rename(os.path.join(folder_path, file), os.path.join(metadata_folder, file))
        elif file.endswith("_FailedChecks.csv"):
            os.rename(os.path.join(folder_path, file), os.path.join(failed_folder, file))

def scan_required(scc_info): # check if scan is required for given SCC
    evidence_methods = [method.lower() for method in scc_info.get('Evidence Methods', [])]
    return any(method in ['automated', 'manual-auto info'] for method in evidence_methods)
def download_reports_for_owner_gui():
    if not progress_file or not project_dir:
        error_label.config(text="Please select a valid progress.json file and project directory.")
        return
    
    client = src.Tenable.api_client()
    owner_id = "0029702"
    
    def download_thread():
        src.Tenable.report_operations.download_status_text.delete(1.0, tk.END)
        src.Tenable.report_operations.download_status_text.insert(tk.END, "Downloading reports...\n")
        
        try:
            src.report_operations.download_reports_for_owner(client, owner_id, project_dir)
            src.Tenable.report_operations.download_status_text.insert(tk.END, "Download completed successfully.")
        except Exception as e:
            src.Tenable.download_status_text.insert(tk.END, f"Error downloading reports: {str(e)}")
    
    # Run the download in a separate thread to avoid freezing the GUI
    import threading
    threading.Thread(target=download_thread).start()
def check_reports_collected(scc_name, project_dir):
    info_path = os.path.join(project_dir, scc_name, "Manual", "Automated Info")
    passfall_path = os.path.join(project_dir, scc_name, "Automated")
    
    info_files = glob.glob(os.path.join(info_path, f"{scc_name}-Info*.pdf"))
    passfall_files = glob.glob(os.path.join(passfall_path, f"{scc_name}-PassFail*.csv"))
    
    return len(info_files) > 0 and len(passfall_files) > 0

# General 
def clear_frames(): # General - supports various screen loading
    for frame in (welcome_screen, options_screen, dashboard_screen, scans_screen):
        frame.pack_forget() # hide all specified frames
def sort_docs():
    if supporting_docs_dir and attestation_dir and bpers_dir:
        # Perform sorting actions with the selected directories
        pass
    else:
        error_label.config(text="Please select all required directories.")


############################
##### BEGIN GUI SETUP ######
#root = ThemedTk(theme="equilux")  # Dark Mode, performance impacted
root = ThemedTk(theme="")
root.title("TDL on Easy Mode")
root.geometry("800x600")

## Welcome screen ##############################################################################################
welcome_screen = ttk.Frame(root)
# Welcome label
welcome_label = ttk.Label(welcome_screen, text="Welcome to the TDL wizard! \n\n This program is your companion through the TDL process. \n\n Where would you like to start?")
welcome_label.pack(pady=20)
# New project button
new_project_button = ttk.Button(welcome_screen, text="Start New Project", command=start_new_project)
new_project_button.pack(pady=10)
# Existing project button
existing_project_button = ttk.Button(welcome_screen, text="Update Existing Project", command=update_existing_project)
existing_project_button.pack(pady=10)
## END Welcome Screen #########################################################################################


## Options screen #############################################################################################
options_screen = ttk.Frame(root)
options_screen.pack(fill="both", expand=True)
# Main label
options_label = ttk.Label(options_screen, text="What would you like to do?", font=("Arial", 16, "bold"))
options_label.pack(pady=20)
# Button frame for main actions
button_frame = ttk.Frame(options_screen)
button_frame.pack(pady=20)
# Options - Row 1 - Pull Information 
pull_info_frame = ttk.LabelFrame(button_frame, text="Pull Information", padding=10)
pull_info_frame.pack(side="left", padx=20)
pull_info_button = ttk.Button(pull_info_frame, text="Pull", width=15, command=pull_information)
pull_info_button.pack(pady=5)
pull_info_status = ttk.Label(pull_info_frame, text="Not done")
pull_info_status.pack()
# Options - Row 1 - Build TDL Directories
build_dirs_frame = ttk.LabelFrame(button_frame, text="Build TDL Directories", padding=10)
build_dirs_frame.pack(side="left", padx=20)
build_dirs_button = ttk.Button(build_dirs_frame, text="Build", width=15, command=build_dirs)
build_dirs_button.pack(pady=5)
build_dirs_status = ttk.Label(build_dirs_frame, text="Not done")
build_dirs_status.pack()
# Options - Row 1 - Build Templates
build_templates_frame = ttk.LabelFrame(button_frame, text="Build Templates", padding=10)
build_templates_frame.pack(side="left", padx=20)
build_templates_button = ttk.Button(build_templates_frame, text="Build", width=15, command=build_templates)
build_templates_button.pack(pady=5)
build_templates_status = ttk.Label(build_templates_frame, text="Not done")
build_templates_status.pack()
# Options - Row 1 - Gather and Sort Documents 
gather_docs_frame = ttk.LabelFrame(button_frame, text="Gather and Sort Documents", padding=10)
gather_docs_frame.pack(side="left", padx=20)
gather_docs_button = ttk.Button(gather_docs_frame, text="Gather", width=15, command=gather_docs)
gather_docs_button.pack(pady=5)
gather_docs_status = ttk.Label(gather_docs_frame, text="Not done")
gather_docs_status.pack()
# Options - Row 1 - Generate MD Files
generate_md_frame = ttk.LabelFrame(button_frame, text="Generate MD Files", padding=10)
generate_md_frame.pack(side="left", padx=20)
generate_md_button = ttk.Button(generate_md_frame, text="Generate", width=15, command=generate_md_files)
generate_md_button.pack(pady=5)
generate_md_status = ttk.Label(generate_md_frame, text="Not done")
generate_md_status.pack()
# Options - Row 1 - Update Document Tracker
update_document_validation_frame = ttk.LabelFrame(button_frame, text="Update Document Tracker", padding=10)
update_document_validation_frame.pack(side="left", padx=20)
update_document_validation_button = ttk.Button(update_document_validation_frame, text="Update", width=15, command=update_document_validation)
update_document_validation_button.pack(pady=5)
update_document_validation_status = ttk.Label(update_document_validation_frame, text="Not done")
update_document_validation_status.pack()

# Options - Row 2 - Frame
additional_buttons_frame = ttk.Frame(options_screen)
additional_buttons_frame.pack(pady=20)
# Options - Row 2 - Button - Output Progress
output_progress_button = ttk.Button(additional_buttons_frame, text="Output progress", width=15, command=output_progress)
output_progress_button.pack(side="left", padx=10)
# Options - Row 2 - Button - Add or redo an SCC
add_redo_scc_button = ttk.Button(additional_buttons_frame, text="Add or redo an SCC", width=20, command=add_or_redo_scc)
add_redo_scc_button.pack(side="left", padx=10)
# Options - Row 2 - Button - Remove an SCC
remove_scc_button = ttk.Button(additional_buttons_frame, text="Remove an SCC", width=20, command=remove_scc)
remove_scc_button.pack(side="left", padx=10)
# Options - Row 2 - Button - Sync
sync_button = ttk.Button(additional_buttons_frame, text="Sync", width=15, command=sync_button_click)
sync_button.pack(side="left", padx=10)

# Options - Selected Dirs - Frame
directory_labels_frame = ttk.LabelFrame(options_screen, text="Selected Directories", padding=10)
directory_labels_frame.pack(pady=40)
directory_canvas = tk.Canvas(directory_labels_frame, width=800)
directory_canvas.pack(side="left", fill="both", expand=True)
directory_scrollbar = ttk.Scrollbar(directory_labels_frame, orient="vertical", command=directory_canvas.yview)
directory_scrollbar.pack(side="right", fill="y")
directory_canvas.configure(yscrollcommand=directory_scrollbar.set)
directory_canvas.bind("<Configure>", lambda e: directory_canvas.configure(scrollregion=directory_canvas.bbox("all")))
directory_frame = ttk.Frame(directory_canvas, width=800)
directory_canvas.create_window((0, 0), window=directory_frame, anchor="nw")
# Options - Selected Dirs - BPERs
bpers_frame = ttk.Frame(directory_frame)
bpers_frame.pack(anchor="w", pady=5)
bpers_dir_button = ttk.Button(bpers_frame, text="Select", command=select_bpers_directory)
bpers_dir_button.pack(side="left", padx=10)
bpers_dir_label = ttk.Label(bpers_frame, text="BPERs Directory: Not selected")
bpers_dir_label.pack(side="left")
# Options - Selected Dirs - Attestations
attestation_frame = ttk.Frame(directory_frame)
attestation_frame.pack(anchor="w", pady=5)
attestation_dir_button = ttk.Button(attestation_frame, text="Select", command=select_attestation_directory)
attestation_dir_button.pack(side="left", padx=10)
attestation_dir_label = ttk.Label(attestation_frame, text="Attestation Directory: Not selected")
attestation_dir_label.pack(side="left")
# Options - Selected Dirs - SupDocs
supporting_docs_frame = ttk.Frame(directory_frame)
supporting_docs_frame.pack(anchor="w", pady=5)
supporting_docs_dir_button = ttk.Button(supporting_docs_frame, text="Select", command=select_supporting_docs_directory)
supporting_docs_dir_button.pack(side="left", padx=10)
supporting_docs_dir_label = ttk.Label(supporting_docs_frame, text="Supporting Documents Directory: Not selected")
supporting_docs_dir_label.pack(side="left")
# Options - Selected Dirs - SCCs
scc_frame = ttk.Frame(directory_frame)
scc_frame.pack(anchor="w", pady=5)
scc_dir_button = ttk.Button(scc_frame, text="Select", command=select_scc_directory)
scc_dir_button.pack(side="left", padx=10)
scc_dir_label = ttk.Label(scc_frame, text="SCC Directory: Not selected")
scc_dir_label.pack(side="left")
# Options - Selected Dirs - Progress File
progress_file_frame = ttk.Frame(directory_frame)
progress_file_frame.pack(anchor="w", pady=5)
progress_file_button = ttk.Button(progress_file_frame, text="Select", command=select_progress_file)
progress_file_button.pack(side="left", padx=10)
progress_file_label = ttk.Label(progress_file_frame, text="Progress File: Not selected")
progress_file_label.pack(side="left")
# Options - Selected Dirs - Project Dir
project_dir_frame = ttk.Frame(directory_frame)
project_dir_frame.pack(anchor="w", pady=5)
project_dir_button = ttk.Button(project_dir_frame, text="Select", command=select_project_directory)
project_dir_button.pack(side="left", padx=10)
project_dir_label = ttk.Label(project_dir_frame, text="Project Directory: Not selected")
project_dir_label.pack(side="left")
# Options - Selected Dirs - Templates
template_dir_frame = ttk.Frame(directory_frame)
template_dir_frame.pack(anchor="w", pady=5)
template_dir_button = ttk.Button(template_dir_frame, text="Select", command=select_template_directory)
template_dir_button.pack(side="left", padx=10)
template_dir_label = ttk.Label(template_dir_frame, text="Template Directory: Not selected")
template_dir_label.pack(side="left")

# Options - Buttons - Navigation
dashboard_scans_frame = ttk.Frame(options_screen)
dashboard_scans_frame.pack(pady=10)
dashboard_button = ttk.Button(dashboard_scans_frame, text="Dashboard", width=15, command=show_dashboard)
dashboard_button.pack(side="left", padx=10)
scans_button = ttk.Button(dashboard_scans_frame, text="Scans", width=15, command=show_scans)
scans_button.pack(side="left", padx=10)

# Options - Errors - Error label  
error_label = ttk.Label(root, text="", foreground="red")
error_label.pack(pady=10)
## END Options Screen ############################################################################################################

## Dashboard screen ##############################################################################################################
dashboard_screen = ttk.Frame(root)

# Dashboard - Button - Back button
back_button = ttk.Button(dashboard_screen, text="Back", width=15, command=show_options)
back_button.pack(side="bottom", padx=10, pady=10)

# Dashboard - frame - Section 1 (SCC List) 
section1_frame = ttk.Frame(dashboard_screen)
section1_frame.pack(side="left", fill="both", expand=True)
section1_label = ttk.Label(section1_frame, text="SCC List (Click Me!)", font=("Arial", 12, "bold"))
section1_label.pack(pady=10)
scc_list_frame = ttk.Frame(section1_frame)
scc_list_frame.pack(fill="both", expand=True)
scc_listbox = tk.Listbox(scc_list_frame, font=("Arial", 10), selectmode="single")
scc_listbox.pack(side="left", fill="both", expand=True)
scc_scrollbar = ttk.Scrollbar(scc_list_frame, orient="vertical", command=scc_listbox.yview)
scc_scrollbar.pack(side="right", fill="y")
scc_listbox.config(yscrollcommand=scc_scrollbar.set)
scc_listbox.bind("<Double-Button-1>", open_scc_markdown_file)

# Dashboard - frame - Section 2 (Items Not Gathered)
section2_frame = ttk.Frame(dashboard_screen)
section2_frame.pack(side="left", fill="both", expand=True)
section2_label = ttk.Label(section2_frame, text="Items Not Gathered", font=("Arial", 12, "bold"))
section2_label.pack(pady=10)
# Dashboard - area - Items Not Gathered (Attestations) 
not_gathered_attestations_label = ttk.Label(section2_frame, text="Attestations", font=("Arial", 10, "bold"))
not_gathered_attestations_label.pack(pady=5)
not_gathered_attestations_listbox = tk.Listbox(section2_frame, font=("Arial", 10), selectmode="multiple")
not_gathered_attestations_listbox.pack(fill="both", expand=True)
attestations_buttons_frame = ttk.Frame(section2_frame)
attestations_buttons_frame.pack(pady=5)
mark_attestation_false_positive_button = ttk.Button(attestations_buttons_frame, text="Mark as False Positive", command=lambda: mark_as_false_positive("Attestations"))
mark_attestation_false_positive_button.pack(side="left", padx=5)
manually_link_attestations_button = ttk.Button(attestations_buttons_frame, text="Assign Match", command=lambda: manually_link_files("Attestations"))
manually_link_attestations_button.pack(side="left", padx=5)
# Dashboard - area - Items Not Gathered (BPERs)
not_gathered_bpers_label = ttk.Label(section2_frame, text="BPERs", font=("Arial", 10, "bold"))
not_gathered_bpers_label.pack(pady=5)
not_gathered_bpers_listbox = tk.Listbox(section2_frame, font=("Arial", 10), selectmode="multiple")
not_gathered_bpers_listbox.pack(fill="both", expand=True)
bpers_buttons_frame = ttk.Frame(section2_frame)
bpers_buttons_frame.pack(pady=5)
mark_bper_false_positive_button = ttk.Button(bpers_buttons_frame, text="Mark as False Positive", command=lambda: mark_as_false_positive("BPERs"))
mark_bper_false_positive_button.pack(side="left", padx=5)
manually_link_bpers_button = ttk.Button(bpers_buttons_frame, text="Assign Match", command=lambda: manually_link_files("BPERs"))
manually_link_bpers_button.pack(side="left", padx=5)
# Dashboard - area - Items Not Gathered (Documents)
not_gathered_documents_label = ttk.Label(section2_frame, text="Documents", font=("Arial", 10, "bold"))
not_gathered_documents_label.pack(pady=5)
not_gathered_documents_listbox = tk.Listbox(section2_frame, font=("Arial", 10), selectmode="multiple")
not_gathered_documents_listbox.pack(fill="both", expand=True)
documents_buttons_frame = ttk.Frame(section2_frame)
documents_buttons_frame.pack(pady=5)
mark_document_false_positive_button = ttk.Button(documents_buttons_frame, text="Mark as False Positive", command=lambda: mark_as_false_positive("Documents"))
mark_document_false_positive_button.pack(side="left", padx=5)
manually_link_documents_button = ttk.Button(documents_buttons_frame, text="Assign Match", command=lambda: manually_link_files("Documents"))
manually_link_documents_button.pack(side="left", padx=5)

# Dashboard - area- Section 3 (Dates and Chart)
section3_frame = ttk.Frame(dashboard_screen)
section3_frame.pack(side="left", fill="both", expand=True)
section3_label = ttk.Label(section3_frame, text="Dates and Chart", font=("Arial", 12))
section3_label.pack(pady=10)
last_info_pull_label = ttk.Label(section3_frame, text="Last Info Pull: N/A", font=("Arial", 10))
last_info_pull_label.pack(pady=5)
last_doc_pull_label = ttk.Label(section3_frame, text="Last Doc Pull: N/A", font=("Arial", 10))
last_doc_pull_label.pack(pady=5)
last_checklist_generated_label = ttk.Label(section3_frame, text="Last Checklist Generated: N/A", font=("Arial", 10))
last_checklist_generated_label.pack(pady=5)

# Placeholder for the pie chart
pie_chart_label = ttk.Label(section3_frame, text="", font=("Arial", 10), justify="center")
pie_chart_label.pack(pady=10)
## END Dashboard Screen #######################################################################################################################


## Scans screen ###############################################################################################################################
scans_screen = ttk.Frame(root)

# Scans - frame
panes_frame = ttk.Frame(scans_screen)
panes_frame.pack(fill="both", expand=True, padx=20, pady=10)

# Scans - area - Left pane (Inventory Information)
left_pane = ttk.Frame(panes_frame, width=250)
left_pane.pack(side="left", fill="both", expand=True, padx=(0, 10))
inventory_content = ttk.Frame(left_pane)
inventory_content.pack(fill="both", expand=True)
inventory_label = ttk.Label(inventory_content, text="Inventories", font=("Arial", 14, "bold"))
inventory_label.pack(pady=10)
inventory_canvas = tk.Canvas(inventory_content)
inventory_canvas.pack(side="left", fill="both", expand=True)
inventory_scrollbar = ttk.Scrollbar(inventory_content, orient="vertical", command=inventory_canvas.yview)
inventory_scrollbar.pack(side="right", fill="y")
inventory_canvas.configure(yscrollcommand=inventory_scrollbar.set)
inventory_canvas.bind("<Configure>", lambda e: inventory_canvas.configure(scrollregion=inventory_canvas.bbox("all")))
inventory_frame = ttk.Frame(inventory_canvas)
inventory_canvas.create_window((0, 0), window=inventory_frame, anchor="nw")
# Scans - button - check inventories
check_inventories_button = ttk.Button(left_pane, text="Check Inventories", width=20, command=check_inventories)
check_inventories_button.pack(side="bottom", pady=10)

# Scans -area - Middle pane (Scan Status)
middle_pane = ttk.Frame(panes_frame, width=250)
middle_pane.pack(side="left", fill="both", expand=True, padx=10)
scan_content = ttk.Frame(middle_pane)
scan_content.pack(fill="both", expand=True)
scan_status_label = ttk.Label(scan_content, text="Scan Status", font=("Arial", 14, "bold"))
scan_status_label.pack(pady=10)
scan_status_canvas = tk.Canvas(scan_content)
scan_status_canvas.pack(side="left", fill="both", expand=True)
scan_status_scrollbar = ttk.Scrollbar(scan_content, orient="vertical", command=scan_status_canvas.yview)
scan_status_scrollbar.pack(side="right", fill="y")
scan_status_canvas.configure(yscrollcommand=scan_status_scrollbar.set)
scan_status_canvas.bind("<Configure>", lambda e: scan_status_canvas.configure(scrollregion=scan_status_canvas.bbox("all")))
scan_status_frame = ttk.Frame(scan_status_canvas)
scan_status_canvas.create_window((0, 0), window=scan_status_frame, anchor="nw")
# Scans - button - initiate scans
initiate_scans_button = ttk.Button(middle_pane, text="Initiate Scans", width=20, command=initiate_scans)
initiate_scans_button.pack(side="bottom", pady=10)

# Scans - area - Right pane (Report Status)
right_pane = ttk.Frame(panes_frame, width=250)
right_pane.pack(side="left", fill="both", expand=True, padx=(10, 0))
report_content = ttk.Frame(right_pane)
report_content.pack(fill="both", expand=True)
report_status_label = ttk.Label(report_content, text="Report Status", font=("Arial", 14, "bold"))
report_status_label.pack(pady=10)
report_status_canvas = tk.Canvas(report_content)
report_status_canvas.pack(side="left", fill="both", expand=True)
report_status_scrollbar = ttk.Scrollbar(report_content, orient="vertical", command=report_status_canvas.yview)
report_status_scrollbar.pack(side="right", fill="y")
report_status_canvas.configure(yscrollcommand=report_status_scrollbar.set)
report_status_canvas.bind("<Configure>", lambda e: report_status_canvas.configure(scrollregion=report_status_canvas.bbox("all")))
report_status_frame = ttk.Frame(report_status_canvas)
report_status_canvas.create_window((0, 0), window=report_status_frame, anchor="nw")
# Scans - button - refresh report status
button_frame = ttk.Frame(right_pane)
button_frame.pack(side="bottom", pady=10)
refresh_report_status_button = ttk.Button(button_frame, text="Refresh Report Status", width=20, command=refresh_report_status)
refresh_report_status_button.pack(side="left", padx=5)
# Scans - button - gather reports
gather_reports_button = ttk.Button(button_frame, text="Gather Reports", width=20, command=gather_reports)
gather_reports_button.pack(side="left", padx=5)

# Scans - button - back 
back_button = ttk.Button(scans_screen, text="Back", width=15, command=show_options)
back_button.pack(side="bottom", padx=10, pady=10)
## END Scans Screen ############################################################################################################################

#######################
#### START THE GUI ####
show_welcome()
root.mainloop()
#######################
