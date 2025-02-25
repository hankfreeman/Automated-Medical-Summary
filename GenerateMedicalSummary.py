import os
import json
import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.shared import Pt, Inches
from collections import defaultdict
import time
import anthropic
import re
from collections import defaultdict
import win32com.client
from datetime import datetime
import docx
from docx import Document
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
from tkinter import ttk


# Initialize the AI client globally
anthropic_client = anthropic.Anthropic(
    api_key=""
)

import psutil
import win32com.client
import pythoncom
import win32api
import time
from contextlib import contextmanager

def kill_office_processes():
    """Kill all running Office-related processes that might interfere with automation."""
    office_processes = [
        "WINWORD.EXE",      # Microsoft Word
        "EXCEL.EXE",        # Microsoft Excel
        "OUTLOOK.EXE",      # Microsoft Outlook
        "POWERPNT.EXE",     # Microsoft PowerPoint
        "MSPUB.EXE",        # Microsoft Publisher
        "MSACCESS.EXE",     # Microsoft Access
        "ONENOTE.EXE",      # Microsoft OneNote
        "GROOVE.EXE",       # Microsoft SharePoint Workspace
        "dllhost.exe",      # COM Surrogate processes
        "Microsoft.Office*"  # Any other Office processes
    ]
    
    killed_processes = []
    for proc in psutil.process_iter(['pid', 'name', 'username']):
        try:
            # Check if process name matches any Office process
            if any(proc.info['name'].upper().startswith(p.upper()) for p in office_processes):
                proc.kill()
                killed_processes.append(proc.info['name'])
                print(f"✅ Killed process: {proc.info['name']} (PID: {proc.info['pid']})")
                time.sleep(0.1)  # Give system time to clean up
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess) as e:
            print(f"⚠️ Could not kill process: {str(e)}")
    
    return killed_processes

def cleanup_com_objects():
    """Clean up COM objects and reset COM system."""
    try:
        pythoncom.CoUninitialize()  # Uninitialize the current thread's COM
        time.sleep(0.5)  # Give system time to clean up
        pythoncom.CoInitialize()    # Reinitialize COM for the current thread
        print("✅ COM objects cleaned up successfully")
    except Exception as e:
        print(f"⚠️ Error cleaning up COM objects: {str(e)}")

@contextmanager
def word_cleanup_context():
    """Context manager for proper Word application cleanup."""
    word_app = None
    try:
        # Kill any existing Word processes first
        kill_office_processes()
        cleanup_com_objects()
        
        # Create new Word instance
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        yield word_app
        
    finally:
        if word_app:
            try:
                word_app.Quit()
                del word_app
            except:
                pass
            
        # Final cleanup
        cleanup_com_objects()
        time.sleep(1)  # Give system time for final cleanup

def check_system_resources():
    """Check system resources that might affect Word automation."""
    try:
        # Check memory usage
        memory = psutil.virtual_memory()
        print(f"Memory Usage: {memory.percent}%")
        if memory.percent > 90:
            print("⚠️ Warning: High memory usage may affect Word automation")

        # Check CPU usage
        cpu_percent = psutil.cpu_percent(interval=1)
        print(f"CPU Usage: {cpu_percent}%")
        if cpu_percent > 80:
            print("⚠️ Warning: High CPU usage may affect Word automation")

        # Check disk usage
        disk = psutil.disk_usage('/')
        print(f"Disk Usage: {disk.percent}%")
        if disk.percent > 90:
            print("⚠️ Warning: Low disk space may affect Word automation")

        # List active Office-related processes
        office_processes = []
        for proc in psutil.process_iter(['pid', 'name']):
            if "WINWORD" in proc.info['name'].upper() or "OFFICE" in proc.info['name'].upper():
                office_processes.append(proc.info)
        
        if office_processes:
            print("\nActive Office processes:")
            for proc in office_processes:
                print(f"- {proc['name']} (PID: {proc['pid']})")

    except Exception as e:
        print(f"Error checking system resources: {str(e)}")

# Example usage:
# with word_cleanup_context() as word_app:
#     word_doc = word_app.Documents.Add()
#     # ... work with Word document ...
#     word_doc.Save()
#     word_doc.Close()

def close_existing_word_instances():
    """Ensures all Microsoft Word instances are properly closed before starting a new one."""
    for process in psutil.process_iter(attrs=['pid', 'name']):
        if process.info['name'] and "WINWORD.EXE" in process.info['name']:
            try:
                process.kill()  # Forcefully terminate Word
                print(f"✅ Closed lingering Word process: {process.info['pid']}")
            except Exception as e:
                print(f"⚠️ Could not close Word process {process.info['pid']}: {e}")

# Run this before starting Word
close_existing_word_instances()

def browse_project_folder():
    """Allow user to select a project folder."""
    global project_folder, json_folder, output_word_doc, master_json_file
    project_folder = filedialog.askdirectory(title="Select Project Folder")
    
    if not project_folder:
        return  # User canceled selection

    # Dynamically set the paths
    json_folder = os.path.join(project_folder, "Work Product", "ai-analyses")
    output_word_doc = os.path.join(json_folder, "Medical_Summary.docx")
    master_json_file = os.path.join(json_folder, "master_json_output.json")

    # Update GUI display
    folder_label.config(text=f"Selected: {project_folder}")

import time

def call_ai_with_retry(messages, max_retries=3):
    """Call AI API with automatic retries on overload."""
    for attempt in range(max_retries):
        try:
            response = anthropic_client.messages.create(
                model="claude-3-5-sonnet-latest",
                max_tokens=4096,
                temperature=0,
                messages=messages
            )
            return response.content[0].text.strip()
        
        except anthropic.errors.APIError as e:
            print(f"[WARNING] AI API error: {e}")
            if "overloaded" in str(e).lower():
                wait_time = 5 * (attempt + 1)  # Exponential backoff
                print(f"[INFO] Retrying AI call in {wait_time} seconds...")
                time.sleep(wait_time)
            else:
                raise  # Raise non-overload errors

    raise RuntimeError("❌ AI failed after multiple retries due to overload.")



def update_progress(step_index):
    """Update the progress bar and status message."""
    progress_var.set((step_index + 1) / len(progress_steps) * 100)
    status_label.config(text=progress_steps[step_index])
    root.update_idletasks()

def start_processing():
    """Start the medical summary process in a separate thread."""
    event_date = event_date_entry.get().strip()  # Ensure correct input

    if not re.match(r'^\d{1,2}/\d{1,2}/\d{4}$', event_date):
        messagebox.showerror("Error", "Invalid event date format! Please use MM/DD/YYYY.")
        return

    print(f"[DEBUG] Event Date Set from GUI: {event_date}")  # ✅ Ensure event_date is being captured


    if not project_folder:
        messagebox.showerror("Error", "Please select a project folder!")
        return

    if not event_date:
        messagebox.showerror("Error", "Please enter the event date!")
        return
        
    # Disable buttons during processing
    browse_button.config(state=tk.DISABLED)
    start_button.config(state=tk.DISABLED)

    # Run the processing in a separate thread to keep GUI responsive
    threading.Thread(target=run_medical_summary, args=(event_date,), daemon=True).start()

import re

import re
import json

def create_provider_name_mapping_via_ai(provider_names):
    """Uses AI to return a dictionary mapping raw provider names to their standardized versions."""
    
    unique_providers = list(set(provider_names))  # Get unique names to condense

    messages = [
        {
            "role": "user",
            "content": (
                "Standardize the following provider names into their most accurate versions.\n\n"
                "### INPUT PROVIDER NAMES ###\n"
                f"{json.dumps(unique_providers, indent=2)}\n\n"
                "### REQUIRED OUTPUT FORMAT ###\n"
                "{\n"
                '    "Raw Provider Name 1": "Standardized Provider Name 1",\n'
                '    "Raw Provider Name 2": "Standardized Provider Name 2",\n'
                '    ...\n'
                "}\n\n"
                "### RULES ###\n"
                "- Keep the most official name.\n"
                "- Remove unnecessary titles (PLLC, LLC, DBA, etc.).\n"
                "- Ensure consistency in naming across encounters.\n"
                "- Return **ONLY** the JSON object, no extra text."
            )
        }
    ]

    try:
        # ✅ Use AI call with retry logic to prevent overload errors
        ai_response = call_ai_with_retry(messages)

        print("\n[DEBUG] Raw AI Response:")
        print(ai_response)

        # Direct JSON parsing attempt
        try:
            provider_mapping = json.loads(ai_response)
            if isinstance(provider_mapping, dict):
                print("\n[DEBUG] Successfully parsed AI response as JSON.")
                return provider_mapping
            else:
                raise ValueError("Parsed data is not a dictionary.")

        except json.JSONDecodeError:
            print("[ERROR] AI did not return valid JSON. Trying regex extraction...")

            # Regex extraction attempt (to handle cases where AI returns text before JSON)
            json_match = re.search(r'\{(?:[^{}]|(?R))*\}', ai_response, re.DOTALL)
            if json_match:
                json_string = json_match.group()
                try:
                    provider_mapping = json.loads(json_string)
                    print("\n[DEBUG] Successfully extracted JSON from AI response.")
                    return provider_mapping
                except json.JSONDecodeError as e:
                    print(f"\n[ERROR] Failed to parse extracted JSON: {e}")

        print(f"[ERROR] AI response format is incorrect: {type(ai_response)}")
        return {}  # Return an empty dictionary on failure

    except Exception as e:
        print(f"[ERROR] AI provider mapping request failed: {e}")
        return {}





def consolidate_provider_names_via_ai(provider_names):
    """Uses AI to normalize a list of provider names and return a mapping of raw → standardized names."""

    unique_providers = list(set(provider_names))  # Remove duplicates

    messages = [
        {
            "role": "user",
            "content": (
                "Standardize and condense these provider names into the best possible representations:\n\n"
                f"{json.dumps(unique_providers, indent=2)}\n\n"
                "### OUTPUT FORMAT ###\n"
                "Return a JSON dictionary mapping raw provider names to the best standardized name:\n"
                "{\n"
                '  "Raw Provider Name 1": "Best Standardized Name",\n'
                '  "Raw Provider Name 2": "Best Standardized Name",\n'
                "  ...\n"
                "}\n"
                "### RULES ###\n"
                "- Preserve provider details accurately.\n"
                "- Remove unnecessary business suffixes (LLC, PLLC, DBA, etc.).\n"
                "- If the name contains a person's name (e.g., 'Ajay Aggarwal, M.D.'), retain it correctly.\n"
                "- Ensure consistency in naming across all encounters.\n"
                "Return ONLY the JSON object, with no additional text."
            )
        }
    ]

    try:
        response = anthropic_client.messages.create(
            model="claude-3-5-sonnet-latest",
            max_tokens=500,
            temperature=0,
            messages=messages
        )

        ai_response = response.content[0].text.strip()

        # Try parsing AI response as JSON
        try:
            cleaned_mapping = json.loads(ai_response)
            if isinstance(cleaned_mapping, dict):
                return cleaned_mapping  # ✅ Successfully parsed dictionary

        except json.JSONDecodeError:
            print(f"[ERROR] AI provider response is not valid JSON: {ai_response}")

        return {name: name for name in unique_providers}  # Fail gracefully by returning original names

    except Exception as e:
        print(f"[ERROR] AI provider consolidation failed: {e}")
        return {name: name for name in unique_providers}  # Return original names if AI fails



def normalize_provider_name(provider_name, provider_mapping, cache={}):
    """Uses AI-generated mapping to clean provider names efficiently."""
    
    if not provider_name or provider_name == "Unknown Provider":
        return "Unknown Provider"

    # Use cached result if available
    if provider_name in cache:
        return cache[provider_name]

    # Try to get the cleaned name from the AI-generated mapping
    cleaned_name = provider_mapping.get(provider_name, provider_name)  # Default to original if not found
    
    # Cache and return the cleaned name
    cache[provider_name] = cleaned_name
    return cleaned_name



def condense_encounters_via_ai(provider_date_groups):
    """Uses AI to condense encounters, ensuring normalized provider names are used."""
    condensed_encounters = []

    for (provider, date), encounters in provider_date_groups.items():
        if not encounters:
            continue

        print(f"[INFO] Processing encounters for {provider} on {date}")

        # Prepare data for AI processing
        cleaned_encounters = clean_json_data(encounters)
        encounters_json = json.dumps(cleaned_encounters, ensure_ascii=False, indent=2)

        try:
            messages = [
                {
                    "role": "user",
                    "content": (
                        f"Condense these medical encounters from {provider} on {date} into a single summary.\n\n"
                        f"### RAW ENCOUNTERS ###\n{encounters_json}\n\n"
                        "### REQUIRED OUTPUT FORMAT ###\n"
                        "Return a JSON object with this structure:\n"
                        "{\n"
                        f'    "Date of Service": "{date}",\n'
                        f'    "Provider/Facility Name": "{provider}",\n'
                        '    "Type of Record": "string",\n'
                        '    "Notes/Summary": "string",\n'
                        '    "Source File": "string",\n'
                        '    "Source Page": "number or string"\n'
                        "}\n"
                        "Return only the JSON object, no additional text."
                        "### RULES ###\n"
                        "1. Preserve provider names as accurately as possible.\n"
                        "2. If multiple providers exist for the same date, DO NOT replace them with a single different provider.\n"
                        "3. If there are variations of the same provider name (e.g., 'Dr. Ajay Aggarwal MD' vs. 'Ajay Aggarwal'), merge them into the most complete and formal version.\n"
                        "4. If an encounter belongs to a specific provider, it MUST remain under that provider in the final output.\n"

                    )
                }
            ]

            response = anthropic_client.messages.create(
                model="claude-3-5-sonnet-latest",
                max_tokens=4096,
                temperature=0,
                messages=messages
            )

            ai_response = response.content[0].text.strip()

            try:
                condensed_data = json.loads(ai_response)
                if isinstance(condensed_data, dict):
                    condensed_encounters.append(condensed_data)
                    continue

            except json.JSONDecodeError:
                print(f"[ERROR] Failed to parse AI response for {provider} on {date}")

        except Exception as e:
            print(f"[ERROR] AI processing failed for {provider} on {date}: {str(e)}")

    return condensed_encounters



def run_medical_summary(event_date):
    """Main function to process encounters and generate the medical summary."""
    try:
        # Load JSON
        master_json, sorted_dates_list = load_json_files()

        # ✅ Step 1: Extract All Providers & Get AI-Summarized Name
        all_providers = [e.get("Provider/Facility Name", "Unknown Provider") for e in master_json["patient_summary"]["Post-Event Medical History"]["Encounters"]]

        # ✅ Step 2: Create AI-based provider name mapping
        provider_mapping = create_provider_name_mapping_via_ai(all_providers)

        # ✅ Step 3: Apply Provider Mapping to Encounters
        for encounter in master_json["patient_summary"]["Post-Event Medical History"]["Encounters"]:
            provider_raw = encounter.get("Provider/Facility Name", "Unknown Provider").strip()
            encounter["Provider/Facility Name"] = provider_mapping.get(provider_raw, provider_raw)

        # ✅ Step 4: Group encounters by date while using standardized provider names
        date_groups = group_encounters_by_date(master_json, sorted_dates_list, provider_mapping)

        # ✅ Step 5: Condense encounters by date
        condensed_encounters = condense_encounters_via_ai(date_groups)

        # ✅ Step 6: AI Deduplicate Categories
        cleaned_categories = deduplicate_categories_via_ai(master_json["patient_summary"]["Post-Event Medical History"])

        # ✅ Step 7: Create Final JSON
        master_json = create_final_master_json(
            condensed_encounters,
            cleaned_categories,
            master_json["patient_summary"]["Post-Event Medical History"].get("Records Reviewed", [])
        )

        # ✅ Step 8: Save Master JSON
        with open(master_json_file, "w", encoding="utf-8") as f:
            json.dump(master_json, f, indent=4)

        # ✅ Step 9: Generate Word Report (Pass provider_mapping)
        create_medical_summary(output_word_doc, master_json, event_date, provider_mapping)

        print("✅ Medical Summary successfully generated!")

    except Exception as e:
        print(f"❌ [ERROR] {str(e)}")




VALID_RECORD_TYPES = [
"Consultation", "Discharge Summary", "Emergency Department", "EMS",
"History & Physical", "Hospital", "Imaging/Diagnostic", "Incident Report",
"Initial Evaluation", "Office Visit", "Operative Report", "Phone Call/Email",
"Police Report", "Procedure", "Progress Note", "Telemedicine Visit",
"Workers' Compensation Exam", "Other"
]

def convert_date_for_comparison(date_str):
    """Convert a date string to a comparable format (YYYY, MM, DD tuple)."""
    if not date_str or date_str.strip() == "" or date_str.lower() == "unknown date":
        return None  # Return None if the date is empty or unknown
    
    try:
        month, day, year = map(int, date_str.split('/'))  # Ensure MM/DD/YYYY format
        return (year, month, day)  # Return a tuple for easy comparison
    except ValueError:
        print(f"[ERROR] Invalid date format encountered: {date_str}")
        return None



def is_pre_event(date_str, event_date):
    if not event_date:
        raise ValueError("event_date is not provided!")  # Prevent undefined errors
    if not date_str or date_str.lower() == "unknown date":
        return False  # Default unknown dates to post-event

    comparison_date = convert_date_for_comparison(date_str)
    event_date_comparison = convert_date_for_comparison(event_date)  # Ensure `event_date` is updated correctly

    print(f"[DEBUG] Running is_pre_event(): {date_str} ({comparison_date}) vs event date {event_date} ({event_date_comparison})")

    if comparison_date and event_date_comparison:
        result = comparison_date < event_date_comparison  # Compare tuples correctly
        print(f"[DEBUG] is_pre_event() result: {result}")
        return result

    return False




def process_item(word_doc, item, key, encounters, footnote_references, footnote_counter):
    """Process a single item for the medical summary."""
    # Split date and text if present
    parts = item.split(", ", 1)
    if len(parts) == 2:
        date_str = convert_to_long_date(parts[0])
        text = parts[1]
    else:
        date_str = "Unknown Date"
        text = parts[0]

    # Find source details
    source_file = "Unknown Source"
    provider = "Unknown Provider"
    page_number = "Unknown"

    for encounter in encounters:
        if text in str(encounter.get(key, [])):
            source_file = str(encounter.get("Source File", "Unknown Source"))
            provider = str(encounter.get("Provider/Facility Name", "Unknown Provider"))
            page_number = str(encounter.get("Source Page", "Unknown"))
            break

    # Format text based on category
    if key == "Diagnoses":
        formatted_text = text
    elif key == "Medications":
        formatted_text = format_medication(text)
    else:
        formatted_text = f"{date_str}, {text}"

    # Create a new paragraph for each item
    range_end = word_doc.Range()
    range_end.Collapse(0)  # Move to end
    
    # 1. Insert text WITH the space
    range_end.Text = f"\t• {formatted_text} "
    
    # 2. Calculate position before the space
    text_without_space = f"\t• {formatted_text}"
    footnote_position = range_end.Start + len(text_without_space)
    footnote_range = word_doc.Range(footnote_position, footnote_position)
    
    # 3. Add footnote
    footnote_text = f"{provider}, {source_file}, {date_str}, Page {page_number}."
    if footnote_text not in footnote_references:
        footnote = word_doc.Footnotes.Add(footnote_range, str(footnote_counter))
        footnote.Range.Text = footnote_text
        footnote_references[footnote_text] = footnote_counter
        footnote_counter += 1
    
    # 4. Add paragraph break
    range_end.InsertParagraphAfter()
    
    # Return the updated counter
    return footnote_counter

def load_json_files():
    """Load and consolidate all JSON files from the ai-analyses folder, retaining sources for each item."""
    master_json = {
        "patient_summary": {
            "Pre-Event Medical History": defaultdict(list),
            "Post-Event Medical History": {
                "Encounters": [],
                "Diagnoses": [],
                "Medications": [],
                "Imaging/Diagnostics": [],
                "Procedures": [],
                "Rehabilitation": [],
                "Work Status/Restrictions": [],
                "Workers' Compensation Records": [],
                "Disability Applications/Awards": [],
                "Records Reviewed": []
            }
        }
    }
    
    records_reviewed = defaultdict(list)  # ✅ Track each source file and its date range
    unique_dates = set()

    for json_file in os.listdir(json_folder):
        if json_file.endswith(".json"):
            json_path = os.path.join(json_folder, json_file)
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Aggregate data into master structure
            for page_num, page_data in data.get("page_analyses", {}).items():
                for entry in page_data.get("entries", []):
                    master_json["patient_summary"]["Post-Event Medical History"]["Encounters"].append(entry)

                    # Extract relevant metadata
                    date_of_service = entry.get("Date of Service", "Unknown Date")
                    provider = entry.get("Provider/Facility Name", "Unknown Provider")
                    source_file = entry.get("Source File", "Unknown Source")
    
                    if date_of_service != "Unknown Date" and source_file != "Unknown Source":
                        records_reviewed[source_file].append(date_of_service)

                    if date_of_service != "Unknown Date":
                        unique_dates.add(date_of_service)

                    # Categorizing structured data with sources
                    for category in ["Diagnoses", "Medications", "Imaging/Diagnostics", "Procedures", "Rehabilitation"]:
                        if category in entry and entry[category]:
                            for item in entry[category]:
                                # Construct source information
                                provider = entry.get("Provider/Facility Name", "").strip() or "Unknown Provider"
                                record_type = entry.get("Type of Record", "Unknown Record Type")
                                date_of_service = entry.get("Date of Service", "Unknown Date")
                                source_file = entry.get("Source File", "Unknown Source")
                                page_number = entry.get("Source Page", "Unknown Page")
                                
                                source_info = f"{provider}, {record_type}, {date_of_service}, {source_file}, Page {page_number}"
                                
                                if source_info.strip() == "Unknown Provider, Unknown Record Type, Unknown Date, Unknown Source, Unknown Page":
                                    print(f"⚠️ Skipping invalid source info for {item}: {source_info}")
                                    continue  # Skip if the source info is empty or invalid
                                
                                # Find existing entry
                                found = False
                                for existing_item in master_json["patient_summary"]["Post-Event Medical History"][category]:
                                    if existing_item["text"].lower().strip() == item.lower().strip():  # ✅ Case-insensitive matching
                                        existing_item["sources"].add(source_info)
                                        found = True
                                        print(f"✅ Updated sources for {item}: {existing_item['sources']}")
                                        break
                                
                                # If new item, add it with source
                                if not found:
                                    master_json["patient_summary"]["Post-Event Medical History"][category].append(
                                        {"text": item.strip(), "sources": {source_info}}
                                    )
                                    print(f"🆕 New {category} entry added: {item.strip()} with sources: {source_info}")


    # Format records reviewed

    for source, dates in records_reviewed.items():
        sorted_dates = sorted(set(dates))
        date_range = f"{sorted_dates[0]} - {sorted_dates[-1]}" if len(sorted_dates) > 1 else sorted_dates[0]
        master_json["patient_summary"]["Post-Event Medical History"]["Records Reviewed"].append({
            "Source File": source,
            "Date Range": date_range
        })

    for entry in page_data.get("entries", []):
        print(f"[DEBUG] Processing Encounter: {entry}")  # Ensure ALL encounters are processed
        if "Date of Service" not in entry or "Provider/Facility Name" not in entry:
            print(f"⚠️ Skipping encounter due to missing data: {entry}")
            continue
        master_json["patient_summary"]["Post-Event Medical History"]["Encounters"].append(entry)


    
    sorted_dates_list = sorted(unique_dates)
    return master_json, sorted_dates_list

def get_deduplication_examples(category):
    """Get category-specific examples"""
    examples = {
        "Diagnoses": [
            {"text": "Cervical Disc Herniation", "date": "12/19/2020"},
            {"text": "Neck Pain", "date": "12/19/2020"},
            {"text": "Cervical Radiculopathy", "date": "12/19/2020"}
        ],
        "Procedures": [
            {"text": "Cervical ESI", "date": "01/16/2021"},
            {"text": "Cervical Epidural Steroid Injection", "date": "01/16/2021"}
        ],
        "Medications": [
            {"text": "Lidocaine 1%", "date": "01/16/2021"},
            {"text": "lidocaine 1 percent", "date": "01/16/2021"}
        ]
    }
    return json.dumps(examples.get(category, []), indent=2)

def format_medication(text):
    """Format medication text, removing Unknown fields."""
    if "Unknown" in text:
        parts = text.split(',')
        # Keep only the medication name if everything else is Unknown
        return parts[0].strip()
    
    # If we have valid information, keep the full formatted text
    return text


def group_encounters_by_date(master_json, sorted_dates_list, provider_mapping):
    """Groups encounters by Date of Service while ensuring correct provider names."""
    date_groups = defaultdict(list)

    for encounter in master_json["patient_summary"]["Post-Event Medical History"]["Encounters"]:
        date_of_service = encounter.get("Date of Service", "Unknown Date").strip()
        provider_raw = encounter.get("Provider/Facility Name", "Unknown Provider").strip()

        # ✅ Normalize provider name using AI-generated mapping
        provider = normalize_provider_name(provider_raw, provider_mapping)

        # ✅ Store corrected provider name in encounter
        encounter["Provider/Facility Name"] = provider

        # ✅ Ensure date is valid before grouping
        if date_of_service in sorted_dates_list:
            date_groups[(provider, date_of_service)].append(encounter)

    return date_groups






def extract_date_and_text(item):
    """Extract date and text from an item, handling various date formats and positions."""
    if not isinstance(item, str):
        return None, item

    # Common date patterns
    patterns = [
        r'(\d{1,2}/\d{1,2}/\d{2,4})',  # MM/DD/YY or MM/DD/YYYY
        r'on (\d{1,2}/\d{1,2}/\d{2,4})',  # "on MM/DD/YYYY"
        r'(\d{1,2}/\d{1,2}/\d{2,4})$'  # date at end of string
    ]

    text = item
    found_date = None

    # Try each pattern
    for pattern in patterns:
        match = re.search(pattern, item)
        if match:
            found_date = match.group(1)
            # Remove the date and any surrounding text like "on" from the text
            text = re.sub(f"(on )?{re.escape(found_date)}", "", text)
            break

    # Clean up the text
    text = text.strip().strip(',').strip()
    return found_date, text

def clean_json_data(data):
    """Clean and sanitize JSON data to remove invalid characters."""
    if isinstance(data, dict):
        return {k: clean_json_data(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [clean_json_data(x) for x in data]
    elif isinstance(data, str):
        # Remove control characters and normalize whitespace
        cleaned = "".join(char for char in data if ord(char) >= 32 or char in "\n\r\t")
        cleaned = " ".join(cleaned.split())
        return cleaned
    else:
        return data

def create_final_master_json(condensed_encounters, cleaned_categories, records_reviewed):
    """Merges AI-condensed encounters and AI-cleaned medical categories into a final structured JSON."""
    master_json = {
        "patient_summary": {
            "Pre-Event Medical History": defaultdict(list),
            "Post-Event Medical History": {
                "Encounters": condensed_encounters,
                "Diagnoses": cleaned_categories.get("Diagnoses", []),
                "Medications": cleaned_categories.get("Medications", []),
                "Imaging/Diagnostics": cleaned_categories.get("Imaging/Diagnostics", []),
                "Procedures": cleaned_categories.get("Procedures", []),
                "Rehabilitation": cleaned_categories.get("Rehabilitation", []),
                "Work Status/Restrictions": cleaned_categories.get("Work Status/Restrictions", []),
                "Workers' Compensation Records": cleaned_categories.get("Workers' Compensation Records", []),
                "Disability Applications/Awards": cleaned_categories.get("Disability Applications/Awards", []),
                "Records Reviewed": records_reviewed
            }
        }
    }

    return master_json

def convert_to_long_date(date_str):
    """Convert a date string from MM/DD/YYYY to long format (Month DD, YYYY)."""
    if not date_str or date_str.strip() == "" or date_str.lower() == "unknown date":
        return "Unknown Date"
    
    try:
        # Handle date ranges (e.g., "12/19/2020 - 02/13/2021")
        if ' - ' in date_str:
            start_date, end_date = date_str.split(' - ')
            long_start = convert_to_long_date(start_date.strip())
            long_end = convert_to_long_date(end_date.strip())
            return f"{long_start} - {long_end}"
        
        # Parse the date
        month_map = {
            1: 'January', 2: 'February', 3: 'March', 4: 'April',
            5: 'May', 6: 'June', 7: 'July', 8: 'August',
            9: 'September', 10: 'October', 11: 'November', 12: 'December'
        }
        
        # Split the date components safely
        month, day, year = map(int, date_str.split('/'))
        
        # Format the long date
        return f"{month_map[month]} {day}, {year}"
    except Exception:
        return "Unknown Date"


def parse_date(date_str):
    """Convert MM/DD/YYYY string to a comparable date object; return None for invalid dates."""
    try:
        return datetime.strptime(date_str, "%m/%d/%Y")
    except (ValueError, TypeError):
        return None  # Return None for invalid dates

def create_medical_summary(output_path, master_json, event_date, provider_mapping):
    """Creates a structured Word document for the medical summary with footnotes using win32com."""

    VALID_RECORD_TYPES = [
        "Consultation", "Discharge Summary", "Emergency Department", "EMS",
        "History & Physical", "Hospital", "Imaging/Diagnostic", "Incident Report",
        "Initial Evaluation", "Office Visit", "Operative Report", "Phone Call/Email",
        "Police Report", "Procedure", "Progress Note", "Telemedicine Visit",
        "Workers' Compensation Exam", "Other"
    ]

    category_titles = {
        "Diagnoses": "Diagnoses",
        "Medications": "Medications",
        "Imaging/Diagnostics": "Imaging/Diagnostics",
        "Procedures": "Procedures",
        "Rehabilitation": "Rehabilitation",
        "Work Status/Restrictions": "Work Status/Restrictions",
        "Workers' Compensation Records": "Workers' Compensation Records",
        "Disability Applications/Awards": "Disability Applications/Awards",
        }
    
    # Initialize footnote tracking dictionary
    footnote_references = {}

    
    # Initialize Word application
    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False  # Set to True for debugging

    try:
        # Create a new document
        word_doc = word_app.Documents.Add()

        # Add Title
        doc_range = word_doc.Range()
        doc_range.Text = "Medical Summary\n"
        doc_range.ParagraphFormat.Alignment = 0  # Left-aligned
        doc_range.InsertParagraphAfter()

        # Extract data
        encounters = master_json["patient_summary"]["Post-Event Medical History"]["Encounters"]
        categorized_data = master_json["patient_summary"]["Post-Event Medical History"]
        records_reviewed = master_json["patient_summary"]["Post-Event Medical History"]["Records Reviewed"]

        # Initialize footnote tracking
        footnote_counter = 1

        ### 🔹 Add Chronological Summary Table
        if encounters:
            summary_range = word_doc.Range()
            summary_range.Collapse(0)  # **Fix for Add.Range error**
            summary_range.InsertAfter("\nChronological Summary\n")
            summary_range.Font.Bold = True
            summary_range.ParagraphFormat.Alignment = 0  # Left-aligned
            summary_range.InsertParagraphAfter()

            # Create a table for encounters
            num_rows = len(encounters) + 1
            table = word_doc.Tables.Add(summary_range, num_rows, 4)
            table.Style = "Table Grid"

            # Set table headers
            headers = ["Date", "Provider", "Record Type", "Summary"]
            for col in range(4):
                table.Cell(1, col + 1).Range.Text = headers[col]
                table.Cell(1, col + 1).Range.Bold = True  # Bold headers

            # Sort encounters by actual date values
            sorted_encounters = sorted(
                encounters, 
                key=lambda x: parse_date(x.get("Date of Service", "Unknown Date")) or datetime.min  # Use min date for unknown
            )
            
            # Populate table rows
            row = 2
            for encounter in sorted_encounters:
                date_str = (str(encounter.get("Date of Service", "Unknown Date")))
                provider = encounter.get("Provider/Facility Name", "Unknown")
                record_types = encounter.get("Type of Record", ["Unknown"])
                record_type = record_types[0] if isinstance(record_types, list) else record_types
                summary_text = str(encounter.get("Notes/Summary", "No summary provided."))

                # ✅ Set cell values
                table.Cell(row, 1).Range.Text = date_str
                table.Cell(row, 2).Range.Text = provider
                table.Cell(row, 3).Range.Text = record_type
                summary_cell = table.Cell(row, 4).Range
                summary_cell.Text = summary_text

               # ✅ Convert date to long format
                date_str = convert_to_long_date(str(encounter.get("Date of Service", "Unknown Date")))
                
                # ✅ Extract Provider & Record Type
                provider = encounter.get("Provider/Facility Name", "Unknown")
                record_types = encounter.get("Type of Record", ["Unknown"])
                record_type = record_types[0] if isinstance(record_types, list) else record_types
                
                # ✅ Ensure footnotes are formatted the same way as itemized categories
                footnote_text = f"{provider}, {record_type}, {date_str}"
                
                # ✅ Insert footnote at the end of the summary cell (same line)
                footnote_position = summary_cell.End - 1
                footnote_range = word_doc.Range(footnote_position, footnote_position)
                
                # ✅ Add the footnote
                if footnote_text not in footnote_references:
                    footnote = word_doc.Footnotes.Add(footnote_range, str(footnote_counter))
                    footnote.Range.Text = footnote_text
                    footnote_references[footnote_text] = footnote_counter
                    footnote_counter += 1


                row += 1  # ✅ Move to next row
        
        # 🔹 Separate pre-event and post-event data
        # 🔹 Separate pre-event and post-event data
        pre_event_data = {key: [] for key in category_titles}
        post_event_data = {key: [] for key in category_titles}
        
        for key, items in categorized_data.items():
            if key in category_titles:
                print(f"[DEBUG] Processing category: {key} with {len(items)} items")  
        
                for item in items:
                    item_date = item.get("Date of Service", "Unknown Date")
        
                    # 🔹 NEW: Extract date from `source` field if "Unknown Date"
                    if item_date == "Unknown Date":
                        source_text = ", ".join(item.get("sources", []))  # Convert set to string
                        date_match = re.search(r'\b\d{1,2}/\d{1,2}/\d{4}\b', source_text)  # Find MM/DD/YYYY format
                        
                        if date_match:
                            item_date = date_match.group()  # Extract matched date
                            print(f"[DEBUG] Extracted date from source: {item_date}")
        
                    print(f"[DEBUG] Found item: {item.get('text', 'Unknown Item')} with date {item_date}")
        
                    result = is_pre_event(item_date, event_date)
                    print(f"[DEBUG] is_pre_event() called for {item_date} -> Result: {result}")
        
                    if result:
                        pre_event_data[key].append(item)
                    else:
                        post_event_data[key].append(item)
        
        
        # 🔹 ADD PRE-EVENT MEDICAL HISTORY (if applicable)
        if any(pre_event_data.values()):
            pre_event_header = word_doc.Range()
            pre_event_header.Collapse(0)
            pre_event_header.InsertAfter("\nPre-Event Medical History")
            pre_event_header.Font.Bold = True
            pre_event_header.Font.Underline = True
            pre_event_header.ParagraphFormat.Alignment = 0  # Left-aligned
            pre_event_header.InsertParagraphAfter()
        
            for key, title in category_titles.items():
                items = pre_event_data.get(key, [])
                if items:
                    category_range = word_doc.Range()
                    category_range.Collapse(0)
                    category_range.InsertAfter(f"\n{title}")
                    category_range.Font.Bold = True
                    category_range.Font.Underline = True
                    category_range.ParagraphFormat.Alignment = 0
                    category_range.InsertParagraphAfter()
        
                    # ✅ Use the corrected function with provider mapping
                    for item in sorted(items, key=lambda x: x["text"]):
                        footnote_counter = process_item_with_sources(word_doc, item, key, footnote_references, footnote_counter, provider_mapping)
        
        
        # 🔹 ADD POST-EVENT MEDICAL HISTORY (if applicable)
        if any(post_event_data.values()):
            post_event_header = word_doc.Range()
            post_event_header.Collapse(0)
            post_event_header.InsertAfter("\nPost-Event Medical History")
            post_event_header.Font.Bold = True
            post_event_header.Font.Underline = True
            post_event_header.ParagraphFormat.Alignment = 0
            post_event_header.InsertParagraphAfter()
        
            for key, title in category_titles.items():
                items = post_event_data.get(key, [])
                if items:
                    category_range = word_doc.Range()
                    category_range.Collapse(0)
                    category_range.InsertAfter(f"\n{title}")
                    category_range.Font.Bold = True
                    category_range.Font.Underline = True
                    category_range.ParagraphFormat.Alignment = 0
                    category_range.InsertParagraphAfter()
        
                    # ✅ Use the corrected function with provider mapping
                    for item in sorted(items, key=lambda x: x["text"]):
                        footnote_counter = process_item_with_sources(word_doc, item, key, footnote_references, footnote_counter, provider_mapping)


        ### 🔹 ADD RECORDS REVIEWED
        if records_reviewed:
            records_range = word_doc.Range()
            records_range.Collapse(0)  # **Fix for Add.Range error**
            records_range.InsertAfter("\nRecords Reviewed")
            records_range.Font.Bold = True
            records_range.Font.Underline = False  # Not underlined
            records_range.ParagraphFormat.Alignment = 0  # Left-aligned
            records_range.InsertParagraphAfter()

            # Extract providers and their corresponding dates
            # Extract providers and their corresponding dates
            provider_dates = defaultdict(set)
            
            for encounter in encounters:
                provider = encounter.get("Provider/Facility Name", "Unknown Provider").strip()
                date_of_service = encounter.get("Date of Service", "Unknown Date").strip()
            
                # Ensure only valid dates are added
                if provider and provider != "Unknown Provider" and date_of_service != "Unknown Date":
                    if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", date_of_service):  # ✅ Validate MM/DD/YYYY format
                        provider_dates[provider].add(date_of_service)
                    else:
                        print(f"⚠️ Skipping invalid date: {date_of_service} for provider: {provider}")
            
            # Sort and format provider records
            for provider, dates in provider_dates.items():
                valid_dates = [date for date in dates if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", date)]  # ✅ Filter valid dates
                
                if valid_dates:
                    sorted_dates = sorted(valid_dates, key=lambda x: datetime.strptime(x, "%m/%d/%Y"))
                    formatted_dates = [convert_to_long_date(date) for date in sorted_dates]
            
                    if len(formatted_dates) > 1:
                        date_range = f"{formatted_dates[0]} through {formatted_dates[-1]}"
                    else:
                        date_range = formatted_dates[0]
            
                    records_range.InsertAfter(f"\t• {provider}, {date_range}\n")
                else:
                    print(f"⚠️ No valid dates for provider: {provider}, skipping entry.")


        # Save the document
        word_doc.SaveAs(output_path)
        word_doc.Close()
        print(f"✅ Medical summary saved successfully at: {output_path}")

    except Exception as e:
        print(f"❌ Error creating Word document: {str(e)}")
        raise

    finally:
        word_app.Quit()


def process_item_with_sources(word_doc, item, key, footnote_references, footnote_counter, provider_mapping):
    """Process a single item while ensuring each mention gets a footnote with the correct provider name from the mapping."""

    text = item.get("text", "Unknown Item").strip()
    sources = list(item.get("sources", []))  # Convert set to list

    print(f"\n[DEBUG] Processing item: {text}")
    print(f"[DEBUG] Sources found: {sources}")

    # ✅ If no sources exist, provide a default
    if not sources:
        sources = ["Unknown Provider, Unknown Record, Unknown Date"]

    # ✅ Extract only the first source
    first_source = sources[0]

    # ✅ Extract provider from source
    source_parts = first_source.split(", ")

    provider_raw = source_parts[0] if len(source_parts) > 0 else "Unknown Provider"
    
    # ✅ Use the provider_mapping dictionary to replace with the best standardized name
    provider_clean = provider_mapping.get(provider_raw, provider_raw)

    # ✅ Extract record type dynamically
    record_type_match = re.search(r'\b(?:' + '|'.join(map(re.escape, VALID_RECORD_TYPES)) + r')\b', first_source)
    record_type = record_type_match.group() if record_type_match else "Unknown Record"
    
    # ✅ Extract date using regex (MM/DD/YYYY format)
    date_match = re.search(r'\b\d{1,2}/\d{1,2}/\d{4}\b', first_source)
    date = date_match.group() if date_match else "Unknown Date"

    # ✅ Convert date to long format
    date_long_format = convert_to_long_date(date)

    # ✅ Construct correct footnote text
    footnote_text = f"{provider_clean}, {record_type}, {date_long_format}"

    # ✅ Insert the text with bullet formatting
    range_end = word_doc.Range()
    range_end.Collapse(0)
    range_end.Text = f"\t• {text} "

    # ✅ Ensure only one footnote per item
    footnote_range = word_doc.Range(range_end.Start + len(range_end.Text) - 1, range_end.Start + len(range_end.Text) - 1)

    try:
        footnote = word_doc.Footnotes.Add(footnote_range, str(footnote_counter))
        footnote.Range.Text = footnote_text
        print(f"[INFO] Added footnote #{footnote_counter} for: {text} | Source: {footnote_text}")
        footnote_counter += 1
    except Exception as e:
        print(f"[ERROR] Failed to add footnote for {text}: {e}")

    # ✅ Add paragraph break
    range_end.InsertParagraphAfter()

    return footnote_counter


def deduplicate_categories_via_ai(deduplicated_data):
    """Uses AI to strictly deduplicate and consolidate similar medical entries while retaining the most relevant source."""
    cleaned_categories = {}

    for category in [
        "Diagnoses", "Imaging/Diagnostics", "Medications", "Procedures", 
        "Rehabilitation", "Work Status/Restrictions", "Workers' Compensation Records",
        "Disability Applications/Awards"
    ]:
        items = deduplicated_data.get(category, [])
        if not items:
            cleaned_categories[category] = []
            continue

        print(f"\n[INFO] Processing {category}...")

        # Pre-process items to extract text and maintain sources
        processed_items = defaultdict(lambda: {"text": "", "sources": set()})

        for item in items:
            if isinstance(item, dict):
                text = item.get("text", "").strip()
                sources = item.get("sources", set())

                if text:
                    processed_items[text]["text"] = text
                    processed_items[text]["sources"].update(sources)

        processed_items_list = [
            {"text": data["text"], "sources": list(data["sources"])}
            for data in processed_items.values()
        ]

        print(f"\n[DEBUG] BEFORE AI Processing for {category}:")
        for item in processed_items_list:
            print(f"- {item['text']} | Sources: {item['sources']}")

        messages = [
            {
                "role": "user",
                "content": get_category_specific_prompt(category, processed_items_list)
            }
        ]

        try:
            response = anthropic_client.messages.create(
                    model="claude-3-5-sonnet-latest",
                    max_tokens=4096,
                    temperature=0,
                    messages=messages
                )

            ai_response = response.content[0].text.strip()


            print(f"\n[DEBUG] RAW AI Response for {category}:")
            print(ai_response)

            # Attempt to directly parse the AI response as JSON
            cleaned_data = None
            try:
                cleaned_data = json.loads(ai_response)
                if not isinstance(cleaned_data, list):
                    raise ValueError("Parsed JSON is not a list. Attempting regex extraction.")
            except (json.JSONDecodeError, ValueError):
                print(f"\n[WARNING] Direct JSON parsing failed. Attempting regex extraction...")

                # Extract JSON array using regex
                array_match = re.search(r'\[\s*\{[^]]*\}\s*\]', ai_response, re.DOTALL)
                if array_match:
                    json_string = array_match.group()
                    try:
                        cleaned_data = json.loads(json_string)
                    except json.JSONDecodeError as e:
                        print(f"\n[ERROR] JSON extraction failed: {e}")
                        cleaned_data = None

            if cleaned_data is None:
                print(f"\n[ERROR] No valid JSON found in AI response for {category}. Using original items.")
                cleaned_categories[category] = processed_items_list
                continue

            formatted_items = []
            for item in cleaned_data:
                text = item.get("text", "").strip()
            
                # ✅ Extract **only the first source** from AI response correctly
                first_source = item.get("source") or item.get("sources", ["Unknown Provider, Unknown Record, Unknown Date"])[0]

                if text:
                    formatted_items.append({
                        "text": text,
                        "sources": [first_source]  # ✅ Only keep the **first** source
                    })

            print(f"\n[DEBUG] AFTER AI Processing for {category}:")
            for item in formatted_items:
                print(f"- {item['text']} | Sources: {item['sources']}")

            cleaned_categories[category] = sorted(formatted_items, key=lambda x: x["text"])

        except Exception as e:
            print(f"\n[ERROR] Processing error for {category}: {str(e)}")
            cleaned_categories[category] = processed_items_list

    print("\n[DEBUG] FINAL Cleaned Categories before saving:")
    for category, items in cleaned_categories.items():
        print(f"\nCategory: {category}")
        for item in items:
            print(f"- {item['text']} | Sources: {item['sources']}")

    return cleaned_categories

def get_category_specific_prompt(category, processed_items):
    """Generate a prompt that forces strict deduplication and source retention."""

    formatted_items = json.dumps(processed_items, indent=2)

    base_prompt = (
        f"Deduplicate and strictly consolidate this {category} list while keeping only the most relevant sources.\n\n"
        f"### INPUT ###\n{formatted_items}\n\n"
    )

    if category == "Diagnoses":
        return base_prompt + """
            ### REQUIREMENTS ###
            1. Return a list of objects with this format:
            [
                {"text": "Cervical Disc Herniation with Radiculopathy", "source": "Most Relevant Source"}
            ]
            
            2. Follow these STRICT rules:
            - Merge identical and closely related diagnoses (e.g., "Cervical Herniation" + "Neck Pain" → "Cervical Disc Herniation")
            - Standardize medical terminology (use "Cervical" instead of "Neck")
            - REMOVE any duplicate or unnecessary variations
            - KEEP ONLY the most detailed, authoritative source
            """

    elif category == "Medications":
        return base_prompt + """
            ### REQUIREMENTS ###
            1. Return a list of objects with this format:
            [
                {"text": "Lidocaine 1% - Local Anesthetic", "source": "Most Relevant Source"}
            ]
            
            2. Follow these STRICT rules:
            - Merge identical medications (case-insensitive)
            - Remove redundant dosage repetitions
            - Keep the medication's purpose when available
            - KEEP ONLY the most complete and authoritative source
            """

    elif category == "Procedures":
        return base_prompt + """
            ### REQUIREMENTS ###
            1. Return a list of objects with this format:
            [
                {"text": "Lumbar Epidural Steroid Injection", "source": "Most Relevant Source"}
            ]
            
            2. STRICT RULES:
            - Keep ONLY **one** most relevant source.
            - Do NOT return multiple sources in a list—only return the best one.
            - Merge identical or very similar procedures.
            - Use precise and standardized medical terminology.
        """


    else:
        return base_prompt + f"""
            ### REQUIREMENTS ###
            1. Return a STRICTLY DEDUPLICATED list of objects with this format:
            {get_deduplication_examples(category)}
            
            2. STRICT RULES:
            - Remove duplicate, redundant, or closely similar entries
            - Keep ONLY the most complete, detailed, and authoritative sources
            - Standardize terminology
            - DO NOT return explanatory text, only the JSON array
            """

def process_item(word_doc, item, key, encounters, footnote_references, footnote_counter):
    """Process a single item for the medical summary."""
    # Split date and text if present
    parts = item.split(", ", 1)
    if len(parts) == 2:
        date_str = convert_to_long_date(parts[0])
        text = parts[1]
    else:
        date_str = "Unknown Date"
        text = parts[0]

    # Find source details
    source_file = "Unknown Source"
    provider = "Unknown Provider"
    page_number = "Unknown"

    for encounter in encounters:
        if text in str(encounter.get(key, [])):
            source_file = str(encounter.get("Source File", "Unknown Source"))
            provider = str(encounter.get("Provider/Facility Name", "Unknown Provider"))
            page_number = str(encounter.get("Source Page", "Unknown"))
            break

    # Format text based on category
    if key == "Diagnoses":
        formatted_text = text
    elif key == "Medications":
        formatted_text = format_medication(text)
    else:
        formatted_text = f"{date_str}, {text}"

    # Create a new paragraph for each item
    range_end = word_doc.Range()
    range_end.Collapse(0)  # Move to end
    
    # 1. Insert text WITH the space
    range_end.Text = f"\t• {formatted_text} "
    
    # 2. Calculate position before the space
    text_without_space = f"\t• {formatted_text}"
    footnote_position = range_end.Start + len(text_without_space)
    footnote_range = word_doc.Range(footnote_position, footnote_position)
    
    # 3. Add footnote
    footnote_text = f"{provider}, {source_file}, {date_str}, Page {page_number}."
    if footnote_text not in footnote_references:
        footnote = word_doc.Footnotes.Add(footnote_range, str(footnote_counter))
        footnote.Range.Text = footnote_text
        footnote_references[footnote_text] = footnote_counter
        footnote_counter += 1
    
    # 4. Add paragraph break
    range_end.InsertParagraphAfter()
    
    # Return the updated counter
    return footnote_counter

def save_master_json_to_txt(master_json, output_path):
    """Save the entire master JSON file as a formatted text file for debugging."""
    try:
        # Convert any sets to lists to make JSON serializable
        def convert_sets(obj):
            if isinstance(obj, set):
                return list(obj)  # Convert set to list
            elif isinstance(obj, dict):
                return {k: convert_sets(v) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [convert_sets(i) for i in obj]
            return obj

        cleaned_json = convert_sets(master_json)

        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(cleaned_json, f, indent=4, ensure_ascii=False)

        print(f"✅ Master JSON saved successfully at: {output_path}")

    except Exception as e:
        print(f"❌ [ERROR] Could not save master JSON: {str(e)}")

def convert_sets_to_lists(data):
    """ Recursively converts sets to lists in a dictionary. """
    if isinstance(data, dict):
        return {key: convert_sets_to_lists(value) for key, value in data.items()}
    elif isinstance(data, list):
        return [convert_sets_to_lists(item) for item in data]
    elif isinstance(data, set):
        return list(data)
    else:
        return data

def main():
    """Load JSON files from folder, aggregate data, summarize via AI, and generate Word medical summary."""
    try:
        master_json, sorted_dates_list = load_json_files()  

        # Ensure AI summarization is used
        date_groups = group_encounters_by_date(master_json, sorted_dates_list, provider_mapping)
        condensed_encounters = condense_encounters_via_ai(date_groups)  

        # Use AI deduplication for all categories
        cleaned_categories = deduplicate_categories_via_ai(
            master_json["patient_summary"]["Post-Event Medical History"]
        )
        

        # Create final master JSON
        master_json = create_final_master_json(
            condensed_encounters, 
            cleaned_categories, 
            master_json["patient_summary"]["Post-Event Medical History"]["Records Reviewed"]
        )

                # Convert before saving
        master_json = convert_sets_to_lists(master_json)
        
        # Debugging before saving
        print(json.dumps(master_json, indent=4))

        print("[DEBUG] All Encounters Collected for Chronology:")
        for encounter in master_json["patient_summary"]["Post-Event Medical History"]["Encounters"]:
            print(encounter)

        
        # Save the final JSON
        with open("master_json_output.json", "w", encoding="utf-8") as f:
            json.dump(master_json, f, indent=4)


        # ✅ Save master JSON to a text file for debugging
        # ✅ Save master JSON to a text file for debugging
        save_master_json_to_txt(master_json, "C:\\Users\\hfreeman\\Desktop\\master_json_debug.txt")


        # Generate the Word document
        create_medical_summary(OUTPUT_WORD_DOC, master_json)
        print("✅ Medical summary successfully generated.")

    except Exception as e:
        print(f"❌ [ERROR] Error in main process: {str(e)}")
        raise

import tkinter as tk
from tkinter import filedialog, messagebox, ttk  # ✅ Import ttk

import threading

# Initialize the Tkinter root window
root = tk.Tk()
root.title("Medical Summary Generator")
root.geometry("500x300")  # Set window size

# Define GUI variables
progress_var = tk.DoubleVar()
progress_steps = [
    "Loading JSON files...",
    "Processing encounters...",
    "AI summarization...",
    "Deduplicating data...",
    "Generating medical summary...",
    "Saving output...",
    "Completed!"
]

# Create UI Elements
frame = tk.Frame(root, padx=20, pady=20)
frame.pack(pady=20)

folder_label = tk.Label(frame, text="No folder selected", wraplength=400)
folder_label.pack()

browse_button = tk.Button(frame, text="Select Project Folder", command=browse_project_folder)
browse_button.pack(pady=5)

event_date_label = tk.Label(frame, text="Enter Event Date (MM/DD/YYYY):")
event_date_label.pack()

event_date_entry = tk.Entry(frame)
event_date_entry.pack(pady=5)

start_button = tk.Button(frame, text="Start Processing", command=start_processing)
start_button.pack(pady=10)

# ✅ Fix the Progress Bar issue
progress_bar = ttk.Progressbar(frame, variable=progress_var, length=400)
progress_bar.pack(pady=10)

status_label = tk.Label(frame, text="Idle")
status_label.pack()

# Run the Tkinter main loop
root.mainloop()

