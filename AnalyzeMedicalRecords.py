import os
import json
import hashlib
import PyPDF2
import anthropic
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from collections import defaultdict
import json
import re

# Allowed Record Types
VALID_RECORD_TYPES = [
    "Consultation", "Discharge Summary", "Emergency Department", "EMS",
    "History & Physical", "Hospital", "Imaging/Diagnostic", "Incident Report",
    "Initial Evaluation", "Office Visit", "Operative Report", "Phone Call/Email",
    "Police Report", "Procedure", "Progress Note", "Telemedicine Visit",
    "Workers' Compensation Exam", "Other"
]

class AnalysisManager:
    def __init__(self, project_folder: str):
        """Initialize Analysis Manager"""
        self.medical_folder = os.path.join(project_folder, "Source Data", "From Client", "Medical")
        self.analyses_dir = os.path.join(project_folder, "Work Product", "ai-analyses")
        os.makedirs(self.analyses_dir, exist_ok=True)

    def _get_analysis_path(self, pdf_path: str) -> str:
        """Generate path for analysis file in the ai-analyses folder"""
        pdf_hash = self._get_file_hash(pdf_path)
        return os.path.join(self.analyses_dir, f"{pdf_hash}.json")

    def _get_file_hash(self, file_path: str) -> str:
        """Calculate MD5 hash of file"""
        with open(file_path, 'rb') as f:
            file_hash = hashlib.md5()
            while chunk := f.read(8192):
                file_hash.update(chunk)
        return file_hash.hexdigest()

    def has_analysis(self, pdf_path: str) -> bool:
        """Check if analysis exists for PDF"""
        return os.path.exists(self._get_analysis_path(pdf_path))

    def create_analysis(self, pdf_path: str, processor) -> dict:
        """Create new analysis for a PDF and save it."""
        print(f"[DEBUG] Starting analysis for {pdf_path}")  # Debugging print
    
        results = processor.process_pdf(pdf_path)
    
        if not results:
            print(f"[WARNING] No results extracted for {pdf_path}. Skipping save.")
            return {}
    
        print(f"[DEBUG] Extracted data: {json.dumps(results, indent=2)}")  # Debugging print
    
        analysis = {
            "document_info": {
                "file_name": os.path.basename(pdf_path),
                "file_path": pdf_path,
                "file_hash": self._get_file_hash(pdf_path),
                "creation_date": datetime.now().isoformat(),
                "last_modified": datetime.now().isoformat(),
                "version": 1
            },
            "page_analyses": defaultdict(lambda: {"entries": []})
        }
    
        for entry in results:
            page = entry["Source Page"]
            analysis["page_analyses"][page]["entries"].append(entry)
    
        print(f"[DEBUG] Calling save_analysis() for {pdf_path}")  # Debugging print
        self.save_analysis(pdf_path, analysis)
    
        return analysis


    def save_analysis(self, pdf_path: str, analysis: dict) -> None:
        """Save analysis to file"""
        analysis_path = self._get_analysis_path(pdf_path)
    
        # Debugging prints
        print(f"[DEBUG] Attempting to save analysis for: {pdf_path}")
        print(f"[DEBUG] Expected save path: {analysis_path}")
    
        try:
            with open(analysis_path, 'w', encoding='utf-8') as f:
                json.dump(analysis, f, indent=2, ensure_ascii=False)
            print(f"[SUCCESS] Analysis saved successfully: {analysis_path}")
    
        except Exception as e:
            print(f"[ERROR] Failed to save analysis for {pdf_path}: {str(e)}")


import anthropic

class PDFProcessor:
    def __init__(self):
        self.previous_page_info = None
        self.client = anthropic.Anthropic(
            api_key=""  # Replace with actual key
        )
        self.previous_page_info = None
        self.date_patterns = [
            r'(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})',  # Matches 1/30/21, 01-30-2021, etc.
            r'(?i)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+(\d{1,2}),?\s+(\d{2,4})'  # Matches January 30, 2021, Jan 30 21, etc.
        ]
        self.month_map = {
            'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 
            'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
            'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
        }

    def standardize_date(self, date_str):
        """Standardize date format to MM/DD/YYYY"""
        if not date_str or date_str == "":
            return ""

        # Handle date ranges
        if ' - ' in date_str:
            start_date, end_date = date_str.split(' - ')
            return f"{self.standardize_date(start_date)} - {self.standardize_date(end_date)}"

        for pattern in self.date_patterns:
            match = re.search(pattern, date_str)
            if match:
                groups = match.groups()
                if len(groups) == 3:
                    if pattern == self.date_patterns[0]:  # Numeric format
                        month, day, year = groups
                    else:  # Text month format
                        month = self.month_map[groups[0].lower()[:3]]
                        day = groups[1]
                        year = groups[2]

                    # Standardize month and day to 2 digits
                    month = month.zfill(2)
                    day = str(int(day)).zfill(2)

                    # Standardize year to 4 digits
                    if len(year) == 2:
                        year = '20' + year if int(year) < 50 else '19' + year
                    
                    return f"{month}/{day}/{year}"

        return date_str  # Return original if no pattern matches

    def process_pdf(self, pdf_path: str) -> list:
        """Processes an entire PDF file with enhanced page continuation logic."""
        results = []
        self.previous_page_info = None  # Reset for new PDF

        try:
            with open(pdf_path, 'rb') as file:
                reader = PyPDF2.PdfReader(file)
                total_pages = len(reader.pages)

                for page_num in range(total_pages):
                    try:
                        page = reader.pages[page_num]
                        page_entries = self.process_page(page, page_num + 1, os.path.basename(pdf_path))

                        if page_entries:
                            results.extend(page_entries)

                        print(f"Processed page {page_num + 1}/{total_pages}")
                    except Exception as e:
                        print(f"Error on page {page_num + 1}: {str(e)}")
                        continue

                return results
        except Exception as e:
            print(f"Error processing PDF {pdf_path}: {str(e)}")
            return []

    def is_continuation_page(self, page_text: str) -> bool:
        """Determine if the current page is likely a continuation of the previous page."""
        if not self.previous_page_info:
            return False  # No previous page, so it can't be a continuation
    
        # Check if page starts with lowercase letter or continuation punctuation
        first_char = page_text.strip()[0] if page_text.strip() else ''
        starts_with_continuation = first_char.islower() or first_char in [',', ';', ')']
    
        # Check if the page lacks a provider and date of service
        missing_provider = not self.previous_page_info.get("Provider/Facility Name")
        missing_date = not self.previous_page_info.get("Date of Service")
    
        # If BOTH provider and date are missing, trigger continuation
        is_missing_key_info = missing_provider and missing_date
    
        # Check if page lacks typical headers but has medical content
        has_header_elements = any(indicator in page_text.lower()[:200] for indicator in [
            'date:', 'patient:', 'name:', 'dr.', 'hospital:', 'clinic:'
        ])
        has_medical_content = any(term in page_text.lower() for term in [
            'diagnosis', 'treatment', 'medication', 'patient', 'prescribed',
            'examination', 'assessment', 'symptoms'
        ])
    
        # ✅ If both Provider and Date are missing → it's a continuation
        if is_missing_key_info:
            return (starts_with_continuation or (not has_header_elements and has_medical_content))
    
        return False
    
    def merge_continuation_data(self, current_entry: dict) -> dict:
        """Merge continuation page data with previous page information."""
        if not self.previous_page_info:
            return current_entry
    
        merged_entry = current_entry.copy()
    
        # ✅ If Provider is missing but Date exists, inherit the last known Provider
        if not merged_entry.get("Provider/Facility Name") and self.previous_page_info.get("Provider/Facility Name"):
            merged_entry["Provider/Facility Name"] = self.previous_page_info["Provider/Facility Name"]
    
        # ✅ If both Provider & Date were missing, it was marked as a continuation
        # Inherit missing critical fields
        for field in ["Date of Service", "Type of Record"]:
            if not merged_entry.get(field) and self.previous_page_info.get(field):
                merged_entry[field] = self.previous_page_info[field]
    
        # Append Notes/Summary if it seems to be a continuation
        if merged_entry.get("Notes/Summary"):
            prev_summary = self.previous_page_info.get("Notes/Summary", "")
            if prev_summary:
                merged_entry["Notes/Summary"] = f"{prev_summary} (continued) {merged_entry['Notes/Summary']}"
    
        # Merge list fields
        for field in ["Diagnoses", "Imaging/Diagnostics", "Medications", "Procedures", 
                      "Rehabilitation", "Work Status/Restrictions", "Workers' Compensation",
                      "Disability Applications/Awards"]:
            if merged_entry.get(field) and self.previous_page_info.get(field):
                merged_entry[field] = list(set(self.previous_page_info[field] + merged_entry[field]))
    
        return merged_entry


    def process_page(self, page, page_num: int, file_name: str) -> list:
        """Process a single PDF page with enhanced continuation detection and date standardization."""
        try:
            page_text = page.extract_text()
            is_continuation = self.is_continuation_page(page_text)
    
            continuation_context = ""
            if is_continuation and self.previous_page_info:
                continuation_context = (
                    "NOTE: This appears to be a continuation page. "
                    "If appropriate, use this context from the previous page:\n"
                    f"Previous Date of Service: {self.previous_page_info.get('Date of Service', '')}\n"
                    f"Previous Provider: {self.previous_page_info.get('Provider/Facility Name', '')}\n"
                    f"Previous Record Type: {self.previous_page_info.get('Type of Record', '')}\n"
                )
    
            messages = [
                {
                    "role": "user",
                    "content": (
                        f"{continuation_context}\n"
                        f"Analyze the following **medical record page (Page {page_num})** and extract relevant "
                        "medical information **strictly in JSON format**.\n\n"
                        f"TEXT CONTENT:\n{page_text}\n\n"
                        "### JSON OUTPUT FORMAT REQUIREMENTS\n"
                        "Your response must be a **valid JSON list of objects**, where each object follows this exact structure:\n\n"
                        "[\n"
                        "    {\n"
                        '        "Date of Service": "",\n'
                        '        "Provider/Facility Name": "",\n'
                        '        "Type of Record": "",\n'
                        '        "Notes/Summary": "",\n'
                        '        "Diagnoses": [""],\n'
                        '        "Imaging/Diagnostics": [""],\n'
                        '        "Medications": [""],\n'
                        '        "Procedures": [""],\n'
                        '        "Rehabilitation": [""],\n'
                        '        "Work Status/Restrictions": [""],\n'
                        '        "Workers\' Compensation": [""],\n'
                        '        "Disability Applications/Awards": [""],\n'
                        f'        "Source Page": {page_num},\n'
                        f'        "Source File": "{file_name}",\n'
                        f'        "Processing Date": "{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}"\n'
                        "    }\n"
                        "]\n\n"
                        "### STRICT INSTRUCTIONS\n"
                        "- **Return only valid JSON** as shown above. Do not include any explanations or extra text.\n"
                        "- **If this page contains no relevant medical content, return an empty list (`[]`)**.\n"
                        "- **Do not modify field names**. Always use the exact JSON structure above.\n"
                        "- **If a field is missing or not relevant, leave it blank (`\"\"`) or empty (`[]`). Do not guess or make assumptions.**\n"
                        f"- **For `Type of Record`, only use one of these exact values:** {', '.join(VALID_RECORD_TYPES)}. If unknown, leave it blank.\n"
                        "- **If multiple distinct medical events exist on this page, return them as separate JSON objects within the list.**\n"
                        "- **All list values (e.g., `Diagnoses`, `Medications`) must contain only strings, not dictionaries.**\n"
                        "- **Ensure that all list values are simple, plain text descriptions. DO NOT return nested dictionaries, objects, or extra key-value pairs inside lists.**\n"
                        "- **All dates must be in `MM/DD/YYYY` format. If a date range exists, use `MM/DD/YYYY - MM/DD/YYYY`.**\n"
                        "- **Example list values (correct format):**\n"
                        '  - `"Diagnoses": ["Cervical sprain", "Lower back pain", "Headaches"]`\n'
                        '  - `"Medications": ["Ibuprofen 800 mg", "Gabapentin 300 mg"]`\n'
                        "- **Example list values (incorrect format, must be avoided):**\n"
                        '  - `"Diagnoses": [{"Condition": "Cervical sprain", "Location": "Neck"}]` ❌ (NO NESTED OBJECTS!)\n'
                        '  - `"Medications": [{"Name": "Gabapentin", "Dosage": "300 mg"}]` ❌ (NO KEY-VALUE PAIRS!)\n'
                    )
                }
            ]
    
            try:
                # Send to Claude API
                response = self.client.messages.create(
                    model="claude-3-5-sonnet-latest",
                    max_tokens=4096,
                    temperature=0,
                    messages=messages
                )
            
                response_text = response.content[0].text.strip()
            
                # 🚨 Debugging: Print raw response
                print(f"[DEBUG] API Response for page {page_num}: {response_text}")
            
                # ✅ If response is empty, retry
                if not response_text:
                    print(f"[ERROR] Claude returned an empty response for page {page_num}. Retrying...")
                    return self.retry_page_analysis(page, page_num, file_name)
            
                parsed_data = self._parse_response(response_text, page_num, file_name)

                # Process each entry
                processed_entries = []
                for entry in parsed_data:
                    entry['Date of Service'] = self.standardize_date(entry['Date of Service'])
    
                    if is_continuation:
                        entry = self.merge_continuation_data(entry)
    
                    processed_entries.append(entry)
    
                # Update previous page info for next iteration
                if processed_entries:
                    self.previous_page_info = processed_entries[-1]
    
                return processed_entries
    
            except Exception as e:
                print(f"[ERROR] API error on page {page_num}: {str(e)}")
                return []
    
        except Exception as e:
            print(f"[ERROR] General error in page processing for page {page_num}: {str(e)}")
            return []


    def _parse_response(self, response_text: str, page_num: int, file_name: str) -> list:
        """Extract JSON from the API response and handle errors."""
        entries = []
    
        # Handle empty responses
        if not response_text or not response_text.strip():
            print(f"[ERROR] Empty response from API for page {page_num}. Skipping this page.")
            return []
    
        try:
            # 🔹 Use regex to extract JSON only
            match = re.search(r"\[\s*\{.*?\}\s*\]", response_text, re.DOTALL)
            if not match:
                print(f"[ERROR] No valid JSON detected in response for page {page_num}. Skipping.")
                print(f"[DEBUG] Raw response:\n{response_text}")  # Log full response for debugging
                return []
    
            json_text = match.group(0)  # Extract JSON portion
    
            # Convert response text to JSON
            data = json.loads(json_text)
    
            if isinstance(data, list):  # Ensure it's a list
                for entry in data:
                    entry.update({
                        "Source Page": page_num,
                        "Source File": file_name,
                        "Processing Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    })
                    entries.append(entry)
            else:
                print(f"[ERROR] Parsed JSON is not a list on page {page_num}. Skipping.")
    
        except json.JSONDecodeError as e:
            print(f"[ERROR] JSON parsing error on page {page_num}: {str(e)}")
            print(f"[DEBUG] Malformed JSON response:\n{response_text}")  # Log full response for debugging
    
        return entries



import threading

class AnalyzeMedicalRecordsApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Analyze Medical Records")
        self.root.geometry("600x400")
        self.analysis_manager = None
        self.setup_ui()

    def setup_ui(self):
        """Set up the UI for the app"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title_label = ttk.Label(main_frame, text="Medical Records Analyzer", font=("Arial", 16, "bold"))
        title_label.pack(pady=10)

        project_frame = ttk.LabelFrame(main_frame, text="Project Selection", padding="5")
        project_frame.pack(fill=tk.X, pady=5)

        self.project_path_var = tk.StringVar()
        path_label = ttk.Label(project_frame, textvariable=self.project_path_var, wraplength=500)
        path_label.pack(fill=tk.X, pady=5)

        select_btn = ttk.Button(project_frame, text="Select Project Folder", command=self.select_folder)
        select_btn.pack(pady=5)

        self.progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="5")
        self.progress_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.progress_bar = ttk.Progressbar(self.progress_frame, length=400, mode="determinate")
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)

        self.progress_label = ttk.Label(self.progress_frame, text="Progress: 0% (0 pages remaining)")
        self.progress_label.pack()

        self.estimated_time_label = ttk.Label(self.progress_frame, text="Estimated time remaining: --:--")
        self.estimated_time_label.pack()

        self.progress_text = tk.Text(self.progress_frame, height=10, wrap=tk.WORD, state='disabled')
        self.progress_text.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.progress_frame, command=self.progress_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.progress_text.configure(yscrollcommand=scrollbar.set)

    def log_progress(self, message: str):
        """Add message to progress text dynamically in the main thread"""
        self.root.after(0, self._update_progress_text, message)

    def _update_progress_text(self, message: str):
        """Helper function to safely update the text box"""
        self.progress_text.configure(state='normal')
        self.progress_text.insert(tk.END, message + "\n")
        self.progress_text.see(tk.END)
        self.progress_text.configure(state='disabled')

    def update_progress_bar(self, progress, pages_remaining, estimated_time):
        """Update the progress bar and labels in the main thread"""
        self.root.after(0, self._update_progress_ui, progress, pages_remaining, estimated_time)

    def _update_progress_ui(self, progress, pages_remaining, estimated_time):
        """Helper function to safely update the progress bar"""
        self.progress_bar["value"] = progress
        self.progress_label.config(text=f"Progress: {int(progress)}% ({pages_remaining} pages remaining)")
        self.estimated_time_label.config(
            text=f"Estimated time remaining: {int(estimated_time // 60)}m {int(estimated_time % 60)}s"
        )

    def select_folder(self):
        """Select a project folder"""
        folder_path = filedialog.askdirectory(title="Select Project Folder")
        if folder_path:
            self.project_path_var.set(folder_path)
            self.analysis_manager = AnalysisManager(folder_path)
            
            # Run file processing in a separate thread to keep the UI responsive
            threading.Thread(target=self.analyze_folder, args=(folder_path,), daemon=True).start()

    def analyze_folder(self, folder_path):
        """Process all PDFs in the medical records folder with real-time UI updates"""
        medical_folder = self.analysis_manager.medical_folder
        pdf_files = [f for f in os.listdir(medical_folder) if f.lower().endswith('.pdf')]
    
        print(f"[DEBUG] Found {len(pdf_files)} PDF files in {medical_folder}")  # Debugging print
    
        if not pdf_files:
            print("[WARNING] No PDFs found! Aborting analysis.")
            return
    
        total_pages = 0
        pdf_page_counts = {}
    
        # Get total page count for progress tracking
        for pdf_file in pdf_files:
            pdf_path = os.path.join(medical_folder, pdf_file)
            try:
                with open(pdf_path, 'rb') as file:
                    reader = PyPDF2.PdfReader(file)
                    page_count = len(reader.pages)
                    pdf_page_counts[pdf_path] = page_count
                    total_pages += page_count
                    print(f"[DEBUG] {pdf_file} has {page_count} pages")  # Debugging print
            except Exception as e:
                print(f"[ERROR] Skipping {pdf_file}: Error reading file ({str(e)})")
    
        if total_pages == 0:
            print("[ERROR] No valid PDFs with pages found. Aborting.")
            return
    
        processed_pages = 0
        start_time = datetime.now()
    
        for pdf_file in pdf_files:
            pdf_path = os.path.join(medical_folder, pdf_file)
    
            # Debugging print
            print(f"[DEBUG] Checking if analysis exists for {pdf_path}")
    
            if not self.analysis_manager.has_analysis(pdf_path):
                print(f"[DEBUG] No existing analysis found. Starting analysis for {pdf_path}")
    
                self.log_progress(f"Analyzing: {pdf_file}")
    
                processor = PDFProcessor()
    
                # CALLING CREATE_ANALYSIS
                self.analysis_manager.create_analysis(pdf_path, processor)  
                
                self.log_progress(f"Completed analysis of {pdf_file}")
                print(f"[DEBUG] Completed analysis for {pdf_file}")
    
            else:
                print(f"[DEBUG] Analysis already exists for {pdf_file}, skipping.")
    
        self.root.after(0, messagebox.showinfo, "Complete", "All PDFs have been analyzed.")


    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = AnalyzeMedicalRecordsApp()
    app.run()
