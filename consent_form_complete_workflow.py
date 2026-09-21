import pandas as pd
import os
from pathlib import Path
import re
from datetime import datetime
import json
import shutil

class ConsentFormCompleteWorkflow:
    def __init__(self, excel_path, consent_forms_dir, output_dir=None):
        """
        Initialize the complete workflow with all required paths
        """
        self.excel_path = excel_path
        self.consent_forms_dir = consent_forms_dir
        
        # If output_dir not provided, create OUTPUT folder in same directory as script
        if output_dir is None:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            self.output_dir = os.path.join(script_dir, "OUTPUT")
        else:
            self.output_dir = output_dir
        
        # Set up subdirectories
        self.reports_dir = os.path.join(self.output_dir, "Reports")
        self.findings_dir = os.path.join(self.output_dir, "Findings")
        
        self.df = None
        self.results = {}
        self.matched_excel_file = None
        
        # Create directories if they don't exist
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.reports_dir, exist_ok=True)
        os.makedirs(self.findings_dir, exist_ok=True)
        
    def load_excel(self):
        """Load the Excel file"""
        try:
            self.df = pd.read_excel(self.excel_path)
            
            # Normalize column names to uppercase for consistency
            self.df.columns = self.df.columns.str.strip().str.upper()
            
            print(f"✓ Excel file loaded successfully: {len(self.df)} records found")
            print(f"✓ Columns: {list(self.df.columns)}")
            return True
        except Exception as e:
            print(f"✗ Error loading Excel file: {e}")
            return False
    
    def normalize_string(self, text):
        """Normalize string for comparison - remove special chars, convert to lowercase"""
        if pd.isna(text):
            return ""
        text = str(text).lower()
        # Remove special characters but keep spaces and underscores
        text = re.sub(r'[^a-z0-9\s_]', '', text)
        # Replace multiple spaces with single space
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def get_first_name(self, acct_desc):
        """Get the first word/name from ACCT_DESC (formerly CLIENT_SHORT)"""
        if pd.isna(acct_desc):
            return ""
        
        normalized = self.normalize_string(acct_desc)
        names = normalized.split()
        
        # Return first name if it exists and is longer than 2 characters
        if names and len(names[0]) > 2:
            return names[0]
        return ""
    
    def get_last_9_digits(self, acct_no):
        """Extract last 9 digits from ACCT_NO"""
        if pd.isna(acct_no):
            return ""
        
        # Convert to string and remove all non-digits
        digits_only = re.sub(r'\D', '', str(acct_no))
        
        # Return last 9 digits if available
        if len(digits_only) >= 9:
            return digits_only[-9:]
        return digits_only
    
    def find_9plus_digit_sequences(self, text):
        """Find all sequences of 9 or more consecutive digits in text"""
        # Find all sequences of digits
        digit_sequences = re.findall(r'\d+', text)
        
        # Filter sequences with 9+ digits and return their last 9 digits
        results = []
        for seq in digit_sequences:
            if len(seq) >= 9:
                results.append(seq[-9:])
        
        return results
    
    def normalize_performed_by(self, performed_by_value):
        """Normalize the 'PERFORMED BY' value - convert invalid/missing values to 'Unknown'"""
        if pd.isna(performed_by_value):
            return "Unknown"
        
        performed_by_str = str(performed_by_value).strip()
        
        # If empty or 'nan' string, treat as Unknown
        if not performed_by_str or performed_by_str.lower() == 'nan':
            return "Unknown"
        
        return performed_by_str
    
    def text_to_underscore_format(self, text):
        """Convert text to underscore-separated format for matching"""
        if pd.isna(text):
            return ""
        
        normalized = self.normalize_string(text)
        # Replace spaces with underscores
        return normalized.replace(' ', '_')
    
    def check_match(self, pdf_name, row):
        """
        Check if PDF name matches any combination of Excel fields
        Returns: (matched, match_type)
        """
        pdf_normalized = self.normalize_string(pdf_name)
        
        # Get normalized values from row
        client_no = self.normalize_string(row['CLIENT_NO'])
        acct_desc = self.normalize_string(row['ACCT_DESC'])  # Formerly CLIENT_SHORT
        acct_no = self.normalize_string(row['ACCT_NO'])
        # ACCT_TYPE_DESC (formerly ACCT_DESC) - NO MATCHING LOGIC
        
        # Get first name from ACCT_DESC
        first_name = self.get_first_name(row['ACCT_DESC'])
        
        # Get last 9 digits of ACCT_NO for advanced matching
        last_9_digits = self.get_last_9_digits(row['ACCT_NO'])
        
        # Get underscore-separated format for ACCT_DESC
        acct_desc_underscore = self.text_to_underscore_format(row['ACCT_DESC'])
        
        # Check various combinations
        
        # 1. Exact CLIENT_NO match
        if client_no and client_no in pdf_normalized:
            return True, f"CLIENT_NO: {client_no}"
        
        # 2. ACCT_NO - Enhanced matching (3 ways)
        if acct_no:
            # 2A. Full ACCT_NO anywhere in PDF
            if acct_no in pdf_normalized:
                return True, f"ACCT_NO: {acct_no}"
            
            # 2B. Last 9 digits match - find 9+ digit sequences in PDF
            if last_9_digits and len(last_9_digits) == 9:
                pdf_digit_sequences = self.find_9plus_digit_sequences(pdf_normalized)
                if last_9_digits in pdf_digit_sequences:
                    return True, f"ACCT_NO_LAST_9: {last_9_digits}"
            
            # 2C. Check underscore-separated format with 9+ consecutive digits
            # This handles cases like "abc_123456789_xyz.pdf"
            if last_9_digits and len(last_9_digits) == 9:
                # Check if the last 9 digits appear as part of underscore-separated segments
                if last_9_digits in pdf_normalized:
                    return True, f"ACCT_NO_UNDERSCORE_9: {last_9_digits}"
        
        # 3. First name from ACCT_DESC (ONLY FIRST NAME)
        if first_name and first_name in pdf_normalized:
            return True, f"FIRST_NAME: {first_name}"
        
        # 4. ACCT_DESC full match
        if acct_desc and acct_desc in pdf_normalized:
            return True, f"ACCT_DESC: {acct_desc}"
        
        # 5. ACCT_DESC underscore-separated match
        if acct_desc_underscore and acct_desc_underscore in pdf_normalized:
            return True, f"ACCT_DESC_UNDERSCORE: {acct_desc_underscore}"
        
        # 6. ACCT_NO + ACCT_DESC combination
        if acct_no and acct_desc:
            combined = f"{acct_no}{acct_desc}".replace(" ", "")
            if combined in pdf_normalized.replace(" ", ""):
                return True, f"ACCT_NO_ACCT_DESC: {acct_no}_{acct_desc}"
        
        # 7. ACCT_DESC + ACCT_NO combination
        if acct_desc and acct_no:
            combined = f"{acct_desc}{acct_no}".replace(" ", "")
            if combined in pdf_normalized.replace(" ", ""):
                return True, f"ACCT_DESC_ACCT_NO: {acct_desc}_{acct_no}"
        
        # 8. CLIENT_NO + ACCT_DESC combination
        if client_no and acct_desc:
            combined = f"{client_no}{acct_desc}".replace(" ", "")
            if combined in pdf_normalized.replace(" ", ""):
                return True, f"CLIENT_NO_ACCT_DESC: {client_no}_{acct_desc}"
        
        # 9. CLIENT_NO + ACCT_NO combination
        if client_no and acct_no:
            combined = f"{client_no}{acct_no}".replace(" ", "")
            if combined in pdf_normalized.replace(" ", ""):
                return True, f"CLIENT_NO_ACCT_NO: {client_no}_{acct_no}"
        
        # 10. CLIENT_NO + ACCT_DESC + ACCT_NO combination
        if client_no and acct_desc and acct_no:
            combined = f"{client_no}{acct_desc}{acct_no}".replace(" ", "")
            if combined in pdf_normalized.replace(" ", ""):
                return True, f"CLIENT_NO_ACCT_DESC_ACCT_NO"
        
        return False, None
    
    def get_all_branch_folders(self):
        """Get all branch folders from consent forms directory"""
        branch_folders = set()
        
        if not os.path.exists(self.consent_forms_dir):
            return branch_folders
        
        # List all directories in consent forms directory
        for item in os.listdir(self.consent_forms_dir):
            item_path = os.path.join(self.consent_forms_dir, item)
            if os.path.isdir(item_path):
                # Normalize branch folder name to 4 digits if it's numeric
                branch_name = item
                if branch_name.replace('.0', '').isdigit():
                    branch_name = branch_name.replace('.0', '')
                    # Try to convert to int to remove leading zeros, then pad to 4 digits
                    try:
                        branch_num = int(branch_name)
                        branch_folders.add(branch_num)
                    except:
                        branch_folders.add(branch_name)
                else:
                    branch_folders.add(branch_name)
        
        return branch_folders
    
    def find_pdfs_in_branch(self, branch_folder):
        """Find all PDF files in the branch folder and subfolders"""
        pdf_files = []
        
        # Try multiple formats: with leading zeros and without
        branch_str = str(branch_folder)
        possible_folders = [
            branch_str,           # e.g., "32"
            branch_str.zfill(4),  # e.g., "0032"
            branch_str.zfill(3),  # e.g., "032"
            branch_str.zfill(2),  # e.g., "32"
        ]
        
        # Remove duplicates while preserving order
        seen = set()
        possible_folders = [x for x in possible_folders if not (x in seen or seen.add(x))]
        
        for folder_name in possible_folders:
            branch_path = Path(self.consent_forms_dir) / folder_name
            
            if branch_path.exists():
                # Recursively find all PDF files
                for pdf_file in branch_path.rglob("*.pdf"):
                    pdf_files.append(pdf_file)
                break  # Found the folder, no need to try other formats
        
        return pdf_files
    
    def match_branch(self, branch):
        """Match all records for a specific branch"""
        print(f"\n{'='*80}")
        print(f"Processing Branch: {branch}")
        print(f"{'='*80}")
        
        # Filter records for this branch
        branch_records = self.df[self.df['BRANCH'] == branch]
        total_records = len(branch_records)
        
        print(f"Total records in Excel for this branch: {total_records}")
        
        # Find PDFs in branch folder
        pdf_files = self.find_pdfs_in_branch(branch)
        print(f"Total PDF files found in branch folder: {len(pdf_files)}")
        
        if len(pdf_files) == 0:
            print(f"⚠ Warning: No PDF files found for branch {branch}")
        
        # Results for this branch
        matched_records = []
        unmatched_records = []
        
        # Check each record against PDFs
        for idx, row in branch_records.iterrows():
            matched = False
            matched_pdf = None
            match_type = None
            
            for pdf_file in pdf_files:
                is_match, m_type = self.check_match(pdf_file.name, row)
                if is_match:
                    matched = True
                    matched_pdf = pdf_file.name
                    match_type = m_type
                    break
            
            record_info = {
                'CLIENT_NO': row['CLIENT_NO'],
                'ACCT_DESC': row['ACCT_DESC'],  # Formerly CLIENT_SHORT
                'ACCT_NO': row['ACCT_NO'],
                'ACCT_TYPE_DESC': row['ACCT_TYPE_DESC'],  # Formerly ACCT_DESC
                'ACCT_STATUS': row['ACCT_STATUS'],  # NEW
                'PERFORMED BY': self.normalize_performed_by(row['PERFORMED BY']),  # Normalize here
                'CONVERSION DATE': row['CONVERSION DATE'],  # NEW - pass through
                'MATCHED_PDF': matched_pdf,
                'MATCH_TYPE': match_type
            }
            
            if matched:
                matched_records.append(record_info)
            else:
                unmatched_records.append(record_info)
        
        # Store results
        self.results[branch] = {
            'total_records': total_records,
            'matched_count': len(matched_records),
            'unmatched_count': len(unmatched_records),
            'matched_records': matched_records,
            'unmatched_records': unmatched_records,
            'included_in_assignment': 'Included',  # Mark as included since it's in Excel
            'performed_by': matched_records[0]['PERFORMED BY'] if matched_records else (unmatched_records[0]['PERFORMED BY'] if unmatched_records else 'Unknown')
        }
        
        # Print summary
        print(f"\n✓ Matched: {len(matched_records)}")
        print(f"✗ Unmatched: {len(unmatched_records)}")
        
        return self.results[branch]
    
    def process_not_included_branches(self):
        """Process branches that have PDF folders but are NOT in Excel"""
        # Get all branch folders from PDF directory
        all_pdf_branches = self.get_all_branch_folders()
        
        # Get branches that are in Excel
        excel_branches = set(self.df['BRANCH'].unique())
        
        # Find branches that are NOT in Excel
        not_included_branches = all_pdf_branches - excel_branches
        
        print(f"\nFound {len(not_included_branches)} branches NOT in Excel but have PDF folders")
        
        # Process each not-included branch
        for branch in not_included_branches:
            print(f"\n{'='*80}")
            print(f"Processing NOT INCLUDED Branch: {branch}")
            print(f"{'='*80}")
            
            # Find PDFs in this branch
            pdf_files = self.find_pdfs_in_branch(branch)
            print(f"Total PDF files found in branch folder: {len(pdf_files)}")
            
            # Store results for this branch
            self.results[branch] = {
                'total_records': 0,  # No records in Excel
                'matched_count': 0,  # Cannot match since no Excel records
                'unmatched_count': 0,
                'matched_records': [],
                'unmatched_records': [],
                'included_in_assignment': 'Not Included',
                'pdf_count': len(pdf_files),  # Track PDF count for reference
                'performed_by': 'N/A'  # Not applicable for branches not in Excel
            }
            
            print(f"✓ Branch marked as 'Not Included in Assignment'")
            print(f"  PDF files in folder: {len(pdf_files)}")
    
    def process_all_branches(self):
        """Process all branches in the Excel file - STEP 1: MATCHING"""
        print("\n" + "="*80)
        print("STEP 1: MATCHING CONSENT FORMS TO RECORDS")
        print("="*80)
        
        if self.df is None:
            print("✗ No data loaded. Please load Excel file first.")
            return False
        
        # Get unique branches
        branches = self.df['BRANCH'].unique()
        print(f"\nTotal branches found: {len(branches)}")
        print(f"Branches: {list(branches)}")
        
        # Process each branch
        for branch in branches:
            self.match_branch(branch)
        
        # Process branches NOT in Excel but have PDF folders
        self.process_not_included_branches()
        
        # Generate summary report
        self.generate_summary_report()
        
        # Generate detailed report
        self.generate_detailed_report()
        
        return True
    
    def generate_summary_report(self):
        """Generate a summary report of all branches"""
        print(f"\n\n{'='*80}")
        print("MATCHING SUMMARY REPORT")
        print(f"{'='*80}\n")
        
        total_records = 0
        total_matched = 0
        total_unmatched = 0
        
        summary_data = []
        
        for branch, data in self.results.items():
            total_records += data['total_records']
            total_matched += data['matched_count']
            total_unmatched += data['unmatched_count']
            
            summary_data.append({
                'Branch': branch,
                'Performed By': data.get('performed_by', 'Unknown'),
                'Total Records': data['total_records'],
                'Matched': data['matched_count'],
                'Unmatched': data['unmatched_count'],
                'Match Rate': f"{(data['matched_count']/data['total_records']*100):.2f}%" if data['total_records'] > 0 else "0%",
                'Included in Assignment': data.get('included_in_assignment', 'Included')
            })
            
            print(f"Branch: {branch}")
            print(f"  Performed By: {data.get('performed_by', 'Unknown')}")
            print(f"  Total Records: {data['total_records']}")
            print(f"  Matched: {data['matched_count']}")
            print(f"  Unmatched: {data['unmatched_count']}")
            print(f"  Match Rate: {(data['matched_count']/data['total_records']*100):.2f}%" if data['total_records'] > 0 else "0%")
            print(f"  Included in Assignment: {data.get('included_in_assignment', 'Included')}")
            print()
        
        print(f"{'='*80}")
        print(f"GRAND TOTAL")
        print(f"{'='*80}")
        print(f"Total Records: {total_records}")
        print(f"Total Matched: {total_matched}")
        print(f"Total Unmatched: {total_unmatched}")
        print(f"Overall Match Rate: {(total_matched/total_records*100):.2f}%" if total_records > 0 else "0%")
        
        # Save summary to CSV
        summary_df = pd.DataFrame(summary_data)
        summary_file = os.path.join(self.reports_dir, f"matching_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        summary_df.to_csv(summary_file, index=False)
        print(f"\n✓ Summary report saved to: {summary_file}")
    
    def generate_detailed_report(self):
        """Generate detailed report with all matched and unmatched records"""
        
        # Prepare detailed data
        detailed_data = []
        
        for branch, data in self.results.items():
            # Add matched records
            for record in data['matched_records']:
                # Create hyperlink path for matched PDFs
                performed_by = record.get('PERFORMED BY', 'Unknown')
                branch_code = str(branch).strip()
                if branch_code.replace('.0', '').isdigit():
                    branch_code = branch_code.replace('.0', '').zfill(4)
                
                pdf_name = record['MATCHED_PDF']
                
                # Relative path from Reports folder to Findings folder
                # Reports and Findings are both in OUTPUT folder
                relative_path = f"../Findings/{performed_by}/{branch_code}/{pdf_name}"
                
                detailed_data.append({
                    'BRANCH': branch,
                    'STATUS': 'MATCHED',
                    'CLIENT_NO': record['CLIENT_NO'],
                    'ACCT_DESC': record['ACCT_DESC'],  # Formerly CLIENT_SHORT
                    'ACCT_NO': record['ACCT_NO'],
                    'ACCT_TYPE_DESC': record['ACCT_TYPE_DESC'],  # Formerly ACCT_DESC
                    'ACCT_STATUS': record['ACCT_STATUS'],  # NEW
                    'PERFORMED BY': record['PERFORMED BY'],  # NEW (normalized to uppercase)
                    'CONVERSION DATE': record['CONVERSION DATE'],  # NEW - pass through
                    'MATCHED_PDF': record['MATCHED_PDF'],
                    'MATCH_TYPE': record['MATCH_TYPE'],
                    'OPEN_CONSENT_FORM': relative_path  # NEW - will be converted to hyperlink
                })
            
            # Add unmatched records
            for record in data['unmatched_records']:
                detailed_data.append({
                    'BRANCH': branch,
                    'STATUS': 'UNMATCHED',
                    'CLIENT_NO': record['CLIENT_NO'],
                    'ACCT_DESC': record['ACCT_DESC'],  # Formerly CLIENT_SHORT
                    'ACCT_NO': record['ACCT_NO'],
                    'ACCT_TYPE_DESC': record['ACCT_TYPE_DESC'],  # Formerly ACCT_DESC
                    'ACCT_STATUS': record['ACCT_STATUS'],  # NEW
                    'PERFORMED BY': record['PERFORMED BY'],  # NEW (normalized to uppercase)
                    'CONVERSION DATE': record['CONVERSION DATE'],  # NEW - pass through
                    'MATCHED_PDF': '',
                    'MATCH_TYPE': '',
                    'OPEN_CONSENT_FORM': ''  # Empty for unmatched
                })
        
        # Save to Excel with multiple sheets
        detailed_file = os.path.join(self.reports_dir, f"matching_detailed_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        
        with pd.ExcelWriter(detailed_file, engine='openpyxl') as writer:
            # All records sheet
            all_df = pd.DataFrame(detailed_data)
            all_df.to_excel(writer, sheet_name='All Records', index=False)
            
            # Add hyperlinks to the OPEN_CONSENT_FORM column for All Records sheet
            workbook = writer.book
            worksheet = writer.sheets['All Records']
            self._add_hyperlinks_to_sheet(worksheet, all_df, 'OPEN_CONSENT_FORM')
            
            # Matched records only
            matched_df = all_df[all_df['STATUS'] == 'MATCHED']
            matched_df.to_excel(writer, sheet_name='Matched Records', index=False)
            worksheet_matched = writer.sheets['Matched Records']
            self._add_hyperlinks_to_sheet(worksheet_matched, matched_df, 'OPEN_CONSENT_FORM')
            
            # Unmatched records only
            unmatched_df = all_df[all_df['STATUS'] == 'UNMATCHED']
            unmatched_df.to_excel(writer, sheet_name='Unmatched Records', index=False)
            
            # Branch-wise sheets
            for branch in self.results.keys():
                branch_df = all_df[all_df['BRANCH'] == branch]
                safe_sheet_name = str(branch)[:31]  # Excel sheet name limit
                branch_df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
                worksheet_branch = writer.sheets[safe_sheet_name]
                self._add_hyperlinks_to_sheet(worksheet_branch, branch_df, 'OPEN_CONSENT_FORM')
        
        print(f"✓ Detailed report saved to: {detailed_file}")
        
        # Store the matched excel file path for next step
        self.matched_excel_file = detailed_file
        
        # Also save JSON for programmatic access
        json_file = os.path.join(self.reports_dir, f"matching_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        
        # Convert numpy int64 keys to regular Python int for JSON serialization
        json_results = {str(k): v for k, v in self.results.items()}
        
        # Convert Pandas Timestamps to JSON-serializable format
        json_results = self._convert_timestamps_to_json(json_results)
        
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(json_results, f, indent=2, ensure_ascii=False)
        print(f"✓ JSON results saved to: {json_file}")
    
    def _convert_timestamps_to_json(self, obj):
        """
        Recursively convert Pandas Timestamp objects to ISO format strings for JSON serialization
        """
        if isinstance(obj, pd.Timestamp):
            return obj.isoformat()
        elif isinstance(obj, dict):
            return {key: self._convert_timestamps_to_json(value) for key, value in obj.items()}
        elif isinstance(obj, list):
            return [self._convert_timestamps_to_json(item) for item in obj]
        elif isinstance(obj, tuple):
            return tuple(self._convert_timestamps_to_json(item) for item in obj)
        else:
            return obj
    
    def _add_hyperlinks_to_sheet(self, worksheet, dataframe, column_name):
        """Add Excel hyperlinks to a specific column in a worksheet"""
        try:
            # Find the column index for OPEN_CONSENT_FORM
            if column_name not in dataframe.columns:
                return
            
            col_idx = list(dataframe.columns).index(column_name) + 1  # +1 because Excel is 1-indexed
            
            # Iterate through rows and add hyperlinks
            for row_idx in range(2, len(dataframe) + 2):  # Start from row 2 (after header)
                cell = worksheet.cell(row=row_idx, column=col_idx)
                hyperlink_path = cell.value
                
                if hyperlink_path and hyperlink_path.strip() and hyperlink_path != '':
                    # Create hyperlink
                    cell.hyperlink = hyperlink_path
                    cell.value = "Open PDF"
                    cell.style = "Hyperlink"
        except Exception as e:
            print(f"Warning: Could not add hyperlinks: {e}")
    
    def copy_matched_pdfs(self):
        """Copy matched PDF files to destination - STEP 2: COPYING"""
        print("\n\n" + "="*80)
        print("STEP 2: COPYING MATCHED PDF FILES")
        print("="*80)
        print()
        
        if not self.matched_excel_file or not os.path.exists(self.matched_excel_file):
            print("✗ Error: No matched Excel file found. Please run matching first.")
            return False
        
        # Read the matched records sheet
        matched_df = pd.read_excel(self.matched_excel_file, sheet_name='Matched Records')
        
        print(f"Total matched records to copy: {len(matched_df)}")
        print()
        
        # Counter for tracking
        total_files = 0
        success_count = 0
        error_count = 0
        errors = []
        
        # Process each row
        for index, row in matched_df.iterrows():
            try:
                # Get branch code and normalize to 4 digits
                branch_code = str(row['BRANCH']).strip()
                
                # Handle numeric branch codes
                if branch_code.replace('.0', '').isdigit():
                    branch_code = branch_code.replace('.0', '')
                    branch_code = branch_code.zfill(4)  # Pad with leading zeros to make it 4 digits
                
                # Get PDF name
                pdf_name = str(row['MATCHED_PDF']).strip()
                
                # Get "Performed by" value for folder organization - normalize it
                performed_by = self.normalize_performed_by(row['PERFORMED BY'])
                
                # Skip if no valid data
                if pd.isna(row['BRANCH']) or pd.isna(row['MATCHED_PDF']) or pdf_name == 'nan':
                    continue
                
                # Try to find the source file in various branch folder formats
                source_file = None
                branch_formats = [branch_code, branch_code.lstrip('0') or '0']
                
                for branch_format in branch_formats:
                    test_folder = os.path.join(self.consent_forms_dir, branch_format)
                    if os.path.exists(test_folder):
                        # Search recursively for the PDF
                        for root, dirs, files in os.walk(test_folder):
                            if pdf_name in files:
                                source_file = os.path.join(root, pdf_name)
                                break
                        if source_file:
                            break
                
                if not source_file:
                    error_msg = f"[ERROR] Not found: Branch {branch_code}, PDF: {pdf_name}"
                    print(error_msg)
                    errors.append(error_msg)
                    error_count += 1
                    continue
                
                # Destination paths: Findings/{Performed by}/{Branch}/
                dest_folder = os.path.join(self.findings_dir, performed_by, branch_code)
                dest_file = os.path.join(dest_folder, pdf_name)
                
                # Create destination folder
                os.makedirs(dest_folder, exist_ok=True)
                
                # Copy the file
                shutil.copy2(source_file, dest_file)
                print(f"[OK] Copied: {performed_by}/{branch_code}/{pdf_name}")
                success_count += 1
                
                total_files += 1
                
            except Exception as e:
                error_msg = f"[ERROR] Error processing row {index}: {str(e)}"
                print(error_msg)
                errors.append(error_msg)
                error_count += 1
        
        # Summary
        print("\n" + "="*80)
        print("COPY SUMMARY")
        print("="*80)
        print(f"Total files processed: {total_files}")
        print(f"Successfully copied: {success_count}")
        print(f"Errors: {error_count}")
        
        if errors:
            print("\nErrors encountered:")
            for error in errors[:10]:  # Show first 10 errors
                print(f"  {error}")
            if len(errors) > 10:
                print(f"  ... and {len(errors) - 10} more errors")
        
        return True
    
    def verify_copy_results(self):
        """Verify the copy operation - STEP 3: VERIFICATION"""
        print("\n\n" + "="*80)
        print("STEP 3: VERIFICATION OF COPIED FILES")
        print("="*80)
        print()
        
        # Count files in destination folder
        total_copied = 0
        performer_folders = []
        
        if not os.path.exists(self.findings_dir):
            print(f"✗ Error: Findings directory not found: {self.findings_dir}")
            return False
        
        # Count files organized by Performed by -> Branch
        for performed_by in os.listdir(self.findings_dir):
            performed_by_path = os.path.join(self.findings_dir, performed_by)
            if os.path.isdir(performed_by_path):
                for branch in os.listdir(performed_by_path):
                    branch_path = os.path.join(performed_by_path, branch)
                    if os.path.isdir(branch_path):
                        file_count = len([f for f in os.listdir(branch_path) if f.endswith('.pdf')])
                        total_copied += file_count
                        performer_folders.append((performed_by, branch, file_count))
        
        # Read Excel to count expected (only matched records)
        if not self.matched_excel_file or not os.path.exists(self.matched_excel_file):
            print("✗ Error: No matched Excel file found.")
            return False
        
        matched_df = pd.read_excel(self.matched_excel_file, sheet_name='Matched Records')
        total_expected = len(matched_df)
        
        print("="*80)
        print("VERIFICATION RESULTS")
        print("="*80)
        print(f"Total matched records (expected): {total_expected}")
        print(f"Total PDF files copied (actual): {total_copied}")
        print(f"Total folders created: {len(performer_folders)}")
        
        if total_expected == total_copied:
            print("\n✓✓✓ SUCCESS! All files copied correctly! ✓✓✓")
        else:
            difference = abs(total_expected - total_copied)
            if total_copied < total_expected:
                print(f"\n⚠ WARNING: {difference} files are MISSING!")
            else:
                print(f"\n⚠ WARNING: {difference} EXTRA files were copied!")
        
        print("\nPerformed by / Branch folders and file counts:")
        for performed_by, branch, count in sorted(performer_folders):
            print(f"  {performed_by}/{branch}: {count} files")
        
        return True
    
    def run_complete_workflow(self):
        """Run the complete workflow: Match -> Copy -> Verify"""
        print("\n" + "="*80)
        print("CONSENT FORM COMPLETE WORKFLOW")
        print("="*80)
        print(f"\nExcel File: {self.excel_path}")
        print(f"Consent Forms Directory: {self.consent_forms_dir}")
        print(f"Output Directory: {self.output_dir}")
        print(f"  - Reports Directory: {self.reports_dir}")
        print(f"  - Findings Directory: {self.findings_dir}")
        print()
        
        # Check if files exist
        if not os.path.exists(self.excel_path):
            print(f"✗ Error: Excel file not found at {self.excel_path}")
            return False
        
        if not os.path.exists(self.consent_forms_dir):
            print(f"✗ Error: Consent forms directory not found at {self.consent_forms_dir}")
            return False
        
        # STEP 1: Load and Match
        if not self.load_excel():
            return False
        
        if not self.process_all_branches():
            return False
        
        # STEP 2: Copy matched PDFs
        if not self.copy_matched_pdfs():
            return False
        
        # STEP 3: Verify results
        if not self.verify_copy_results():
            return False
        
        print("\n" + "="*80)
        print("WORKFLOW COMPLETED SUCCESSFULLY!")
        print("="*80)
        print(f"\nAll reports saved to: {self.reports_dir}")
        print(f"All matched PDFs copied to: {self.findings_dir}")
        print("\nGenerated files:")
        print("  - matching_summary_*.csv: Branch-wise summary")
        print("  - matching_detailed_report_*.xlsx: Detailed report with multiple sheets and hyperlinks")
        print("  - matching_results_*.json: JSON output for programmatic access")
        print(f"  - Copied PDFs in: {self.findings_dir}")
        print("    Structure: Findings/{Performed by}/{Branch}/PDFs")
        print()
        
        return True


def main():
    """Main function to run the complete consent form workflow"""
    
    print("="*80)
    print("CONSENT FORM COMPLETE AUTOMATION SYSTEM")
    print("="*80)
    print()
    
    # Get user inputs
    print("Please provide the following paths:")
    print()
    
    excel_file = input("1. Enter the path to Excel file (with client records): ").strip().strip('"').strip("'")
    if not excel_file:
        print("✗ Error: Excel file path cannot be empty")
        return
    
    consent_forms_dir = input("2. Enter the path to Consent Forms directory (source): ").strip().strip('"').strip("'")
    if not consent_forms_dir:
        print("✗ Error: Consent Forms directory path cannot be empty")
        return
    
    # Optional: Allow custom output directory or use default (OUTPUT folder next to script)
    print("\n3. OUTPUT folder location:")
    print("   Press Enter to create 'OUTPUT' folder in the same directory as this script")
    print("   OR enter a custom path for OUTPUT folder:")
    output_dir = input("   > ").strip().strip('"').strip("'")
    
    if not output_dir:
        output_dir = None  # Will use default (OUTPUT folder next to script)
    
    # Initialize and run workflow
    workflow = ConsentFormCompleteWorkflow(
        excel_file,
        consent_forms_dir,
        output_dir
    )
    
    # Run the complete workflow
    workflow.run_complete_workflow()


if __name__ == "__main__":
    main()
