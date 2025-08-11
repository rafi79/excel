import streamlit as st
import json
import os
from typing import List, Dict, Any, Optional, Tuple
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter, column_index_from_string
import google.generativeai as genai
import tempfile
import logging
from io import BytesIO
import pandas as pd
import math
import re
from collections import Counter, defaultdict
from datetime import datetime
import numpy as np
import copy

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Page config
st.set_page_config(
    page_title="📊 Excel JSON AI Editor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

class ExcelToJSONProcessor:
    """Converts Excel files to JSON format for LLM processing"""
    
    def __init__(self):
        self.original_workbooks = {}
        self.json_data = {}
        self.file_paths = {}
        self.metadata = {}
    
    def excel_to_json(self, uploaded_files) -> Dict[str, Any]:
        """Convert Excel files to structured JSON"""
        results = {}
        
        for uploaded_file in uploaded_files:
            try:
                # Save to temporary file
                with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_file:
                    tmp_file.write(uploaded_file.read())
                    tmp_path = tmp_file.name
                
                self.file_paths[uploaded_file.name] = tmp_path
                
                # Load workbook
                workbook = load_workbook(tmp_path, data_only=False)
                self.original_workbooks[uploaded_file.name] = workbook
                
                # Convert to JSON structure
                file_json = {
                    'file_name': uploaded_file.name,
                    'sheets': {},
                    'metadata': {
                        'total_sheets': len(workbook.sheetnames),
                        'sheet_names': workbook.sheetnames
                    }
                }
                
                for sheet_name in workbook.sheetnames:
                    sheet_json = self._sheet_to_json(workbook[sheet_name], sheet_name)
                    file_json['sheets'][sheet_name] = sheet_json
                
                self.json_data[uploaded_file.name] = file_json
                results[uploaded_file.name] = file_json
                
            except Exception as e:
                logger.error(f"Error processing {uploaded_file.name}: {str(e)}")
                results[uploaded_file.name] = {'error': str(e)}
        
        return results
    
    def _sheet_to_json(self, sheet, sheet_name: str) -> Dict[str, Any]:
        """Convert Excel sheet to JSON structure"""
        try:
            if not sheet.max_row or sheet.max_row == 0:
                return {
                    'sheet_name': sheet_name,
                    'rows': [],
                    'metadata': {
                        'max_row': 0,
                        'max_col': 0,
                        'total_cells': 0
                    }
                }
            
            max_row = sheet.max_row
            max_col = sheet.max_column or 0
            
            # Convert all cells to JSON
            rows = []
            for row_num in range(1, max_row + 1):
                row_data = {
                    'row_number': row_num,
                    'cells': {}
                }
                
                has_data = False
                for col_num in range(1, max_col + 1):
                    cell = sheet.cell(row=row_num, column=col_num)
                    col_letter = get_column_letter(col_num)
                    cell_ref = f"{col_letter}{row_num}"
                    
                    cell_data = {
                        'cell_ref': cell_ref,
                        'value': None,
                        'formula': None,
                        'data_type': cell.data_type,
                        'row': row_num,
                        'column': col_num,
                        'column_letter': col_letter
                    }
                    
                    if cell.value is not None:
                        cell_data['value'] = str(cell.value)
                        has_data = True
                        
                        # Store formula if it's a formula cell
                        if cell.data_type == 'f':
                            cell_data['formula'] = str(cell.value)
                            # Try to get calculated value
                            try:
                                calc_wb = load_workbook(self.file_paths[list(self.file_paths.keys())[0]], data_only=True)
                                calc_cell = calc_wb[sheet_name].cell(row=row_num, column=col_num)
                                if calc_cell.value is not None:
                                    cell_data['calculated_value'] = str(calc_cell.value)
                            except:
                                pass
                    
                    row_data['cells'][cell_ref] = cell_data
                
                if has_data:
                    rows.append(row_data)
            
            return {
                'sheet_name': sheet_name,
                'rows': rows,
                'metadata': {
                    'max_row': max_row,
                    'max_col': max_col,
                    'total_cells': max_row * max_col,
                    'data_rows': len(rows)
                }
            }
            
        except Exception as e:
            logger.error(f"Error converting sheet {sheet_name}: {str(e)}")
            return {
                'sheet_name': sheet_name,
                'rows': [],
                'metadata': {'error': str(e)}
            }
    
    def json_to_excel(self, json_data: Dict[str, Any], original_file_name: str) -> bytes:
        """Convert modified JSON back to Excel file"""
        try:
            # Create new workbook
            new_workbook = Workbook()
            
            # Remove default sheet
            if 'Sheet' in new_workbook.sheetnames:
                new_workbook.remove(new_workbook['Sheet'])
            
            file_data = json_data
            
            for sheet_name, sheet_data in file_data['sheets'].items():
                # Create new worksheet
                ws = new_workbook.create_sheet(title=sheet_name)
                
                # Populate cells from JSON
                for row_data in sheet_data['rows']:
                    for cell_ref, cell_data in row_data['cells'].items():
                        if cell_data['value'] is not None:
                            row_num = cell_data['row']
                            col_num = cell_data['column']
                            
                            cell = ws.cell(row=row_num, column=col_num)
                            
                            # Set value (handle formulas)
                            if cell_data.get('formula'):
                                cell.value = cell_data['formula']
                            else:
                                # Try to convert to appropriate type
                                value = cell_data['value']
                                try:
                                    if value.replace('.', '').replace('-', '').isdigit():
                                        value = float(value) if '.' in value else int(value)
                                except:
                                    pass  # Keep as string
                                cell.value = value
            
            # Save to bytes
            output = BytesIO()
            new_workbook.save(output)
            output.seek(0)
            return output.getvalue()
            
        except Exception as e:
            logger.error(f"Error converting JSON to Excel: {str(e)}")
            raise e

class LLMJSONProcessor:
    """Processes JSON data with LLM and returns modified JSON"""
    
    def __init__(self, api_key: str):
        genai.configure(api_key=api_key)
        self.model = "gemini-2.0-flash-exp"
    
    def process_json_with_llm(self, json_data: Dict[str, Any], user_request: str) -> Dict[str, Any]:
        """Process JSON data with LLM instructions"""
        try:
            # Convert JSON to LLM-friendly format
            llm_format = self._json_to_llm_format(json_data)
            
            # Create LLM prompt
            prompt = self._create_llm_prompt(llm_format, user_request)
            
            # Get LLM response
            model = genai.GenerativeModel(self.model)
            response = model.generate_content(prompt)
            
            # Parse LLM response and apply changes
            modified_json = self._apply_llm_changes(json_data, response.text, user_request)
            
            return {
                'success': True,
                'modified_json': modified_json,
                'llm_response': response.text,
                'changes_applied': True
            }
            
        except Exception as e:
            return {
                'success': False,
                'error': str(e),
                'modified_json': json_data
            }
    
    def _json_to_llm_format(self, json_data: Dict[str, Any]) -> str:
        """Convert JSON to LLM-readable tabular format"""
        llm_text = []
        
        for file_name, file_data in json_data.items():
            if 'error' in file_data:
                continue
                
            llm_text.append(f"FILE: {file_name}")
            llm_text.append(f"Total Sheets: {file_data['metadata']['total_sheets']}")
            llm_text.append(f"Sheet Names: {', '.join(file_data['metadata']['sheet_names'])}")
            llm_text.append("")
            
            for sheet_name, sheet_data in file_data['sheets'].items():
                llm_text.append(f"SHEET: {sheet_name}")
                llm_text.append(f"Rows: {sheet_data['metadata']['data_rows']}")
                
                # Create table format for LLM
                llm_text.append("CELL DATA:")
                
                for row_data in sheet_data['rows'][:20]:  # Limit to first 20 rows for LLM
                    for cell_ref, cell_data in row_data['cells'].items():
                        if cell_data['value']:
                            if cell_data.get('formula'):
                                llm_text.append(f"  {cell_ref}: {cell_data['value']} (Formula: {cell_data['formula']})")
                            else:
                                llm_text.append(f"  {cell_ref}: {cell_data['value']}")
                
                llm_text.append("")
        
        return "\n".join(llm_text)
    
    def _create_llm_prompt(self, llm_format: str, user_request: str) -> str:
        """Create comprehensive LLM prompt"""
        return f"""
You are an Excel editing AI. You can view Excel data in JSON format and provide specific editing instructions.

EXCEL DATA:
{llm_format}

USER REQUEST: {user_request}

Your task is to analyze the Excel data and provide specific editing instructions to fulfill the user's request.

IMPORTANT RULES:
1. You must provide specific cell references (like A1, B5, C10) for edits
2. You must specify the exact new values to be placed in cells
3. Format your response as JSON with this structure:

{{
    "analysis": "Your analysis of what needs to be changed",
    "changes": [
        {{
            "file_name": "filename.xlsx",
            "sheet_name": "Sheet1",
            "cell_ref": "A1",
            "old_value": "current value",
            "new_value": "new value",
            "reason": "why this change is needed"
        }}
    ],
    "summary": "Summary of all changes made"
}}

CRITICAL: Respond ONLY with valid JSON. Do not include any text outside the JSON structure.
"""
    
    def _apply_llm_changes(self, original_json: Dict[str, Any], llm_response: str, user_request: str) -> Dict[str, Any]:
        """Apply LLM-suggested changes to JSON data"""
        try:
            # Parse LLM response
            llm_response_clean = llm_response.strip()
            if llm_response_clean.startswith('```json'):
                llm_response_clean = llm_response_clean[7:]
            if llm_response_clean.endswith('```'):
                llm_response_clean = llm_response_clean[:-3]
            
            llm_data = json.loads(llm_response_clean)
            
            # Create a deep copy for modifications
            modified_json = copy.deepcopy(original_json)
            
            # Apply each change
            changes_made = []
            for change in llm_data.get('changes', []):
                try:
                    file_name = change['file_name']
                    sheet_name = change['sheet_name']
                    cell_ref = change['cell_ref']
                    new_value = change['new_value']
                    
                    # Find and update the cell in JSON
                    if file_name in modified_json:
                        if sheet_name in modified_json[file_name]['sheets']:
                            sheet_data = modified_json[file_name]['sheets'][sheet_name]
                            
                            # Find the row containing this cell
                            for row_data in sheet_data['rows']:
                                if cell_ref in row_data['cells']:
                                    old_value = row_data['cells'][cell_ref]['value']
                                    row_data['cells'][cell_ref]['value'] = str(new_value)
                                    
                                    changes_made.append({
                                        'location': f"{file_name}/{sheet_name}/{cell_ref}",
                                        'old_value': old_value,
                                        'new_value': str(new_value)
                                    })
                                    break
                
                except Exception as e:
                    logger.error(f"Error applying change: {e}")
                    continue
            
            # Add metadata about changes
            if not hasattr(modified_json, '_changes'):
                for file_name in modified_json:
                    if isinstance(modified_json[file_name], dict):
                        modified_json[file_name]['_changes'] = changes_made
            
            return modified_json
            
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse LLM response as JSON: {e}")
            return original_json
        except Exception as e:
            logger.error(f"Error applying changes: {e}")
            return original_json

class ExcelJSONViewer:
    """Creates viewers for JSON-based Excel data"""
    
    @staticmethod
    def create_json_viewer(json_data: Dict[str, Any], changes_made: List[Dict] = None):
        """Create a viewer for JSON Excel data"""
        if not json_data:
            st.info("No data to display")
            return
        
        # Show file selector
        file_names = [f for f in json_data.keys() if not f.startswith('_')]
        if not file_names:
            st.info("No valid files found")
            return
        
        selected_file = st.selectbox("📁 Select File:", file_names, key="json_viewer_file")
        
        if selected_file and selected_file in json_data:
            file_data = json_data[selected_file]
            
            if 'error' in file_data:
                st.error(f"Error in file: {file_data['error']}")
                return
            
            # Show file metadata
            st.subheader(f"📊 {selected_file}")
            col1, col2 = st.columns(2)
            col1.metric("Total Sheets", file_data['metadata']['total_sheets'])
            col2.metric("Available Sheets", len(file_data['sheets']))
            
            # Sheet selector
            sheet_names = list(file_data['sheets'].keys())
            selected_sheet = st.selectbox("📋 Select Sheet:", sheet_names, key="json_viewer_sheet")
            
            if selected_sheet:
                sheet_data = file_data['sheets'][selected_sheet]
                
                # Show sheet metadata
                col1, col2, col3 = st.columns(3)
                col1.metric("Data Rows", sheet_data['metadata']['data_rows'])
                col2.metric("Max Columns", sheet_data['metadata']['max_col'])
                col3.metric("Total Cells", sheet_data['metadata']['total_cells'])
                
                # Create DataFrame for display
                ExcelJSONViewer._display_sheet_data(sheet_data, changes_made, selected_file, selected_sheet)
    
    @staticmethod
    def _display_sheet_data(sheet_data: Dict[str, Any], changes_made: List[Dict], file_name: str, sheet_name: str):
        """Display sheet data in Excel-like format"""
        try:
            # Get all cell references to determine table size
            all_cells = {}
            max_row = 0
            max_col = 0
            
            for row_data in sheet_data['rows']:
                for cell_ref, cell_data in row_data['cells'].items():
                    if cell_data['value']:
                        all_cells[cell_ref] = cell_data
                        max_row = max(max_row, cell_data['row'])
                        max_col = max(max_col, cell_data['column'])
            
            if not all_cells:
                st.info("No data in this sheet")
                return
            
            # Create changes lookup
            changes_lookup = set()
            if changes_made:
                for change in changes_made:
                    if f"{file_name}/{sheet_name}" in change['location']:
                        cell_ref = change['location'].split('/')[-1]
                        changes_lookup.add(cell_ref)
            
            # Limit display size
            display_rows = min(max_row, 50)
            display_cols = min(max_col, 15)
            
            # Create table data
            table_data = []
            
            # Header row
            header = ["Row"] + [get_column_letter(c) for c in range(1, display_cols + 1)]
            
            # Data rows
            for r in range(1, display_rows + 1):
                row = [str(r)]
                for c in range(1, display_cols + 1):
                    cell_ref = f"{get_column_letter(c)}{r}"
                    
                    if cell_ref in all_cells:
                        value = all_cells[cell_ref]['value']
                        # Highlight changed cells
                        if cell_ref in changes_lookup:
                            value = f"🔥 {value}"
                        row.append(value)
                    else:
                        row.append("")
                
                table_data.append(row)
            
            # Create DataFrame
            df = pd.DataFrame(table_data, columns=header)
            
            # Display
            st.dataframe(
                df,
                use_container_width=True,
                height=600,
                hide_index=True
            )
            
            # Show changes summary
            if changes_lookup:
                st.success(f"🔥 **Changed cells**: {', '.join(sorted(changes_lookup))}")
            
            # Show sample formulas
            formulas = []
            for cell_ref, cell_data in all_cells.items():
                if cell_data.get('formula'):
                    formulas.append((cell_ref, cell_data['formula']))
                    if len(formulas) >= 10:
                        break
            
            if formulas:
                with st.expander(f"🔬 Formulas Found ({len(formulas)})"):
                    for cell_ref, formula in formulas:
                        st.code(f"{cell_ref}: {formula}")
        
        except Exception as e:
            st.error(f"Error displaying sheet data: {str(e)}")

# API Key
API_KEY = "AIzaSyCSDx-q3PgkvMQktdi4tScbT1wOLgZ9jQg"

# Initialize session state
if 'excel_json_processor' not in st.session_state:
    st.session_state.excel_json_processor = ExcelToJSONProcessor()
if 'llm_processor' not in st.session_state:
    st.session_state.llm_processor = LLMJSONProcessor(API_KEY)
if 'json_data' not in st.session_state:
    st.session_state.json_data = {}
if 'modified_json_data' not in st.session_state:
    st.session_state.modified_json_data = {}
if 'files_processed' not in st.session_state:
    st.session_state.files_processed = False
if 'llm_response' not in st.session_state:
    st.session_state.llm_response = ""
if 'changes_made' not in st.session_state:
    st.session_state.changes_made = []

def main():
    st.title("📊 Excel JSON AI Editor")
    st.markdown("**Excel → JSON → LLM Processing → Excel Pipeline**")
    
    # Sidebar
    with st.sidebar:
        st.header("🔧 System Status")
        
        if st.session_state.files_processed:
            file_count = len(st.session_state.json_data)
            st.success(f"✅ {file_count} files processed")
            
            # Show JSON data stats
            total_sheets = 0
            total_rows = 0
            for file_data in st.session_state.json_data.values():
                if isinstance(file_data, dict) and 'metadata' in file_data:
                    total_sheets += file_data['metadata'].get('total_sheets', 0)
                    for sheet_data in file_data.get('sheets', {}).values():
                        total_rows += sheet_data.get('metadata', {}).get('data_rows', 0)
            
            st.info(f"📊 {total_sheets} sheets, {total_rows} data rows")
            
            if st.session_state.changes_made:
                st.info(f"✏️ {len(st.session_state.changes_made)} changes made")
        else:
            st.warning("⚠️ No files processed")
        
        st.divider()
        
        # Pipeline status
        st.header("🔄 Pipeline Status")
        
        # Step 1: Excel → JSON
        if st.session_state.files_processed:
            st.success("✅ 1. Excel → JSON")
        else:
            st.info("⏳ 1. Excel → JSON")
        
        # Step 2: LLM Processing
        if st.session_state.modified_json_data:
            st.success("✅ 2. LLM Processing")
        else:
            st.info("⏳ 2. LLM Processing")
        
        # Step 3: JSON → Excel
        if st.session_state.modified_json_data:
            st.success("✅ 3. JSON → Excel")
        else:
            st.info("⏳ 3. JSON → Excel")
        
        st.divider()
        
        # Quick actions
        if st.session_state.files_processed:
            st.header("🚀 Quick Actions")
            
            # Download modified Excel
            if st.session_state.modified_json_data:
                for file_name in st.session_state.modified_json_data:
                    if not file_name.startswith('_'):
                        try:
                            excel_bytes = st.session_state.excel_json_processor.json_to_excel(
                                {file_name: st.session_state.modified_json_data[file_name]},
                                file_name
                            )
                            
                            st.download_button(
                                f"📥 Download {file_name}",
                                data=excel_bytes,
                                file_name=f"modified_{file_name}",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"download_{file_name}"
                            )
                        except Exception as e:
                            st.error(f"Error preparing {file_name}: {str(e)}")
            
            # Download JSON data
            if st.session_state.json_data:
                json_str = json.dumps(st.session_state.json_data, indent=2)
                st.download_button(
                    "📥 Download JSON Data",
                    data=json_str,
                    file_name="excel_data.json",
                    mime="application/json"
                )
        
        # Clear all
        if st.button("🗑️ Clear All", type="secondary"):
            for key in ['json_data', 'modified_json_data', 'files_processed', 'llm_response', 'changes_made']:
                st.session_state[key] = {} if 'data' in key else [] if 'changes' in key else False if 'processed' in key else ""
            st.session_state.excel_json_processor = ExcelToJSONProcessor()
            st.rerun()
    
    # Main content
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.header("📁 Step 1: Upload Excel Files")
        uploaded_files = st.file_uploader(
            "Upload Excel files for JSON conversion",
            type=['xlsx', 'xls'],
            accept_multiple_files=True,
            help="Files will be converted to JSON for LLM processing"
        )
        
        if uploaded_files:
            st.success(f"📄 {len(uploaded_files)} file(s) ready")
            
            if st.button("🔄 Convert Excel → JSON", type="primary"):
                with st.spinner("🔄 Converting Excel files to JSON..."):
                    try:
                        json_data = st.session_state.excel_json_processor.excel_to_json(uploaded_files)
                        st.session_state.json_data = json_data
                        st.session_state.files_processed = True
                        
                        st.success("✅ Excel files converted to JSON!")
                        
                        # Show conversion summary
                        total_sheets = sum(data.get('metadata', {}).get('total_sheets', 0) 
                                         for data in json_data.values() 
                                         if isinstance(data, dict))
                        st.info(f"📊 Converted {len(json_data)} files with {total_sheets} sheets")
                        
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"❌ Error converting files: {str(e)}")
    
    with col2:
        st.header("🤖 Step 2: LLM Processing")
        
        user_request = st.text_area(
            "Tell AI what changes to make:",
            placeholder="""Examples:
• Replace 'BEUMER India Pvt. Ltd.' with 'BEUMER Bangladesh Pvt. Ltd.' in all cells
• Change all instances of 'India' to 'Bangladesh'
• Update company name in header cells
• Find and replace specific text across all sheets
• Modify values in specific cells""",
            height=120,
            help="The LLM will process JSON data and return modified JSON"
        )
        
        if st.button("🎯 Process with LLM", type="primary", disabled=not (st.session_state.files_processed and user_request)):
            with st.spinner("🤖 LLM is processing JSON data..."):
                try:
                    result = st.session_state.llm_processor.process_json_with_llm(
                        st.session_state.json_data, 
                        user_request
                    )
                    
                    if result['success']:
                        st.session_state.modified_json_data = result['modified_json']
                        st.session_state.llm_response = result['llm_response']
                        
                        # Extract changes made
                        changes_made = []
                        for file_name, file_data in result['modified_json'].items():
                            if isinstance(file_data, dict) and '_changes' in file_data:
                                changes_made.extend(file_data['_changes'])
                        
                        st.session_state.changes_made = changes_made
                        
                        st.success("✅ LLM processing completed!")
                        st.info(f"📝 Made {len(changes_made)} changes")
                        
                    else:
                        st.error(f"❌ LLM processing failed: {result['error']}")
                    
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
    
    # Step 3: Display Results
    if st.session_state.files_processed:
        st.divider()
        st.header("📊 Step 3: View Results")
        
        tab1, tab2, tab3 = st.tabs(["📋 Original JSON Data", "🔄 Modified JSON Data", "📝 LLM Response"])
        
        with tab1:
            st.subheader("📋 Original Excel Data (JSON Format)")
            if st.session_state.json_data:
                ExcelJSONViewer.create_json_viewer(st.session_state.json_data)
            else:
                st.info("No original data available")
        
        with tab2:
            st.subheader("🔄 Modified Excel Data (JSON Format)")
            if st.session_state.modified_json_data:
                ExcelJSONViewer.create_json_viewer(
                    st.session_state.modified_json_data, 
                    st.session_state.changes_made
                )
                
                # Show changes summary
                if st.session_state.changes_made:
                    st.subheader("✏️ Changes Made:")
                    for change in st.session_state.changes_made:
                        st.success(f"✅ **{change['location']}**: `{change['old_value']}` → `{change['new_value']}`")
            else:
                st.info("No modified data available - run LLM processing first")
        
        with tab3:
            st.subheader("📝 LLM Analysis & Response")
            if st.session_state.llm_response:
                st.text_area(
                    "LLM Response:",
                    value=st.session_state.llm_response,
                    height=300,
                    disabled=True
                )
            else:
                st.info("No LLM response available")
    
    # Instructions
    if not st.session_state.files_processed:
        st.divider()
        st.header("💡 How the Pipeline Works")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.subheader("🔄 Step 1: Excel → JSON")
            st.markdown("""
            **Conversion Process:**
            - Reads Excel files with openpyxl
            - Preserves formulas and formatting
            - Converts to structured JSON
            - Each cell becomes a JSON object
            - Maintains cell references (A1, B5, etc.)
            """)
        
        with col2:
            st.subheader("🤖 Step 2: LLM Processing")
            st.markdown("""
            **LLM Analysis:**
            - Converts JSON to readable format
            - LLM analyzes tabular data
            - Understands user requests
            - Generates specific edit instructions
            - Returns structured JSON response
            - Applies changes to JSON data
            """)
        
        with col3:
            st.subheader("📊 Step 3: JSON → Excel")
            st.markdown("""
            **Back to Excel:**
            - Converts modified JSON to Excel
            - Recreates workbook structure
            - Preserves formulas and types
            - Maintains cell references
            - Generates downloadable file
            - Shows visual changes
            """)
        
        st.subheader("🎯 Advantages of This Approach")
        
        advantages = [
            "**Complete Data Preservation**: All formulas, formatting, and structure maintained",
            "**LLM Compatibility**: JSON format is perfect for LLM understanding",
            "**Precise Editing**: Exact cell-level modifications with full control",
            "**Scalability**: Can handle large Excel files efficiently",
            "**Transparency**: Every change is tracked and visible",
            "**Reversibility**: Original data is preserved throughout the process"
        ]
        
        for advantage in advantages:
            st.markdown(f"✅ {advantage}")
        
        st.subheader("📝 Example Workflow")
        st.markdown("""
        **Scenario**: Change company name from "BEUMER India Pvt. Ltd." to "BEUMER Bangladesh Pvt. Ltd."
        
        1. **Excel → JSON**: 
           ```json
           {
             "A1": {"value": "BEUMER India Pvt. Ltd.", "cell_ref": "A1"}
           }
           ```
        
        2. **LLM Processing**: 
           - LLM analyzes: "Find cells containing 'BEUMER India Pvt. Ltd.'"
           - LLM generates: "Change A1 from 'BEUMER India Pvt. Ltd.' to 'BEUMER Bangladesh Pvt. Ltd.'"
        
        3. **JSON → Excel**: 
           ```json
           {
             "A1": {"value": "BEUMER Bangladesh Pvt. Ltd.", "cell_ref": "A1"}
           }
           ```
           - Converts back to Excel with changes applied
        """)

if __name__ == "__main__":
    main()
