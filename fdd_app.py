import streamlit as st
import pandas as pd
import json
import warnings
import re
import os
import datetime
import time
from tqdm import tqdm
from fdd_utils.mappings import (
    KEY_TO_SECTION_MAPPING,
    KEY_TERMS_BY_KEY,
)
from fdd_utils.category_config import (
    DISPLAY_NAME_MAPPING_DEFAULT,
    DISPLAY_NAME_MAPPING_NB_NJ,
)
from fdd_utils.excel_processing import (
    detect_latest_date_column,
    parse_accounting_table,
    process_and_filter_excel,
    get_worksheet_sections_by_keys,
    create_improved_table_markdown
)
from fdd_utils.data_utils import (
    get_tab_name,
    get_financial_keys,
    get_key_display_name,
    format_date_to_dd_mmm_yyyy,
    load_config_files
)
from fdd_utils.content_utils import (
    display_bs_content_by_key,
    clean_content_quotes,
    load_json_content,
    parse_markdown_to_json,
    get_content_from_json,
    generate_content_from_session_storage,
    generate_markdown_from_ai_results,
    convert_sections_to_markdown,
    update_bs_content_with_agent_corrections,
    read_bs_content_by_key
)
from fdd_utils.display_utils import (
    display_ai_content_by_key,
    display_ai_prompt_by_key,
    display_sequential_agent_results,
    display_before_after_comparison,
    display_step_by_step_comparison,
    display_validation_comparison,
    calculate_content_similarity,
    show_text_differences
)
from fdd_utils.general_utils import (
    write_prompt_debug_content,
    calculate_text_metrics,
    format_file_size,
    safe_json_load,
    safe_json_dump,
    create_backup_file,
    validate_file_exists,
    get_file_modification_time,
    log_processing_step,
    validate_text_content
)
from pathlib import Path
from tabulate import tabulate
import urllib3
import shutil

# Disable Python bytecode generation to prevent __pycache__ issues
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
from common.pptx_export import export_pptx, merge_presentations
# Import assistant modules at module level to prevent runtime import issues
from common.assistant import process_keys, QualityAssuranceAgent, DataValidationAgent, PatternValidationAgent, find_financial_figures_with_context_check, get_tab_name, get_financial_figure, load_ip, ProofreadingAgent


import uuid
import tempfile

# Suppress warnings
urllib3.disable_warnings()
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=UserWarning)
warnings.filterwarnings('ignore', message='Data Validation extension is not supported and will be removed', category=UserWarning, module='openpyxl')

# Suppress Streamlit file watcher errors
import logging
logging.getLogger('streamlit.watcher.event_based_path_watcher').setLevel(logging.ERROR)
logging.getLogger('streamlit.watcher.util').setLevel(logging.ERROR)

# AI Agent Logging System
class AIAgentLogger:
    """Single-file JSON logging per session (inputs, outputs, errors)."""
    
    def __init__(self):
        self.logs = {'agent1': {}, 'agent2': {}, 'agent3': {}}
        self.session_logs = []
        self.log_dir = Path("logging")
        self.log_dir.mkdir(exist_ok=True)
        self.session_id = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        # Consolidated session JSON file (only file we keep)
        self.session_file = self.log_dir / f"session_{self.session_id}.json"
        
    def _write_to_file(self, message):
        """No-op: we only keep consolidated JSON logs."""
        return
    
    def _save_json_log(self, log_entry, log_type):
        """Disabled: per-entry JSON files are not used."""
        return
    
    def _make_json_serializable(self, obj):
        """Convert DataFrame and other non-serializable objects to strings"""
        if hasattr(obj, 'to_dict'):  # DataFrame
            return str(obj)
        elif isinstance(obj, dict):
            return {k: self._make_json_serializable(v) for k, v in obj.items()}
        elif isinstance(obj, list):
            return [self._make_json_serializable(item) for item in obj]
        else:
            try:
                json.dumps(obj)  # Test if serializable
                return obj
            except (TypeError, ValueError):
                return str(obj)
    
    def _save_key_log(self, agent_name, key, input_entry, output_entry):
        """Disabled: per-key JSON files are not used."""
        return
        
    def _save_to_consolidated_log(self, log_entry):
        """Save log entry to consolidated session file (1 file per session with all agents and keys)"""
        try:
            session_file = self.session_file
            # Load existing session data or create new
            session_data = {'session_info': {'session_id': self.session_id, 'started': datetime.datetime.now().isoformat()}, 'logs': []}
            if session_file.exists():
                try:
                    with open(session_file, 'r', encoding='utf-8') as f:
                        session_data = json.load(f)
                except (json.JSONDecodeError, FileNotFoundError):
                    pass
            
            # Add new log entry
            session_data['logs'].append(log_entry)
            session_data['session_info']['last_updated'] = datetime.datetime.now().isoformat()
            session_data['session_info']['total_logs'] = len(session_data['logs'])
            
            # Save updated session data
            with open(session_file, 'w', encoding='utf-8') as f:
                json.dump(session_data, f, indent=2, ensure_ascii=False)
                
        except Exception as e:
            print(f"Error saving to consolidated log: {e}")
    
    def log_agent_input(self, agent_name, key, system_prompt, user_prompt, context_data=None, actual_prompts=None):
        """Log agent input prompts and data to consolidated session file"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Create structured log entry for consolidated logging
        log_entry = {
            'timestamp': timestamp,
            'agent': agent_name.upper(),
            'key': key,
            'type': 'INPUT',
            'prompts': {
                'system_prompt': {
                    'content': system_prompt,
                    'length': len(system_prompt),
                    'token_estimate': len(system_prompt.split()) * 1.3  # Rough token estimate
                },
                'user_prompt': {
                    'content': user_prompt,
                    'length': len(user_prompt),
                    'token_estimate': len(user_prompt.split()) * 1.3
                }
            },
            'context_data': {
                'content': str(context_data) if context_data else None,
                'length': len(str(context_data)) if context_data else 0
            },
            'actual_prompts': actual_prompts,  # Store the actual parsed prompts
            'session_id': getattr(self, 'session_id', 'default')
        }
        
        if agent_name not in self.logs:
            self.logs[agent_name] = {}
        if key not in self.logs[agent_name]:
            self.logs[agent_name][key] = []
            
        self.logs[agent_name][key].append(log_entry)
        self.session_logs.append(log_entry)
        
        # Save to consolidated session file
        self._save_to_consolidated_log(log_entry)
        
        # Per-entry JSON disabled
        
        # Store input entry for paired logging
        if not hasattr(self, 'pending_inputs'):
            self.pending_inputs = {}
        self.pending_inputs[f"{agent_name}_{key}"] = log_entry
        
        # Minimal text log for quick reference
        self._write_to_file(f"📝 [{timestamp}] {agent_name.upper()} INPUT → {key} (Est. {log_entry['prompts']['system_prompt']['token_estimate'] + log_entry['prompts']['user_prompt']['token_estimate']:.0f} tokens)")
        
        # Silent processing - no Streamlit UI updates
    
    def log_agent_output(self, agent_name, key, output, processing_time=0):
        """Log agent output and processing details to JSON files with clear structure"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        is_success = output and not str(output).startswith("Error")
        
        # Create structured log entry for JSON
        log_entry = {
            'timestamp': timestamp,
            'agent': agent_name.upper(),
            'key': key,
            'type': 'OUTPUT',
            'output': {
                'content': output,
                'length': len(str(output)),
                'format': 'json' if isinstance(output, dict) else 'text'
            },
            'processing': {
                'time_seconds': processing_time,
                'status': 'success' if is_success else 'error',
                'is_success': is_success
            },
            'session_id': getattr(self, 'session_id', 'default')
        }
        
        if agent_name not in self.logs:
            self.logs[agent_name] = {}
        if key not in self.logs[agent_name]:
            self.logs[agent_name][key] = []
            
        self.logs[agent_name][key].append(log_entry)
        self.session_logs.append(log_entry)
        
        # Per-entry JSON disabled
        
        # Create key log if input exists
        input_key = f"{agent_name}_{key}"
        if hasattr(self, 'pending_inputs') and input_key in self.pending_inputs:
            input_entry = self.pending_inputs[input_key]
            # Per-key JSON disabled
            # Remove from pending after key log is created
            del self.pending_inputs[input_key]
        
        # Write summary to text log
        status_icon = "✅" if is_success else "❌"
        self._write_to_file(f"{status_icon} [{timestamp}] {agent_name.upper()} OUTPUT ← {key} ({processing_time:.2f}s)")
        self._write_to_file(f"   Length: {len(str(output))} chars | Status: {'Success' if is_success else 'Error'}")
        
        # No Streamlit display during processing (silent logging)
    
    def log_error(self, agent_name, key, error_msg):
        """Log errors during processing to file"""
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        log_entry = {
            'timestamp': timestamp,
            'agent': agent_name,
            'key': key,
            'type': 'ERROR',
            'error': str(error_msg)
        }
        
        if agent_name not in self.logs:
            self.logs[agent_name] = {}
        if key not in self.logs[agent_name]:
            self.logs[agent_name][key] = []
            
        self.logs[agent_name][key].append(log_entry)
        self.session_logs.append(log_entry)
        
        # Write to file
        self._write_to_file(f"\n❌ ERROR - [{timestamp}] {agent_name.upper()} - {key}")
        self._write_to_file(f"ERROR: {error_msg}")
        self._write_to_file("")
        
        # Display error in Streamlit
        st.error(f"❌ {agent_name.upper()} error for {key}: {error_msg}")
    
    def save_logs_to_json(self):
        """Save all logs to JSON file for structured access"""
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        json_file = self.log_dir / f"ai_agents_{timestamp}.json"
        
        try:
            with open(json_file, 'w', encoding='utf-8') as f:
                json.dump(self.logs, f, indent=2, default=str, ensure_ascii=False)
            return json_file
        except Exception as e:
            print(f"Failed to save JSON logs: {e}")
            return None
    
    def display_session_summary(self):
        """Display simple summary in Streamlit"""
        if not self.session_logs:
            st.info("No AI agent activity logged yet.")
            return
            
        st.markdown("### 📊 AI Processing Summary")
        st.info(f"📁 Detailed logs saved to: `{self.session_file}`")
        
        # Count by agent
        agent_counts = {}
        for log in self.session_logs:
            agent = log['agent']
            if agent not in agent_counts:
                agent_counts[agent] = {'inputs': 0, 'outputs': 0, 'errors': 0}
            agent_counts[agent][log['type'].lower() + 's'] += 1
        
        # Display counts
        col1, col2, col3 = st.columns(3)
        for i, (agent, counts) in enumerate(agent_counts.items()):
            with [col1, col2, col3][i % 3]:
                st.metric(f"{agent.upper()}", 
                         f"I:{counts['inputs']} O:{counts['outputs']} E:{counts['errors']}")
        
        # Download consolidated session log
        try:
            if self.session_file.exists():
                with open(self.session_file, 'r', encoding='utf-8') as f:
                    st.download_button(
                        label="📥 Download Session Log (JSON)",
                        data=f.read(),
                        file_name=self.session_file.name,
                        mime='application/json',
                        type="secondary"
                    )
        except Exception:
            pass

# Initialize global logger
if 'ai_logger' not in st.session_state:
    st.session_state.ai_logger = AIAgentLogger()

# Load configuration files
# load_config_files moved to fdd_utils


def main():
    
    # Configure Streamlit page and sanitize deprecated options
    from common.ui import configure_streamlit_page
    configure_streamlit_page()
    st.title("🏢 Real Estate DD Report Writer")

    # Add navigation description
    st.info("📋 **Welcome!** Please navigate to the left sidebar to upload your databook and configure input data.")

    # Sidebar for controls
    with st.sidebar:
        # File uploader with default file option
        uploaded_file = st.file_uploader(
            "Upload Excel File (Optional)",
            type=['xlsx', 'xls'],
            help="Upload your financial data Excel file or use the default databook.xlsx"
        )

        # Use default file if no file is uploaded (show immediately under uploader)
        if uploaded_file is None:
            default_file_path = "databook.xlsx"
            if os.path.exists(default_file_path):
                st.caption(f"Using default file: {default_file_path}")
                # Create a proper mock uploaded file object for the default file
                class MockUploadedFile:
                    def __init__(self, file_path):
                        self.name = file_path
                        self.file_path = file_path
                        self._file = None
                    
                    def read(self, size=-1):
                        if self._file is None:
                            self._file = open(self.file_path, 'rb')
                        return self._file.read(size)
                    
                    def getbuffer(self):
                        with open(self.file_path, 'rb') as f:
                            return f.read()
                    
                    def seek(self, offset, whence=0):
                        if self._file is None:
                            self._file = open(self.file_path, 'rb')
                        return self._file.seek(offset, whence)
                    
                    def tell(self):
                        if self._file is None:
                            return 0
                        return self._file.tell()
                    
                    def seekable(self):
                        return True
                    
                    def close(self):
                        if self._file:
                            self._file.close()
                            self._file = None
                
                uploaded_file = MockUploadedFile(default_file_path)
            else:
                st.error(f"❌ Default file not found: {default_file_path}")
                st.info("Please upload an Excel file to get started.")
                st.stop()

        # (Removed duplicate provider/model UI to avoid two model selectors)

        # Entity name input with auto-mapping
        entity_input = st.text_input(
            "Enter Entity Name",
            value="",
            placeholder="e.g., Company Name Limited, Entity Name Corp",
            help="Enter the full entity name to configure processing"
        )
        
        # Clear session state when entity changes
        if 'last_entity_input' in st.session_state:
            if st.session_state['last_entity_input'] != entity_input:
                # Entity has changed, clear the cached data
                if 'ai_data' in st.session_state:
                    del st.session_state['ai_data']
                if 'filtered_keys_for_ai' in st.session_state:
                    del st.session_state['filtered_keys_for_ai']
                # Reset processing state when entity changes
                if 'processing_started' in st.session_state:
                    del st.session_state['processing_started']
                # Entity change detected, session cleared
        
        # Store current entity input for next comparison
        st.session_state['last_entity_input'] = entity_input
        
        # Entity Selection Mode (Single vs Multiple)
        st.markdown("---")
        st.markdown("**Entity Mode Selection:**")
        entity_mode = st.radio(
            "Choose entity processing mode:",
            ["Single Entity", "Multiple Entity"],
            index=0,  # Default to Single Entity
            help="Single Entity: Process one entity table | Multiple Entity: Detect and filter multiple entity tables in the same sheet"
        )

        # Convert to internal format
        entity_mode_internal = 'single' if entity_mode == "Single Entity" else 'multiple'
        st.session_state['entity_mode'] = entity_mode_internal

        st.info(f"📊 **Selected Mode:** {entity_mode} - {'Process single entity table' if entity_mode_internal == 'single' else 'Automatically detect and filter multiple entity tables'}")

        # Entity mode is now properly passed through the system

        # Show entity matching status if we have processed data
        if 'ai_data' in st.session_state and st.session_state['ai_data']:
            ai_data = st.session_state['ai_data']
            if 'entity_keywords' in ai_data:
                with st.expander("🔍 Entity Detection Status", expanded=False):
                    st.markdown("**Your Entity Keywords:**")
                    entity_keywords = ai_data.get('entity_keywords', [])
                    if entity_keywords:
                        for keyword in entity_keywords:
                            st.markdown(f"• `{keyword}`")
                    else:
                        st.warning("No entity keywords found")

                    st.markdown("**📋 Note:** If you're using Multiple Entity mode and it's falling back to single mode, check that:")
                    st.markdown("• Your entered entity name matches exactly what's in your Excel file")
                    st.markdown("• The entity name appears in row 0 or early rows of your Excel sheet")
                    st.markdown("• Try using just the base entity name (first word only)")


        

        
        # Auto-extract base entity and generate comprehensive entity keywords
        if entity_input:
            # Extract base entity name (first word)
            base_entity = entity_input.split()[0] if entity_input.split() else None
            
            # Generate comprehensive entity keywords from the input
            words = entity_input.split()
            entity_keywords = []

            # Always include the base entity
            entity_keywords.append(base_entity)

            # Generate all possible combinations of the input entity
            if len(words) >= 2:
                # Add two-word combinations
                for i in range(1, len(words)):
                    entity_keywords.append(f"{base_entity} {words[i]}")

                # Add three-word combinations if available
                if len(words) >= 3:
                    for i in range(1, len(words)-1):
                        for j in range(i+1, len(words)):
                            entity_keywords.append(f"{base_entity} {words[i]} {words[j]}")

            # For multiple entity mode, we want to be more inclusive
            # The entity detection logic will handle discovering all entities in the Excel file
            
            # Use full entity name for processing
            selected_entity = entity_input
            
            # Show entity info with first two words for display
            if entity_input:
                words = entity_input.split()
                # Use first two words, or first word if only one word
                display_name = ' '.join(words[:2]) if len(words) >= 2 else words[0] if words else entity_input
            else:
                display_name = base_entity
            st.info(f"📋 Entity: {display_name}")
            # Entity keywords generated successfully
        else:
            selected_entity = None
            entity_keywords = []
            st.warning("⚠️ Please enter an entity name to start processing")
        
        # Check if entity is provided (file can be default)
        if not selected_entity:
            st.warning("Please select an entity first")
            st.stop()

        # Generate entity_helpers dynamically from the input entity name
        if selected_entity:
            words = selected_entity.split()
        else:
            # Fallback if somehow we got here without an entity
            words = ["Unknown"]
        if len(words) > 1:
            # Use all words after the first as potential suffixes
            entity_helpers = ",".join(words[1:]) + ","
        else:
            # Default suffix for single-word entities
            entity_helpers = "Limited,"
        
        # Generate entity_suffixes from entity_helpers for backward compatibility
        entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]


        
        # Financial Statement Type Selection
        st.markdown("---")
        statement_type_options = ["Balance Sheet", "Income Statement", "All"]
        statement_type_display = st.radio(
            "Financial Statement Type",
            statement_type_options,
            help="Select the type of financial statement to process"
        )
        
        # Map display names back to internal codes
        statement_type_mapping = {
            "Balance Sheet": "BS",
            "Income Statement": "IS", 
            "All": "ALL"
        }
        statement_type = statement_type_mapping[statement_type_display]
        
        if uploaded_file is not None:
            if hasattr(uploaded_file, 'name') and uploaded_file.name != "databook.xlsx":
                st.success(f"Uploaded {uploaded_file.name}")
            # Store uploaded file in session state for later use
            st.session_state['uploaded_file'] = uploaded_file
            
            # AI Provider Selection (models defined in fdd_utils/config.json)
            ai_mode_options = ["Local AI", "Open AI", "DeepSeek", "Offline"]
            mode_display = st.selectbox(
                "Select Mode", 
                ai_mode_options,
                index=0,  # Local AI default
                help="Choose AI provider. Models are taken from fdd_utils/config.json"
            )
            
            # Show API configuration status
            config, _, _, _ = load_config_files()
            if config:
                if mode_display == "Open AI":
                    if config.get('OPENAI_API_KEY') and config.get('OPENAI_API_BASE'):
                        st.success("✅ OpenAI configured")
                        model = config.get('OPENAI_CHAT_MODEL', 'Not configured')
                        st.info(f"🤖 Model: {model}")
                    else:
                        st.warning("⚠️ OpenAI not configured. Add OPENAI_API_KEY and OPENAI_API_BASE in fdd_utils/config.json")
                elif mode_display == "DeepSeek":
                    if config.get('DEEPSEEK_API_KEY') and config.get('DEEPSEEK_API_BASE'):
                        st.success("✅ DeepSeek configured")
                        model = config.get('DEEPSEEK_CHAT_MODEL', 'Not configured')
                        st.info(f"🤖 Model: {model}")
                    else:
                        st.warning("⚠️ DeepSeek not configured. Add DEEPSEEK_API_KEY and DEEPSEEK_API_BASE in fdd_utils/config.json")
                elif mode_display == "Local AI":
                    if config.get('LOCAL_AI_API_BASE') and config.get('LOCAL_AI_ENABLED'):
                        st.success("✅ Local AI configured")
                        model = config.get('LOCAL_AI_CHAT_MODEL', 'Not configured')
                        endpoint = config.get('LOCAL_AI_API_BASE', 'Not configured')
                        st.info(f"🏠 Model: {model}")
                        st.info(f"🔗 Endpoint: {endpoint}")
                    else:
                        st.warning("⚠️ Local AI not configured. Configure LOCAL_AI_* in fdd_utils/config.json")

            # Show table mapping information
            st.subheader("📊 Table Mapping Configuration")
            config, mapping, pattern, prompts = load_config_files()
            if mapping:
                with st.expander("View Sheet Name Mappings", expanded=False):
                    # Create a table to show mapping information
                    mapping_data = []
                    for key, values in mapping.items():
                        if isinstance(values, list):
                            mapping_data.append({
                                "Financial Key": key,
                                "Mapped Sheet Names": ", ".join(values[:3]) + ("..." if len(values) > 3 else "")
                            })

                    if mapping_data:
                        st.table(mapping_data)
                    else:
                        st.info("No mapping data available")

            # Map display names to internal mode names
            provider_mapping = {
                "Open AI": "Open AI",
                "Local AI": "Local AI",
                "DeepSeek": "DeepSeek"
            }
            mode = f"AI Mode - {provider_mapping[mode_display]}"
            st.session_state['selected_mode'] = mode
            st.session_state['ai_model'] = mode_display
            st.session_state['selected_provider'] = provider_mapping[mode_display]
            st.session_state['use_local_ai'] = (mode_display == "Local AI")
            st.session_state['use_openai'] = (mode_display == "Open AI")
            
            # Performance statistics - moved below Select Mode
            st.markdown("---")


    # Main area for results
    if uploaded_file is not None:

        # Add a process button to control when processing starts
        if 'processing_started' not in st.session_state:
            st.session_state['processing_started'] = False

        if not st.session_state['processing_started']:
            st.markdown("### 🎯 Ready to Process")
            st.info("📋 Configuration loaded. Click 'Start Processing' to begin data analysis and AI processing.")

            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                start_processing = st.button(
                    "🚀 Start Processing",
                    type="primary",
                    use_container_width=True,
                    key="btn_start_processing",
                    help="Begin data processing and AI analysis"
                )

            if start_processing:
                st.session_state['processing_started'] = True
                st.rerun()  # Refresh to show processing interface

            # Don't show processing interface until button is clicked
            st.stop()

        # --- View Table Section ---
        config, mapping, pattern, prompts = load_config_files()
        
        # Use the pre-generated entity keywords
        if 'entity_keywords' not in locals() or not entity_keywords:
            # Get the manual entity mode selection
            entity_mode = st.session_state.get('entity_mode', 'single')

            # Generate entity keywords based on the selected entity
            entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]

            # Create entity keywords based on the selected entity
            entity_keywords = [selected_entity]  # Always include the base entity name
            for suffix in entity_suffixes:
                if suffix != selected_entity:  # Avoid duplicates
                    entity_keywords.append(f"{selected_entity} {suffix}")

            if not entity_keywords:
                entity_keywords = [selected_entity]
            
            # Entity processing completed
            print(f"🎯 APP DEBUG - Entity Setup Complete:")
            print(f"   👤 selected_entity: '{selected_entity}'")
            print(f"   🔑 entity_keywords: {entity_keywords}")
            print(f"   📋 entity_mode: {entity_mode}")
            print(f"   📄 entity_suffixes: {entity_suffixes}")
        
        # Handle different statement types
        if statement_type == "BS":
            st.markdown("### Balance Sheet")

            # Always process the data first and store it in session state for both sections
            # Check if we need to reprocess (entity changed or no data)
            entity_changed = st.session_state.get('last_processed_entity') != selected_entity
            needs_processing = 'ai_data' not in st.session_state or 'sections_by_key' not in st.session_state['ai_data'] or entity_changed

            if needs_processing:
                # Process Excel file and store in session state
                with st.spinner("🔄 Processing Excel file..."):
                    print(f"\n{'='*80}")
                    print(f"🔄 STARTING EXCEL PROCESSING FOR BALANCE SHEET")
                    print(f"🏢 Entity: {selected_entity}")
                    print(f"📁 File size: {len(uploaded_file.getvalue()) if hasattr(uploaded_file, 'getvalue') else 'Unknown'} bytes")
                    print(f"⏱️  Processing with 30-second timeout...")
                    print(f"{'='*80}\n")
                    start_excel_time = time.time()

                    # Add timeout protection for Excel processing (cross-platform)
                    import threading

                    result_container = {}
                    exception_container = {}

                    def process_excel_with_timeout():
                        try:
                            result = get_worksheet_sections_by_keys(
                                uploaded_file=uploaded_file,
                                tab_name_mapping=mapping,
                                entity_name=selected_entity,
                                entity_suffixes=entity_suffixes,
                                entity_keywords=entity_keywords,
                                entity_mode=entity_mode,  # Use the manual entity mode selection
                                debug=True  # Set to True for debugging
                            )
                            result_container['result'] = result
                        except Exception as e:
                            exception_container['exception'] = e

                    # Start processing in a separate thread
                    processing_thread = threading.Thread(target=process_excel_with_timeout)
                    processing_thread.daemon = True
                    processing_thread.start()

                    # Wait for completion with timeout
                    processing_thread.join(timeout=30)

                    if processing_thread.is_alive():
                        # Thread is still running after timeout
                        print("❌ Excel processing timed out after 30 seconds")

                        # Provide option to continue without Excel data
                        if st.button("⚠️ Continue Without Excel Data", key="continue_without_excel_bs"):
                            st.warning("⚠️ Continuing without Excel data. Some features may be limited.")
                            sections_by_key = {}  # Empty data structure
                            st.session_state['excel_processing_skipped'] = True
                        else:
                            st.error("❌ Excel processing timed out. Click 'Continue Without Excel Data' to proceed or try a smaller file.")
                            st.stop()
                            return
                    elif 'exception' in exception_container:
                        # Exception occurred during processing
                        raise exception_container['exception']
                    else:
                        # Processing completed successfully
                        sections_by_key = result_container['result']

                    excel_processing_time = time.time() - start_excel_time
                    print(f"\n{'='*60}")
                    print(f"✅ EXCEL PROCESSING COMPLETED!")
                    print(f"⏱️  Processing time: {excel_processing_time:.2f}s")
                    print(f"📊 Found {len(sections_by_key)} financial keys with data")
                    print(f"💡 The detailed tab-by-tab processing results are shown above.")
                    print(f"{'='*60}\n")

                    # High-level debug only
                    total_sections = sum(len(sections) for sections in sections_by_key.values())


                    # Store in session state for AI section to use
                    if 'ai_data' not in st.session_state:
                        st.session_state['ai_data'] = {}
                    st.session_state['ai_data']['sections_by_key'] = sections_by_key
                    st.session_state['ai_data']['entity_name'] = selected_entity
                    st.session_state['ai_data']['entity_keywords'] = entity_keywords
                    st.session_state['last_processed_entity'] = selected_entity  # Track last processed entity
                    print(f"💾 Stored processed data for Balance Sheet entity: {selected_entity}")
            else:
                # Use the data that was already processed
                sections_by_key = st.session_state['ai_data']['sections_by_key']
                print(f"♻️ Reusing previously processed Balance Sheet data for entity: {selected_entity}")
            
            from common.ui_sections import render_balance_sheet_sections
            render_balance_sheet_sections(
                sections_by_key,
                get_key_display_name,
                selected_entity,
                format_date_to_dd_mmm_yyyy,
            )
        
        elif statement_type == "IS":
            st.markdown("### Income Statement")
            
            # Always process the data first and store it in session state for both sections
            if 'ai_data' not in st.session_state or 'sections_by_key' not in st.session_state['ai_data']:
                # Process Excel file and store in session state
                with st.spinner("🔄 Processing Excel file for Income Statement..."):
                    print(f"\n{'='*80}")
                    print(f"🔄 STARTING EXCEL PROCESSING FOR INCOME STATEMENT")
                    print(f"🏢 Entity: {selected_entity}")
                    print(f"📁 File size: {len(uploaded_file.getvalue()) if hasattr(uploaded_file, 'getvalue') else 'Unknown'} bytes")
                    print(f"⏱️  Processing with 30-second timeout...")
                    print(f"{'='*80}\n")
                    start_excel_time = time.time()

                    # Add timeout protection for Excel processing (cross-platform)
                    import threading

                    result_container = {}
                    exception_container = {}

                    def process_excel_with_timeout():
                        try:
                            result = get_worksheet_sections_by_keys(
                                uploaded_file=uploaded_file,
                                tab_name_mapping=mapping,
                                entity_name=selected_entity,
                                entity_suffixes=entity_suffixes,
                                entity_keywords=entity_keywords,
                                entity_mode=entity_mode,  # Use the manual entity mode selection
                                debug=True  # Set to True for debugging
                            )
                            result_container['result'] = result
                        except Exception as e:
                            exception_container['exception'] = e

                    # Start processing in a separate thread
                    processing_thread = threading.Thread(target=process_excel_with_timeout)
                    processing_thread.daemon = True
                    processing_thread.start()

                    # Wait for completion with timeout
                    processing_thread.join(timeout=30)

                    if processing_thread.is_alive():
                        # Thread is still running after timeout
                        print("❌ Income Statement Excel processing timed out after 30 seconds")

                        # Provide option to continue without Excel data
                        if st.button("⚠️ Continue Without Excel Data", key="continue_without_excel_is"):
                            st.warning("⚠️ Continuing without Excel data. Some features may be limited.")
                            sections_by_key = {}  # Empty data structure
                            st.session_state['excel_processing_skipped'] = True
                        else:
                            st.error("❌ Excel processing timed out. Click 'Continue Without Excel Data' to proceed or try a smaller file.")
                            st.stop()
                            return
                    elif 'exception' in exception_container:
                        # Exception occurred during processing
                        raise exception_container['exception']
                    else:
                        # Processing completed successfully
                        sections_by_key = result_container['result']

                    excel_processing_time = time.time() - start_excel_time
                    print(f"\n{'='*60}")
                    print(f"✅ INCOME STATEMENT EXCEL PROCESSING COMPLETED!")
                    print(f"⏱️  Processing time: {excel_processing_time:.2f}s")
                    print(f"📊 Found {len(sections_by_key)} financial keys with data")
                    print(f"💡 The detailed tab-by-tab processing results are shown above.")
                    print(f"{'='*60}\n")
                    
                    # High-level debug only
                    total_sections = sum(len(sections) for sections in sections_by_key.values())

                    
                    # Store in session state for AI section to use
                    if 'ai_data' not in st.session_state:
                        st.session_state['ai_data'] = {}
                    st.session_state['ai_data']['sections_by_key'] = sections_by_key
                    st.session_state['ai_data']['entity_name'] = selected_entity
                    st.session_state['ai_data']['entity_keywords'] = entity_keywords
            else:
                # Use the data that was already processed
                sections_by_key = st.session_state['ai_data']['sections_by_key']
            
            from common.ui_sections import render_income_statement_sections
            render_income_statement_sections(
                sections_by_key,
                get_key_display_name,
                selected_entity,
                format_date_to_dd_mmm_yyyy,
            )
        
        elif statement_type == "ALL":
            # Combined view - show both BS and IS data
            st.markdown("### Combined Financial Statements")
            
            # Use the data that was already processed
            if 'ai_data' in st.session_state and 'sections_by_key' in st.session_state['ai_data']:
                sections_by_key = st.session_state['ai_data']['sections_by_key']
            else:
                # Fallback: process data if not already done
                with st.spinner("🔄 Processing Excel file for combined view..."):
                    print(f"🔍 DEBUG: Starting Excel processing for combined view - {selected_entity}")
                    start_excel_time = time.time()

                    # Add timeout protection for Excel processing (cross-platform)
                    import threading

                    result_container = {}
                    exception_container = {}

                    def process_excel_with_timeout():
                        try:
                            result = get_worksheet_sections_by_keys(
                                uploaded_file=uploaded_file,
                                tab_name_mapping=mapping,
                                entity_name=selected_entity,
                                entity_suffixes=entity_suffixes,
                                entity_keywords=entity_keywords,
                                entity_mode=entity_mode,  # Use the manual entity mode selection
                                debug=False
                            )
                            result_container['result'] = result
                        except Exception as e:
                            exception_container['exception'] = e

                    # Start processing in a separate thread
                    processing_thread = threading.Thread(target=process_excel_with_timeout)
                    processing_thread.daemon = True
                    processing_thread.start()

                    # Wait for completion with timeout
                    processing_thread.join(timeout=30)

                    if processing_thread.is_alive():
                        # Thread is still running after timeout
                        print("❌ Combined view Excel processing timed out after 30 seconds")
                        sections_by_key = {}  # Use empty data structure
                        st.session_state['excel_processing_skipped'] = True
                    elif 'exception' in exception_container:
                        # Exception occurred during processing
                        raise exception_container['exception']
                    else:
                        # Processing completed successfully
                        sections_by_key = result_container['result']

                    excel_processing_time = time.time() - start_excel_time
                    print(f"✅ Combined view Excel processing completed in {excel_processing_time:.2f}s")
            
            # Show combined data using the combined rendering function
            from common.ui_sections import render_combined_sections
            render_combined_sections(
                sections_by_key,
                get_key_display_name,
                selected_entity,
                format_date_to_dd_mmm_yyyy,
            )

        # --- AI Processing Section (Bottom) ---
        # Check AI configuration status (updated for DeepSeek as default)
        try:
            config, _, _, _ = load_config_files()
            if config:
                any_provider_configured = (
                    (config.get('DEEPSEEK_API_KEY') and config.get('DEEPSEEK_API_BASE')) or
                    (config.get('OPENAI_API_KEY') and config.get('OPENAI_API_BASE')) or
                    (config.get('LOCAL_AI_API_BASE') and config.get('LOCAL_AI_ENABLED')) or
                    (config.get('SERVER_AI_API_BASE') and (config.get('SERVER_AI_API_KEY') or config.get('LOCAL_AI_API_KEY')))
                )
                if not any_provider_configured:
                    st.warning("⚠️ AI Mode: No provider configured. Will use fallback mode with test data.")
        except Exception:
            st.warning("⚠️ AI Mode: Configuration not found. Will use fallback mode.")
        
        # --- AI Processing & Results Section ---
        st.markdown("---")
        st.markdown("## 🤖 AI Processing & Results")
        
        # Initialize session state for AI data if not exists
        if 'ai_data' not in st.session_state:
            st.session_state['ai_data'] = {}
        
        # Prepare data for AI processing
        if uploaded_file is not None:
            try:
                # Load configuration files
                config, mapping, pattern, prompts = load_config_files()
                if not all([config, mapping, pattern]):
                    st.error("❌ Failed to load configuration files")
                    return
                
                # Process entity configuration based on mode
                entity_mode = st.session_state.get('entity_mode', 'multiple')
                
                if entity_mode == 'single':
                    # For single entity mode, use only the base entity name
                    entity_suffixes = []
                    entity_keywords = [selected_entity]
                else:
                    # For multiple entity mode, use the full entity helpers
                    entity_suffixes = [s.strip() for s in entity_helpers.split(',') if s.strip()]
                    
                    # FIX: Don't duplicate the entity name if it's already in the suffix
                    entity_keywords = []
                    for suffix in entity_suffixes:
                        if suffix == selected_entity:
                            # If suffix is the same as selected_entity, just use selected_entity
                            entity_keywords.append(selected_entity)
                        else:
                            # Otherwise, combine them
                            entity_keywords.append(f"{selected_entity} {suffix}")
                    
                    if not entity_keywords:
                        entity_keywords = [selected_entity]
                    

                
                # Use the data that was already processed by the first section
                if 'ai_data' in st.session_state and 'sections_by_key' in st.session_state['ai_data']:
                    sections_by_key = st.session_state['ai_data']['sections_by_key']

                else:
                    # Fallback: process data if not already done
                    with st.spinner("🔄 Processing Excel file for AI..."):
                        sections_by_key = get_worksheet_sections_by_keys(
                            uploaded_file=uploaded_file,
                            tab_name_mapping=mapping,
                            entity_name=selected_entity,
                            entity_suffixes=entity_suffixes,
                            entity_keywords=entity_keywords,
                            entity_mode='auto',  # Always use auto mode for intelligent detection
                            debug=False
                        )

                st.success("✅ Excel processing completed for AI")
                
                # Get keys with data
                keys_with_data = [key for key, sections in sections_by_key.items() if sections]
                
                # Filter keys based on statement type (include synonyms for non-Haining)
                bs_keys = [
                    "Cash", "AR", "Prepayments", "OR", "Other CA", "Other NCA", "IP", "NCA",
                    "AP", "Taxes payable", "OP", "Capital", "Reserve"
                ]
                is_keys = [
                    "OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", "Other Income",
                    "Non-operating Income", "Non-operating Exp", "Income tax", "LT DTA"
                ]
                
                if statement_type == "BS":
                    filtered_keys_for_ai = [key for key in keys_with_data if key in bs_keys]
                elif statement_type == "IS":
                    filtered_keys_for_ai = [key for key in keys_with_data if key in is_keys]
                elif statement_type == "ALL":
                    # For ALL mode, process all keys that have data, not just predefined ones
                    filtered_keys_for_ai = keys_with_data
                else:
                    filtered_keys_for_ai = keys_with_data
                
                # Store the filtered keys in session state for use in other sections
                st.session_state['filtered_keys_for_ai'] = filtered_keys_for_ai
                st.session_state['current_statement_type'] = statement_type
                
                if not filtered_keys_for_ai:
                    st.warning("No data found for AI processing with the selected statement type.")
                    return
                
                # Store uploaded file data in session state for agents
                st.session_state['uploaded_file_data'] = uploaded_file.getbuffer()
                
                # Prepare AI data
                temp_ai_data = {
                    'entity_name': selected_entity,
                    'entity_keywords': entity_keywords,
                    'sections_by_key': sections_by_key,
                    'pattern': pattern,
                    'mapping': mapping,
                    'config': config
                }
                
                # Store in session state
                st.session_state['ai_data'] = temp_ai_data
                st.session_state['filtered_keys_for_ai'] = filtered_keys_for_ai
                
                # Initialize agent states if not exists
                if 'agent_states' not in st.session_state:
                    st.session_state['agent_states'] = {
                        'agent1_completed': False,
                        'agent2_completed': False, 
                        'agent3_completed': False,
                        'agent1_results': {},
                        'agent2_results': {},
                        'agent3_results': {},
                        'agent1_success': False,
                        'agent2_success': False,
                        'agent3_success': False
                    }
                
                # AI Processing Buttons - Direct language selection
                st.markdown("### 🤖 AI Report Generation")
                agent_states = st.session_state.get('agent_states', {})

                # Two columns for English and Chinese buttons
                col_eng, col_chi = st.columns(2)

                with col_eng:
                    run_eng_clicked = st.button(
                        "🇺🇸 Generate English Report",
                        type="primary",
                        use_container_width=True,
                        key="btn_ai_eng",
                        help="Generate AI report in English (Content Generation + Proofreading)"
                    )

                with col_chi:
                    run_chi_clicked = st.button(
                        "🇨🇳 生成中文报告",
                        type="primary",
                        use_container_width=True,
                        key="btn_ai_chi",
                        help="Generate AI report in Chinese (内容生成 + 校对 + 翻译)"
                    )

                    # Handle Chinese AI processing
                    if run_chi_clicked:
                        # Chinese AI button processing
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        eta_text = st.empty()

                        # Clear previous translation flags and prepare for new translation
                        if 'translation_completed' in st.session_state:
                            del st.session_state['translation_completed']
                        if 'refresh_needed' in st.session_state:
                            del st.session_state['refresh_needed']
                        if 'last_translation_time' in st.session_state:
                            del st.session_state['last_translation_time']

                        # Clear any existing agent3 results to ensure clean slate
                        if 'agent_states' in st.session_state and 'agent3_results' in st.session_state['agent_states']:
                            print("🧹 Clearing existing agent3_results for fresh translation")
                            st.session_state['agent_states']['agent3_results'] = {}

                        # Clear content store agent3 entries
                        if 'ai_content_store' in st.session_state:
                            content_store = st.session_state['ai_content_store']
                            for key in list(content_store.keys()):
                                if 'agent3_content' in content_store[key]:
                                    print(f"🧹 Clearing existing agent3_content for {key}")
                                    del content_store[key]['agent3_content']

                        status_text.text("🤖 初始化中文AI处理…")
                        progress_bar.progress(10)

                # Handle English AI processing
                if run_eng_clicked:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    eta_text = st.empty()
                    try:
                        status_text.text("🤖 Initializing…")
                        progress_bar.progress(10)

                        # Get selected language for AI processing
                        selected_language = st.session_state.get('selected_language', 'English')

                        # Handle different statement types
                        current_statement_type = st.session_state.get('current_statement_type', 'BS')

                        # Initialize variables to avoid "not associated with a value" errors
                        agent1_results = {}
                        agent1_success = False

                        # Define key lists for BS and IS
                        bs_key_list = [
                            "Cash", "AR", "Prepayments", "OR", "Other CA", "Other NCA", "IP", "NCA",
                            "AP", "Taxes payable", "OP", "Capital", "Reserve"
                        ]
                        is_key_list = [
                            "OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", "Other Income",
                            "Non-operating Income", "Non-operating Exp", "Income tax", "LT DTA"
                        ]

                        if current_statement_type == "ALL":
                            # For ALL, process both BS and IS + translation + proofreading
                            ext = {'bar': progress_bar, 'status': status_text, 'combined': {'stages': 4, 'stage_index': 0, 'start_time': time.time()}}

                            # Process all keys for ALL mode
                            status_text.text("📊 Processing all financial data...")
                            progress_bar.progress(20)
                            st.session_state['current_statement_type'] = 'ALL'
                            agent1_results_bs = run_agent_1_simple(filtered_keys_for_ai, temp_ai_data, external_progress=ext, language=selected_language)

                            # Store all results in session state
                            if agent1_results_bs:
                                # Initialize content store if not exists
                                if 'ai_content_store' not in st.session_state:
                                    st.session_state['ai_content_store'] = {}

                                # Store all results
                                for key, result in agent1_results_bs.items():
                                    if key not in st.session_state['ai_content_store']:
                                        st.session_state['ai_content_store'][key] = {}
                                    st.session_state['ai_content_store'][key]['agent1_content'] = result
                                    st.session_state['ai_content_store'][key]['current_content'] = result
                                    st.session_state['ai_content_store'][key]['agent1_timestamp'] = datetime.datetime.now().isoformat()

                            # Generate content files
                            if agent1_results_bs:
                                status_text.text("📝 Generating content files...")
                                progress_bar.progress(90)
                                generate_content_from_session_storage(selected_entity)

                            # Content files are already generated above
                            pass

                            # Use all results
                            agent1_results = agent1_results_bs
                            agent1_success = bool(agent1_results and any(agent1_results.values()))

                            # Reset to ALL for future operations
                            st.session_state['current_statement_type'] = 'ALL'

                        elif current_statement_type == "BS":
                            # Handle Balance Sheet processing
                            ext = {'bar': progress_bar, 'status': status_text, 'combined': {'stages': 2, 'stage_index': 0, 'start_time': time.time()}}

                            # Filter BS keys
                            bs_filtered_keys = [key for key in filtered_keys_for_ai if key in bs_key_list]

                            if not bs_filtered_keys:
                                status_text.text("⚠️ No Balance Sheet keys found to process")
                                agent1_results = {}
                                agent1_success = False
                            else:
                                status_text.text("📊 Processing Balance Sheet data...")
                                progress_bar.progress(30)
                                agent1_results = run_agent_1_simple(bs_filtered_keys, temp_ai_data, external_progress=ext, language=selected_language)
                                agent1_success = bool(agent1_results and any(agent1_results.values()))

                                # Store BS results in session state
                                if agent1_results:
                                    if 'ai_content_store' not in st.session_state:
                                        st.session_state['ai_content_store'] = {}

                                    for key, result in agent1_results.items():
                                        if key not in st.session_state['ai_content_store']:
                                            st.session_state['ai_content_store'][key] = {}
                                        st.session_state['ai_content_store'][key]['agent1_content'] = result
                                        st.session_state['ai_content_store'][key]['current_content'] = result
                                        st.session_state['ai_content_store'][key]['agent1_timestamp'] = datetime.datetime.now().isoformat()

                                # Generate content files
                                if agent1_results:
                                    status_text.text("📝 Generating Balance Sheet content files...")
                                    progress_bar.progress(90)
                                    generate_content_from_session_storage(selected_entity)

                        elif current_statement_type == "IS":
                            # Handle Income Statement processing
                            ext = {'bar': progress_bar, 'status': status_text, 'combined': {'stages': 2, 'stage_index': 0, 'start_time': time.time()}}

                            # Filter IS keys
                            is_filtered_keys = [key for key in filtered_keys_for_ai if key in is_key_list]

                            if not is_filtered_keys:
                                status_text.text("⚠️ No Income Statement keys found to process")
                                agent1_results = {}
                                agent1_success = False
                            else:
                                status_text.text("📊 Processing Income Statement data...")
                                progress_bar.progress(30)
                                agent1_results = run_agent_1_simple(is_filtered_keys, temp_ai_data, external_progress=ext, language=selected_language)
                                agent1_success = bool(agent1_results and any(agent1_results.values()))

                                # Store IS results in session state
                                if agent1_results:
                                    if 'ai_content_store' not in st.session_state:
                                        st.session_state['ai_content_store'] = {}

                                    for key, result in agent1_results.items():
                                        if key not in st.session_state['ai_content_store']:
                                            st.session_state['ai_content_store'][key] = {}
                                        st.session_state['ai_content_store'][key]['agent1_content'] = result
                                        st.session_state['ai_content_store'][key]['current_content'] = result
                                        st.session_state['ai_content_store'][key]['agent1_timestamp'] = datetime.datetime.now().isoformat()

                                # Generate content files
                                if agent1_results:
                                    status_text.text("📝 Generating Income Statement content files...")
                                    progress_bar.progress(90)
                                    generate_content_from_session_storage(selected_entity)

                        else:
                            # Unknown statement type
                            status_text.text(f"⚠️ Unknown statement type: {current_statement_type}")
                            agent1_results = {}
                            agent1_success = False
                        
                        st.session_state['agent_states']['agent1_results'] = agent1_results
                        st.session_state['agent_states']['agent1_completed'] = True
                        st.session_state['agent_states']['agent1_success'] = agent1_success
                        
                        progress_bar.progress(100)
                        status_text.text("✅ AI done" if agent1_success else "❌ AI failed")
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        progress_bar.progress(100)
                        status_text.text(f"❌ AI failed: {e}")
                        time.sleep(1)
                        st.rerun()

                # Handle Chinese AI processing
                if run_chi_clicked:
                    # Chinese AI processing started
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    eta_text = st.empty()
                    try:
                        status_text.text("🤖 初始化中文AI处理…")
                        progress_bar.progress(10)
                        # Progress bar initialized

                        selected_language = '中文'
                        st.session_state['selected_language'] = selected_language

                        # Get AI data from session state
                        temp_ai_data = st.session_state.get('ai_data', {})
                        filtered_keys_for_ai = st.session_state.get('filtered_keys_for_ai', [])

                        # Check if required data is available
                        if not temp_ai_data or not filtered_keys_for_ai:
                            st.error("❌ AI数据未准备好。请先上传文件并选择实体。")
                            return

                        # Update progress and status
                        progress_bar.progress(15)
                        status_text.text(f"📊 找到 {len(filtered_keys_for_ai)} 个财务科目，准备处理...")

                        # Debug: Check logger availability
                        if 'ai_logger' not in st.session_state:
                            st.warning("⚠️ AI日志记录器未初始化，正在初始化...")
                            from fdd_utils.enhanced_logging_config import AIAgentLogger
                            st.session_state.ai_logger = AIAgentLogger()

                        # Handle different statement types
                        current_statement_type = st.session_state.get('current_statement_type', 'BS')

                        if current_statement_type == "ALL":
                            # For ALL, process both BS and IS + translation + proofreading
                            ext = {'bar': progress_bar, 'status': status_text, 'combined': {'stages': 4, 'stage_index': 0, 'start_time': time.time()}}

                            # Define key lists for BS and IS
                            bs_key_list = [
                                "Cash", "AR", "Prepayments", "OR", "Other CA", "Other NCA", "IP", "NCA",
                                "AP", "Taxes payable", "OP", "Capital", "Reserve"
                            ]
                            is_key_list = [
                                "OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", "Other Income",
                                "Non-operating Income", "Non-operating Exp", "Income tax", "LT DTA"
                            ]

                            # Process all keys for ALL mode
                            status_text.text("📊 处理所有财务数据...")
                            progress_bar.progress(20)
                            st.session_state['current_statement_type'] = 'ALL'
                            # For Chinese reports, generate English content first, then proofread, then translate
                            agent1_language = 'English' if selected_language == '中文' else selected_language
                            agent1_results_bs = run_agent_1_simple(filtered_keys_for_ai, temp_ai_data, external_progress=ext, language=agent1_language)

                            # Store all results in session state
                            if agent1_results_bs:
                                # Initialize content store if not exists
                                if 'ai_content_store' not in st.session_state:
                                    st.session_state['ai_content_store'] = {}

                                # Store BS results
                                for key, result in agent1_results_bs.items():
                                    if key in bs_key_list:
                                        st.session_state['ai_content_store'][f"bs_{key}"] = result

                                # Switch to IS processing
                                ext['combined']['stage_index'] = 1
                                status_text.text("📈 处理损益表...")
                                progress_bar.progress(40)

                                # Filter IS keys
                                is_filtered_keys = [key for key in filtered_keys_for_ai if key in is_key_list]
                                agent1_results_is = run_agent_1_simple(is_filtered_keys, temp_ai_data, external_progress=ext, language=agent1_language)

                                if agent1_results_is:
                                    # Store IS results
                                    for key, result in agent1_results_is.items():
                                        if key in is_key_list:
                                            st.session_state['ai_content_store'][f"is_{key}"] = result

                                # Proofreader for both
                                status_text.text("🧐 运行全面校对...")
                                progress_bar.progress(80)

                                # Combine all results for translation
                                combined_results = {}
                                combined_results.update(agent1_results_bs or {})
                                combined_results.update(agent1_results_is or {})

                                # New Chinese Logic: Generate English → Proofread English → Translate to Chinese
                                if selected_language == '中文':
                                    # Step 1: Proofread the English content first
                                    ext['combined']['stage_index'] = 2
                                    status_text.text("🧐 校对英文内容...")
                                    progress_bar.progress(70)
                                    proofread_english_results = run_ai_proofreader(filtered_keys_for_ai, combined_results, temp_ai_data, external_progress=ext, language='English')

                                    # Step 2: Then translate the proofread English content to Chinese
                                    ext['combined']['stage_index'] = 3
                                    status_text.text("🌐 翻译为中文...")
                                    progress_bar.progress(85)

                                                                    # ENHANCED LOGGING: Create button-specific logger
                                from fdd_utils.logging_config import log_ai_processing, create_button_logger
                                button_logger = create_button_logger().get_button_logger("chinese_ai")
                                button_logger.info("Starting Chinese AI translation process")
                                button_logger.info(f"Processing {len(filtered_keys_for_ai)} keys")
                                button_logger.info(f"Proofread results available: {len(proofread_english_results) if proofread_english_results else 0}")



                                # FAST: Prepare translation input with minimal processing
                                translation_input = {}

                                # Pre-fetch agent1 results to avoid repeated session state access
                                agent1_state = st.session_state.get('agent_states', {}).get('agent1_results', {})

                                for key in filtered_keys_for_ai:
                                    if key in proofread_english_results and isinstance(proofread_english_results[key], dict):
                                        translation_input[key] = proofread_english_results[key]
                                    elif key in agent1_state:
                                        translation_input[key] = agent1_state[key]



                                # Check if we need translation (only for non-English languages)
                                if selected_language == '中文':
                                    translated_results = run_chinese_translator(filtered_keys_for_ai, translation_input, temp_ai_data, external_progress=ext)
                                    proof_results = translated_results  # Final results are the translated content

                                    # Check if translation actually produced Chinese content
                                    if translated_results:
                                        # Verify that at least some content is in Chinese
                                        has_chinese_content = False
                                        has_translation_failure = False

                                        for key, result in translated_results.items():
                                            content = result.get('content', '') if isinstance(result, dict) else str(result)
                                            if isinstance(result, dict) and result.get('translation_failed'):
                                                has_translation_failure = True
                                                break
                                            elif content and any('\u4e00' <= char <= '\u9fff' for char in content):
                                                has_chinese_content = True
                                                break

                                        if has_translation_failure:
                                            st.error("❌ 翻译完全失败。请检查AI配置和网络连接。")
                                            status_text.text("❌ 翻译失败")
                                        elif not has_chinese_content:
                                            st.warning("⚠️ 翻译可能失败，内容仍为英文。请检查AI配置。")
                                            status_text.text("⚠️ 翻译警告：内容可能仍为英文")
                                    else:
                                        proof_results = None
                                else:
                                    # English version: Skip translation, go directly to proofread
                                    translated_results = combined_results  # No translation needed
                                    ext['combined']['stage_index'] = 2
                                    status_text.text("🧐 运行校对...")
                                    progress_bar.progress(70)
                                    proof_results = run_ai_proofreader(filtered_keys_for_ai, combined_results, temp_ai_data, external_progress=ext, language=selected_language)

                                if proof_results:
                                    # Storing translation results


                                    st.session_state['agent_states']['agent3_results'] = proof_results
                                    st.session_state['agent_states']['agent3_completed'] = True
                                    st.session_state['agent_states']['agent3_success'] = bool(proof_results)

                                    # UPDATE AI CONTENT STORE WITH TRANSLATED CONTENT FOR PPTX EXPORT
                                    if 'ai_content_store' not in st.session_state:
                                        st.session_state['ai_content_store'] = {}

                                    for key, result in proof_results.items():
                                        if isinstance(result, dict):
                                            translated_content = result.get('corrected_content', '') or result.get('content', '')
                                            if translated_content and result.get('is_chinese', False):
                                                if key not in st.session_state['ai_content_store']:
                                                    st.session_state['ai_content_store'][key] = {}
                                                st.session_state['ai_content_store'][key]['current_content'] = translated_content
                                                st.session_state['ai_content_store'][key]['agent3_content'] = translated_content
                                                st.session_state['ai_content_store'][key]['agent3_timestamp'] = time.time()


                                    print(f"✅ AI content store updated for PPTX export")

                            progress_bar.progress(100)
                            status_text.text("✅ 所有处理完成")

                            # Set refresh flag for UI to know translation is complete
                            st.session_state['refresh_needed'] = True
                            st.session_state['translation_completed'] = True
                            st.session_state['last_translation_time'] = time.time()
                            st.session_state['force_reload_agent3'] = True  # Force UI to reload agent3_results

                            # Force immediate refresh to show Chinese content
                            print("🔄 Forcing UI refresh to display Chinese content...")
                            print(f"📊 agent3_results stored with {len(proof_results)} keys: {list(proof_results.keys())}")
                            time.sleep(1)
                            st.rerun()

                        else:
                            # Single statement type processing
                            ext = {'bar': progress_bar, 'status': status_text, 'combined': {'stages': 3, 'stage_index': 0, 'start_time': time.time()}}
                            # For Chinese reports, generate English content first, then proofread, then translate
                            agent1_language = 'English' if selected_language == '中文' else selected_language
                            agent1_results = run_agent_1_simple(filtered_keys_for_ai, temp_ai_data, external_progress=ext, language=agent1_language)
                            st.session_state['agent_states']['agent1_results'] = agent1_results
                            st.session_state['agent_states']['agent1_completed'] = True
                            st.session_state['agent_states']['agent1_success'] = bool(agent1_results)

                            # New Chinese Logic: Generate English → Proofread English → Translate to Chinese
                            if selected_language == '中文':
                                # Step 1: Proofread the English content first
                                ext['combined']['stage_index'] = 1
                                status_text.text("🧐 校对英文内容...")
                                proofread_english_results = run_ai_proofreader(filtered_keys_for_ai, agent1_results, temp_ai_data, external_progress=ext, language='English')

                                # Step 2: Then translate the proofread English content to Chinese
                                ext['combined']['stage_index'] = 2
                                status_text.text("🌐 翻译为中文...")



                                # FAST: Prepare translation input with minimal processing
                                translation_input = {}

                                # Pre-fetch agent1 results to avoid repeated session state access
                                agent1_state = st.session_state.get('agent_states', {}).get('agent1_results', {})

                                for key in filtered_keys_for_ai:
                                    if key in proofread_english_results and isinstance(proofread_english_results[key], dict):
                                        translation_input[key] = proofread_english_results[key]
                                    elif key in agent1_state:
                                        translation_input[key] = agent1_state[key]



                                translated_results = run_chinese_translator(filtered_keys_for_ai, translation_input, temp_ai_data, external_progress=ext)

                                proof_results = translated_results  # Final results are the translated content

                                # Check if translation actually produced Chinese content
                                if translated_results:
                                    # Verify that at least some content is in Chinese
                                    has_chinese_content = False
                                    has_translation_failure = False

                                    for key, result in translated_results.items():
                                        content = result.get('content', '') if isinstance(result, dict) else str(result)
                                        if isinstance(result, dict) and result.get('translation_failed'):
                                            has_translation_failure = True
                                            break
                                        elif content and any('\u4e00' <= char <= '\u9fff' for char in content):
                                            has_chinese_content = True
                                            break

                                    if has_translation_failure:
                                        st.error("❌ 翻译完全失败。请检查AI配置和网络连接。")
                                        status_text.text("❌ 翻译失败")
                                    elif not has_chinese_content:
                                        st.warning("⚠️ 翻译可能失败，内容仍为英文。请检查AI配置。")
                                        status_text.text("⚠️ 翻译警告：内容可能仍为英文")

                            else:
                                # English version: Generate → Proofread
                                status_text.text("🧐 运行校对...")
                                ext['combined']['stage_index'] = 1
                                proof_results = run_ai_proofreader(filtered_keys_for_ai, agent1_results, temp_ai_data, external_progress=ext, language=selected_language)
                                translated_results = agent1_results  # No translation for English

                            # STORE FINAL RESULTS (translated_results for Chinese, proof_results for English)
                            final_results = translated_results if 'translated_results' in locals() and selected_language == '中文' else proof_results
                            st.session_state['agent_states']['agent3_results'] = final_results
                            st.session_state['agent_states']['agent3_completed'] = True
                            st.session_state['agent_states']['agent3_success'] = bool(final_results)

                            # UPDATE AI CONTENT STORE WITH FINAL RESULTS FOR PPTX EXPORT
                            print(f"\n🔄 UPDATING AI CONTENT STORE WITH FINAL RESULTS...")
                            if 'ai_content_store' not in st.session_state:
                                st.session_state['ai_content_store'] = {}

                            for key, result in final_results.items():
                                if isinstance(result, dict):
                                    final_content = result.get('corrected_content', '') or result.get('content', '')
                                    if final_content:
                                        if key not in st.session_state['ai_content_store']:
                                            st.session_state['ai_content_store'][key] = {}
                                        st.session_state['ai_content_store'][key]['current_content'] = final_content
                                        if selected_language == '中文':
                                            st.session_state['ai_content_store'][key]['agent3_content'] = final_content
                                            st.session_state['ai_content_store'][key]['agent3_timestamp'] = time.time()
                                        print(f"   ✅ Updated {key} with final content ({len(final_content)} chars)")
                                    else:
                                        print(f"   ⚠️  {key} has no content to update")
                                else:
                                    print(f"   ⚠️  {key} result is not a dict: {type(result)}")

                            print(f"✅ AI content store updated for PPTX export")

                            # LOG COMPLETION
                            try:
                                button_logger.info("Chinese AI translation process completed successfully")
                                button_logger.info(f"Total keys processed: {len(final_results)}")
                                button_logger.info(f"Session state updated with agent3_results")
                                button_logger.info("Ready for PowerPoint export with Chinese content")
                            except Exception:
                                pass

                            # Generate content files after combined processing
                            if proof_results:
                                status_text.text("📝 生成内容文件...")
                                progress_bar.progress(95)
                                generate_content_from_session_storage(selected_entity)

                            progress_bar.progress(100)
                            status_text.text("✅ 中文AI处理完成")
                            time.sleep(1)
                    except Exception as e:
                        st.error(f"❌ 中文AI处理失败: {e}")
                        progress_bar.progress(0)
                        status_text.text("❌ 处理失败")

                # Clean UI with direct language selection buttons

                # All legacy button handlers removed

                
            except Exception as e:
                st.error(f"❌ Failed to prepare AI data: {e}")
        else:
            st.info("Please upload an Excel file first.")
        
        # --- AI Results Display ---
        # Check if any agent has run
        agent_states = st.session_state.get('agent_states', {})
        any_agent_completed = any([
            agent_states.get('agent1_completed', False),
            agent_states.get('agent2_completed', False),
            agent_states.get('agent3_completed', False)
        ])
        
        if any_agent_completed:
            # Get available keys
            filtered_keys = st.session_state.get('filtered_keys_for_ai', [])
            
            if filtered_keys:
                # Create tabs for each key with content-aware display names
                tab_labels = []
                for key in filtered_keys:
                    # Try to get content for language detection
                    content_for_detection = None

                    # Check agent3_results first
                    agent3_results_all = agent_states.get('agent3_results', {}) or {}
                    if key in agent3_results_all:
                        pr = agent3_results_all[key]
                        if isinstance(pr, dict):
                            # Use the same logic as in the display section
                            translated_content = pr.get('translated_content', '')
                            corrected_content = pr.get('corrected_content', '') or pr.get('content', '')
                            content_for_detection = translated_content if translated_content and pr.get('is_chinese', False) else corrected_content

                    # Fallback to agent1_results if no agent3 content
                    if not content_for_detection:
                        agent1_results_all = agent_states.get('agent1_results', {}) or {}
                        if key in agent1_results_all:
                            agent1_content = agent1_results_all[key]
                            if isinstance(agent1_content, dict):
                                content_for_detection = agent1_content.get('content', str(agent1_content))
                            else:
                                content_for_detection = str(agent1_content)

                    # Use content-aware display name
                    display_name = get_key_display_name(key, content=content_for_detection)
                    tab_labels.append(display_name)

                key_tabs = st.tabs(tab_labels)
                
                # Display results for each key in its tab
                for i, key in enumerate(filtered_keys):
                    with key_tabs[i]:
                        # Display key results
                        # Show Compliance (Proofreader) first if available
                        agent3_results_all = agent_states.get('agent3_results', {}) or {}
                        agent3_final_content = None

                        # Check for agent3_final content from JSON file if not in session state
                        if key not in agent3_results_all:
                            agent3_final_content = get_content_from_json(key)

                        if key in agent3_results_all:
                            pr = agent3_results_all[key]

                            # PRIORITY: Check for translated Chinese content first
                            translated_content = pr.get('translated_content', '')
                            corrected_content = pr.get('corrected_content', '') or pr.get('content', '')

                            # Use translated content if available and it's actually Chinese, otherwise use corrected content
                            final_content = translated_content if translated_content and pr.get('is_chinese', False) else corrected_content

                            # Determine content to display

                            # Check for Chinese characters in final content
                            chinese_chars = sum(1 for char in final_content if '\u4e00' <= char <= '\u9fff')
                            english_chars = sum(1 for char in final_content if char.isascii() and char.isalnum())
                            total_chars = len(final_content)

                            # Content analysis for display

                            # Content is ready for display

                            # Check if this is a translation failure
                            if isinstance(pr, dict) and pr.get('translation_failed'):
                                st.error(f"❌ 翻译失败: {corrected_content}")
                                # Show original English content if available
                                original_content = pr.get('original_content', '')
                                if original_content:
                                    st.info("📝 显示原始英文内容:")
                                    st.markdown(original_content)
                            elif corrected_content:
                                st.markdown(corrected_content)
                        elif agent3_final_content:
                            # Display agent3_final content from JSON (without label)
                            st.markdown(agent3_final_content)
                            agent3_results_all[key] = {"content": agent3_final_content}  # Add to results for expander logic

                        # AI1 Results (collapsible if proofreader exists)
                        with st.expander("📝 AI: Generation (details)", expanded=key not in agent3_results_all and not agent3_final_content):
                            agent1_results = agent_states.get('agent1_results', {}) or {}
                            if key in agent1_results and agent1_results[key]:
                                content = agent1_results[key]
                                
                                # Handle both string and dict content
                                if isinstance(content, dict):
                                    content_str = content.get('content', str(content))
                                else:
                                    content_str = str(content)
                                
                                st.markdown("**Generated Content:**")
                                st.markdown(content_str)
                                
                                # Metadata
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("Characters", len(content_str))
                                with col2:
                                    st.metric("Words", len(content_str.split()))
                                with col3:
                                    st.metric("Status", "✅ Generated" if content else "❌ Failed")
                            else:
                                st.info("No AI results available. Run AI first.")

                        # If Proofreader made changes or found issues, show a compact summary
                        if key in agent3_results_all:
                            pr = agent3_results_all[key]
                            issues = pr.get('issues', []) or []
                            changed = bool(pr.get('corrected_content'))
                            if issues or changed:
                                with st.expander("🧐 AI Proofreader: Changes & Notes", expanded=False):
                                    if changed and not issues:
                                        st.markdown("- Corrected content applied")
                                    if issues:
                                        st.markdown("- Detected issues (reference only):")
                                        for issue in issues:
                                            st.write(f"  • {issue}")
                                    runs = pr.get('translation_runs', 0)
                                    if runs:
                                        st.markdown(f"- Heuristic translation applied: {runs} run(s)")
                        if key not in agent3_results_all:
                            st.info("No compliance results available. Run AI Proofreader.")
            else:
                st.info("No financial keys available for results display.")
        else:
            st.info("No AI agents have run yet. Use the buttons above to start processing.")
        
        # --- PowerPoint Generation Section (Bottom) ---
        st.markdown("---")
        st.subheader("📊 PowerPoint Generation")

        # Define the export function that combines prepare and download
        def export_pptx_with_download(selected_entity, statement_type, language='english'):
            try:
                # Show language-specific progress message
                if language == 'chinese':
                    st.info("📊 开始生成中文 PowerPoint 演示文稿...")
                else:
                    st.info("📊 Generating English PowerPoint presentation...")

                # Get the project name based on selected entity (use first two words)
                if selected_entity:
                    words = selected_entity.split()
                    # Use first two words, or first word if only one word
                    project_name = ' '.join(words[:2]) if len(words) >= 2 else words[0] if words else selected_entity
                else:
                    project_name = selected_entity

                # Check for template file in common locations
                possible_templates = [
                    "fdd_utils/template.pptx",
                    "template.pptx",
                    "old_ver/template.pptx",
                    "common/template.pptx"
                ]

                template_path = None
                for template in possible_templates:
                    if os.path.exists(template):
                        template_path = template
                        break

                if not template_path:
                    st.error("❌ PowerPoint template not found. Please ensure 'template.pptx' exists in the fdd_utils/ directory.")
                    st.info("💡 You can copy a template file from the old_ver/ directory or create a new one.")
                    return

                # Add language suffix to filename
                language_suffix = "_CN" if language == 'chinese' else "_EN"

                # Define output path with timestamp in fdd_utils/output directory
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                output_filename = f"{project_name}_{statement_type.upper()}_{timestamp}{language_suffix}.pptx"
                output_path = f"fdd_utils/output/{output_filename}"

                # Ensure output directory exists
                os.makedirs("fdd_utils/output", exist_ok=True)

                # 1. Get the correct filtered keys for export
                ai_data = st.session_state.get('ai_data', {})
                sections_by_key = ai_data.get('sections_by_key', {})
                entity_name = ai_data.get('entity_name', selected_entity)
                keys_with_data = [key for key, sections in sections_by_key.items() if sections]

                # Dynamic BS key selection (as in your old logic)
                bs_keys = [
                    "Cash", "AR", "Prepayments", "OR", "Other CA", "IP", "Other NCA",
                    "AP", "Taxes payable", "OP", "Capital", "Reserve"
                ]
                if entity_name in ['Ningbo', 'Nanjing']:
                    bs_keys = [key for key in bs_keys if key != "Reserve"]

                is_keys = [
                    "OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", "Other Income",
                    "Non-operating Income", "Non-operating Exp", "Income tax", "LT DTA"
                ]

                if statement_type == "BS":
                    filtered_keys = [key for key in keys_with_data if key in bs_keys]
                elif statement_type == "IS":
                    filtered_keys = [key for key in keys_with_data if key in is_keys]
                else:  # ALL
                    filtered_keys = keys_with_data

                # 2. Use appropriate content file based on statement type
                # Note: Content files should contain narrative content from AI processing, not table data

                # Get the Excel file path for embedding data
                excel_file_path = None
                # Get uploaded_file from session state or use default
                current_uploaded_file = st.session_state.get('uploaded_file', None)
                if current_uploaded_file is not None:
                    try:
                        # Save uploaded file to temporary location for PowerPoint processing
                        import tempfile
                        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
                            temp_file.write(current_uploaded_file.getvalue())
                            excel_file_path = temp_file.name
                        print(f"💾 Saved uploaded file to temporary location: {excel_file_path}")
                    except Exception as e:
                        print(f"⚠️ Failed to save uploaded file: {e}")
                        # Fallback to default file
                        if os.path.exists("databook.xlsx"):
                            excel_file_path = "databook.xlsx"
                else:
                    # Fallback to default file
                    if os.path.exists("databook.xlsx"):
                        excel_file_path = "databook.xlsx"

                # Handle different statement types
                if statement_type == "IS":
                    # Income Statement only
                    markdown_path = "fdd_utils/is_content.md"
                    if not os.path.exists(markdown_path):
                        st.error(f"❌ Content file not found: {markdown_path}")
                        st.info("💡 Please run AI processing first to generate content for PowerPoint export.")
                        return

                    # Use language-specific export settings
                    if language == 'chinese':
                        st.info("📊 使用中文优化设置生成演示文稿...")
                        # Chinese-specific settings are handled in pptx_export.py functions

                    export_pptx(
                        template_path=template_path,
                        markdown_path=markdown_path,
                        output_path=output_path,
                        project_name=project_name,
                        excel_file_path=excel_file_path,
                        language=language
                    )

                elif statement_type == "BS":
                    # Balance Sheet only
                    markdown_path = "fdd_utils/bs_content.md"
                    if not os.path.exists(markdown_path):
                        st.error(f"❌ Content file not found: {markdown_path}")
                        st.info("💡 Please run AI processing first to generate content for PowerPoint export.")
                        return

                    # Use language-specific export settings
                    if language == 'chinese':
                        st.info("📊 使用中文优化设置生成演示文稿...")
                        # Chinese-specific settings are handled in pptx_export.py functions

                    export_pptx(
                        template_path=template_path,
                        markdown_path=markdown_path,
                        output_path=output_path,
                        project_name=project_name,
                        excel_file_path=excel_file_path,
                        language=language
                    )

                else:  # ALL - Generate BS first, then IS, then merge
                    st.info("🔄 Generating combined Balance Sheet and Income Statement presentation...")

                    # Check if both content files exist
                    bs_markdown_path = "fdd_utils/bs_content.md"
                    is_markdown_path = "fdd_utils/is_content.md"

                    if not os.path.exists(bs_markdown_path):
                        st.error(f"❌ Balance Sheet content file not found: {bs_markdown_path}")
                        st.info("💡 Please run AI processing first to generate content for PowerPoint export.")
                        return

                    if not os.path.exists(is_markdown_path):
                        st.error(f"❌ Income Statement content file not found: {is_markdown_path}")
                        st.info("💡 Please run AI processing first to generate content for PowerPoint export.")
                        return

                    # Generate temporary files for BS and IS
                    import tempfile
                    import shutil

                    with tempfile.TemporaryDirectory() as temp_dir:
                        bs_temp_path = os.path.join(temp_dir, "bs_temp.pptx")
                        is_temp_path = os.path.join(temp_dir, "is_temp.pptx")

                        # Generate BS presentation
                        if language == 'chinese':
                            st.info("📊 使用中文优化设置生成资产负债表...")
                        else:
                            st.info("📊 Generating Balance Sheet section...")
                        export_pptx(
                            template_path=template_path,
                            markdown_path=bs_markdown_path,
                            output_path=bs_temp_path,
                            project_name=project_name,
                            excel_file_path=excel_file_path
                        )

                        # Generate IS presentation
                        if language == 'chinese':
                            st.info("📈 使用中文优化设置生成损益表...")
                        else:
                            st.info("📈 Generating Income Statement section...")
                        export_pptx(
                            template_path=template_path,
                            markdown_path=is_markdown_path,
                            output_path=is_temp_path,
                            project_name=project_name,
                            excel_file_path=excel_file_path
                        )

                        # Merge the presentations
                        if language == 'chinese':
                            st.info("🔗 合并中文演示文稿...")
                        else:
                            st.info("🔗 Merging presentations...")
                        merge_presentations(bs_temp_path, is_temp_path, output_path)

                        if language == 'chinese':
                            st.success("✅ 中文组合演示文稿生成成功!")
                        else:
                            st.success("✅ Combined presentation generated successfully!")

                # Store session state and show download
                st.session_state['pptx_exported'] = True
                st.session_state['pptx_filename'] = output_filename
                st.session_state['pptx_path'] = output_path

                # Immediately show download button after successful generation
                if os.path.exists(output_path):
                    with open(output_path, "rb") as file:
                        if language == 'chinese':
                            download_label = f"📥 下载中文 PowerPoint: {output_filename}"
                        else:
                            download_label = f"📥 Download English PowerPoint: {output_filename}"

                        st.download_button(
                            label=download_label,
                            data=file.read(),
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )

                    if language == 'chinese':
                        st.success(f"✅ 中文 PowerPoint 生成完成: {output_filename}")
                    else:
                        st.success(f"✅ English PowerPoint generated successfully: {output_filename}")

            except FileNotFoundError as e:
                st.error(f"❌ Template file not found: {e}")
            except Exception as e:
                st.error(f"❌ Export failed: {e}")
                st.error(f"Error details: {str(e)}")

        # Create combined buttons for English and Chinese exports
        col1, col2 = st.columns([1, 1])

        # English PPTX Export Button (combines prepare and download)
        with col1:
            if st.button("📊 Export English PPTX", type="primary", use_container_width=True):
                export_pptx_with_download(selected_entity, statement_type, language='english')

        # Chinese PPTX Export Button (combines prepare and download)
        with col2:
            if st.button("📊 Export Chinese PPTX", type="primary", use_container_width=True):
                export_pptx_with_download(selected_entity, statement_type, language='chinese')

        # Download buttons are now integrated into the export functions above
        




# Helper function to parse and display bs_content.md by key
# display_bs_content_by_key moved to fdd_utils.content_utils

# clean_content_quotes moved to fdd_utils.content_utils

# display_ai_content_by_key moved to fdd_utils.display_utils

# JSON Content Access Helper Functions
# load_json_content moved to fdd_utils.content_utils

# parse_markdown_to_json moved to fdd_utils.content_utils

# get_content_from_json moved to fdd_utils.content_utils

# Offline content functions removed - system now only uses AI-generated content

# Removed offline content fallback function

# Removed get_offline_content function

# Removed get_offline_content_fallback function

# Offline validation functions removed - system now only uses AI-generated content

def generate_content_from_session_storage(entity_name):
    """Generate content files (JSON + Markdown) from session state storage (PERFORMANCE OPTIMIZED)"""
    try:
        # Get content from session state storage (fastest method)
        content_store = st.session_state.get('ai_content_store', {})
        current_statement_type = st.session_state.get('current_statement_type', 'BS')
        

        
        if not content_store:
            st.error("❌ No AI-generated content available. Please run AI processing first.")
            return
        
        # Get current statement type from session state
        current_statement_type = st.session_state.get('current_statement_type', 'BS')
        
        # Get category mappings from centralized config
        from fdd_utils.category_config import get_category_mapping
        category_mapping, name_mapping = get_category_mapping(current_statement_type, entity_name)
        
        # Generate JSON content from session storage (for AI2 easy access)
        json_content = {
            'metadata': {
                'entity_name': entity_name,
                'generated_timestamp': datetime.datetime.now().isoformat(),
                'session_id': getattr(st.session_state.get('ai_logger'), 'session_id', 'default'),
                'total_keys': len(content_store)
            },
            'categories': {},
            'keys': {}
        }
        
        # Removed: st.info(f"📊 Generating content files from session storage for {len(content_store)} keys")
        
        # Filter content store based on current statement type
        if current_statement_type == "IS":
            # For IS, only process IS-related keys
            is_keys = ['OI', 'OC', 'Tax and Surcharges', 'GA', 'Fin Exp', 'Cr Loss', 'Other Income', 'Non-operating Income', 'Non-operating Exp', 'Income tax', 'LT DTA']
            filtered_content_store = {k: v for k, v in content_store.items() if k in is_keys}
        elif current_statement_type == "BS":
            # For BS, only process BS-related keys
            bs_keys = ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA', 'Other NCA', 'IP', 'NCA', 'AP', 'Taxes payable', 'OP', 'Capital', 'Reserve']
            filtered_content_store = {k: v for k, v in content_store.items() if k in bs_keys}
        else:
            # For ALL, use all content
            filtered_content_store = content_store
        

        
        # Process content by category
        for category, items in category_mapping.items():
            json_content['categories'][category] = []
            
            for item in items:
                full_name = name_mapping[item]
                
                # Get latest content from session storage (could be Agent 1, 2, or 3 version)
                if item in filtered_content_store:
                    key_data = filtered_content_store[item]

                    latest_content = key_data.get('current_content', key_data.get('agent1_content', ''))
                    
                    # Determine content source
                    if 'agent3_content' in key_data:
                        content_source = "agent3_final"
                        source_timestamp = key_data.get('agent3_timestamp')
                        # Use agent3_content if available and it's Chinese
                        agent3_content = key_data['agent3_content']
                        if agent3_content and any('\u4e00' <= char <= '\u9fff' for char in agent3_content):
                            latest_content = agent3_content
                    elif 'agent2_content' in key_data:
                        content_source = "agent2_validated"
                        source_timestamp = key_data.get('agent2_timestamp')
                    else:
                        content_source = "agent1_original"
                        source_timestamp = key_data.get('agent1_timestamp')
                    
                    # Removed version display message for cleaner UI
                else:
                    latest_content = f"No information available for {item}"
                    content_source = "none"
                    source_timestamp = None
                    st.write(f"  • {item}: No content found")
                
                # Clean the content
                cleaned_content = clean_content_quotes(latest_content)
                
                # Add to JSON structure
                key_info = {
                    'key': item,
                    'display_name': full_name,
                    'content': cleaned_content,
                    'content_source': content_source,
                    'source_timestamp': source_timestamp,
                    'length': len(cleaned_content),
                    'category': category
                }
                
                json_content['categories'][category].append(key_info)
                json_content['keys'][item] = key_info
        
        # Get current statement type to determine file names
        current_statement_type = st.session_state.get('current_statement_type', 'BS')
        
        # Save JSON format (for AI2 easy access)
        if current_statement_type == "IS":
            json_file_path = 'fdd_utils/is_content.json'
            md_file_path = 'fdd_utils/is_content.md'
        elif current_statement_type == "BS":
            json_file_path = 'fdd_utils/bs_content.json'
            md_file_path = 'fdd_utils/bs_content.md'
        else:  # ALL - create comprehensive file
            json_file_path = 'fdd_utils/all_content.json'
            md_file_path = 'fdd_utils/all_content.md'
            
        with open(json_file_path, 'w', encoding='utf-8') as file:
            json.dump(json_content, file, indent=2, ensure_ascii=False)
        
        # Also generate markdown for PowerPoint compatibility
        markdown_lines = []
        for category, items in category_mapping.items():
            markdown_lines.append(f"## {category}\n")
            for item in items:
                full_name = name_mapping[item]
                key_info = json_content['keys'].get(item)
                if key_info:
                    cleaned_content = key_info['content']
                else:
                    cleaned_content = f"No information available for {item}"
                markdown_lines.append(f"### {full_name}\n{cleaned_content}\n")
        
        markdown_text = "\n".join(markdown_lines)
        
        # For ALL statement type, create comprehensive content file
        if current_statement_type == "ALL":
            # Also create BS and IS specific files for PowerPoint compatibility
            bs_keys = ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA', 'Other NCA', 'IP', 'NCA', 'AP', 'Taxes payable', 'OP', 'Capital', 'Reserve']
            is_keys = ['OI', 'OC', 'Tax and Surcharges', 'GA', 'Fin Exp', 'Cr Loss', 'Other Income', 'Non-operating Income', 'Non-operating Exp', 'Income tax', 'LT DTA']
            
            # Create BS content file
            bs_content_store = {k: v for k, v in content_store.items() if k in bs_keys}
            if bs_content_store:
                bs_json_content = {
                    'metadata': {
                        'generated_at': datetime.datetime.now().strftime('%Y-%m-%d'),
                        'format_version': '1.0',
                        'description': 'Balance Sheet content data'
                    },
                    'categories': {},
                    'keys': {}
                }
                
                # Process BS content
                for category, items in category_mapping.items():
                    bs_json_content['categories'][category] = []
                    for item in items:
                        if item in bs_content_store:
                            full_name = name_mapping[item]
                            key_data = bs_content_store[item]
                            latest_content = key_data.get('current_content', key_data.get('agent1_content', ''))
                            cleaned_content = clean_content_quotes(latest_content)

                            # Skip items with no information available
                            if not cleaned_content or "no information available" in cleaned_content.lower():
                                continue

                            key_info = {
                                'key': item,
                                'display_name': full_name,
                                'content': cleaned_content,
                                'content_source': 'agent1_original',
                                'source_timestamp': key_data.get('agent1_timestamp'),
                                'length': len(cleaned_content),
                                'category': category
                            }

                            bs_json_content['categories'][category].append(key_info)
                            bs_json_content['keys'][item] = key_info
                
                # Save BS content files
                with open('fdd_utils/bs_content.json', 'w', encoding='utf-8') as file:
                    json.dump(bs_json_content, file, indent=2, ensure_ascii=False)
                
                # Generate BS markdown
                bs_markdown_lines = []
                for category, items in category_mapping.items():
                    bs_markdown_lines.append(f"## {category}\n")
                    for item in items:
                        if item in bs_content_store:
                            full_name = name_mapping[item]
                            key_info = bs_json_content['keys'].get(item)
                            if key_info and key_info['content'] and "no information available" not in key_info['content'].lower():
                                cleaned_content = key_info['content']
                                bs_markdown_lines.append(f"### {full_name}\n{cleaned_content}\n")
                            # Skip items with no information or empty content
                
                bs_markdown_text = "\n".join(bs_markdown_lines)
                with open('fdd_utils/bs_content.md', 'w', encoding='utf-8') as file:
                    file.write(bs_markdown_text)
            
            # Create IS content file
            is_content_store = {k: v for k, v in content_store.items() if k in is_keys}
            if is_content_store:
                is_json_content = {
                    'metadata': {
                        'generated_at': datetime.datetime.now().strftime('%Y-%m-%d'),
                        'format_version': '1.0',
                        'description': 'Income Statement content data'
                    },
                    'categories': {},
                    'keys': {}
                }
                
                # Process IS content
                for category, items in category_mapping.items():
                    is_json_content['categories'][category] = []
                    for item in items:
                        if item in is_content_store:
                            full_name = name_mapping[item]
                            key_data = is_content_store[item]
                            latest_content = key_data.get('current_content', key_data.get('agent1_content', ''))
                            cleaned_content = clean_content_quotes(latest_content)

                            # Skip items with no information available
                            if not cleaned_content or "no information available" in cleaned_content.lower():
                                continue

                            key_info = {
                                'key': item,
                                'display_name': full_name,
                                'content': cleaned_content,
                                'content_source': 'agent1_original',
                                'source_timestamp': key_data.get('agent1_timestamp'),
                                'length': len(cleaned_content),
                                'category': category
                            }

                            is_json_content['categories'][category].append(key_info)
                            is_json_content['keys'][item] = key_info
                
                # Save IS content files
                with open('fdd_utils/is_content.json', 'w', encoding='utf-8') as file:
                    json.dump(is_json_content, file, indent=2, ensure_ascii=False)
                
                # Generate IS markdown
                is_markdown_lines = []
                for category, items in category_mapping.items():
                    is_markdown_lines.append(f"## {category}\n")
                    for item in items:
                        if item in is_content_store:
                            full_name = name_mapping[item]
                            key_info = is_json_content['keys'].get(item)
                            if key_info and key_info['content'] and "no information available" not in key_info['content'].lower():
                                cleaned_content = key_info['content']
                                is_markdown_lines.append(f"### {full_name}\n{cleaned_content}\n")
                            # Skip items with no information or empty content
                
                is_markdown_text = "\n".join(is_markdown_lines)
                with open('fdd_utils/is_content.md', 'w', encoding='utf-8') as file:
                    file.write(is_markdown_text)
        
        # Save markdown format (for PowerPoint export)
        with open(md_file_path, 'w', encoding='utf-8') as file:
            file.write(markdown_text)
        
        # Removed: st.success(f"✅ Generated {json_file_path} (AI-friendly) and {md_file_path} (PowerPoint-compatible)")
        return True
        
    except Exception as e:
        st.error(f"Error generating content from session storage: {e}")
        return False

def generate_markdown_from_ai_results(ai_results, entity_name):
    """Generate markdown content file from AI results following the old version pattern"""
    try:
        # Get current statement type from session state
        current_statement_type = st.session_state.get('current_statement_type', 'BS')
        
        # Define category mappings based on entity name and statement type
        if current_statement_type == "IS":
            # Income Statement categories
            if entity_name in ['Ningbo', 'Nanjing']:
                name_mapping = DISPLAY_NAME_MAPPING_NB_NJ
                category_mapping = {
                    'Revenue': ['OI', 'Other Income', 'Non-operating Income'],
                    'Expenses': ['OC', 'GA', 'Fin Exp', 'Cr Loss', 'Non-operating Exp'],
                    'Taxes': ['Tax and Surcharges', 'Income tax'],
                    'Other': ['LT DTA']
                }
            else:  # Haining and others
                name_mapping = DISPLAY_NAME_MAPPING_DEFAULT
                category_mapping = {
                    'Revenue': ['OI', 'Other Income', 'Non-operating Income'],
                    'Expenses': ['OC', 'GA', 'Fin Exp', 'Cr Loss', 'Non-operating Exp'],
                    'Taxes': ['Tax and Surcharges', 'Income tax'],
                    'Other': ['LT DTA']
                }
        else:
            # Balance Sheet categories (default)
            if entity_name in ['Ningbo', 'Nanjing']:
                name_mapping = DISPLAY_NAME_MAPPING_NB_NJ
                category_mapping = {
                    'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                    'Non-current Assets': ['IP', 'Other NCA'],
                    'Liabilities': ['AP', 'Taxes payable', 'OP'],
                    'Equity': ['Capital']
                }
            else:  # Haining and others
                name_mapping = DISPLAY_NAME_MAPPING_DEFAULT
                category_mapping = {
                    'Current Assets': ['Cash', 'AR', 'Prepayments', 'OR', 'Other CA'],
                    'Non-current Assets': ['IP', 'Other NCA'],
                    'Liabilities': ['AP', 'Taxes payable', 'OP'],
                    'Equity': ['Capital', 'Reserve']
                }
        
        # Generate markdown content
        markdown_lines = []
        for category, items in category_mapping.items():
            markdown_lines.append(f"## {category}\n")
            for item in items:
                full_name = name_mapping[item]
                info = ai_results.get(item, f"No information available for {item}")
                
                # Clean the content
                cleaned_info = clean_content_quotes(info)
                
                markdown_lines.append(f"### {full_name}\n{cleaned_info}\n")
        
        markdown_text = "\n".join(markdown_lines)
        
        # Get current statement type to determine file name
        current_statement_type = st.session_state.get('current_statement_type', 'BS')
        
        # Write to appropriate file
        if current_statement_type == "IS":
            file_path = 'fdd_utils/is_content.md'
        else:  # BS or ALL
            file_path = 'fdd_utils/bs_content.md'
            
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(markdown_text)
        
        return True
        
    except Exception as e:
        print(f"Error generating markdown: {e}")
        return False

# display_ai_prompt_by_key moved to fdd_utils.display_utils

# --- For AI1/2/3 prompt/debug output, use a separate file ---
def write_prompt_debug_content(filtered_keys, sections_by_key):
    with open("fdd_utils/bs_prompt_debug.md", "w", encoding="utf-8") as f:
        for key in filtered_keys:
            if key in sections_by_key and sections_by_key[key]:
                f.write(f"## {get_key_display_name(key)}\n")
                for section in sections_by_key[key]:
                    df = section['data']
                    df_clean = df.dropna(axis=1, how='all')
                    for idx, row in df_clean.iterrows():
                        row_str = " | ".join(str(x) for x in row if pd.notna(x) and str(x).strip() != "None")
                        if row_str:
                            f.write(f"- {row_str}\n")
                f.write("\n")

# --- In your AI1/2/3 or debug logic, call write_prompt_debug_content instead of writing to bs_content.md ---
# Example usage:
# write_prompt_debug_content(filtered_keys, sections_by_key)

# --- For PowerPoint export, always use bs_content.md ---
# (No changes needed here, just ensure you do NOT overwrite bs_content.md in prompt/debug logic)
# export_pptx(
#     template_path=template_path,
#     markdown_path="utils/bs_content.md",
#     output_path=output_path,
#     project_name=project_name
# )

def run_agent_1(filtered_keys, ai_data, external_progress=None, language='English'):
    """Run Agent 1: Content Generation for all keys"""
    try:
        import time

        logger = st.session_state.ai_logger
        # Keep content section minimal; avoid extra headings that duplicate the main status
        
        # Get data from ai_data
        entity_name = ai_data.get('entity_name', '')
        entity_keywords = ai_data.get('entity_keywords', [])
        
        # Create temporary file for processing
        temp_file_path = None
        try:
            if 'uploaded_file_data' in st.session_state:
                # Use a unique filename to avoid conflicts
                unique_filename = f"databook_{uuid.uuid4().hex[:8]}.xlsx"
                temp_file_path = os.path.join(tempfile.gettempdir(), unique_filename)
                
                with open(temp_file_path, 'wb') as tmp_file:
                    tmp_file.write(st.session_state['uploaded_file_data'])
                # Temp file created silently
            else:
                # Fallback: use existing databook.xlsx
                if os.path.exists('databook.xlsx'):
                    temp_file_path = 'databook.xlsx'
                    st.info("📄 Using existing databook.xlsx")
                else:
                    st.error("❌ No databook available for processing")
                    return {}
        except Exception as e:
            st.error(f"❌ Error creating temporary file: {e}")
            return {}
        
        # Get the actual prompts that will be sent to AI by calling process_keys
        # We need to capture the real prompts with table data
        try:
        # Load prompts from prompts.json file
            with open('fdd_utils/prompts.json', 'r', encoding='utf-8') as f:
                prompts_config = json.load(f)

            # Select language-specific prompts
            language_key = 'chinese' if language == '中文' else 'english'
            system_prompts = prompts_config.get('system_prompts', {}).get(language_key, {})



            actual_system_prompt = system_prompts.get('Agent 1', '')
            if not actual_system_prompt:
                # Fallback to English if language-specific prompt not found
                print(f"⚠️ WARNING: No Agent 1 prompt found for {language_key}, falling back to English")
                actual_system_prompt = prompts_config.get('system_prompts', {}).get('english', {}).get('Agent 1', '')



                if not actual_system_prompt:
                    from fdd_utils.prompt_templates import get_fallback_system_prompt
                    actual_system_prompt = get_fallback_system_prompt()
            
            # Add entity placeholder instructions to system prompt
            from fdd_utils.prompt_templates import get_entity_instructions
            actual_system_prompt += get_entity_instructions().format(entity_name=entity_name)
        except Exception as e:
            actual_system_prompt = "Error loading prompts.json"
            print(f"Error loading prompts: {e}")
        
        # Get sections data for each key to build actual user prompts
        sections_by_key = ai_data.get('sections_by_key', {})
        
        # Log Agent 1 input for all keys with actual prompts
        for key in filtered_keys:
            # Get the actual table data for this key with proper cleaning
            key_sections = sections_by_key.get(key, [])
            key_sections_str = []
            
            for section in key_sections:
                if isinstance(section, dict):
                    # Handle dict serialization
                    try:
                        key_sections_str.append(json.dumps(section, indent=2, default=str))
                    except:
                        key_sections_str.append(str(section))
                elif hasattr(section, 'to_string'):  # DataFrame
                    # Clean the DataFrame first - remove columns that are all None/NaN
                    df = section.copy()
                    for col in list(df.columns):
                        if df[col].isna().all() or (df[col].astype(str) == 'None').all():
                            df = df.drop(columns=[col])

                        
                        # Convert to string with proper formatting
                        df_str = df.to_string(index=False, na_rep='')
                        lines = df_str.split('\n')
                        cleaned_lines = []
                        for line in lines:
                            # Skip lines that are mostly empty or just separators
                            if line.strip() and not line.strip().replace('|', '').replace('-', '').replace(' ', '').replace('+', '') == '':
                                cleaned_lines.append(line)
                        key_sections_str.append('\n'.join(cleaned_lines))
                else:
                    key_sections_str.append(str(section))
            
            key_tables = "\n".join(key_sections_str) if key_sections_str else "No table data available"
            
            # Build the actual user prompt that matches what gets sent to AI
            pattern = ai_data.get('pattern', {}).get(key, {})
            pattern_json = json.dumps(pattern, indent=2)
            
            # Get financial figure for this key
    
            financial_figures = find_financial_figures_with_context_check(temp_file_path, get_tab_name(entity_name), None, convert_thousands=False)
            financial_figure_info = f"{key}: {get_financial_figure(financial_figures, key)}"
            
            # Build the actual user prompt using templates from prompts.json, then inject dynamic data
            try:
                with open('fdd_utils/prompts.json', 'r', encoding='utf-8') as f:
                    prompts_cfg = json.load(f)
                user_prompts_cfg = prompts_cfg.get('user_prompts', {})
                generic_cfg = prompts_cfg.get('generic_prompt', {})
                # Pick key-specific config if available; else generic
                key_prompt_cfg = user_prompts_cfg.get(key, generic_cfg)
            except Exception:
                key_prompt_cfg = {}

            title = key_prompt_cfg.get('title', f'{get_key_display_name(key)} Analysis for {entity_name}')
            description = key_prompt_cfg.get('description', f'Analyze {get_key_display_name(key)} using worksheet data')
            analysis_points = key_prompt_cfg.get('analysis_points', [])
            key_questions = key_prompt_cfg.get('key_questions', [])
            data_sources = key_prompt_cfg.get('data_sources', [])

            # Compose user prompt template then inject dynamic dataset and patterns
            prompt_lines = [
                f"{title}",
                f"{description}",
                "",
                f"AVAILABLE PATTERNS: {pattern_json}",
                f"FINANCIAL FIGURE: {financial_figure_info}",
                f"DATA SOURCE: {key_tables}",
                "",
            ]

            if analysis_points:
                prompt_lines.append("REQUIRED ANALYSIS:")
                for i, p in enumerate(analysis_points, 1):
                    prompt_lines.append(f"{i}. {p.replace('[ENTITY_NAME]', entity_name)}")
                prompt_lines.append("")

            if key_questions:
                prompt_lines.append("KEY QUESTIONS:")
                for q in key_questions:
                    prompt_lines.append(f"- {q.replace('[ENTITY_NAME]', entity_name)}")
                prompt_lines.append("")

            if data_sources:
                prompt_lines.append("DATA SOURCES:")
                for s in data_sources:
                    prompt_lines.append(f"- {s}")
                prompt_lines.append("")

            # Output requirements from centralized template
            from fdd_utils.prompt_templates import get_output_requirements
            prompt_lines += get_output_requirements()

            actual_user_prompt = "\n".join(prompt_lines)
            
            context_data = {
                'entity': entity_name,
                'key': key,
                'financial_figure': financial_figure_info,
                'table_data_length': len(key_tables),
                'patterns_count': len(pattern) if pattern else 0
            }
            
            # Store the actual prompts that will be sent to AI
            actual_prompts = {
                'system_prompt': actual_system_prompt,
                'user_prompt': actual_user_prompt,
                'context': context_data,
                'table_data': key_tables[:2000] + "..." if len(key_tables) > 2000 else key_tables,  # Show more structured data
                'structured_tables': []  # Will store parsed structured tables
            }
            
            # Parse structured tables from the key_tables content
            try:
                # Extract structured table information from the markdown content
                table_sections = key_tables.split('## ')
                for section in table_sections[1:]:  # Skip first empty section
                    lines = section.strip().split('\n')
                    if len(lines) > 0:
                        table_info = {
                            'table_name': lines[0],
                            'entity': 'Unknown',
                            'date': 'Unknown',
                            'currency': 'CNY',
                            'multiplier': 1,
                            'items': [],
                            'total': 'Unknown'
                        }
                        
                        for line in lines[1:]:
                            if line.startswith('**Entity:**'):
                                table_info['entity'] = line.replace('**Entity:**', '').strip()
                            elif line.startswith('**Date:**'):
                                table_info['date'] = line.replace('**Date:**', '').strip()
                            elif line.startswith('**Currency:**'):
                                table_info['currency'] = line.replace('**Currency:**', '').strip()
                            elif line.startswith('**Multiplier:**'):
                                table_info['multiplier'] = int(line.replace('**Multiplier:**', '').strip())
                            elif line.startswith('**Total:**'):
                                table_info['total'] = line.replace('**Total:**', '').strip()
                            elif line.startswith('- ') and ':' in line:
                                # Parse item line like "- Deposits with banks: 9076000"
                                item_parts = line[2:].split(': ')
                                if len(item_parts) == 2:
                                    table_info['items'].append({
                                        'description': item_parts[0],
                                        'amount': item_parts[1]
                                    })
                            
                            actual_prompts['structured_tables'].append(table_info)
            except Exception as e:
                print(f"Error parsing structured tables for logging: {e}")
            
            logger.log_agent_input('agent1', key, actual_system_prompt, actual_user_prompt, context_data, actual_prompts)
        
        # Process ALL keys at once with proper tqdm progress (1/9, 2/9, etc.)
        start_time = time.time()
        
        # Create/Reuse a single progress bar and status text
        if external_progress:
            progress_bar = external_progress.get('bar')
            status_text = external_progress.get('status')
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
        
        # Create progress callback for Streamlit with ETA (supports combined 2-stage runs)
        start_time = time.time()
        total = max(1, len(filtered_keys))
        def update_progress(progress, message):
            try:
                # If running in combined mode, map this stage's progress into overall combined progress
                combined = external_progress.get('combined') if external_progress else None
                if combined and isinstance(combined, dict):
                    stages = combined.get('stages', 2)
                    stage_index = combined.get('stage_index', 0)
                    stage_weight = 1.0 / max(1, stages)
                    combined_progress = min(1.0, max(0.0, stage_index * stage_weight + progress * stage_weight))
                    progress_bar.progress(combined_progress)
                else:
                    progress_bar.progress(progress)
            except Exception:
                pass
            try:
                # Estimate ETA from progress
                combined = external_progress.get('combined') if external_progress else None
                now = time.time()
                if combined and isinstance(combined, dict):
                    # Combined ETA from overall combined progress
                    start_all = combined.get('start_time') or start_time
                    stages = combined.get('stages', 2)
                    stage_index = combined.get('stage_index', 0)
                    stage_weight = 1.0 / max(1, stages)
                    overall_progress = max(0.001, min(0.999, stage_index * stage_weight + progress * stage_weight))
                    elapsed_all = now - start_all
                    remaining = int(elapsed_all * (1 - overall_progress) / overall_progress)
                else:
                    elapsed = now - start_time
                    remaining = int(elapsed * (1 - progress) / max(progress, 0.001))
                mins, secs = divmod(max(0, remaining), 60)
                eta_str = f"ETA {mins:02d}:{secs:02d}" if progress < 1 else "ETA 00:00"

                # Enhanced status message with more details
                # Extract current key from various message formats
                key_patterns = [
                    r'Processing (\w+)',
                    r'Loading data for (\w+)',
                    r'Processing (\w+) data',
                    r'AI generating (\w+)',
                    r'Sending (\w+) to',
                    r'Processing (\w+) response',
                    r'Completed (\w+)'
                ]

                current_key = "Unknown"
                for pattern in key_patterns:
                    match = re.search(pattern, message)
                    if match:
                        current_key = match.group(1)
                        break

                # Extract more details from the message and determine phase
                if "Loading data" in message:
                    phase = "📊 Data Loading"
                elif "Processing data" in message:
                    phase = "📈 Data Processing"
                elif "AI generating" in message:
                    phase = "🤖 AI Generation"
                elif "Sending" in message and "to" in message:
                    phase = "📤 AI Request"
                elif "Processing response" in message:
                    phase = "📥 Response Processing"
                elif "Completed" in message:
                    phase = "✅ Completed"
                elif "AI processing completed" in message:
                    phase = "🎉 All Done"
                    current_key = "All Keys"
                else:
                    phase = "🔄 Processing"

                # Show current key and phase information
                enhanced_message = f"{phase} — {current_key} — {eta_str}"

                # Add processing statistics if available
                if progress > 0 and progress < 1:
                    processed_keys = int(progress * total)
                    enhanced_message += f" ({processed_keys}/{total} keys)"

                status_text.text(enhanced_message)
            except Exception:
                pass
        
        # Get processed table data from session state
        processed_table_data = ai_data.get('sections_by_key', {})
        
        # Get AI model settings from session state
        use_local_ai = st.session_state.get('use_local_ai', False)
        use_openai = st.session_state.get('use_openai', False)

        # Get selected language for AI processing
        selected_language = st.session_state.get('selected_language', 'English')
        language_key = 'chinese' if selected_language == '中文' else 'english'

        # Clear JSON cache to ensure fresh prompts are loaded
        from common.assistant import clear_json_cache
        clear_json_cache()




        results = process_keys(
            keys=filtered_keys,  # All keys at once
            entity_name=entity_name,
            entity_helpers=entity_keywords,
            input_file=temp_file_path,
            mapping_file="fdd_utils/mapping.json",
            pattern_file="fdd_utils/pattern.json",
            config_file='fdd_utils/config.json',
            prompts_file='fdd_utils/prompts.json',
            use_ai=True,
            progress_callback=update_progress,
            processed_table_data=processed_table_data,
            use_local_ai=use_local_ai,
            use_openai=use_openai,
            language=language_key
        )
        
        # Process results
        for key in filtered_keys:
            result = results[key]
        
        processing_time = time.time() - start_time
        
        # Log Agent 1 output for each key with pattern information
        for key in filtered_keys:
            key_result = results.get(key, {})
            if isinstance(key_result, dict):
                content = key_result.get('content', f"No result generated for {key}")
                pattern_used = key_result.get('pattern_used', 'Unknown')
                table_data = key_result.get('table_data', '')
                financial_figure = key_result.get('financial_figure', 0)
            else:
                content = key_result
                pattern_used = 'Unknown'
                table_data = ''
                financial_figure = 0
            
            # Create enhanced output with pattern information
            enhanced_output = {
                'content': content,
                'pattern_used': pattern_used,
                'table_data': table_data,
                'financial_figure': financial_figure,
                'entity_name': entity_name
            }
            
            logger.log_agent_output('agent1', key, enhanced_output, processing_time / len(filtered_keys))
        
        # Return results
        
        st.success(f"🎉 Agent 1 completed all {len(filtered_keys)} keys in {processing_time:.2f}s")
        return results
            
    except RuntimeError as e:
        # AI-specific errors
        if "AI processing is required" in str(e):
            st.error("❌ **AI Processing Required**")
            st.error("This application requires AI services to function. Please check your configuration.")
            st.error(f"Error: {e}")
        elif "AI services are not available" in str(e):
            st.error("❌ **AI Services Unavailable**")
            st.error("DeepSeek AI services are not available. Please check your internet connection and API configuration.")
            st.error(f"Error: {e}")
        else:
            st.error(f"❌ **AI Error**: {e}")
        return {}
    except Exception as e:
        st.error(f"❌ **Unexpected Error**: {e}")
        return {}
    finally:
        # Clean up temp file if created
        try:
            if temp_file_path and temp_file_path != 'databook.xlsx' and os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
        except PermissionError:
            # On Windows, file may still be locked momentarily; retry once
            try:
                import time
                time.sleep(0.2)
                if os.path.exists(temp_file_path):
                    os.unlink(temp_file_path)
            except Exception:
                pass

def run_chinese_translator(filtered_keys, agent1_results, ai_data, external_progress=None):
    """Simple Chinese Translation Agent: Process proofread content one by one using AI"""
    try:
        import json
        import time
        from common.assistant import generate_response, load_config, initialize_ai_services

        # Initialize is_cli flag properly
        is_cli = True  # Default to CLI mode

        # Determine if we're in Streamlit context
        try:
            import streamlit as st
            _ = st.session_state
            is_cli = False
        except (ImportError, AttributeError):
            is_cli = True

        # Initialize processing



        # Setup tqdm progress bar
        if is_cli:
            progress_bar = tqdm(total=len(filtered_keys), desc="🌐 中文翻译", unit="key")
        else:
            progress_bar = None

        # Get AI model settings - try session state regardless of CLI mode
        use_local_ai = False
        use_openai = False

        try:
            # Try to get settings from session state (works even in CLI mode if Streamlit is imported)
            import streamlit as st
            use_local_ai = st.session_state.get('use_local_ai', False)
            use_openai = st.session_state.get('use_openai', False)

            # Also check selected_provider for more specific control
            selected_provider = st.session_state.get('selected_provider')
            if selected_provider == 'Open AI':
                use_openai = True
                use_local_ai = False
            elif selected_provider == 'Local AI' or selected_provider == 'Server AI':
                use_local_ai = True
                use_openai = False

        except Exception as e:
            # Fallback to config defaults silently

            # If still no settings, try to detect from config
            if not use_local_ai and not use_openai:
                try:
                    config_details = load_config('fdd_utils/config.json')
                    if config_details.get('LOCAL_AI_API_BASE'):
                        use_local_ai = True
                    elif config_details.get('OPENAI_API_KEY'):
                        use_openai = True
                except Exception as e:
                    # Config loading failed, continue with defaults
                    pass

        # Get AI data
        entity_name = ai_data.get('entity_name', '')
        entity_keywords = ai_data.get('entity_keywords', [])
        sections_by_key = ai_data.get('sections_by_key', {})

        # Load configuration
        config_details = load_config('fdd_utils/config.json')

        oai_client, _ = initialize_ai_services(config_details, use_local=use_local_ai, use_openai=use_openai)

            # system_prompt moved to higher scope above

        # Get model name
        if use_local_ai:
            model = config_details.get('LOCAL_AI_CHAT_MODEL', 'local-model')
        elif use_openai:
            model = config_details.get('OPENAI_CHAT_MODEL', 'gpt-4o-mini-2024-07-18')
        else:
            model = config_details.get('DEEPSEEK_CHAT_MODEL', 'deepseek-chat')

        # Create temporary file for processing
        temp_file_path = None
        try:
            if not is_cli and 'uploaded_file_data' in st.session_state:
                unique_filename = f"databook_{uuid.uuid4().hex[:8]}.xlsx"
                temp_file_path = os.path.join(tempfile.gettempdir(), unique_filename)
                with open(temp_file_path, 'wb') as tmp_file:
                    tmp_file.write(st.session_state['uploaded_file_data'])
            else:
                if os.path.exists('databook.xlsx'):
                    temp_file_path = 'databook.xlsx'
                else:
                    if not is_cli:
                        st.error("❌ No databook available for processing")
                    else:
                        print("❌ No databook available for processing")
                    return agent1_results
        except Exception as e:
            if not is_cli:
                st.error(f"❌ Error creating temporary file: {e}")
            else:
                print(f"❌ Error creating temporary file: {e}")
                temp_file_path = None

        # Prepare processed table data (only if temp_file_path is available)
        # OPTIMIZATION: Skip expensive Excel processing during translation
        # The translation function only needs text content, not table data
        # Table data processing is only needed for PPTX generation, not translation
        processed_table_data = {}

        # Create minimal empty data structure (fast)
        # This satisfies the process_keys API without expensive Excel processing
        for key in filtered_keys:
            processed_table_data[key] = {}

        # NOTE: Excel processing skipped for performance - translation only needs text content
        # PPTX generation will handle Excel processing separately if needed

        # Load system prompt from centralized template (move to higher scope)
        from fdd_utils.prompt_templates import get_translation_prompts
        prompts = get_translation_prompts()
        system_prompt = prompts["chinese_translator_system"]

        translated_results = {}
        start_time = time.time()

        # Enhanced progress tracking for translation
        if external_progress:
            def update_translation_progress(current_idx, total, current_key, status_msg=""):
                try:
                    progress_pct = (current_idx + 1) / total
                    bar = external_progress.get('bar')
                    status = external_progress.get('status')

                    if bar:
                        bar.progress(progress_pct)

                    if status:
                        eta_str = ""
                        if current_idx > 0:
                            elapsed = time.time() - start_time
                            avg_time = elapsed / (current_idx + 1)
                            remaining = total - current_idx - 1
                            eta_seconds = int(avg_time * remaining)
                            mins, secs = divmod(eta_seconds, 60)
                            eta_str = f" ETA {mins:02d}:{secs:02d}" if eta_seconds > 0 else ""

                        enhanced_msg = f"🌐 翻译中: {current_key} ({current_idx + 1}/{total}){eta_str}"
                        if status_msg:
                            enhanced_msg += f" - {status_msg}"

                        status.text(enhanced_msg)
                except Exception as e:
                    print(f"Progress update error: {e}")

            progress_callback = update_translation_progress
        else:
            progress_callback = None

        translated_results = {}
        start_time = time.time()

        for idx, key in enumerate(filtered_keys):
            try:
                # DEBUG: Show we're processing this key
                print(f"\n🔄 [{idx+1}/{len(filtered_keys)}] Processing key: {key}")

                # Get the proofread content
                content = agent1_results.get(key, '')

                if isinstance(content, dict):
                    # First try corrected_content (from proofreader), then content (from agent1)
                    content_text = content.get('corrected_content', '') or content.get('content', '')
                else:
                    content_text = str(content)

                # FILTER OUT SUMMARY SECTIONS - Don't translate summary content
                summary_keywords = ['summary', 'conclusion', 'overall', 'in summary', 'to summarize', 'key findings']
                if content_text and any(keyword in content_text.lower() for keyword in summary_keywords):
                    print(f"⚠️  Skipping {key} - contains summary-like content")
                    translated_results[key] = agent1_results.get(key, {})
                    if is_cli and progress_bar:
                        progress_bar.update(1)
                    continue

                if not content_text:
                    print(f"⚠️  Empty content for {key}, skipping")
                    translated_results[key] = agent1_results.get(key, {})
                    if is_cli and progress_bar:
                        progress_bar.update(1)
                    continue

                # Skip table information for simple translation
                table_info = ""  # Not needed for simple translation

                # Update progress
                if progress_callback:
                    progress_callback(idx, len(filtered_keys), key, "开始翻译")
                elif is_cli and progress_bar:
                    progress_bar.update(1)

                # CLEAN OUTPUT: Show only before/after translation
                print(f"\n🔄 [{idx+1}/{len(filtered_keys)}] 翻译中: {key}")
                print(f"📝 BEFORE: {content_text[:80]}{'...' if len(content_text) > 80 else ''}")

                # Use prompt from centralized template (NO SUMMARY SECTION)
                user_prompt = prompts["chinese_translator_user"].replace("{content_text}", content_text)

                # LOG AI PROCESSING
                try:
                    from fdd_utils.logging_config import log_ai_processing
                    ai_logger = log_ai_processing(
                        button_name="chinese_ai",
                        agent_name="Chinese Translator",
                        key=key,
                        system_prompt=system_prompt,
                        user_prompt=user_prompt
                    )
                    ai_logger.debug(f"Original content: {content_text}")
                except Exception as log_error:
                    print(f"⚠️ Logging error: {log_error}")

                # TIMING: Record time before AI call
                debug_start_time = time.time()

                # TIMING: Record time before AI call
                ai_call_start = time.time()
                print(f"⏱️  Starting AI call at {ai_call_start:.2f}s")

                # Call AI for translation
                translated_content = generate_response(
                    user_query=user_prompt,
                    system_prompt=system_prompt,
                    oai_client=oai_client,
                    context_content=table_info,
                    openai_chat_model=model,
                    entity_name=entity_name,
                    use_local_ai=use_local_ai
                )

                # LOG AI RESPONSE
                try:
                    if translated_content:
                        ai_logger.info(f"Translation successful - Response length: {len(translated_content)} chars")
                        ai_logger.debug(f"Translated content: {translated_content}")
                    else:
                        ai_logger.error(f"Translation failed - No response received")
                except Exception:
                    pass  # Ignore logging errors

                # CLEAN OUTPUT: Show only AFTER translation result
                if translated_content:
                    print(f"🌐 AFTER:  {translated_content[:80]}{'...' if len(translated_content) > 80 else ''}")
                else:
                    print(f"❌ FAILED: No translation received for {key}")

                # Clean the response
                if translated_content:
                    translated_content = translated_content.strip()
                    if translated_content.startswith('"') and translated_content.endswith('"'):
                        translated_content = translated_content[1:-1]

                    # Debug output for translation result
                    print(f"✅ 翻译完成: {key}")
                    print(f"🌐 译文预览: {translated_content[:50]}..." if len(translated_content) > 50 else f"🌐 译文: {translated_content}")

                    # Update progress to show completion
                    if progress_callback:
                        progress_callback(idx, len(filtered_keys), key, "翻译完成")

                    # Quality check for Chinese content
                        chinese_chars = sum(1 for char in translated_content if '\u4e00' <= char <= '\u9fff')
                        english_chars = sum(1 for char in translated_content if char.isascii() and char.isalnum())
                        total_chars = chinese_chars + english_chars

                        if total_chars > 0:
                            chinese_ratio = chinese_chars / total_chars
                            print(f"📊 中文占比: {chinese_ratio:.1%} ({chinese_chars}/{total_chars} 字符)")

                            if chinese_ratio < 0.3:
                                print(f"⚠️ 警告: 中文占比过低 ({chinese_ratio:.2%}) - 可能翻译失败")
                            elif chinese_ratio > 0.7:
                                print(f"✅ 良好: 中文占比正常 ({chinese_ratio:.1%})")
                            else:
                                print(f"ℹ️ 一般: 中文占比中等 ({chinese_ratio:.1%})")

                        # Additional check for common English words that should be translated
                        english_words = ['the', 'and', 'for', 'with', 'from', 'that', 'have', 'been', 'were']
                        found_english_words = [word for word in english_words if word in translated_content.lower()]
                        if found_english_words:
                            print(f"⚠️ 发现英文词汇: {', '.join(found_english_words)}")
                    else:
                        print(f"⚠️ 警告: 无法分析字符类型")

                    # Show detailed comparison
                    print(f"🔍 对比分析:")
                    print(f"   原文长度: {len(content_text)} 字符")
                    print(f"   译文长度: {len(translated_content)} 字符")
                    print(f"   长度变化: {len(translated_content) - len(content_text)} 字符")
                    print(f"{'─' * 60}")

                # Store result with explicit Chinese content - ENSURE UI CAN FIND IT
                result_data = agent1_results.get(key, {})
                if isinstance(result_data, dict):
                    # Store the translated content in ALL relevant fields to ensure maximum compatibility
                    # Priority: translated_content > corrected_content > content
                    result_data['content'] = translated_content or content_text
                    result_data['corrected_content'] = translated_content or content_text  # ← KEY FIX: UI looks here first
                    result_data['translated_content'] = translated_content  # ← Store Chinese content here
                    result_data['original_content'] = content_text
                    result_data['translated'] = True
                    result_data['is_chinese'] = bool(translated_content and any('\u4e00' <= char <= '\u9fff' for char in translated_content))
                else:
                    result_data = {
                        'content': translated_content or content_text,
                        'corrected_content': translated_content or content_text,  # ← KEY FIX: UI looks here first
                        'translated_content': translated_content,  # ← Store Chinese content here
                        'original_content': content_text,
                        'translated': True,
                        'is_chinese': bool(translated_content and any('\u4e00' <= char <= '\u9fff' for char in translated_content))
                    }
                translated_results[key] = result_data

                # DEBUG: Show what we're storing
                print(f"💾 STORING RESULT FOR {key}:")
                print(f"   📝 Original length: {len(content_text)} chars")
                print(f"   🌐 Translated length: {len(translated_content) if translated_content else 0} chars")
                print(f"   ✅ Is Chinese: {result_data.get('is_chinese', False)}")
                print(f"   📊 Content preview: {result_data['content'][:100]}..." if len(result_data['content']) > 100 else f"   📊 Content: {result_data['content']}")

            except Exception as e:
                if is_cli:
                    print(f"Error translating {key}: {e}")
                translated_results[key] = agent1_results.get(key, {})
        # Return translated results

        # Final summary and close progress bar
        total_processed = len([k for k in translated_results.keys() if translated_results[k]])
        success_rate = total_processed / len(filtered_keys) if filtered_keys else 0

        summary_msg = f"✅ 中文翻译完成 - 成功处理 {total_processed}/{len(filtered_keys)} 个项目 ({success_rate:.1%})"
        print(f"📈 Success rate: {success_rate:.1%}")
        print(f"⏱️  Total time: {time.time() - start_time:.2f} seconds")

        if is_cli and progress_bar:
            progress_bar.close()
            print(f"\n{summary_msg}")
            print(f"🔍 翻译质量统计:")
            print(f"{'─' * 80}")
            for key in translated_results:
                if translated_results[key]:
                    content = translated_results[key].get('content', '') if isinstance(translated_results[key], dict) else str(translated_results[key])
                    chinese_chars = sum(1 for char in content if '\u4e00' <= char <= '\u9fff')
                    total_chars = len(content)
                    if total_chars > 0:
                        ratio = chinese_chars / total_chars
                        status = "✅" if ratio > 0.5 else "⚠️" if ratio > 0.2 else "❌"
                        print(f"  {status} {key}: {ratio:.1%} 中文 ({chinese_chars}/{total_chars} 字符)")

                        # Show sample of Chinese content if available
                        if ratio > 0.5:
                            print(f"      🌐 中文预览: {content[:100]}{'...' if len(content) > 100 else ''}")
                        else:
                            print(f"      ❓ 内容可能仍为英文: {content[:100]}{'...' if len(content) > 100 else ''}")
            print(f"{'─' * 80}")

            # Show overall translation success summary
            successful_translations = 0
            total_keys = len(translated_results)
            for key in translated_results:
                if translated_results[key]:
                    content = translated_results[key].get('content', '') if isinstance(translated_results[key], dict) else str(translated_results[key])
                    chinese_chars = sum(1 for char in content if '\u4e00' <= char <= '\u9fff')
                    total_chars = len(content)
                    if total_chars > 0 and chinese_chars / total_chars > 0.5:
                        successful_translations += 1

            success_rate = successful_translations / total_keys if total_keys > 0 else 0
            print(f"🎯 翻译成功率: {successful_translations}/{total_keys} ({success_rate:.1%})")
        elif external_progress and external_progress.get('status'):
            external_progress['status'].text(summary_msg)

        print(f"\n{'='*60}")
        print(f"🎯 翻译任务完成总结:")
        print(f"   总项目数: {len(filtered_keys)}")
        print(f"   成功翻译: {total_processed}")
        print(f"   成功率: {success_rate:.1%}")
        print(f"   耗时: {time.time() - start_time:.1f} 秒")
        print(f"{'='*60}")

        return translated_results

    except Exception as e:
        if is_cli:
            print(f"❌ Chinese translation error: {e}")
        else:
            try:
                st.error(f"❌ Chinese translation error: {e}")
            except Exception:
                print(f"❌ Chinese translation error: {e}")
        if is_cli and progress_bar:
            progress_bar.close()

        # Return a clear indication of translation failure instead of original results
        failed_translation_results = {}
        for key in filtered_keys:
            failed_translation_results[key] = {
                'content': f"❌ 翻译失败: {str(e)}",
                'translation_failed': True,
                'original_content': agent1_results.get(key, '')
            }
        return failed_translation_results

def run_ai_proofreader(filtered_keys, agent1_results, ai_data, external_progress=None, language='English'):
    """Run AI Proofreader for all keys (Compliance, Figures, Entities, Grammar)."""
    try:
        # Initialize is_cli flag properly
        is_cli = True  # Default to CLI mode

        # Check if we're in a Streamlit environment (not CLI)
        try:
            import streamlit as st
            # Try to access Streamlit session state - if it succeeds, we're in Streamlit
            _ = st.session_state
            is_cli = False
        except (ImportError, AttributeError):
            # If Streamlit is not available or session state doesn't exist, we're in CLI
            is_cli = True

        # Only show verbose output in CLI mode
        if is_cli:
            print(f"🔍 run_ai_proofreader called with {len(filtered_keys)} keys")

        # Get logger only if available
        logger = None
        if not is_cli:
            try:
                logger = st.session_state.ai_logger
            except:
                pass

        # Model/provider selection - try session state regardless of CLI mode
        use_local_ai = False
        use_openai = False

        try:
            # Try to get settings from session state (works even in CLI mode if Streamlit is imported)
            import streamlit as st
            use_local_ai = st.session_state.get('use_local_ai', False)
            use_openai = st.session_state.get('use_openai', False)

            # Also check selected_provider for more specific control
            selected_provider = st.session_state.get('selected_provider')
            if selected_provider == 'Open AI':
                use_openai = True
                use_local_ai = False
            elif selected_provider == 'Local AI' or selected_provider == 'Server AI':
                use_local_ai = True
                use_openai = False

        except Exception as e:
            # Fallback to config detection
            try:
                from common.assistant import load_config
                config_details = load_config('fdd_utils/config.json')
                if config_details.get('LOCAL_AI_API_BASE'):
                    use_local_ai = True
                elif config_details.get('OPENAI_API_KEY'):
                    use_openai = True
            except Exception:
                pass

        proof_agent = ProofreadingAgent(use_local_ai=use_local_ai, use_openai=use_openai, language=language)

        results = {}
        entity_name = ai_data.get('entity_name', '')
        sections_by_key = ai_data.get('sections_by_key', {})

        # Enhanced progress tracking for proofreading
        if external_progress:
            def update_proofreader_progress(current_idx, total, current_key, status_msg=""):
                try:
                    progress_pct = (current_idx + 1) / total
                    bar = external_progress.get('bar')
                    status = external_progress.get('status')

                    if bar:
                        bar.progress(progress_pct)

                    if status:
                        eta_str = ""
                        if current_idx > 0:
                            elapsed = time.time() - start_time
                            avg_time = elapsed / (current_idx + 1)
                            remaining = total - current_idx - 1
                            eta_seconds = int(avg_time * remaining)
                            mins, secs = divmod(eta_seconds, 60)
                            eta_str = f" ETA {mins:02d}:{secs:02d}" if eta_seconds > 0 else ""

                        enhanced_msg = f"🔍 校对中: {current_key} ({current_idx + 1}/{total}){eta_str}"
                        if status_msg:
                            enhanced_msg += f" - {status_msg}"

                        status.text(enhanced_msg)
                except Exception as e:
                    print(f"Progress update error: {e}")

            progress_callback = update_proofreader_progress
        else:
            progress_callback = None

        # Setup tqdm progress bar for CLI processing
        if is_cli:
            progress_bar = tqdm(total=len(filtered_keys), desc="🔍 校对", unit="key")
        else:
            progress_bar = None

        start_time = time.time()

        # Process each key with proofreading
        for idx, key in enumerate(filtered_keys):
            try:
                # Update progress
                if progress_callback:
                    progress_callback(idx, len(filtered_keys), key, "开始校对")
                elif is_cli and progress_bar:
                    progress_bar.set_description(f"🔍 校对 {key} ({idx+1}/{len(filtered_keys)})")

                content = agent1_results.get(key, '')
                if isinstance(content, dict):
                    content_text = content.get('content', '')
                else:
                    content_text = str(content)

                # Debug output for all modes
                print(f"\n🔍 [{idx+1}/{len(filtered_keys)}] 校对中: {key}")
                print(f"📝 内容预览: {content_text[:50]}..." if len(content_text) > 50 else f"📝 内容: {content_text}")

                if is_cli and progress_bar:
                    progress_bar.update(1)

                if not content_text:
                    results[key] = agent1_results.get(key, {})
                    continue

                # Get sections for this key
                key_sections = sections_by_key.get(key, [])

                # DEBUG: Print proofreading input
                print(f"\n{'='*80}")

                print(f"{'='*80}")
                print(f"📝 CONTENT TO PROOFREAD ({len(content_text)} chars):")
                print(f"{content_text}")
                print(f"🔧 ENTITY NAME: {entity_name}")
                print(f"{'='*80}")

                # Process the content with proofreading agent
                proofread_result = proof_agent.proofread(content_text, key_sections, entity_name)

                # DEBUG: Print proofreading output
                print(f"\n{'='*80}")

                print(f"{'='*80}")
                if isinstance(proofread_result, dict):
                    print(f"📤 PROOFREAD RESULT:")
                    for k, v in proofread_result.items():
                        if k == 'corrected_content' and v:
                            print(f"  {k}: {v[:100]}..." if len(str(v)) > 100 else f"  {k}: {v}")
                        else:
                            print(f"  {k}: {v}")
                else:
                    print(f"📤 PROOFREAD RESULT: {proofread_result}")
                print(f"{'='*80}")

                # Debug output for proofreading result
                print(f"✅ 校对完成: {key}")
                if isinstance(proofread_result, dict):
                    corrected_content = proofread_result.get('corrected_content', '')
                    issues_found = len(proofread_result.get('issues', []))
                    print(f"🔧 校对结果: {issues_found} 个问题发现")
                    if corrected_content:
                        print(f"📝 修正内容预览: {corrected_content[:50]}..." if len(corrected_content) > 50 else f"📝 修正内容: {corrected_content}")

                # Update progress to show completion
                if progress_callback:
                    progress_callback(idx, len(filtered_keys), key, "校对完成")

                # Store result
                result_data = agent1_results.get(key, {})
                if isinstance(result_data, dict):
                    result_data.update(proofread_result)
                else:
                    result_data = proofread_result
                results[key] = result_data

            except Exception as e:
                if is_cli:
                    print(f"Error proofreading {key}: {e}")
                results[key] = agent1_results.get(key, {})

        # Final summary and close progress bar
        total_processed = len([k for k in results.keys() if results[k]])
        success_rate = total_processed / len(filtered_keys) if filtered_keys else 0

        summary_msg = f"✅ 校对完成 - 成功处理 {total_processed}/{len(filtered_keys)} 个项目 ({success_rate:.1%})"

        if is_cli and progress_bar:
            progress_bar.close()
            print(f"\n{summary_msg}")
            print(f"🔍 校对质量统计:")
            for key in results:
                if results[key] and isinstance(results[key], dict):
                    issues = len(results[key].get('issues', []))
                    status = "✅" if issues == 0 else "⚠️" if issues < 3 else "❌"
                    print(f"  {status} {key}: {issues} 个问题")
        elif external_progress and external_progress.get('status'):
            external_progress['status'].text(summary_msg)

        print(f"\n{'─' * 60}")
        print(f"🎯 校对任务完成总结:")
        print(f"   总项目数: {len(filtered_keys)}")
        print(f"   成功校对: {total_processed}")
        print(f"   成功率: {success_rate:.1%}")
        print(f"   耗时: {time.time() - start_time:.1f} 秒")
        print(f"{'─' * 60}")

        return results

    except Exception as e:
        if is_cli:
            print(f"❌ Proofreading error: {e}")
        else:
            try:
                st.error(f"❌ Proofreading error: {e}")
            except Exception:
                print(f"❌ Proofreading error: {e}")
        return agent1_results

def run_ai_proofreader(filtered_keys, agent1_results, ai_data, external_progress=None, language='English'):
    """Run AI Proofreader for all keys (Compliance, Figures, Entities, Grammar)."""
    try:
        # Initialize is_cli flag properly
        is_cli = True  # Default to CLI mode

        # Check if we're in a Streamlit environment (not CLI)
        try:
            import streamlit as st
            # Try to access Streamlit session state - if it succeeds, we're in Streamlit
            _ = st.session_state
            is_cli = False
        except (ImportError, AttributeError):
            # If Streamlit is not available or session state doesn't exist, we're in CLI
            is_cli = True

        # Only show verbose output in CLI mode
        if is_cli:
            print(f"🔍 run_ai_proofreader called with {len(filtered_keys)} keys")

        # Get logger only if available
        logger = None
        if not is_cli:
            try:
                logger = st.session_state.ai_logger
            except:
                pass

        # Model/provider selection - try session state regardless of CLI mode
        use_local_ai = False
        use_openai = False

        try:
            # Try to get settings from session state (works even in CLI mode if Streamlit is imported)
            import streamlit as st
            use_local_ai = st.session_state.get('use_local_ai', False)
            use_openai = st.session_state.get('use_openai', False)

            # Also check selected_provider for more specific control
            selected_provider = st.session_state.get('selected_provider')
            if selected_provider == 'Open AI':
                use_openai = True
                use_local_ai = False
            elif selected_provider == 'Local AI' or selected_provider == 'Server AI':
                use_local_ai = True
                use_openai = False

        except Exception as e:
            # Fallback to config detection
            try:
                from common.assistant import load_config
                config_details = load_config('fdd_utils/config.json')
                if config_details.get('LOCAL_AI_API_BASE'):
                    use_local_ai = True
                elif config_details.get('OPENAI_API_KEY'):
                    use_openai = True
            except Exception:
                pass

        proof_agent = ProofreadingAgent(use_local_ai=use_local_ai, use_openai=use_openai, language=language)

        results = {}
        entity_name = ai_data.get('entity_name', '')
        sections_by_key = ai_data.get('sections_by_key', {})

        # Prepare per-key tables markdown (from processed tables)
        def get_tables_for_key(k):
            key_sections = sections_by_key.get(k, [])
            parts = []
            for section in key_sections:
                if isinstance(section, dict):
                    try:
                        parts.append(json.dumps(section, indent=2, default=str))
                    except Exception:
                        parts.append(str(section))
                elif hasattr(section, 'to_string'):
                    df = section.copy()
                    for col in list(df.columns):
                        if df[col].isna().all() or (df[col].astype(str) == 'None').all():
                            df = df.drop(columns=[col])
                    parts.append(df.to_string(index=False, na_rep=''))
                else:
                    parts.append(str(section))
            return "\n".join(parts)
        
        if is_cli:
            # Use tqdm for command-line progress
            progress_bar = tqdm(total=len(filtered_keys), desc="🤖 AI Proofreader", unit="key")
            status_text = None
        else:
            # Use Streamlit progress
            if external_progress:
                progress_bar = external_progress.get('bar')
                status_text = external_progress.get('status')
            else:
                progress_bar = st.progress(0)
                status_text = st.empty()

        start_time = time.time()
        total = len(filtered_keys)

        for idx, key in enumerate(filtered_keys):
            # Debug output for all modes
            print(f"\n🔍 [{idx+1}/{len(filtered_keys)}] 校对中: {key}")

            if not is_cli:
                elapsed = time.time() - start_time
                # Simple ETA: average time per processed item * remaining
                avg = (elapsed / (idx or 1)) if idx else 0
                remaining = total - idx
                eta_seconds = int(avg * remaining) if idx else 0
                mins, secs = divmod(eta_seconds, 60)
                eta_str = f"ETA {mins:02d}:{secs:02d}" if eta_seconds > 0 else "ETA --:--"
                # Enhanced proofreading progress message
                progress_pct = int((idx + 1) / total * 100)
                enhanced_msg = f"🔍 校对中: {key} ({idx+1}/{total}) {eta_str}"

                # Add compliance status if available from previous results
                if results and key in results:
                    is_compliant = results[key].get('is_compliant', False)
                    status_icon = "✅" if is_compliant else "⚠️"
                    enhanced_msg += f" {status_icon}"

                if status_text:
                    status_text.text(enhanced_msg)
                try:
                    combined = external_progress.get('combined') if external_progress else None
                    if combined and isinstance(combined, dict):
                        stages = combined.get('stages', 2)
                        stage_index = combined.get('stage_index', 1)
                        stage_weight = 1.0 / max(1, stages)
                        progress = (idx+1)/len(filtered_keys)
                        combined_progress = min(1.0, max(0.0, stage_index * stage_weight + progress * stage_weight))
                        progress_bar.progress(combined_progress)
                    else:
                        progress_bar.progress((idx+1)/len(filtered_keys))
                except Exception:
                    pass
            elif is_cli:
                # Update tqdm progress bar
                progress_bar.set_description(f"🔍 校对 {key} ({idx+1}/{len(filtered_keys)})")
                progress_bar.update(1)
            
            try:
                content = agent1_results.get(key, '')
                if isinstance(content, dict):
                    content_text = content.get('content', '')
                    tables_md = content.get('table_data', '') or get_tables_for_key(key)
                else:
                    content_text = str(content)
                    tables_md = get_tables_for_key(key)

                if not content_text:
                    results[key] = {'is_compliant': False, 'issues': ["No Agent 1 content"], 'corrected_content': ''}
                    if is_cli:
                        progress_bar.update(1)
                    continue

                # DEBUG: Print proofreading input
                print(f"\n{'='*80}")

                print(f"{'='*80}")
                print(f"📝 CONTENT TO PROOFREAD ({len(content_text)} chars):")
                print(f"{content_text}")
                print(f"🔧 ENTITY NAME: {entity_name}")
                print(f"{'='*80}")

                # Pass progress_bar to proofread method
                result = proof_agent.proofread(content_text, key, tables_md, entity_name, progress_bar if is_cli else None)

                # DEBUG: Print proofreading output
                print(f"\n{'='*80}")

                print(f"{'='*80}")
                if isinstance(result, dict):
                    print(f"📤 PROOFREAD RESULT:")
                    for k, v in result.items():
                        if k == 'corrected_content' and v:
                            print(f"  {k}: {v[:100]}..." if len(str(v)) > 100 else f"  {k}: {v}")
                        else:
                            print(f"  {k}: {v}")
                else:
                    print(f"📤 PROOFREAD RESULT: {result}")
                print(f"{'='*80}")

                # Debug output for proofreading result
                print(f"✅ 校对完成: {key}")
                if isinstance(result, dict):
                    corrected_content = result.get('corrected_content', '')
                    issues_found = len(result.get('issues', []))
                    is_compliant = result.get('is_compliant', False)
                    print(f"🔧 校对结果: {issues_found} 个问题发现, 合规性: {'✅' if is_compliant else '⚠️'}")
                    if corrected_content:
                        print(f"📝 修正内容预览: {corrected_content[:50]}..." if len(corrected_content) > 50 else f"📝 修正内容: {corrected_content}")

                results[key] = result

                # Log output only if logger is available
                if logger:
                    try:
                        logger.log_agent_output('agent3', key, result, 0)
                    except Exception:
                        pass

                # Update session store with corrected content
                if not is_cli:
                    try:
                        content_store = st.session_state.get('ai_content_store', {})
                        if key not in content_store:
                            content_store[key] = {}
                        corrected = result.get('corrected_content') or content_text
                        content_store[key]['agent3_content'] = corrected
                        content_store[key]['current_content'] = corrected
                        st.session_state['ai_content_store'] = content_store
                    except Exception:
                        pass

            except Exception as e:
                results[key] = {'is_compliant': False, 'issues': [str(e)], 'corrected_content': ''}
                if is_cli and progress_bar:
                    progress_bar.update(1)

        # Final summary and close progress bar
        total_processed = len([k for k in results.keys() if results[k]])
        success_rate = total_processed / len(filtered_keys) if filtered_keys else 0

        summary_msg = f"✅ 校对完成 - 成功处理 {total_processed}/{len(filtered_keys)} 个项目 ({success_rate:.1%})"

        if is_cli and progress_bar:
            progress_bar.close()
            print(f"\n{summary_msg}")
            print(f"🔍 校对质量统计:")
            for key in results:
                if results[key] and isinstance(results[key], dict):
                    issues = len(results[key].get('issues', []))
                    is_compliant = results[key].get('is_compliant', False)
                    status = "✅" if issues == 0 and is_compliant else "⚠️" if issues < 3 else "❌"
                    print(f"  {status} {key}: {issues} 个问题, 合规性: {'是' if is_compliant else '否'}")
        elif not is_cli:
            try:
                st.success(summary_msg)
            except Exception:
                pass

        print(f"\n{'─' * 60}")
        print(f"🎯 校对任务完成总结:")
        print(f"   总项目数: {len(filtered_keys)}")
        print(f"   成功校对: {total_processed}")
        print(f"   成功率: {success_rate:.1%}")
        print(f"   耗时: {time.time() - start_time:.1f} 秒")
        print(f"{'─' * 60}")

        return results
    except Exception as e:
        if is_cli:
            print(f"❌ AI Proofreader Error: {e}")
        else:
            try:
                st.error(f"❌ AI Proofreader Error: {e}")
            except Exception:
                print(f"❌ AI Proofreader Error: {e}")
        return {}

def update_bs_content_with_agent_corrections(corrections_dict, entity_name, agent_name):
    """Update bs_content.md with corrections from Agent 2 or Agent 3"""
    try:
        import re
        
        # Read current bs_content.md
        bs_content_path = 'fdd_utils/bs_content.md'
        if not os.path.exists(bs_content_path):
            st.warning(f"bs_content.md not found at {bs_content_path}")
            return False
        
        with open(bs_content_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Define category mappings based on entity name
        if entity_name in ['Ningbo', 'Nanjing']:
            from fdd_utils.mappings import DISPLAY_NAME_MAPPING_NB_NJ as name_mapping
        else:  # Haining and others
            from fdd_utils.mappings import DISPLAY_NAME_MAPPING_DEFAULT as name_mapping
        
        # Update content for each corrected key
        for key, corrected_content in corrections_dict.items():
            if key in name_mapping:
                section_name = name_mapping[key]
                # Find and replace the section content
                pattern = rf'(### {re.escape(section_name)}\n)(.*?)(?=\n### |\Z)'
                
                # Clean the corrected content
                cleaned_content = clean_content_quotes(corrected_content)
                replacement = f'\\1{cleaned_content}\n'
                
                content = re.sub(pattern, replacement, content, flags=re.DOTALL)
        
        # Write updated content back
        with open(bs_content_path, 'w', encoding='utf-8') as f:
            f.write(content)
        
        # Also update the AI generated copy
        try:
            import shutil
            shutil.copy2(bs_content_path, 'fdd_utils/bs_content_ai_generated.md')
        except Exception as e:
            print(f"Could not update AI reference copy: {e}")
        
        return True
        
    except Exception as e:
        st.error(f"Error updating bs_content.md with {agent_name} corrections: {e}")
        return False

# run_agent_2 function removed as requested

def read_bs_content_by_key():
    """Read BS content by key - simplified since AI2 removed"""
    return {}

# Previous run_agent_2 code removed - this section will be cleaned up
# All content below was part of the old run_agent_2 function and needs to be removed

def run_agent_3(filtered_keys, agent1_results, ai_data):
    # Deprecated legacy agent (kept for backward compatibility); no-op.
    return {}

def read_bs_content_by_key(entity_name):
    """Read bs_content.md and return content organized by key"""
    try:
        import re
        
        # Read current bs_content.md
        bs_content_path = 'fdd_utils/bs_content.md'
        if not os.path.exists(bs_content_path):
            return {}
        
        with open(bs_content_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Define category mappings based on entity name
        if entity_name in ['Ningbo', 'Nanjing']:
            name_mapping = {
                'Cash at bank': 'Cash',
                'Accounts receivables': 'AR',
                'Prepayments': 'Prepayments',
                'Other receivables': 'OR',
                'Other current assets': 'Other CA',
                'Investment properties': 'IP',
                'Other non-current assets': 'Other NCA',
                'Accounts payable': 'AP',
                'Taxes payables': 'Taxes payable',
                'Other payables': 'OP',
                'Capital': 'Capital'
            }
        else:  # Haining and others
            name_mapping = {
                'Cash at bank': 'Cash',
                'Accounts receivables': 'AR',
                'Prepayments': 'Prepayments',
                'Other receivables': 'OR',
                'Other current assets': 'Other CA',
                'Investment properties': 'IP',
                'Other non-current assets': 'Other NCA',
                'Accounts payable': 'AP',
                'Taxes payables': 'Taxes payable',
                'Other payables': 'OP',
                'Capital': 'Capital',
                'Surplus reserve': 'Reserve'
            }
        
        # Extract content by section
        content_by_key = {}
        
        # Split by section headers (### Section Name)
        sections = re.split(r'(^### .+$)', content, flags=re.MULTILINE)
        
        for i in range(1, len(sections), 2):
            if i + 1 < len(sections):
                section_header = sections[i].strip()
                section_content = sections[i + 1].strip()
                
                # Extract section name from header
                section_name = section_header.replace('### ', '').strip()
                
                # Map section name to key
                if section_name in name_mapping:
                    key = name_mapping[section_name]
                    content_by_key[key] = section_content
        
        return content_by_key
        
    except Exception as e:
        print(f"Error reading bs_content.md: {e}")
        return {}

def convert_sections_to_markdown(sections_by_key):
    """Convert sections_by_key format to markdown format expected by process_keys"""
    processed_table_data = {}

    for key, sections in sections_by_key.items():
        if not sections:
            continue

        # Combine all markdown content from sections for this key
        markdown_parts = []
        for section in sections:
            if isinstance(section, dict) and 'markdown' in section:
                markdown_parts.append(section['markdown'])
            elif isinstance(section, str):
                markdown_parts.append(section)

        # Join all sections for this key
        if markdown_parts:
            processed_table_data[key] = '\n\n'.join(markdown_parts)

    return processed_table_data

def run_agent_1_simple(filtered_keys, ai_data, external_progress=None, language='English'):
    """Optimized Agent 1 using process_keys directly - eliminates redundant Excel processing"""
    try:
        import time
        print(f"🔍 run_agent_1_simple called with {len(filtered_keys)} keys, language: {language}")
        from common.assistant import process_keys

        # Convert language parameter to the format expected by process_keys
        language_key = 'chinese' if language == '中文' else 'english'

        # Get data from ai_data
        entity_name = ai_data.get('entity_name', '')
        entity_keywords = ai_data.get('entity_keywords', [])
        print(f"🔍 DEBUG: entity_name='{entity_name}', entity_keywords={entity_keywords}")

        # Get AI model settings
        use_local_ai = st.session_state.get('use_local_ai', False)
        use_openai = st.session_state.get('use_openai', False)

        # Create progress callback if provided
        if external_progress:
            def update_progress(progress, message):
                try:
                    bar = external_progress.get('bar')
                    status = external_progress.get('status')
                    if bar:
                        bar.progress(progress)
                    if status:
                        status.text(message)
                except Exception:
                    pass
            progress_callback = update_progress
        else:
            progress_callback = None

        # Try to use already processed data from sections_by_key
        sections_by_key = ai_data.get('sections_by_key', {})
        processed_table_data = None

        if sections_by_key:
            print(f"✅ Using already processed table data from sections_by_key")
            processed_table_data = convert_sections_to_markdown(sections_by_key)
            print(f"📊 Converted {len(processed_table_data)} keys to markdown format")

            # Create a dummy file path since we won't need to read from file
            temp_file_path = "dummy.xlsx"  # process_keys will use processed_table_data instead
        else:
            print(f"⚠️ No processed data found, falling back to Excel processing")
            # Fallback to original Excel processing
            temp_file_path = None
            try:
                if 'uploaded_file_data' in st.session_state:
                    # Use a unique filename to avoid conflicts
                    unique_filename = f"databook_{uuid.uuid4().hex[:8]}.xlsx"
                    temp_file_path = os.path.join(tempfile.gettempdir(), unique_filename)

                    with open(temp_file_path, 'wb') as tmp_file:
                        tmp_file.write(st.session_state['uploaded_file_data'])
                else:
                    # Fallback: use existing databook.xlsx
                    if os.path.exists('databook.xlsx'):
                        temp_file_path = 'databook.xlsx'
                    else:
                        st.error("❌ No databook available for processing")
                        return {}
            except Exception as e:
                st.error(f"❌ Error creating temporary file: {e}")
                return {}

        # Clear JSON cache to ensure fresh prompts are loaded
        from common.assistant import clear_json_cache
        clear_json_cache()

        # Call process_keys directly - it handles Excel processing efficiently
        results = process_keys(
            keys=filtered_keys,
            entity_name=entity_name,
            entity_helpers=entity_keywords,
            input_file=temp_file_path,
            mapping_file="fdd_utils/mapping.json",
            pattern_file="fdd_utils/pattern.json",
            config_file='fdd_utils/config.json',
            prompts_file='fdd_utils/prompts.json',
            use_ai=True,
            progress_callback=progress_callback,
            processed_table_data=processed_table_data,
            use_local_ai=use_local_ai,
            use_openai=use_openai,
            language=language_key
        )

        # Clean up temp file (only if we created one)
        if temp_file_path and temp_file_path != 'databook.xlsx' and temp_file_path != "dummy.xlsx" and os.path.exists(temp_file_path):
            try:
                os.unlink(temp_file_path)
            except Exception:
                pass

        return results

    except Exception as e:
        print(f"Error in run_agent_1_simple: {e}")
        return {}

def run_agent_3(filtered_keys, agent1_results, ai_data):
    """Run Agent 3: Pattern Compliance for all keys"""
    try:

        import json
        import time
        
        logger = st.session_state.ai_logger
        st.markdown("## 🎯 Agent 3: Pattern Compliance")
        st.write(f"Starting Agent 3 for {len(filtered_keys)} keys...")
        
        # Load prompts from prompts.json
        try:
            with open('fdd_utils/prompts.json', 'r', encoding='utf-8') as f:
                prompts_config = json.load(f)
            agent3_system_prompt = prompts_config.get('system_prompts', {}).get('Agent 3', '')
            st.success("✅ Loaded Agent 3 system prompt from prompts.json")
        except (FileNotFoundError, json.JSONDecodeError) as e:
            st.warning(f"⚠️ Could not load prompts.json: {e}")
            agent3_system_prompt = "Fallback Agent 3 system prompt"
        
        # Get AI model settings from session state
        use_local_ai = st.session_state.get('use_local_ai', False)
        use_openai = st.session_state.get('use_openai', False)
        pattern_agent = PatternValidationAgent(use_local_ai=use_local_ai, use_openai=use_openai)
        results = {}
        
        # Load patterns
        patterns = load_ip("fdd_utils/pattern.json")
        st.write(f"Loaded patterns for {len(patterns)} keys")
        
        # Use improved session state storage for fast content access
        content_store = st.session_state.get('ai_content_store', {})
        bs_content_updates = {}
        
        # Check content availability from session state
        available_keys = []
        for key in filtered_keys:
            if key in content_store and 'current_content' in content_store[key]:
                available_keys.append(key)
        
        if not available_keys:
            st.error("❌ No content available in session state for Agent 3")
            st.info("Make sure Agent 1 and Agent 2 have run successfully")
            return {}
        
        # Removed: st.success(f"✅ Found content for {len(available_keys)} keys in session state storage")
        
        for key in available_keys:
            st.write(f"🔄 Checking pattern compliance for {key} with Agent 3...")
            start_time = time.time()
            
            try:
                # Get the most recent content from session state (Agent 2 corrected or Agent 1 original)
                key_data = content_store[key]
                current_content = key_data.get('current_content', key_data.get('agent1_content', ''))
                
                content_source = "Agent 2 corrected" if 'agent2_content' in key_data else "Agent 1 original"
                st.write(f"📊 Content source for {key}: {content_source} (session state)")
                st.write(f"✅ Content length for {key}: {len(current_content)} characters")
                
                if current_content:
                    # Get patterns for this key
                    key_patterns = patterns.get(key, {})
                    if not key_patterns:
                        st.warning(f"⚠️ No patterns found for {key} in pattern.json")
                        st.info(f"Available pattern keys: {list(patterns.keys())}")

                    st.write(f"Found {len(key_patterns)} patterns for {key}")
                    
                    # Prepare detailed user prompt for pattern compliance
                    user_prompt = f"""
                    AI3 PATTERN COMPLIANCE CHECK:
                    
                    CONTENT TO ANALYZE: {current_content}
                    KEY: {key}
                    AVAILABLE PATTERNS FOR {key}: {json.dumps(key_patterns, indent=2)}
                    """
                    
                    # Store the actual prompts that will be sent to AI
                    actual_prompts = {
                        'system_prompt': agent3_system_prompt,
                        'user_prompt': user_prompt,
                        'context': {
                            'patterns_count': len(key_patterns), 
                            'content_length': len(current_content),
                            'key': key,
                            'patterns': key_patterns
                        }
                    }
                    
                    # Log Agent 3 input with actual prompts
                    logger.log_agent_input('agent3', key, agent3_system_prompt, user_prompt, 
                                         {'patterns_count': len(key_patterns), 'content_length': len(current_content)}, actual_prompts)
                    
                    # Use real AI pattern validation instead of fallback
                    pattern_result = pattern_agent.validate_pattern_compliance(
                        content=current_content,
                        key=key
                    )
                    
                    processing_time = time.time() - start_time
                    
                    # Log Agent 3 output
                    logger.log_agent_output('agent3', key, pattern_result, processing_time)
                    
                    # If Agent 3 found issues and provided corrected content, use it
                    if pattern_result.get('corrected_content') and pattern_result.get('corrected_content') != current_content:
                        corrected_content = pattern_result['corrected_content']
                        bs_content_updates[key] = corrected_content
                        
                        # Update session state storage with Agent 3 corrected content
                        content_store[key]['agent3_content'] = corrected_content
                        content_store[key]['agent3_timestamp'] = time.time()
                        content_store[key]['current_content'] = corrected_content  # Latest version
                        
                        pattern_result['content_updated'] = True
                        st.success(f"✅ Agent 3 improved pattern compliance for {key}")
                    else:
                        # Keep existing content as current (no Agent 3 changes)
                        pattern_result['content_updated'] = False
                        st.info(f"ℹ️ Agent 3 found no pattern improvements needed for {key}")
                    
                    results[key] = pattern_result
                    st.success(f"✅ Agent 3 completed {key} in {processing_time:.2f}s")
                    
                else:
                    error_msg = f"No content available for {key}"
                    st.error(f"❌ {error_msg}")
                    results[key] = {
                        "is_compliant": False,
                        "issues": [error_msg],
                        "pattern_match": "none",
                        "suggestions": ["Run Agent 1 first"],
                        "content_updated": False
                    }
                    logger.log_error('agent3', key, error_msg)
            
            except RuntimeError as e:
                # AI-specific errors
                if "AI services are required" in str(e):
                    error_msg = f"AI services required for pattern validation: {e}"
                    logger.log_error('agent3', key, error_msg)
                    st.error(f"❌ **AI Required**: Agent 3 failed for {key}")
                    st.error(f"Error: {e}")
                else:
                    error_msg = f"AI error: {e}"
                    logger.log_error('agent3', key, error_msg)
                    st.error(f"❌ **AI Error**: Agent 3 failed for {key}: {e}")
                results[key] = {
                    "is_compliant": False,
                    "issues": [f"AI processing error: {e}"],
                    "pattern_match": "error",
                    "suggestions": ["Check AI configuration"],
                    "content_updated": False
                }
            except Exception as e:
                logger.log_error('agent3', key, str(e))
                st.error(f"❌ **Unexpected Error**: Agent 3 failed for {key}: {e}")
                results[key] = {
                    "is_compliant": False,
                    "issues": [f"Processing error: {e}"],
                    "pattern_match": "error",
                    "suggestions": ["Check error details"],
                    "content_updated": False
                }
        
        # Update bs_content.md with Agent 3 corrections if any
        if bs_content_updates:
            update_bs_content_with_agent_corrections(bs_content_updates, ai_data.get('entity_name', ''), "Agent 3")
            # Removed: st.success(f"✅ Agent 3 updated bs_content.md with pattern compliance fixes for {len(bs_content_updates)} keys")
        else:
            st.info("ℹ️ Agent 3 found no pattern compliance improvements needed")
        
        st.success(f"🎉 Agent 3 completed all {len(filtered_keys)} keys")
        return results
        
    except RuntimeError as e:
        # AI-specific errors
        if "AI services are required" in str(e):
            st.error("❌ **AI Services Required for Agent 3**")
            st.error("Agent 3 (Pattern Compliance) requires AI services to function.")
            st.error(f"Error: {e}")
        else:
            st.error(f"❌ **Agent 3 AI Error**: {e}")
        logger = st.session_state.ai_logger
        logger.log_error('agent3', 'general', str(e))
        return {}
    except Exception as e:
        st.error(f"❌ **Agent 3 Unexpected Error**: {e}")
        logger = st.session_state.ai_logger
        logger.log_error('agent3', 'general', str(e))
        return {}

def display_sequential_agent_results(key, filtered_keys, ai_data):
    """Display consolidated AI results in organized tabs with parallel comparison (ENHANCED INTERFACE)"""
    # Single consolidated AI results area with tabs
    st.markdown("## 🤖 AI Processing Results")
    
    # Create main tabs for different views
    main_tabs = st.tabs(["📊 By Agent", "🗂️ By Key", "🔄 Parallel Comparison", "📈 Session Overview"])
    
    # Tab 1: Results organized by Agent (AI1, AI2, AI3)
    with main_tabs[0]:
        st.markdown("### View results organized by AI Agent")
        
        # Agent tabs
        agent_tabs = st.tabs(["🚀 Agent 1: Generation", "📊 Agent 2: Validation"])
        
        # Agent 1 Tab
        with agent_tabs[0]:
            st.markdown("**Focus**: Generate comprehensive financial analysis content")
            
            # Show Agent 1 results for all keys
            agent_states = st.session_state.get('agent_states', {})
            if agent_states.get('agent1_completed', False):
                agent1_results = agent_states.get('agent1_results', {}) or {}
                
                # Key tabs within Agent 1
                if agent1_results:
                    available_keys = [k for k in filtered_keys if k in agent1_results and agent1_results[k]]
                    if available_keys:
                        # Create tab labels with content-aware display names
                        tab_labels = []
                        for k in available_keys:
                            content = agent1_results[k]
                            if isinstance(content, dict):
                                content_str = content.get('content', str(content))
                            else:
                                content_str = str(content)
                            # Use content-aware display name that shows Excel tab names for Chinese content
                            display_name = get_key_display_name(k, content=content_str)
                            tab_labels.append(display_name)

                        key_tabs = st.tabs(tab_labels)
                        
                        for i, key in enumerate(available_keys):
                            with key_tabs[i]:
                                content = agent1_results[key]
                                if content:
                                    # Handle both string and dict content
                                    if isinstance(content, dict):
                                        content_str = content.get('content', str(content))
                                    else:
                                        content_str = str(content)
                                    
                                    # Show metadata
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("Characters", len(content_str))
                                    with col2:
                                        st.metric("Words", len(content_str.split()))
                                    with col3:
                                        st.metric("Entity", ai_data.get('entity_name', ''))
                                    
                                    # Show content
                                    st.markdown("**Generated Content:**")
                                    st.markdown(content_str)
                                else:
                                    st.warning("No content generated")
                    else:
                        st.info("No content available from Agent 1")
                else:
                    st.info("Agent 1 results not available")
            else:
                st.info("⏳ Agent 1 will run when you click 'Process with AI'")
        
        # Agent 2 Tab
        with agent_tabs[1]:
            st.markdown("**Focus**: Validate data accuracy and fix issues")
            
            if agent_states.get('agent2_completed', False):
                agent2_results = agent_states.get('agent2_results', {}) or {}
                
                if agent2_results:
                    available_keys = [k for k in filtered_keys if k in agent2_results]
                    if available_keys:
                        # Create tab labels with content-aware display names
                        tab_labels = []
                        for k in available_keys:
                            content = agent2_results[k]
                            if isinstance(content, dict):
                                content_str = content.get('content', str(content))
                            else:
                                content_str = str(content)
                            # Use content-aware display name that shows Excel tab names for Chinese content
                            display_name = get_key_display_name(k, content=content_str)
                            tab_labels.append(display_name)

                        key_tabs = st.tabs(tab_labels)
                        
                        for i, key in enumerate(available_keys):
                            with key_tabs[i]:
                                validation_result = agent2_results[key]
                                
                                # Show validation metrics
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    score = validation_result.get('score', 0)
                                    st.metric("Validation Score", f"{score}%")
                                with col2:
                                    is_valid = validation_result.get('is_valid', False)
                                    st.metric("Status", "✅ Valid" if is_valid else "❌ Issues")
                                with col3:
                                    issues = validation_result.get('issues', [])
                                    st.metric("Issues Found", len(issues))
                                
                                # Show corrected content if available
                                corrected_content = validation_result.get('corrected_content', '')
                                if corrected_content:
                                    st.markdown("**Validated Content:**")
                                    st.markdown(corrected_content)
                                
                                # Show issues if any
                                if issues:
                                    with st.expander("🚨 Issues Found", expanded=False):
                                        for issue in issues:
                                            st.write(f"• {issue}")
                    else:
                        st.info("No validation results available")
                else:
                    st.info("Agent 2 results not available")
            else:
                st.info("⏳ Agent 2 will run after Agent 1 completes")
        

    
    # Tab 2: Results organized by Key (Cash, AR, etc.)
    with main_tabs[1]:
        st.markdown("### View results organized by Financial Key")
        
        # Get final content from session storage (latest versions)
        content_store = st.session_state.get('ai_content_store', {})
        
        if content_store:
            available_keys = [k for k in filtered_keys if k in content_store]
            if available_keys:
                # Create tab labels with content-aware display names
                tab_labels = []
                for k in available_keys:
                    key_data = content_store[k]
                    current_content = key_data.get('current_content', key_data.get('agent1_content', ''))
                    # Use content-aware display name that shows Excel tab names for Chinese content
                    display_name = get_key_display_name(k, content=current_content)
                    tab_labels.append(display_name)

                key_tabs = st.tabs(tab_labels)
                
                for i, key in enumerate(available_keys):
                    with key_tabs[i]:
                        key_data = content_store[key]
                        current_content = key_data.get('current_content', key_data.get('agent1_content', ''))
                        
                        if current_content:
                            # Determine final version source
                            if 'agent3_content' in key_data:
                                content_source = "Agent 3 (Final - Pattern Compliant)"
                                content_icon = "🎯"
                                processing_steps = "Generated → Validated → Pattern Compliant"
                            elif 'agent2_content' in key_data:
                                content_source = "Agent 2 (Validated - Data Accurate)"
                                content_icon = "📊"
                                processing_steps = "Generated → Validated"
                            else:
                                content_source = "Agent 1 (Original - Generated)"
                                content_icon = "📝"
                                processing_steps = "Generated"
                            
                            # Show metadata
                            col1, col2, col3, col4 = st.columns(4)
                            with col1:
                                st.metric("Final Version", content_source.split()[1])
                            with col2:
                                st.metric("Entity", ai_data.get('entity_name', ''))
                            with col3:
                                st.metric("Characters", len(current_content))
                            with col4:
                                st.metric("Words", len(str(current_content).split()))
                            
                            # Show processing pipeline
                            st.info(f"🔄 Processing Pipeline: {processing_steps}")
                            
                            # Show final content
                            st.markdown(f"**{content_icon} Final Content:**")
                            st.markdown(current_content)
                            
                        else:
                            st.warning("No content available for this key")
            else:
                st.info("No processed keys available")
        else:
            st.info("No content available in session storage")
    
    # Tab 3: Parallel Comparison with Before/After Visualization
    with main_tabs[2]:
        st.markdown("### 🔄 Parallel Agent Comparison & Before/After Changes")
        
        # Get agent states and results with forced refresh
        agent_states = st.session_state.get('agent_states', {})
        agent1_results = agent_states.get('agent1_results', {}) or {}
        agent2_results = agent_states.get('agent2_results', {}) or {}
        agent3_results = agent_states.get('agent3_results', {}) or {}
        content_store = st.session_state.get('ai_content_store', {})

        # Force reload agent3_results if we just completed a translation
        if st.session_state.get('force_reload_agent3', False):
            print("🔄 FORCE RELOADING agent3_results from session state")
            agent_states = st.session_state.get('agent_states', {})
            agent3_results = agent_states.get('agent3_results', {}) or {}
            st.session_state['force_reload_agent3'] = False

        # Debug: Show what's available in session state for troubleshooting
        print(f"🔍 UI DEBUG - agent_states keys: {list(agent_states.keys()) if agent_states else 'None'}")
        print(f"🔍 UI DEBUG - agent3_results keys: {list(agent3_results.keys()) if agent3_results else 'None'}")
        print(f"🔍 UI DEBUG - content_store keys: {list(content_store.keys()) if content_store else 'None'}")
        print(f"🔍 UI DEBUG - translation_completed: {st.session_state.get('translation_completed', False)}")
        print(f"🔍 UI DEBUG - last_translation_time: {st.session_state.get('last_translation_time', 0)}")
        print(f"🔍 UI DEBUG - last_ui_refresh_time: {st.session_state.get('last_ui_refresh_time', 0)}")

        # Show translation status and refresh controls
        translation_completed = st.session_state.get('translation_completed', False)

        # Add manual refresh button for immediate UI update
        col_status, col_refresh = st.columns([3, 1])
        with col_status:
            if translation_completed and agent3_results:
                st.success("✅ Chinese translation completed! Translated content is now available.")
            elif agent3_results:
                st.info("ℹ️ Agent 3 results available. Select a key to view translated content.")
            elif translation_completed:
                st.warning("⚠️ Translation completed but no results found. Please check the translation process.")

        with col_refresh:
            if st.button("🔄 Refresh UI", key="refresh_ui_button", help="Force refresh to show latest translated content"):
                print("🔄 MANUAL UI REFRESH TRIGGERED")
                st.session_state['manual_refresh_triggered'] = time.time()
                st.rerun()

        # Check for new translation results and force refresh if needed
        last_translation_time = st.session_state.get('last_translation_time', 0)
        last_ui_refresh_time = st.session_state.get('last_ui_refresh_time', 0)
        manual_refresh_time = st.session_state.get('manual_refresh_triggered', 0)

        # Force refresh conditions:
        # 1. New translation results detected
        # 2. Manual refresh button pressed
        # 3. UI is stale (more than 5 seconds old)
        should_refresh = (
            (agent3_results and last_translation_time > last_ui_refresh_time) or
            (manual_refresh_time > last_ui_refresh_time) or
            (agent3_results and time.time() - last_ui_refresh_time > 5)  # Auto-refresh every 5 seconds if we have results
        )

        if should_refresh:
            refresh_reason = "new translation" if last_translation_time > last_ui_refresh_time else "manual trigger" if manual_refresh_time > last_ui_refresh_time else "stale UI"
            print(f"🔄 UI DEBUG - Refreshing due to: {refresh_reason} (translation: {last_translation_time}, ui: {last_ui_refresh_time})")

            # Show refresh indicator
            with st.spinner(f"🔄 Refreshing UI - {refresh_reason}..."):
                time.sleep(0.5)  # Brief visual feedback

            st.session_state['last_ui_refresh_time'] = time.time()
            if 'refresh_needed' in st.session_state:
                del st.session_state['refresh_needed']
            st.rerun()

        # Clear stale refresh flags
        if 'refresh_needed' in st.session_state and not agent3_results:
            print("🧹 UI DEBUG - Clearing stale refresh flag")
            del st.session_state['refresh_needed']


        # Key selector for comparison (with dynamic key to force refresh)
        if filtered_keys:
            # Use timestamp to force refresh of selectbox when new content arrives
            selectbox_key = f"parallel_comparison_key_{int(st.session_state.get('last_translation_time', 0))}"

            selected_key = st.selectbox(
                "Select Financial Key for Detailed Comparison:",
                filtered_keys,
                format_func=get_key_display_name,
                key=selectbox_key
            )
            
            if selected_key:
                st.markdown(f"### Analysis for {get_key_display_name(selected_key)}")
                
                # Parallel comparison buttons
                st.markdown("#### 🔄 Choose Comparison Mode:")
                comparison_mode = st.radio(
                    "Comparison Type:",
                    ["Before vs After (AI1 → AI3)", "Step-by-Step (AI1 → AI2 → AI3)", "Agent Validation (AI2 vs AI3)"],
                    horizontal=True,
                    key="comparison_mode"
                )
                
                # Get content for selected key
                agent1_content = agent1_results.get(selected_key, "")
                agent2_content = ""
                agent3_content = ""
                
                # Get validated content from Agent 2
                if selected_key in agent2_results:
                    agent2_data = agent2_results[selected_key]
                    agent2_content = agent2_data.get('corrected_content', '')
                    if not agent2_content and selected_key in content_store:
                        agent2_content = content_store[selected_key].get('agent2_content', '')
                
                # Get final content from Agent 3
                if selected_key in agent3_results:
                    agent3_data = agent3_results[selected_key]
                    print(f"🔍 UI DEBUG - Found {selected_key} in agent3_results")
                    print(f"🔍 UI DEBUG - agent3_data keys: {list(agent3_data.keys()) if isinstance(agent3_data, dict) else 'Not dict'}")

                    if isinstance(agent3_data, dict):
                        # PRIORITY ORDER: translated_content (if Chinese) > corrected_content > content
                        translated_content = agent3_data.get('translated_content', '')
                        corrected_content = agent3_data.get('corrected_content', '')
                        content = agent3_data.get('content', '')

                        print(f"🔍 UI DEBUG - translated_content (first 100): '{translated_content[:100]}...'")
                        print(f"🔍 UI DEBUG - corrected_content (first 100): '{corrected_content[:100]}...'")

                        # Check if translated_content contains Chinese and use it if so
                        if (translated_content and
                            any('\u4e00' <= char <= '\u9fff' for char in translated_content)):
                            print("🔍 UI DEBUG - Using translated_content (contains Chinese)")
                            agent3_content = translated_content
                        elif corrected_content:
                            print("🔍 UI DEBUG - Using corrected_content")
                            agent3_content = corrected_content
                        elif content:
                            print("🔍 UI DEBUG - Using content")
                            agent3_content = content
                        else:
                            agent3_content = ''
                            print("🔍 UI DEBUG - No content found in agent3_data")
                    else:
                        agent3_content = str(agent3_data)
                        print(f"🔍 UI DEBUG - Using raw string content: '{agent3_content[:100]}...'")

                    # Also check content_store as backup (it might have more recent data)
                    content_store_content = ''
                    if selected_key in content_store:
                        content_store_content = content_store[selected_key].get('agent3_content', '')
                        print(f"🔍 UI DEBUG - content_store has agent3_content (first 100 chars): '{content_store_content[:100]}...'")

                        # Use content_store content if it's different and contains Chinese
                        if (content_store_content and
                            content_store_content != agent3_content and
                            any('\u4e00' <= char <= '\u9fff' for char in content_store_content)):
                            print("🔍 UI DEBUG - Using content_store content (more recent/Chinese)")
                            agent3_content = content_store_content

                    if not agent3_content and selected_key in content_store:
                        agent3_content = content_store_content
                        print(f"🔍 UI DEBUG - Fallback from content_store (first 100 chars): '{agent3_content[:100]}...'")
                else:
                    print(f"🔍 UI DEBUG - {selected_key} NOT found in agent3_results")
                    # Try content_store as last resort
                    if selected_key in content_store and 'agent3_content' in content_store[selected_key]:
                        agent3_content = content_store[selected_key]['agent3_content']
                        print(f"🔍 UI DEBUG - Using content_store fallback: '{agent3_content[:100]}...'")

                
                # Default to Agent 1 content if later agents don't have content
                if not agent2_content:
                    agent2_content = agent1_content
                if not agent3_content:
                    agent3_content = agent2_content or agent1_content

                # Detect content language for display
                def detect_language(text):
                    if not text:
                        return "Empty"
                    chinese_chars = sum(1 for char in text if '\u4e00' <= char <= '\u9fff')
                    total_chars = len(text.replace(' ', '').replace('\n', ''))
                    if total_chars == 0:
                        return "Empty"
                    chinese_ratio = chinese_chars / total_chars
                    return "🇨🇳 Chinese" if chinese_ratio > 0.3 else "🇺🇸 English"

                # Show content language indicators and last update time
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Agent 1", detect_language(agent1_content))
                with col2:
                    if agent2_content:
                        st.metric("Agent 2", detect_language(agent2_content))
                    else:
                        st.metric("Agent 2", "Not Available")
                with col3:
                    last_update = st.session_state.get('last_translation_time', 0)
                    if last_update > 0:
                        update_time = time.strftime("%H:%M:%S", time.localtime(last_update))
                        st.metric("Agent 3", f"{detect_language(agent3_content)}", delta=f"Updated: {update_time}")
                    else:
                        st.metric("Agent 3", detect_language(agent3_content))

                # Display comparison based on selected mode
                if comparison_mode == "Before vs After (AI1 → AI3)":
                    display_before_after_comparison(selected_key, agent1_content, agent3_content, agent_states)
                    
                elif comparison_mode == "Step-by-Step (AI1 → AI2 → AI3)":
                    display_step_by_step_comparison(selected_key, agent1_content, agent2_content, agent3_content, agent_states)
                    
                elif comparison_mode == "Agent Validation (AI2 vs AI3)":
                    display_validation_comparison(selected_key, agent2_content, agent3_content, agent2_results, agent3_results)
        
        else:
            st.info("No financial keys available for comparison")
    
    # Tab 4: Session Overview
    with main_tabs[3]:
        st.markdown("### Session Processing Overview")
        
        # Session statistics
        if agent_states:
            col1, col2, col3 = st.columns(3)
            with col1:
                agent1_completed = agent_states.get('agent1_completed', False)
                st.metric("Agent 1", "✅ Completed" if agent1_completed else "⏳ Pending")
            with col2:
                agent2_completed = agent_states.get('agent2_completed', False)
                st.metric("Agent 2", "✅ Completed" if agent2_completed else "⏳ Pending")
            with col3:
                agent3_completed = agent_states.get('agent3_completed', False)
                st.metric("Agent 3", "✅ Completed" if agent3_completed else "⏳ Pending")
            
            # Show logging info
            st.markdown("---")
            st.markdown("### 📋 Logging Information")
            
            # Get logger info
            logger = st.session_state.get('ai_logger')
            if logger:
                session_id = getattr(logger, 'session_id', 'unknown')
                log_file = getattr(logger, 'log_file', 'unknown')
                
                st.info(f"📁 **Session ID**: {session_id}")
                st.info(f"📄 **Detailed logs**: `{log_file}`")
                st.info(f"📊 **Consolidated logs**: `logging/session_{session_id}.json`")
                
                # Token usage summary (if available)
                total_logs = len(logger.session_logs) if hasattr(logger, 'session_logs') else 0
                st.metric("Total Log Entries", total_logs)
            else:
                st.warning("No logging information available")
        else:
            st.info("No processing session data available")

# Helper functions for parallel comparison and before/after visualization

def display_before_after_comparison(key, before_content, after_content, agent_states):
    """Display before (AI1) vs after (AI3) comparison with visual diff"""
    st.markdown("#### 📊 Before vs After Comparison")
    
    # Status indicators
    col1, col2, col3 = st.columns(3)
    with col1:
        agent1_success = agent_states.get('agent1_success', False)
        st.metric("Agent 1", "✅ Success" if agent1_success else "❌ Failed")
    with col2:
        changes_made = before_content != after_content
        st.metric("Changes Made", "✅ Yes" if changes_made else "➖ No")
    with col3:
        agent3_success = agent_states.get('agent3_success', False)
        st.metric("Agent 3", "✅ Success" if agent3_success else "❌ Failed")
    
    # Side-by-side comparison
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("##### 📝 **BEFORE** (Agent 1 - Original)")
        if before_content:
            st.markdown(f"**Length:** {len(str(before_content))} characters, {len(str(before_content).split())} words")
            with st.container():
                st.markdown(before_content)
        else:
            st.warning("No original content available")
    
    with col2:
        st.markdown("##### 🎯 **AFTER** (Agent 3 - Final)")
        if after_content:
            st.markdown(f"**Length:** {len(str(after_content))} characters, {len(str(after_content).split())} words")
            with st.container():
                st.markdown(after_content)
        else:
            st.warning("No final content available")
    
    # Change analysis
    if before_content and after_content:
        st.markdown("---")
        st.markdown("#### 📈 Change Analysis")
        
        # Length comparison
        length_diff = len(after_content) - len(before_content)
        word_diff = len(str(after_content).split()) - len(str(before_content).split())
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Character Change", f"{length_diff:+d}", delta=f"{length_diff/len(before_content)*100:+.1f}%" if before_content else "N/A")
        with col2:
            st.metric("Word Change", f"{word_diff:+d}", delta=f"{word_diff/len(str(before_content).split())*100:+.1f}%" if str(before_content).split() else "N/A")
        with col3:
            similarity = calculate_content_similarity(before_content, after_content)
            st.metric("Similarity", f"{similarity:.1f}%")
        
        # Highlight differences
        if before_content != after_content:
            with st.expander("🔍 Detailed Changes", expanded=False):
                show_text_differences(before_content, after_content)

def display_step_by_step_comparison(key, agent1_content, agent2_content, agent3_content, agent_states):
    """Display step-by-step progression through all agents"""
    st.markdown("#### 🔄 Step-by-Step Agent Progression")
    
    # Progress indicators
    col1, col2, col3 = st.columns(3)
    with col1:
        agent1_success = agent_states.get('agent1_success', False)
        st.metric("🚀 Agent 1", "✅ Generated" if agent1_success else "❌ Failed")
    with col2:
        agent2_success = agent_states.get('agent2_success', False)
        st.metric("📊 Agent 2", "✅ Validated" if agent2_success else "❌ Failed")
    with col3:
        agent3_success = agent_states.get('agent3_success', False)
        st.metric("🎯 Agent 3", "✅ Compliant" if agent3_success else "❌ Failed")
    
    # Agent progression tabs
    step_tabs = st.tabs(["🚀 Step 1: Generation", "📊 Step 2: Validation", "🎯 Step 3: Compliance"])
    
    with step_tabs[0]:
        st.markdown("##### Agent 1: Content Generation")
        if agent1_content:
            st.markdown(f"**Length:** {len(agent1_content)} characters")
            st.markdown(agent1_content)
        else:
            st.warning("No content generated by Agent 1")
    
    with step_tabs[1]:
        st.markdown("##### Agent 2: Data Validation & Correction")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Before Validation:**")
            if agent1_content:
                st.markdown(agent1_content[:200] + "..." if len(agent1_content) > 200 else agent1_content)
            else:
                st.warning("No input content")
        
        with col2:
            st.markdown("**After Validation:**")
            if agent2_content:
                if agent2_content != agent1_content:
                    st.success("✅ Changes made during validation")
                    st.markdown(agent2_content[:200] + "..." if len(agent2_content) > 200 else agent2_content)
                else:
                    st.info("ℹ️ No changes needed")
                    st.markdown("Content validated as accurate")
            else:
                st.warning("No validation output")
    
    with step_tabs[2]:
        st.markdown("##### Agent 3: Pattern Compliance & Final Polish")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Before Compliance Check:**")
            if agent2_content:
                st.markdown(agent2_content[:200] + "..." if len(agent2_content) > 200 else agent2_content)
            else:
                st.warning("No input content")
        
        with col2:
            st.markdown("**After Compliance Check:**")
            if agent3_content:
                if agent3_content != agent2_content:
                    st.success("✅ Pattern compliance improvements made")
                    st.markdown(agent3_content[:200] + "..." if len(agent3_content) > 200 else agent3_content)
                else:
                    st.info("ℹ️ Content already compliant")
                    st.markdown("No pattern improvements needed")
            else:
                st.warning("No compliance output")

def display_validation_comparison(key, agent2_content, agent3_content, agent2_results, agent3_results):
    """Display comparison between Agent 2 and Agent 3 results"""
    st.markdown("#### 🔍 Agent Validation Comparison")
    
    # Get validation details
    agent2_data = agent2_results.get(key, {})
    agent3_data = agent3_results.get(key, {})
    
    # Validation metrics comparison
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("##### 📊 Agent 2: Data Validation")
        validation_score = agent2_data.get('score', 0)
        is_valid = agent2_data.get('is_valid', False)
        issues = agent2_data.get('issues', [])
        
        st.metric("Validation Score", f"{validation_score}%")
        st.metric("Status", "✅ Valid" if is_valid else "❌ Issues Found")
        st.metric("Issues Found", len(issues))
        
        if issues:
            with st.expander("🚨 Data Issues", expanded=False):
                for issue in issues:
                    st.write(f"• {issue}")
        
        if agent2_content:
            with st.expander("📝 Agent 2 Content", expanded=False):
                st.markdown(agent2_content)
    
    with col2:
        st.markdown("##### 🎯 Agent 3: Pattern Compliance")
        is_compliant = agent3_data.get('is_compliant', False)
        compliance_issues = agent3_data.get('issues', [])
        pattern_match = agent3_data.get('pattern_match', 'unknown')
        
        st.metric("Compliance Status", "✅ Compliant" if is_compliant else "⚠️ Issues")
        st.metric("Pattern Match", pattern_match.title())
        st.metric("Pattern Issues", len(compliance_issues))
        
        if compliance_issues:
            with st.expander("🚨 Pattern Issues", expanded=False):
                for issue in compliance_issues:
                    st.write(f"• {issue}")
        
        if agent3_content:
            with st.expander("📝 Agent 3 Content", expanded=False):
                st.markdown(agent3_content)
    
    # Overall comparison
    st.markdown("---")
    st.markdown("#### 📈 Overall Quality Comparison")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        data_quality = "High" if is_valid else "Needs Work"
        st.metric("Data Quality", data_quality)
    with col2:
        pattern_quality = "High" if is_compliant else "Needs Work"
        st.metric("Pattern Quality", pattern_quality)
    with col3:
        overall_status = "✅ Ready" if (is_valid and is_compliant) else "⚠️ Needs Review"
        st.metric("Overall Status", overall_status)

def calculate_content_similarity(text1, text2):
    """Calculate similarity percentage between two texts"""
    if not text1 or not text2:
        return 0.0
    
    # Simple word-based similarity
    words1 = set(text1.lower().split())
    words2 = set(text2.lower().split())
    
    if not words1 and not words2:
        return 100.0
    
    intersection = words1.intersection(words2)
    union = words1.union(words2)
    
    return (len(intersection) / len(union)) * 100 if union else 0.0

def show_text_differences(text1, text2):
    """Show differences between two texts (simplified diff)"""
    if text1 == text2:
        st.info("No differences found")
        return
    
    # Split into sentences for comparison
    sentences1 = [s.strip() for s in text1.split('.') if s.strip()]
    sentences2 = [s.strip() for s in text2.split('.') if s.strip()]
    
    st.markdown("**Changes Summary:**")
    
    # Find added/removed sentences
    added = [s for s in sentences2 if s not in sentences1]
    removed = [s for s in sentences1 if s not in sentences2]
    
    if added:
        st.markdown("**✅ Added:**")
        for sentence in added[:3]:  # Show first 3
            st.write(f"+ {sentence}")
        if len(added) > 3:
            st.write(f"... and {len(added) - 3} more additions")
    
    if removed:
        st.markdown("**❌ Removed:**")
        for sentence in removed[:3]:  # Show first 3
            st.write(f"- {sentence}")
        if len(removed) > 3:
            st.write(f"... and {len(removed) - 3} more removals")
    
    if not added and not removed:
        st.info("Changes are mostly within existing sentences (minor edits)")

if __name__ == "__main__":
    main() 