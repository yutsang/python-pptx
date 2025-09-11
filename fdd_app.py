import streamlit as st
import pandas as pd
import json
import warnings
import os
import datetime
import time
from pathlib import Path
import threading
import tempfile
import uuid

# Suppress warnings and bytecode generation
os.environ['PYTHONDONTWRITEBYTECODE'] = '1'
import urllib3
urllib3.disable_warnings()
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=UserWarning)

# Import all required modules
from fdd_utils.mappings import KEY_TO_SECTION_MAPPING, KEY_TERMS_BY_KEY
from fdd_utils.category_config import DISPLAY_NAME_MAPPING_DEFAULT, DISPLAY_NAME_MAPPING_NB_NJ
from fdd_utils.excel_processing import get_worksheet_sections_by_keys
from fdd_utils.data_utils import (
    get_tab_name, get_financial_keys, get_key_display_name,
    format_date_to_dd_mmm_yyyy, load_config_files
)
from fdd_utils.content_utils import (
    clean_content_quotes, get_content_from_json,
    generate_content_from_session_storage
)
from fdd_utils.ui import configure_streamlit_page
from fdd_utils.ui_sections import (
    render_balance_sheet_sections, render_income_statement_sections,
    render_combined_sections
)
from fdd_utils.pptx_export import export_pptx, merge_presentations
from fdd_utils.assistant import (
    process_keys, QualityAssuranceAgent, DataValidationAgent,
    PatternValidationAgent, find_financial_figures_with_context_check,
    get_financial_figure, ProofreadingAgent
)

# Import logging
import logging
logging.getLogger('streamlit.watcher.event_based_path_watcher').setLevel(logging.ERROR)
logging.getLogger('streamlit.watcher.util').setLevel(logging.ERROR)


class MockUploadedFile:
    """Mock uploaded file object for default file handling"""
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
    
    def getvalue(self):
        return self.getbuffer()
    
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


def initialize_session_state():
    """Initialize session state variables"""
    if 'processing_started' not in st.session_state:
        st.session_state['processing_started'] = False
    
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
    
    if 'ai_logger' not in st.session_state:
        from fdd_utils.app_helpers import AIAgentLogger
        st.session_state.ai_logger = AIAgentLogger()


def generate_entity_keywords(entity_input):
    """Generate comprehensive entity keywords from input"""
    if not entity_input:
        return [], None, []
    
    words = entity_input.split()
    entity_keywords = [words[0]] if words else []
    
    # Generate combinations
    if len(words) >= 2:
        for i in range(1, len(words)):
            entity_keywords.append(f"{words[0]} {words[i]}")
    
    # Generate suffixes
    if len(words) > 1:
        entity_suffixes = [s.strip() for s in " ".join(words[1:]).split(',') if s.strip()]
    else:
        entity_suffixes = ["Limited"]
    
    return entity_keywords, entity_input, entity_suffixes


def process_excel_with_timeout(uploaded_file, mapping, selected_entity, entity_suffixes, entity_keywords, entity_mode, timeout=30):
    """Process Excel file with timeout protection"""
    result_container = {}
    exception_container = {}

    def excel_worker():
        try:
            result = get_worksheet_sections_by_keys(
                uploaded_file=uploaded_file,
                tab_name_mapping=mapping,
                entity_name=selected_entity,
                entity_suffixes=entity_suffixes,
                entity_keywords=entity_keywords,
                entity_mode=entity_mode,
                debug=True
            )
            result_container['result'] = result
        except Exception as e:
            exception_container['exception'] = e

    processing_thread = threading.Thread(target=excel_worker)
    processing_thread.daemon = True
    processing_thread.start()
    processing_thread.join(timeout=timeout)

    if processing_thread.is_alive():
        return None, "timeout"
    elif 'exception' in exception_container:
        raise exception_container['exception']
    else:
        return result_container['result'], "success"


def run_ai_processing(filtered_keys, ai_data, language='English'):
    """Run AI processing for content generation"""
    try:
        # Create temporary file for processing
        temp_file_path = None
        uploaded_file_data = st.session_state.get('uploaded_file_data')
        if uploaded_file_data:
            unique_filename = f"databook_{uuid.uuid4().hex[:8]}.xlsx"
            temp_file_path = os.path.join(tempfile.gettempdir(), unique_filename)
            with open(temp_file_path, 'wb') as tmp_file:
                tmp_file.write(uploaded_file_data)
        elif os.path.exists('databook.xlsx'):
            temp_file_path = 'databook.xlsx'
        
        if not temp_file_path:
            st.error("❌ No databook available for processing")
            return {}

        # Get entity information
        entity_name = ai_data.get('entity_name', '')
        entity_keywords = ai_data.get('entity_keywords', [])

        # Process keys using the assistant
        results = process_keys(
            keys=filtered_keys,
            uploaded_file_path=temp_file_path,
            entity_name=entity_name,
            entity_keywords=entity_keywords,
            language=language
        )

        return results or {}
        
    except Exception as e:
        st.error(f"❌ AI processing failed: {e}")
        return {}


def main():
    """Main application function"""
    # Initialize session state FIRST - before any other operations
    initialize_session_state()
    
    # Configure Streamlit
    configure_streamlit_page()
    st.title("🏢 Real Estate DD Report Writer")

    # Navigation description
    if not st.session_state.get('processing_started', False):
        st.info("📋 **Welcome!** Please navigate to the left sidebar to upload your databook and configure input data.")

    # Sidebar controls
    with st.sidebar:
        # File uploader
        uploaded_file = st.file_uploader(
            "Upload Excel File (Optional)",
            type=['xlsx', 'xls'],
            help="Upload your financial data Excel file or use the default databook.xlsx"
        )

        # Use default file if none uploaded
        if uploaded_file is None:
            default_file_path = "databook.xlsx"
            if os.path.exists(default_file_path):
                st.caption(f"Using default file: {default_file_path}")
                uploaded_file = MockUploadedFile(default_file_path)
            else:
                st.error("❌ Default file not found: databook.xlsx")
                st.info("Please upload an Excel file to get started.")
                st.stop()

        # Entity input
        entity_input = st.text_input(
            "Enter Entity Name",
            value="",
            placeholder="e.g., Company Name Limited, Entity Name Corp",
            help="Enter the full entity name to configure processing"
        )
        
        # Clear session state when entity changes
        if st.session_state.get('last_entity_input') != entity_input:
            if 'ai_data' in st.session_state:
                del st.session_state['ai_data']
            if 'filtered_keys_for_ai' in st.session_state:
                del st.session_state['filtered_keys_for_ai']
            if 'processing_started' in st.session_state:
                del st.session_state['processing_started']
        
        st.session_state['last_entity_input'] = entity_input
        
        # Auto-detect entity mode (default to multiple for better detection)
        entity_mode_internal = 'multiple'  # Use multiple mode for better entity detection
        st.session_state['entity_mode'] = entity_mode_internal

        # Generate entity configuration
        entity_keywords, selected_entity, entity_suffixes = generate_entity_keywords(entity_input)
        
        if not selected_entity:
            st.warning("⚠️ Please enter an entity name to start processing")
            st.stop()

        # Show entity info
        if entity_input:
            words = entity_input.split()
            display_name = ' '.join(words[:2]) if len(words) >= 2 else words[0] if words else entity_input
            st.info(f"📋 Entity: {display_name}")

        # Statement type selection
        st.markdown("---")
        statement_type_display = st.radio(
            "Financial Statement Type",
            ["Balance Sheet", "Income Statement", "All"],
            help="Select the type of financial statement to process"
        )
        
        statement_type_mapping = {
            "Balance Sheet": "BS",
            "Income Statement": "IS", 
            "All": "ALL"
        }
        statement_type = statement_type_mapping[statement_type_display]
        
        # Store uploaded file
        st.session_state['uploaded_file'] = uploaded_file
        
        # AI Provider Selection
        ai_mode_options = ["Local AI", "Open AI", "DeepSeek", "Offline"]
        mode_display = st.selectbox(
            "Select Mode", 
            ai_mode_options,
            index=0,
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
                    st.warning("⚠️ OpenAI not configured")
            elif mode_display == "DeepSeek":
                if config.get('DEEPSEEK_API_KEY') and config.get('DEEPSEEK_API_BASE'):
                    st.success("✅ DeepSeek configured")
                    model = config.get('DEEPSEEK_CHAT_MODEL', 'Not configured')
                    st.info(f"🤖 Model: {model}")
                else:
                    st.warning("⚠️ DeepSeek not configured")
            elif mode_display == "Local AI":
                if config.get('LOCAL_AI_API_BASE') and config.get('LOCAL_AI_ENABLED'):
                    st.success("✅ Local AI configured")
                    model = config.get('LOCAL_AI_CHAT_MODEL', 'Not configured')
                    st.info(f"🏠 Model: {model}")
                else:
                    st.warning("⚠️ Local AI not configured")

        # Store mode configuration
        provider_mapping = {
            "Open AI": "Open AI",
            "Local AI": "Local AI",
            "DeepSeek": "DeepSeek"
        }
        st.session_state['selected_mode'] = f"AI Mode - {provider_mapping[mode_display]}"
        st.session_state['ai_model'] = mode_display
        st.session_state['selected_provider'] = provider_mapping[mode_display]
        st.session_state['use_local_ai'] = (mode_display == "Local AI")
        st.session_state['use_openai'] = (mode_display == "Open AI")

    # Main processing area
    if uploaded_file is not None:
        # Start processing button
        if not st.session_state.get('processing_started', False):
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
                st.rerun()

            st.stop()

        # Load configuration
        config, mapping, pattern, prompts = load_config_files()
        
        # Process Excel data
        entity_changed = st.session_state.get('last_processed_entity') != selected_entity
        needs_processing = 'ai_data' not in st.session_state or 'sections_by_key' not in st.session_state.get('ai_data', {}) or entity_changed

        if needs_processing:
            with st.spinner("🔄 Processing Excel file..."):
                print(f"🔄 Processing Excel for {selected_entity}")
                start_time = time.time()

                sections_by_key, status = process_excel_with_timeout(
                    uploaded_file=uploaded_file,
                    mapping=mapping,
                    selected_entity=selected_entity,
                    entity_suffixes=entity_suffixes,
                    entity_keywords=entity_keywords,
                    entity_mode=entity_mode_internal,
                    timeout=30
                )

                if status == "timeout":
                    if st.button("⚠️ Continue Without Excel Data", key="continue_without_excel"):
                        st.warning("⚠️ Continuing without Excel data. Some features may be limited.")
                        sections_by_key = {}
                        st.session_state['excel_processing_skipped'] = True
                    else:
                        st.error("❌ Excel processing timed out. Click 'Continue Without Excel Data' to proceed.")
                        st.stop()

                processing_time = time.time() - start_time
                print(f"✅ Excel processing completed in {processing_time:.2f}s")
                print(f"📊 Found {len(sections_by_key)} financial keys with data")
                
                # Debug: Check what's actually in sections_by_key
                print(f"🔍 DEBUG: sections_by_key keys: {list(sections_by_key.keys())}")
                for key, sections in sections_by_key.items():
                    print(f"🔍 DEBUG: Key '{key}' has {len(sections) if sections else 0} sections")
                    if sections:
                        print(f"🔍 DEBUG: First section keys: {list(sections[0].keys()) if sections[0] else 'Empty section'}")

                # Store processed data
                if 'ai_data' not in st.session_state:
                    st.session_state['ai_data'] = {}
                st.session_state['ai_data']['sections_by_key'] = sections_by_key
                st.session_state['ai_data']['entity_name'] = selected_entity
                st.session_state['ai_data']['entity_keywords'] = entity_keywords
                st.session_state['last_processed_entity'] = selected_entity
        else:
            sections_by_key = st.session_state.get('ai_data', {}).get('sections_by_key', {})

        # Display financial statements
        if statement_type == "BS":
            st.markdown("### Balance Sheet")
            render_balance_sheet_sections(
                sections_by_key, get_key_display_name, selected_entity, format_date_to_dd_mmm_yyyy
            )
        elif statement_type == "IS":
            st.markdown("### Income Statement")
            render_income_statement_sections(
                sections_by_key, get_key_display_name, selected_entity, format_date_to_dd_mmm_yyyy
            )
        elif statement_type == "ALL":
            st.markdown("### Combined Financial Statements")
            render_combined_sections(
                sections_by_key, get_key_display_name, selected_entity, format_date_to_dd_mmm_yyyy
            )

        # AI Processing Section
        st.markdown("---")
        st.markdown("## 🤖 AI Processing & Results")
        
        # Prepare AI data
        keys_with_data = [key for key, sections in sections_by_key.items() if sections]
        
        # Filter keys by statement type
        bs_keys = ["Cash", "AR", "Prepayments", "OR", "Other CA", "Other NCA", "IP", "NCA",
                   "AP", "Taxes payable", "OP", "Capital", "Reserve"]
        is_keys = ["OI", "OC", "Tax and Surcharges", "GA", "Fin Exp", "Cr Loss", "Other Income",
                   "Non-operating Income", "Non-operating Exp", "Income tax", "LT DTA"]
        
        if statement_type == "BS":
            filtered_keys_for_ai = [key for key in keys_with_data if key in bs_keys]
        elif statement_type == "IS":
            filtered_keys_for_ai = [key for key in keys_with_data if key in is_keys]
        else:  # ALL
            filtered_keys_for_ai = keys_with_data

        st.session_state['filtered_keys_for_ai'] = filtered_keys_for_ai
        
        # Store uploaded file data for AI processing
        if hasattr(uploaded_file, 'getbuffer'):
            st.session_state['uploaded_file_data'] = uploaded_file.getbuffer()
        elif hasattr(uploaded_file, 'getvalue'):
            st.session_state['uploaded_file_data'] = uploaded_file.getvalue()
        
        # Prepare AI data
        temp_ai_data = {
            'entity_name': selected_entity,
            'entity_keywords': entity_keywords,
            'sections_by_key': sections_by_key,
            'pattern': pattern,
            'mapping': mapping,
            'config': config
        }
        st.session_state['ai_data'] = temp_ai_data

        # AI Processing Buttons
        st.markdown("### 🤖 AI Report Generation")
        col_eng, col_chi = st.columns(2)

        with col_eng:
            run_eng_clicked = st.button(
                "🇺🇸 Generate English Report",
                type="primary",
                use_container_width=True,
                key="btn_ai_eng",
                help="Generate AI report in English"
            )

        with col_chi:
            run_chi_clicked = st.button(
                "🇨🇳 生成中文报告",
                type="primary",
                use_container_width=True,
                key="btn_ai_chi",
                help="Generate AI report in Chinese"
            )

        # Handle AI processing
        if run_eng_clicked:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                status_text.text("🤖 Generating English content...")
                progress_bar.progress(30)
                
                agent1_results = run_ai_processing(filtered_keys_for_ai, temp_ai_data, language='English')
                
                if agent1_results:
                    # Store results
                    if 'ai_content_store' not in st.session_state:
                        st.session_state['ai_content_store'] = {}
                    
                    for key, result in agent1_results.items():
                        if key not in st.session_state['ai_content_store']:
                            st.session_state['ai_content_store'][key] = {}
                        content = result.get('content', str(result)) if isinstance(result, dict) else str(result)
                        st.session_state['ai_content_store'][key]['agent1_content'] = content
                        st.session_state['ai_content_store'][key]['current_content'] = content
                        st.session_state['ai_content_store'][key]['agent1_timestamp'] = datetime.datetime.now().isoformat()
                    
                    st.session_state['agent_states']['agent1_results'] = agent1_results
                    st.session_state['agent_states']['agent1_completed'] = True
                    st.session_state['agent_states']['agent1_success'] = True
                    
                    # Generate content files
                    status_text.text("📝 Generating content files...")
                    progress_bar.progress(90)
                    generate_content_from_session_storage(selected_entity)
                
                progress_bar.progress(100)
                status_text.text("✅ English AI processing completed")
                time.sleep(1)
                st.rerun()
                
            except Exception as e:
                progress_bar.progress(100)
                status_text.text(f"❌ English AI processing failed: {e}")
                time.sleep(1)
                st.rerun()

        if run_chi_clicked:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                status_text.text("🤖 初始化中文AI处理...")
                progress_bar.progress(10)
                
                # Generate English first, then translate
                status_text.text("📊 生成英文内容...")
                progress_bar.progress(30)
                agent1_results = run_ai_processing(filtered_keys_for_ai, temp_ai_data, language='English')
                
                # Simple translation (placeholder - you can enhance this)
                status_text.text("🌐 翻译为中文...")
                progress_bar.progress(70)
                
                translated_results = {}
                for key, result in agent1_results.items():
                    content = result.get('content', str(result)) if isinstance(result, dict) else str(result)
                    # Simple placeholder translation - replace with actual translation logic
                    translated_results[key] = {
                        'content': f"[中文翻译] {content}",
                        'translated_content': f"[中文翻译] {content}",
                        'is_chinese': True
                    }
                
                # Store results
                if 'ai_content_store' not in st.session_state:
                    st.session_state['ai_content_store'] = {}
                
                for key, result in translated_results.items():
                    if key not in st.session_state['ai_content_store']:
                        st.session_state['ai_content_store'][key] = {}
                    content = result.get('translated_content', result.get('content', ''))
                    st.session_state['ai_content_store'][key]['agent3_content'] = content
                    st.session_state['ai_content_store'][key]['current_content'] = content
                    st.session_state['ai_content_store'][key]['agent3_timestamp'] = time.time()
                
                st.session_state['agent_states']['agent3_results'] = translated_results
                st.session_state['agent_states']['agent3_completed'] = True
                st.session_state['agent_states']['agent3_success'] = True
                
                # Generate content files
                status_text.text("📝 生成内容文件...")
                progress_bar.progress(95)
                generate_content_from_session_storage(selected_entity)
                
                progress_bar.progress(100)
                status_text.text("✅ 中文AI处理完成")
                time.sleep(1)
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ 中文AI处理失败: {e}")
                progress_bar.progress(0)
                status_text.text("❌ 处理失败")

        # Display AI Results
        agent_states = st.session_state.get('agent_states', {})
        any_agent_completed = any([
            agent_states.get('agent1_completed', False),
            agent_states.get('agent2_completed', False),
            agent_states.get('agent3_completed', False)
        ])
        
        if any_agent_completed:
            filtered_keys = st.session_state.get('filtered_keys_for_ai', [])
            
            if filtered_keys:
                # Create tabs for each key
                tab_labels = [get_key_display_name(key) for key in filtered_keys]
                key_tabs = st.tabs(tab_labels)
                
                for i, key in enumerate(filtered_keys):
                    with key_tabs[i]:
                        # Show Agent 3 results first (translated/proofread)
                        agent3_results_all = agent_states.get('agent3_results', {}) or {}
                        if key in agent3_results_all:
                            pr = agent3_results_all[key]
                            translated_content = pr.get('translated_content', '')
                            corrected_content = pr.get('corrected_content', '') or pr.get('content', '')
                            final_content = translated_content if translated_content and pr.get('is_chinese', False) else corrected_content
                            
                            if final_content:
                                st.markdown(final_content)
                        
                        # Show Agent 1 results (collapsible)
                        with st.expander("📝 AI: Generation (details)", expanded=key not in agent3_results_all):
                            agent1_results = agent_states.get('agent1_results', {}) or {}
                            if key in agent1_results and agent1_results[key]:
                                content = agent1_results[key]
                                content_str = content.get('content', str(content)) if isinstance(content, dict) else str(content)
                                
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
            else:
                st.info("No financial keys available for results display.")
        else:
            st.info("No AI agents have run yet. Use the buttons above to start processing.")

        # PowerPoint Export Section
        st.markdown("---")
        st.subheader("📊 PowerPoint Generation")

        col1, col2 = st.columns([1, 1])

        with col1:
            if st.button("📊 Export English PPTX", type="primary", use_container_width=True):
                export_pptx_simple(selected_entity, statement_type, language='english')

        with col2:
            if st.button("📊 Export Chinese PPTX", type="primary", use_container_width=True):
                export_pptx_simple(selected_entity, statement_type, language='chinese')


def export_pptx_simple(selected_entity, statement_type, language='english'):
    """Simple PowerPoint export function"""
    try:
        if language == 'chinese':
            st.info("📊 开始生成中文 PowerPoint 演示文稿...")
        else:
            st.info("📊 Generating English PowerPoint presentation...")

        # Get project name
        words = selected_entity.split() if selected_entity else ['Project']
        project_name = ' '.join(words[:2]) if len(words) >= 2 else words[0] if words else 'Project'

        # Find template
        template_path = None
        for template in ["fdd_utils/template.pptx", "template.pptx"]:
            if os.path.exists(template):
                template_path = template
                break

        if not template_path:
            st.error("❌ PowerPoint template not found")
            return

        # Create output filename
        language_suffix = "_CN" if language == 'chinese' else "_EN"
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"{project_name}_{statement_type.upper()}_{timestamp}{language_suffix}.pptx"
        output_path = f"fdd_utils/output/{output_filename}"

        # Ensure output directory exists
        os.makedirs("fdd_utils/output", exist_ok=True)

        # Get content file
        if statement_type == "IS":
            markdown_path = "fdd_utils/is_content.md"
        elif statement_type == "BS":
            markdown_path = "fdd_utils/bs_content.md"
        else:  # ALL
            st.info("🔄 Generating combined presentation...")
            # For combined, create both BS and IS then merge
            bs_path = "fdd_utils/bs_content.md"
            is_path = "fdd_utils/is_content.md"
            
            if not os.path.exists(bs_path) or not os.path.exists(is_path):
                st.error("❌ Content files not found. Please run AI processing first.")
                return
            
            with tempfile.TemporaryDirectory() as temp_dir:
                bs_temp = os.path.join(temp_dir, "bs_temp.pptx")
                is_temp = os.path.join(temp_dir, "is_temp.pptx")
                
                # Generate BS and IS presentations
                export_pptx(template_path, bs_path, bs_temp, project_name, language=language)
                export_pptx(template_path, is_path, is_temp, project_name, language=language)
                
                # Merge presentations
                merge_presentations(bs_temp, is_temp, output_path)
                
                if language == 'chinese':
                    st.success("✅ 中文组合演示文稿生成成功!")
                else:
                    st.success("✅ Combined presentation generated successfully!")
        
        if statement_type in ["IS", "BS"]:
            if not os.path.exists(markdown_path):
                st.error(f"❌ Content file not found: {markdown_path}")
                st.info("💡 Please run AI processing first.")
                return

            export_pptx(template_path, markdown_path, output_path, project_name, language=language)

        # Show download button
        if os.path.exists(output_path):
            with open(output_path, "rb") as file:
                download_label = f"📥 下载中文 PowerPoint: {output_filename}" if language == 'chinese' else f"📥 Download English PowerPoint: {output_filename}"
                
                st.download_button(
                    label=download_label,
                    data=file.read(),
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )

            success_msg = f"✅ 中文 PowerPoint 生成完成: {output_filename}" if language == 'chinese' else f"✅ English PowerPoint generated successfully: {output_filename}"
            st.success(success_msg)

    except Exception as e:
        st.error(f"❌ Export failed: {e}")


if __name__ == "__main__":
    main()