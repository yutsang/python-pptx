"""Improved Streamlit application with hexagonal architecture."""

import streamlit as st
import asyncio
import tempfile
import os
from datetime import datetime
from pathlib import Path
from typing import Optional

# For now, we'll create mock implementations to make the app runnable
# In a full implementation, these would import from the actual modules

from enum import Enum
from dataclasses import dataclass
from typing import Optional, Dict, Any, List
from datetime import datetime

# Mock domain entities
class StatementType(Enum):
    BALANCE_SHEET = "balance_sheet"
    INCOME_STATEMENT = "income_statement"
    ALL = "all"

class EntityType(Enum):
    REAL_ESTATE = "real_estate"

class ProcessingStatus(Enum):
    PENDING = "pending"
    IN_PROGRESS = "in_progress"
    COMPLETED = "completed"
    FAILED = "failed"

@dataclass
class ProcessFinancialDataRequest:
    entity_name: str
    statement_type: StatementType
    excel_file_data: bytes
    excel_filename: str
    ai_model: str = "gpt-4o-mini"
    validate_data: bool = True
    check_patterns: bool = True
    generate_report: bool = True
    uploaded_by: Optional[str] = None

@dataclass
class GenerateReportRequest:
    entity_name: str
    statement_type: StatementType
    processing_result_id: str
    project_name: Optional[str] = None
    include_summary: bool = True

@dataclass
class MockProcessingResult:
    id: str
    status: ProcessingStatus
    successful_keys: List[str]
    total_keys_processed: int
    total_processing_time: float
    
    def get_agent_result(self, key: str, agent_type: str):
        return None

# Mock settings
class MockSettings:
    def __init__(self):
        self.supported_entities = ["Haining", "Nanjing", "Ningbo"]

def get_settings():
    return MockSettings()

# Mock use cases  
class ProcessFinancialDataUseCase:
    def __init__(self, **kwargs):
        pass
    
    async def execute(self, request: ProcessFinancialDataRequest):
        # Mock implementation for demo
        import time
        await asyncio.sleep(1)  # Simulate processing
        
        return MockProcessingResult(
            id="mock-123",
            status=ProcessingStatus.COMPLETED,
            successful_keys=["Cash", "AR", "AP"],
            total_keys_processed=3,
            total_processing_time=2.5
        )

class GenerateReportUseCase:
    def __init__(self, **kwargs):
        pass
    
    async def execute(self, request: GenerateReportRequest):
        # Mock implementation
        import tempfile
        temp_file = tempfile.NamedTemporaryFile(suffix='.pptx', delete=False)
        temp_file.write(b"Mock PowerPoint content")
        temp_file.close()
        
        @dataclass
        class MockReportResult:
            file_path: str
        
        return MockReportResult(file_path=temp_file.name)


class StreamlitDueDiligenceApp:
    """Main Streamlit application with clean architecture."""
    
    def __init__(self):
        """Initialize the application with dependency injection."""
        self._setup_dependencies()
        self._configure_streamlit()
    
    def _setup_dependencies(self):
        """Set up dependency injection container."""
        # Mock implementations for demo
        self.settings = get_settings()
        
        # Use Cases with mock implementations
        self.process_usecase = ProcessFinancialDataUseCase()
        self.report_usecase = GenerateReportUseCase()
    
    def _configure_streamlit(self):
        """Configure Streamlit page settings."""
        st.set_page_config(
            page_title="Due Diligence Automation",
            page_icon="📊",
            layout="wide",
            initial_sidebar_state="expanded"
        )
    
    def run(self):
        """Main application entry point."""
        self._render_header()
        self._render_sidebar()
        self._render_main_content()
    
    def _render_header(self):
        """Render application header."""
        st.title("📊 Due Diligence Automation")
        st.markdown("**Enterprise-Grade Financial Data Processing**")
        
        # System status indicator
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("System Status", "🟢 Online")
        with col2:
            st.metric("AI Services", "🟢 Available")
        with col3:
            if 'processing_stats' in st.session_state:
                stats = st.session_state.processing_stats
                st.metric("Reports Processed", stats.get('total_processed', 0))
    
    def _render_sidebar(self):
        """Render sidebar with configuration options."""
        with st.sidebar:
            st.header("⚙️ Configuration")
            
            # File upload
            uploaded_file = st.file_uploader(
                "📁 Upload Excel File",
                type=['xlsx', 'xls'],
                help="Upload your financial data Excel file"
            )
            
            # Entity selection
            entity_name = st.selectbox(
                "🏢 Select Entity",
                options=self.settings.supported_entities,
                help="Choose the entity for data processing"
            )
            
            # Statement type
            statement_type_map = {
                "Balance Sheet": StatementType.BALANCE_SHEET,
                "Income Statement": StatementType.INCOME_STATEMENT,
                "All Statements": StatementType.ALL
            }
            
            statement_display = st.radio(
                "📋 Statement Type",
                options=list(statement_type_map.keys()),
                help="Select the type of financial statement to process"
            )
            statement_type = statement_type_map[statement_display]
            
            # AI Model selection
            ai_model = st.selectbox(
                "🤖 AI Model",
                options=["gpt-4o-mini", "deepseek-chat"],
                help="Choose the AI model for processing"
            )
            
            # Store in session state
            st.session_state.update({
                'uploaded_file': uploaded_file,
                'entity_name': entity_name,
                'statement_type': statement_type,
                'ai_model': ai_model
            })
            
            # Processing options
            st.markdown("---")
            st.subheader("🔧 Processing Options")
            
            validate_data = st.checkbox("Validate Data", value=True)
            check_patterns = st.checkbox("Check Patterns", value=True)
            generate_report = st.checkbox("Generate PowerPoint", value=True)
            
            st.session_state.update({
                'validate_data': validate_data,
                'check_patterns': check_patterns,
                'generate_report': generate_report
            })
    
    def _render_main_content(self):
        """Render main content area."""
        uploaded_file = st.session_state.get('uploaded_file')
        
        if not uploaded_file:
            self._render_welcome_screen()
            return
        
        # File uploaded - show processing interface
        self._render_file_info(uploaded_file)
        
        if st.button("🚀 Process Financial Data", type="primary", use_container_width=True):
            self._process_financial_data(uploaded_file)
        
        # Show results if available
        if 'processing_result' in st.session_state:
            self._render_results()
    
    def _render_welcome_screen(self):
        """Render welcome screen when no file is uploaded."""
        st.markdown("## 👋 Welcome to Due Diligence Automation")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            ### 🎯 **Features**
            - **🤖 AI-Powered Processing**: 3-agent validation pipeline
            - **📊 Smart Pattern Matching**: Intelligent content generation
            - **🔍 Data Validation**: Comprehensive accuracy checks
            - **📋 PowerPoint Export**: Professional report generation
            - **⚡ Enterprise Performance**: Scalable and reliable
            """)
        
        with col2:
            st.markdown("""
            ### 🏢 **Supported Entities**
            - Haining Real Estate
            - Nanjing Properties  
            - Ningbo Developments
            
            ### 📈 **Statement Types**
            - Balance Sheet Analysis
            - Income Statement Review
            - Comprehensive Reports
            """)
        
        st.info("👆 Upload an Excel file in the sidebar to get started")
    
    def _render_file_info(self, uploaded_file):
        """Render information about uploaded file."""
        st.markdown("## 📁 File Information")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Filename", uploaded_file.name)
        with col2:
            file_size = len(uploaded_file.getbuffer()) / 1024  # KB
            st.metric("File Size", f"{file_size:.1f} KB")
        with col3:
            st.metric("File Type", uploaded_file.type)
    
    def _process_financial_data(self, uploaded_file):
        """Process financial data using the use case."""
        try:
            # Create processing request
            request = ProcessFinancialDataRequest(
                entity_name=st.session_state['entity_name'],
                statement_type=st.session_state['statement_type'],
                excel_file_data=uploaded_file.getbuffer(),
                excel_filename=uploaded_file.name,
                ai_model=st.session_state['ai_model'],
                validate_data=st.session_state['validate_data'],
                check_patterns=st.session_state['check_patterns'],
                generate_report=st.session_state['generate_report'],
                uploaded_by="streamlit_user"
            )
            
            # Show processing status
            with st.spinner("🔄 Processing financial data..."):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Execute use case
                result = asyncio.run(
                    self._process_with_progress(request, progress_bar, status_text)
                )
                
                progress_bar.progress(100)
                status_text.success("✅ Processing completed successfully!")
                
                # Store result in session state
                st.session_state['processing_result'] = result
                st.session_state['last_processed'] = datetime.now()
                
                # Update processing stats
                if 'processing_stats' not in st.session_state:
                    st.session_state['processing_stats'] = {'total_processed': 0}
                st.session_state['processing_stats']['total_processed'] += 1
                
                st.rerun()
        
        except Exception as e:
            st.error(f"❌ Processing failed: {str(e)}")
            st.exception(e)
    
    async def _process_with_progress(self, request, progress_bar, status_text):
        """Process with progress updates."""
        # This would use the actual use case implementation
        status_text.text("🔄 Initializing processing...")
        progress_bar.progress(10)
        
        status_text.text("📊 Parsing Excel data...")
        progress_bar.progress(30)
        
        status_text.text("🤖 Running AI Agent 1: Content Generation...")
        progress_bar.progress(50)
        
        status_text.text("🔍 Running AI Agent 2: Data Validation...")
        progress_bar.progress(70)
        
        status_text.text("🎯 Running AI Agent 3: Pattern Compliance...")
        progress_bar.progress(90)
        
        # Execute the actual use case
        result = await self.process_usecase.execute(request)
        return result
    
    def _render_results(self):
        """Render processing results."""
        result = st.session_state['processing_result']
        
        st.markdown("---")
        st.markdown("## 📊 Processing Results")
        
        # Summary metrics
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Status", "🟢 Completed" if result.status == ProcessingStatus.COMPLETED else "🔴 Failed")
        with col2:
            st.metric("Keys Processed", len(result.successful_keys))
        with col3:
            st.metric("Processing Time", f"{result.total_processing_time:.2f}s")
        with col4:
            success_rate = len(result.successful_keys) / result.total_keys_processed * 100 if result.total_keys_processed > 0 else 0
            st.metric("Success Rate", f"{success_rate:.1f}%")
        
        # Agent results tabs
        tab1, tab2, tab3 = st.tabs(["📝 Content Generation", "🔍 Data Validation", "🎯 Pattern Compliance"])
        
        with tab1:
            self._render_agent_results(result, "content_generation")
        
        with tab2:
            self._render_agent_results(result, "data_validation")
        
        with tab3:
            self._render_agent_results(result, "pattern_compliance")
        
        # PowerPoint generation
        if st.session_state.get('generate_report', False):
            self._render_report_generation()
    
    def _render_agent_results(self, result, agent_type: str):
        """Render results for a specific agent."""
        st.markdown(f"### {agent_type.replace('_', ' ').title()} Results")
        
        for key in result.successful_keys:
            with st.expander(f"📊 {key}", expanded=False):
                agent_result = result.get_agent_result(key, agent_type)
                if agent_result:
                    st.write(f"**Processing Time:** {agent_result.processing_time:.2f}s")
                    st.write(f"**Status:** {agent_result.status.value}")
                    
                    if agent_result.content:
                        st.text_area("Generated Content", agent_result.content, height=200, disabled=True)
                    
                    if agent_result.issues:
                        st.warning("Issues found:")
                        for issue in agent_result.issues:
                            st.write(f"- {issue}")
                else:
                    st.info("No results available for this key")
    
    def _render_report_generation(self):
        """Render PowerPoint report generation section."""
        st.markdown("---")
        st.markdown("## 📋 PowerPoint Report Generation")
        
        col1, col2 = st.columns(2)
        
        with col1:
            project_name = st.text_input(
                "Project Name",
                value=st.session_state.get('entity_name', 'Project'),
                help="Name for the PowerPoint report"
            )
        
        with col2:
            include_summary = st.checkbox("Include Summary", value=True)
        
        if st.button("📊 Generate PowerPoint Report", type="secondary"):
            self._generate_powerpoint_report(project_name, include_summary)
    
    def _generate_powerpoint_report(self, project_name: str, include_summary: bool):
        """Generate PowerPoint report."""
        try:
            processing_result = st.session_state['processing_result']
            
            # Create report generation request
            request = GenerateReportRequest(
                entity_name=st.session_state['entity_name'],
                statement_type=st.session_state['statement_type'],
                processing_result_id=processing_result.id,  # This would be set by the use case
                project_name=project_name,
                include_summary=include_summary
            )
            
            with st.spinner("📊 Generating PowerPoint report..."):
                # Execute report generation use case
                report_result = asyncio.run(self.report_usecase.execute(request))
                
                st.success("✅ PowerPoint report generated successfully!")
                
                # Provide download button
                if os.path.exists(report_result.file_path):
                    with open(report_result.file_path, 'rb') as f:
                        st.download_button(
                            label="📥 Download PowerPoint Report",
                            data=f.read(),
                            file_name=os.path.basename(report_result.file_path),
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                        )
        
        except Exception as e:
            st.error(f"❌ Report generation failed: {str(e)}")


def main():
    """Main application entry point."""
    app = StreamlitDueDiligenceApp()
    app.run()


if __name__ == "__main__":
    main() 