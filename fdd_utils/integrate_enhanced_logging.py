"""
Integration script to optionally enhance the existing AI logging system
Usage: Import this in app.py to add enhanced monitoring capabilities
"""

import streamlit as st
from pathlib import Path
import json
from datetime import datetime

def integrate_enhanced_logging():
    """Integrate enhanced logging with the existing system"""
    try:
        # Only import if the user wants enhanced features
        from fdd_utils.enhanced_logging_config import enhance_existing_logger, create_logging_dashboard_data
        
        # Enhance the existing logger if it exists
        if 'ai_logger' in st.session_state:
            st.session_state.ai_logger = enhance_existing_logger(st.session_state.ai_logger)
            return True
        return False
    except ImportError:
        # Enhanced logging not available, continue with existing system
        return False

def create_logging_dashboard():
    """Create a dashboard for viewing logging information"""
    st.markdown("### 📊 AI Agent Logging Dashboard")
    
    # Check for log files
    log_dir = Path("logging")
    if not log_dir.exists():
        st.info("No logging directory found. Start AI processing to create logs.")
        return
    
    # Get recent log files
    log_files = list(log_dir.glob("*.log")) + list(log_dir.glob("*.json"))
    
    if not log_files:
        st.info("No log files found. Process with AI to generate logs.")
        return
    
    # Display recent session logs
    st.markdown("#### 📁 Recent Log Files")
    
    # Sort by modification time (most recent first)
    sorted_logs = sorted(log_files, key=lambda x: x.stat().st_mtime, reverse=True)
    
    for i, log_file in enumerate(sorted_logs[:5]):  # Show last 5 files
        mod_time = datetime.fromtimestamp(log_file.stat().st_mtime)
        file_size = log_file.stat().st_size
        
        with st.expander(f"📄 {log_file.name} ({file_size:,} bytes) - {mod_time.strftime('%Y-%m-%d %H:%M:%S')}"):
            try:
                if log_file.suffix == '.json':
                    # Try to parse as JSON for structured display
                    with open(log_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        st.json(data)
                else:
                    # Display as text for .log files
                    with open(log_file, 'r', encoding='utf-8') as f:
                        content = f.read()
                        if len(content) > 5000:  # Truncate very long files
                            st.text_area("Log Content (truncated)", content[:5000] + "\n... (truncated)", height=300)
                        else:
                            st.text_area("Log Content", content, height=300)
            except Exception as e:
                st.error(f"Error reading file: {e}")

def display_agent_performance():
    """Display performance metrics for AI agents"""
    if 'ai_logger' not in st.session_state:
        st.info("No AI logger found. Start AI processing first.")
        return
    
    logger = st.session_state.ai_logger
    
    # Show current session summary
    if hasattr(logger, 'session_logs') and logger.session_logs:
        st.markdown("#### 🎯 Current Session Performance")
        
        # Count activities by agent
        agent_counts = {}
        for log in logger.session_logs:
            agent = log.get('agent', 'unknown')
            if agent not in agent_counts:
                agent_counts[agent] = {'inputs': 0, 'outputs': 0, 'errors': 0}
            
            log_type = log.get('type', '').lower()
            if log_type == 'input':
                agent_counts[agent]['inputs'] += 1
            elif log_type == 'output':
                agent_counts[agent]['outputs'] += 1
            elif log_type == 'error':
                agent_counts[agent]['errors'] += 1
        
        # Display metrics in columns
        if agent_counts:
            col1, col2, col3 = st.columns(3)
            
            for i, (agent, counts) in enumerate(agent_counts.items()):
                with [col1, col2, col3][i % 3]:
                    st.metric(
                        f"🤖 {agent.upper()}", 
                        f"✅ {counts['outputs']} completions",
                        f"📝 {counts['inputs']} inputs | ❌ {counts['errors']} errors"
                    )
        
        # Show processing timeline
        st.markdown("#### ⏱️ Processing Timeline")
        
        # Create timeline of events
        timeline_data = []
        for log in logger.session_logs[-20:]:  # Show last 20 events
            timeline_data.append({
                'Time': log.get('timestamp', ''),
                'Agent': log.get('agent', ''),
                'Key': log.get('key', ''),
                'Type': log.get('type', ''),
                'Status': '✅' if log.get('type') == 'OUTPUT' else '📝' if log.get('type') == 'INPUT' else '❌'
            })
        
        if timeline_data:
            import pandas as pd
            df = pd.DataFrame(timeline_data)
            st.dataframe(df, use_container_width=True)
    else:
        st.info("No session data available. Process with AI to see performance metrics.")

def export_session_logs():
    """Export current session logs"""
    if 'ai_logger' not in st.session_state:
        st.warning("No AI logger found.")
        return
    
    logger = st.session_state.ai_logger
    
    st.markdown("#### 💾 Export Session Data")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📄 Export JSON Logs", type="secondary"):
            if hasattr(logger, 'save_logs_to_json'):
                json_file = logger.save_logs_to_json()
                if json_file:
                    st.success(f"Logs exported to: {json_file}")
                    
                    # Provide download
                    try:
                        with open(json_file, 'r', encoding='utf-8') as f:
                            st.download_button(
                                label="📥 Download JSON",
                                data=f.read(),
                                file_name=json_file.name,
                                mime='application/json'
                            )
                    except Exception as e:
                        st.error(f"Download failed: {e}")
                else:
                    st.error("Failed to export logs")
    
    with col2:
        if st.button("📋 Export Text Summary", type="secondary"):
            if hasattr(logger, 'session_logs') and logger.session_logs:
                summary_text = f"""
AI Agent Session Summary
Generated: {datetime.now().isoformat()}

Total Events: {len(logger.session_logs)}

Agent Activity:
"""
                # Count by agent
                agent_counts = {}
                for log in logger.session_logs:
                    agent = log.get('agent', 'unknown')
                    if agent not in agent_counts:
                        agent_counts[agent] = {'inputs': 0, 'outputs': 0, 'errors': 0}
                    
                    log_type = log.get('type', '').lower()
                    if log_type == 'input':
                        agent_counts[agent]['inputs'] += 1
                    elif log_type == 'output':
                        agent_counts[agent]['outputs'] += 1
                    elif log_type == 'error':
                        agent_counts[agent]['errors'] += 1
                
                for agent, counts in agent_counts.items():
                    summary_text += f"\n{agent.upper()}:\n"
                    summary_text += f"  - Inputs: {counts['inputs']}\n"
                    summary_text += f"  - Outputs: {counts['outputs']}\n"
                    summary_text += f"  - Errors: {counts['errors']}\n"
                
                st.download_button(
                    label="📥 Download Summary",
                    data=summary_text,
                    file_name=f"ai_session_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                    mime='text/plain'
                )

def show_logging_help():
    """Show help information about the logging system"""
    st.markdown("### ❓ Logging System Help")
    
    with st.expander("📚 Understanding the Logging System", expanded=False):
        st.markdown("""
        #### 🔍 What Gets Logged
        
        **Agent 1 (Content Generation)**:
        - Input: System prompts, user prompts, worksheet data
        - Output: Generated financial narrative content
        - Processing time and content length
        
        **Agent 2 (Data Validation)**:
        - Input: Agent 1 content, expected financial figures
        - Output: Validation results, data accuracy scores
        - Issues found and corrections made
        
        **Agent 3 (Pattern Compliance)**:
        - Input: Current content, pattern templates
        - Output: Compliance check results, pattern improvements
        - Content modifications and style corrections
        
        #### 📁 Log File Locations
        - **Main logs**: `fdd_utils/logs/ai_agents_YYYYMMDD_HHMMSS.log`
        - **JSON data**: `fdd_utils/logs/ai_agents_YYYYMMDD_HHMMSS.json`
        - **Enhanced logs**: `fdd_utils/logs/enhanced/` (if enabled)
        
        #### 🎯 How to Use Logs
        1. **Debugging**: Check error messages and processing details
        2. **Performance**: Monitor agent processing times
        3. **Quality**: Review input/output for each agent
        4. **Auditing**: Track all AI decision-making processes
        """)
    
    with st.expander("🚀 Tips for Better Logging", expanded=False):
        st.markdown("""
        #### 💡 Best Practices
        
        - **Regular Export**: Save logs after each session
        - **Monitor Performance**: Check processing times for optimization
        - **Review Errors**: Address any recurring agent failures
        - **Quality Check**: Validate AI outputs using logged data
        
        #### 🔧 Troubleshooting
        
        **If logs seem empty**:
        - Ensure AI processing completed successfully
        - Check that all 3 agents ran
        - Verify file permissions in logging directory
        
        **If performance is slow**:
        - Review agent processing times in logs
        - Check for repeated errors in specific agents
        - Consider API rate limits or connection issues
        """)

# Main integration function
def setup_enhanced_logging_ui():
    """Setup the enhanced logging UI components"""
    # Add enhanced logging controls to sidebar or main area
    st.markdown("---")
    st.markdown("## 📊 AI Logging & Monitoring")
    
    # Create tabs for different logging views
    log_tabs = st.tabs([
        "📈 Performance", 
        "📁 Log Files", 
        "💾 Export", 
        "❓ Help"
    ])
    
    with log_tabs[0]:
        display_agent_performance()
    
    with log_tabs[1]:
        create_logging_dashboard()
    
    with log_tabs[2]:
        export_session_logs()
    
    with log_tabs[3]:
        show_logging_help() 