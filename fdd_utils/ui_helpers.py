"""
UI helper functions for FDD application
"""

import streamlit as st
import json
import os
from datetime import datetime

def display_ai_content_by_key(key, agent_choice, language='English'):
    """Display AI content by key and agent choice"""
    try:
        # Get content from session state
        agent_states = st.session_state.get('agent_states', {})
        agent1_results = agent_states.get('agent1_results', {})
        agent2_results = agent_states.get('agent2_results', {})
        agent3_results = agent_states.get('agent3_results', {})

        # Display content based on agent choice
        if agent_choice == 'Agent 1':
            if key in agent1_results:
                content = agent1_results[key]
                if isinstance(content, dict):
                    display_content = content.get('content', str(content))
                else:
                    display_content = str(content)
                st.text_area(f"Agent 1 - {key}", display_content, height=200, key=f"agent1_{key}")
            else:
                st.warning(f"No Agent 1 content available for {key}")

        elif agent_choice == 'Agent 2':
            if key in agent2_results:
                content = agent2_results[key]
                if isinstance(content, dict):
                    display_content = content.get('content', str(content))
                else:
                    display_content = str(content)
                st.text_area(f"Agent 2 - {key}", display_content, height=200, key=f"agent2_{key}")
            else:
                st.warning(f"No Agent 2 content available for {key}")

        elif agent_choice == 'Agent 3':
            if key in agent3_results:
                content = agent3_results[key]
                if isinstance(content, dict):
                    display_content = content.get('content', str(content))
                else:
                    display_content = str(content)
                st.text_area(f"Agent 3 - {key}", display_content, height=200, key=f"agent3_{key}")
            else:
                st.warning(f"No Agent 3 content available for {key}")

    except Exception as e:
        st.error(f"Error displaying content: {e}")

def display_ai_prompt_by_key(key, agent_choice, language='English'):
    """Display AI prompt by key and agent choice"""
    try:
        # Load prompts from JSON file
        prompts_file = 'fdd_utils/prompts.json'
        if os.path.exists(prompts_file):
            with open(prompts_file, 'r', encoding='utf-8') as f:
                prompts_config = json.load(f)

            language_key = 'chinese' if language == '中文' else 'english'

            if agent_choice == 'Agent 1':
                system_prompts = prompts_config.get('system_prompts', {}).get(language_key, {})
                system_prompt = system_prompts.get('Agent 1', 'No prompt available')
                st.code(system_prompt, language="text")

            elif agent_choice == 'Agent 2':
                system_prompts = prompts_config.get('system_prompts', {}).get(language_key, {})
                system_prompt = system_prompts.get('Agent 2', 'No prompt available')
                st.code(system_prompt, language="text")

            elif agent_choice == 'Agent 3':
                system_prompts = prompts_config.get('system_prompts', {}).get(language_key, {})
                system_prompt = system_prompts.get('Agent 3', 'No prompt available')
                st.code(system_prompt, language="text")

        else:
            st.error("Prompts file not found")

    except Exception as e:
        st.error(f"Error displaying prompt: {e}")

def display_session_summary():
    """Display simple summary of current session"""
    try:
        agent_states = st.session_state.get('agent_states', {})

        st.markdown("### 📊 Session Summary")

        # Agent 1 status
        agent1_results = agent_states.get('agent1_results', {})
        if agent1_results:
            st.success(f"✅ Agent 1: {len(agent1_results)} keys processed")
        else:
            st.info("ℹ️ Agent 1: Not started")

        # Agent 2 status
        agent2_results = agent_states.get('agent2_results', {})
        if agent2_results:
            st.success(f"✅ Agent 2: {len(agent2_results)} keys processed")
        else:
            st.info("ℹ️ Agent 2: Not started")

        # Agent 3 status
        agent3_results = agent_states.get('agent3_results', {})
        if agent3_results:
            st.success(f"✅ Agent 3: {len(agent3_results)} keys processed")
        else:
            st.info("ℹ️ Agent 3: Not started")

        # Show current statement type
        current_statement_type = st.session_state.get('current_statement_type', 'BS')
        st.info(f"📄 Current Statement Type: {current_statement_type}")

        # Show entity name if available
        entity_name = st.session_state.get('entity_input', '')
        if entity_name:
            st.info(f"🏢 Entity: {entity_name}")

    except Exception as e:
        st.error(f"Error displaying session summary: {e}")

def create_download_link(data, filename, text):
    """Create a download link for data"""
    try:
        import base64
        b64 = base64.b64encode(data.encode()).decode()
        href = f'<a href="data:file/txt;base64,{b64}" download="{filename}">{text}</a>'
        st.markdown(href, unsafe_allow_html=True)
    except Exception as e:
        st.error(f"Error creating download link: {e}")

def show_processing_status(message, progress=None):
    """Show processing status with optional progress bar"""
    try:
        status_text = st.empty()
        status_text.text(message)

        if progress is not None:
            progress_bar = st.progress(progress)
            return status_text, progress_bar
        else:
            return status_text, None
    except Exception as e:
        st.error(f"Error showing processing status: {e}")
        return None, None
