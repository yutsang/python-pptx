"""
Agent utilities for FDD application
AI agent processing functions for content generation, proofreading, and validation
"""

import os
import json
import uuid
import tempfile
import streamlit as st
import pandas as pd
from datetime import datetime
from fdd_utils.prompt_templates import get_fallback_system_prompt, get_entity_instructions
from fdd_utils.content_utils import clean_content_quotes, generate_markdown_from_ai_results
from fdd_utils.general_utils import log_processing_step, calculate_text_metrics
from fdd_utils.data_utils import load_config_files, get_key_display_name

def initialize_agent_processing(filtered_keys, ai_data, language='English'):
    """
    Initialize common processing elements for all agents
    """
    try:
        # Get data from ai_data
        entity_name = ai_data.get('entity_name', '')
        entity_keywords = ai_data.get('entity_keywords', [])

        # Create temporary file for processing
        temp_file_path = None
        try:
            if 'uploaded_file_data' in st.session_state:
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
                    return None, {}
        except Exception as e:
            st.error(f"❌ Error creating temporary file: {e}")
            return None, {}

        # Load configuration
        config, mapping, pattern, prompts = load_config_files()

        return {
            'temp_file_path': temp_file_path,
            'entity_name': entity_name,
            'entity_keywords': entity_keywords,
            'config': config,
            'mapping': mapping,
            'pattern': pattern,
            'prompts': prompts,
            'language': language,
            'language_key': 'chinese' if language == '中文' else 'english'
        }, {}

    except Exception as e:
        st.error(f"Error initializing agent processing: {e}")
        return None, {}

def load_agent_prompts(language_key='english'):
    """
    Load AI agent prompts from configuration
    """
    try:
        with open('fdd_utils/prompts.json', 'r', encoding='utf-8') as f:
            prompts_config = json.load(f)

        system_prompts = prompts_config.get('system_prompts', {}).get(language_key, {})

        actual_system_prompt = system_prompts.get('Agent 1', '')
        if not actual_system_prompt:
            if language_key == 'chinese':
                actual_system_prompt = prompts_config.get('system_prompts', {}).get('english', {}).get('Agent 1', '')
            actual_system_prompt = get_fallback_system_prompt() if not actual_system_prompt else actual_system_prompt

        # Add entity instructions
        actual_system_prompt += get_entity_instructions().format(entity_name='entity_name')

        return actual_system_prompt

    except Exception as e:
        print(f"Error loading agent prompts: {e}")
        return get_fallback_system_prompt()

def prepare_agent_input_data(filtered_keys, sections_by_key, pattern, entity_name, entity_keywords):
    """
    Prepare input data structure for agent processing
    """
    try:
        # Create input data structure for process_keys
        input_data = {
            'keys': filtered_keys,
            'sections_by_key': sections_by_key,
            'pattern': pattern,
            'entity_name': entity_name,
            'entity_keywords': entity_keywords,
            'temp_file_path': None  # Will be set by caller
        }

        return input_data

    except Exception as e:
        print(f"Error preparing agent input data: {e}")
        return {}

def process_agent_results(results, agent_type, logger=None):
    """
    Process and log agent results
    """
    try:
        if not results:
            return {}

        processed_results = {}
        total_keys = len(results)
        successful_keys = 0

        for key, result in results.items():
            if isinstance(result, dict) and 'content' in result:
                content = result['content']

                # Clean and validate content
                content = clean_content_quotes(content)
                metrics = calculate_text_metrics(content)

                processed_results[key] = {
                    'content': content,
                    'metrics': metrics,
                    'processed_at': datetime.now().isoformat(),
                    'agent_type': agent_type
                }

                successful_keys += 1

                # Log successful processing
                if logger:
                    logger.log_agent_output(agent_type, key, content, 0)

        # Log summary
        log_processing_step(
            f"{agent_type} Processing Complete",
            f"{successful_keys}/{total_keys} keys processed successfully"
        )

        return processed_results

    except Exception as e:
        print(f"Error processing agent results: {e}")
        return {}

def validate_agent_output(content, key, pattern=None):
    """
    Validate agent output content
    """
    try:
        if not content:
            return False, "Content is empty"

        # Basic validation
        if len(str(content)) < 10:
            return False, "Content too short"

        if len(str(content)) > 50000:  # 50KB limit
            return False, "Content too long"

        # Pattern compliance check (if pattern available)
        if pattern and key in pattern:
            # Simple pattern check
            pattern_text = str(pattern[key])
            content_lower = str(content).lower()
            pattern_lower = pattern_text.lower()

            # Check for key terms
            key_terms = ['balance', 'represents', 'total', 'amount']
            matches = sum(1 for term in key_terms if term in content_lower)

            if matches < 2:
                return False, "Content may not comply with required pattern"

        return True, "Content validation passed"

    except Exception as e:
        return False, f"Validation error: {e}"

def create_agent_progress_tracker(total_keys, agent_name):
    """
    Create a progress tracker for agent processing
    """
    try:
        progress_data = {
            'agent_name': agent_name,
            'total_keys': total_keys,
            'processed_keys': 0,
            'successful_keys': 0,
            'failed_keys': 0,
            'start_time': datetime.now(),
            'current_key': None,
            'status': 'initialized'
        }

        return progress_data

    except Exception as e:
        print(f"Error creating progress tracker: {e}")
        return {}

def update_agent_progress(progress_data, key=None, success=None, status=None):
    """
    Update agent processing progress
    """
    try:
        if key:
            progress_data['current_key'] = key
            progress_data['processed_keys'] += 1

        if success is not None:
            if success:
                progress_data['successful_keys'] += 1
            else:
                progress_data['failed_keys'] += 1

        if status:
            progress_data['status'] = status

        return progress_data

    except Exception as e:
        print(f"Error updating progress: {e}")
        return progress_data

def finalize_agent_processing(progress_data):
    """
    Finalize agent processing and return summary
    """
    try:
        end_time = datetime.now()
        duration = (end_time - progress_data['start_time']).total_seconds()

        summary = {
            'agent_name': progress_data['agent_name'],
            'total_keys': progress_data['total_keys'],
            'successful_keys': progress_data['successful_keys'],
            'failed_keys': progress_data['failed_keys'],
            'success_rate': (progress_data['successful_keys'] / progress_data['total_keys'] * 100) if progress_data['total_keys'] > 0 else 0,
            'duration_seconds': duration,
            'completion_time': end_time.isoformat()
        }

        return summary

    except Exception as e:
        print(f"Error finalizing agent processing: {e}")
        return {}

# Placeholder functions for main agent processing
# These would contain the actual agent logic when moved from fdd_app.py

def run_content_generation_agent(filtered_keys, ai_data, external_progress=None, language='English'):
    """
    Run content generation agent (Agent 1)
    Placeholder - actual implementation would be moved from fdd_app.py
    """
    try:
        log_processing_step("Starting Content Generation Agent", f"Processing {len(filtered_keys)} keys")
        # Actual implementation would go here
        return {"status": "placeholder", "keys_processed": len(filtered_keys)}
    except Exception as e:
        print(f"Error in content generation agent: {e}")
        return {}

def run_proofreading_agent(filtered_keys, agent1_results, ai_data, external_progress=None, language='English'):
    """
    Run proofreading agent (Agent 2)
    Placeholder - actual implementation would be moved from fdd_app.py
    """
    try:
        log_processing_step("Starting Proofreading Agent", f"Processing {len(filtered_keys)} keys")
        # Actual implementation would go here
        return {"status": "placeholder", "keys_processed": len(filtered_keys)}
    except Exception as e:
        print(f"Error in proofreading agent: {e}")
        return {}

def run_validation_agent(filtered_keys, agent1_results, ai_data, external_progress=None):
    """
    Run validation agent (Agent 3)
    Placeholder - actual implementation would be moved from fdd_app.py
    """
    try:
        log_processing_step("Starting Validation Agent", f"Processing {len(filtered_keys)} keys")
        # Actual implementation would go here
        return {"status": "placeholder", "keys_processed": len(filtered_keys)}
    except Exception as e:
        print(f"Error in validation agent: {e}")
        return {}

def run_translation_agent(filtered_keys, agent1_results, ai_data, external_progress=None):
    """
    Run Chinese translation agent
    Placeholder - actual implementation would be moved from fdd_app.py
    """
    try:
        log_processing_step("Starting Translation Agent", f"Processing {len(filtered_keys)} keys")
        # Actual implementation would go here
        return {"status": "placeholder", "keys_processed": len(filtered_keys)}
    except Exception as e:
        print(f"Error in translation agent: {e}")
        return {}
