"""
AI Assistant functionality moved from common/
Simplified version with essential functions
"""

import json
import os
import httpx
import time
import pandas as pd
from pathlib import Path
import re
from typing import Dict, List, Optional
import logging
import streamlit as st

# Suppress httpx logging
logging.getLogger("httpx").setLevel(logging.WARNING)

# AI-related imports
try:
    from openai import OpenAI
    AI_AVAILABLE = True
except ImportError:
    OpenAI = None
    AI_AVAILABLE = False

def load_config(file_path):
    """Load configuration from a JSON file."""
    try:
        with open(file_path) as config_file:
            config_details = json.load(config_file)
        return config_details
    except Exception as e:
        print(f"Error loading config: {e}")
        return {}

def initialize_ai_services(config_details, use_local=False, use_openai=False):
    """Initialize AI client using config details - supports DeepSeek, OpenAI, and local AI."""
    if not AI_AVAILABLE:
        raise RuntimeError("AI services not available on this machine.")
    
    httpx_client = httpx.Client(verify=False)
    
    if use_openai and config_details.get('OPENAI_API_KEY'):
        return OpenAI(
            api_key=config_details['OPENAI_API_KEY'],
            base_url=config_details.get('OPENAI_API_BASE', 'https://api.openai.com/v1'),
            http_client=httpx_client
        )
    elif use_local and config_details.get('LOCAL_AI_API_BASE'):
        return OpenAI(
            api_key=config_details.get('LOCAL_AI_API_KEY', 'local'),
            base_url=config_details['LOCAL_AI_API_BASE'],
            http_client=httpx_client
        )
    else:
        # Default to DeepSeek
        return OpenAI(
            api_key=config_details.get('DEEPSEEK_API_KEY', ''),
            base_url=config_details.get('DEEPSEEK_API_BASE', 'https://api.deepseek.com'),
            http_client=httpx_client
        )

def get_tab_name(entity_name):
    """Get tab name based on entity name"""
    if 'Haining' in entity_name:
        return "BSHN"
    elif 'Nanjing' in entity_name:
        return "BSNJ"
    elif 'Ningbo' in entity_name:
        return "BSNB"
    else:
        return entity_name

def get_financial_figure(financial_figures, key):
    """Get financial figure for a specific key"""
    return financial_figures.get(key, "Not found")

def find_financial_figures_with_context_check(file_path, tab_name, entity_keywords=None, convert_thousands=False):
    """Find financial figures with context checking"""
    try:
        # Simplified implementation - return empty dict for now
        return {}
    except Exception as e:
        print(f"Error finding financial figures: {e}")
        return {}

def load_ip():
    """Load IP configuration"""
    return "127.0.0.1"

def process_keys(keys, uploaded_file_path, entity_name, entity_keywords, language='English'):
    """Process financial keys and generate AI content"""
    try:
        # Load configuration
        config_details = load_config('fdd_utils/config.json')
        
        # Initialize AI services
        use_local = st.session_state.get('use_local_ai', False)
        use_openai = st.session_state.get('use_openai', False)
        
        client = initialize_ai_services(config_details, use_local, use_openai)
        
        results = {}
        
        for key in keys:
            try:
                # Create a simple prompt for each key
                prompt = f"Analyze the financial data for {key} for entity {entity_name}. Provide a brief analysis in {language}."
                
                # Make AI request
                response = client.chat.completions.create(
                    model=config_details.get('DEEPSEEK_CHAT_MODEL', 'deepseek-chat'),
                    messages=[
                        {"role": "system", "content": "You are a financial analyst."},
                        {"role": "user", "content": prompt}
                    ],
                    max_tokens=500,
                    temperature=0.7
                )
                
                content = response.choices[0].message.content
                results[key] = {"content": content}
                
            except Exception as e:
                print(f"Error processing key {key}: {e}")
                results[key] = {"content": f"Error processing {key}: {str(e)}"}
        
        return results
        
    except Exception as e:
        print(f"Error in process_keys: {e}")
        return {}

class QualityAssuranceAgent:
    """Quality Assurance Agent for content validation"""
    
    def __init__(self, use_local_ai=False, use_openai=False):
        self.use_local_ai = use_local_ai
        self.use_openai = use_openai
    
    def validate_content(self, content, key):
        """Validate content quality"""
        return {"is_valid": True, "issues": []}

class DataValidationAgent:
    """Data Validation Agent"""
    
    def __init__(self, use_local_ai=False, use_openai=False):
        self.use_local_ai = use_local_ai
        self.use_openai = use_openai
    
    def validate_data(self, data):
        """Validate data integrity"""
        return {"is_valid": True, "issues": []}

class PatternValidationAgent:
    """Pattern Validation Agent"""
    
    def __init__(self, use_local_ai=False, use_openai=False):
        self.use_local_ai = use_local_ai
        self.use_openai = use_openai
    
    def validate_patterns(self, content):
        """Validate content patterns"""
        return {"is_valid": True, "issues": []}

class ProofreadingAgent:
    """Proofreading Agent for content correction"""
    
    def __init__(self, use_local_ai=False, use_openai=False, language='English'):
        self.use_local_ai = use_local_ai
        self.use_openai = use_openai
        self.language = language
    
    def proofread_content(self, content, key, language='English'):
        """Proofread and correct content"""
        try:
            # Simple proofreading - in a real implementation, this would use AI
            return {
                "corrected_content": content,
                "issues": [],
                "is_compliant": True
            }
        except Exception as e:
            return {
                "corrected_content": content,
                "issues": [str(e)],
                "is_compliant": False
            }
    
    def proofread(self, content, key, tables_md, entity_name, progress_bar=None):
        """Proofread content with additional context"""
        return self.proofread_content(content, key, self.language)