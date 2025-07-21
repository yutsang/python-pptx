#!/usr/bin/env python3
"""
AI Configuration for Due Diligence Automation

This module handles AI service initialization and configuration.
Place your API keys and AI settings here.
"""

import json
from pathlib import Path

def load_ai_config():
    """Load AI configuration from config.json"""
    try:
        config_path = Path("config/config.json")
        with open(config_path, 'r') as f:
            config = json.load(f)
        return config
    except Exception as e:
        print(f"Error loading AI config: {e}")
        return {}

def initialize_ai_services(config):
    """Initialize AI services based on configuration"""
    try:
        # Check for OpenAI configuration
        if config.get('OPENAI_API_KEY'):
            print("✅ OpenAI API key found")
            # Here you would initialize OpenAI client
            # import openai
            # openai.api_key = config['OPENAI_API_KEY']
            # return openai_client, None
            return "openai_client_placeholder", None
        else:
            print("⚠️ OpenAI API key not configured")
            return None, None
            
    except Exception as e:
        print(f"Error initializing AI services: {e}")
        return None, None

def generate_ai_response(client, system_prompt, user_prompt):
    """Generate AI response using configured service"""
    if client:
        # Placeholder for actual AI API call
        # response = client.chat.completions.create(
        #     model="gpt-4",
        #     messages=[
        #         {"role": "system", "content": system_prompt},
        #         {"role": "user", "content": user_prompt}
        #     ]
        # )
        # return response.choices[0].message.content
        
        return f"[AI Response Placeholder]\nSystem: {system_prompt[:50]}...\nUser: {user_prompt[:50]}...\nResponse: This would be the actual AI-generated content."
    else:
        return f"[Fallback Response]\nNo AI service configured. This is a demo response for the given prompt."

# Configuration template
AI_CONFIG_TEMPLATE = {
    "OPENAI_API_KEY": "your-openai-api-key-here",
    "OPENAI_API_BASE": "https://api.openai.com/v1",
    "CHAT_MODEL": "gpt-4o-mini",
    "TEMPERATURE": 0.3,
    "MAX_TOKENS": 2000,
    
    # Alternative AI providers
    "DEEPSEEK_API_KEY": "your-deepseek-key-here",
    "DEEPSEEK_API_BASE": "https://api.deepseek.com/v1",
    "DEEPSEEK_CHAT_MODEL": "deepseek-chat",
    
    # AI processing settings
    "ENABLE_AI_PROCESSING": True,
    "ENABLE_FALLBACK_MODE": True,
    "AI_TIMEOUT": 30,
    "MAX_RETRIES": 3
}

def create_config_template():
    """Create a template config.json file"""
    config_path = Path("config/config.json")
    
    if not config_path.exists():
        config_path.parent.mkdir(exist_ok=True)
        with open(config_path, 'w') as f:
            json.dump(AI_CONFIG_TEMPLATE, f, indent=2)
        print(f"✅ Created AI config template at {config_path}")
        print("💡 Please add your API keys to config/config.json")
    else:
        print(f"ℹ️ Config file already exists at {config_path}")

if __name__ == "__main__":
    # Create config template if needed
    create_config_template()
    
    # Test configuration
    config = load_ai_config()
    client, _ = initialize_ai_services(config)
    
    if client:
        print("🚀 AI services ready")
    else:
        print("⚠️ AI services in fallback mode") 