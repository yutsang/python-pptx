#!/usr/bin/env python3
"""
AI Configuration for Due Diligence Automation

This module handles AI service initialization and configuration.
Place your API keys and AI settings here.
"""

import json
from pathlib import Path

def load_ai_config():
    """Load AI configuration from config.json with fallbacks"""
    try:
        config = {}
        
        # Try main config/config.json 
        config_path = Path("config/config.json")
        if config_path.exists():
            with open(config_path, 'r') as f:
                config.update(json.load(f))
                print(f"✅ AI config loaded from config/: {len(config)} settings")
        
        # Check environment variables as fallback
        import os
        if not config.get('OPENAI_API_KEY') or config.get('OPENAI_API_KEY') == "":
            env_key = os.getenv('OPENAI_API_KEY')
            if env_key:
                config['OPENAI_API_KEY'] = env_key
                print("✅ OpenAI API key loaded from environment")
        
        # Debug the final config status
        has_openai_key = bool(config.get('OPENAI_API_KEY') and config.get('OPENAI_API_KEY') != "" and config.get('OPENAI_API_KEY') != "your-openai-api-key-here")
        print(f"🔑 API Key Status: {'✅ Available' if has_openai_key else '❌ Not Set'}")
        
        if not has_openai_key:
            print("⚠️ No OpenAI API key found - will use demo mode")
            print("💡 To enable real AI: Set OPENAI_API_KEY in config/config.json or environment variable")
        
        return config
        
    except Exception as e:
        print(f"❌ Error loading config: {e}")
        return {}

def initialize_ai_services(config):
    """Initialize AI services based on configuration"""
    try:
        # Check for OpenAI configuration
        if config.get('OPENAI_API_KEY') and config.get('OPENAI_API_KEY') != "your-openai-api-key-here":
            print("✅ OpenAI API key found")
            try:
                # Try to initialize actual OpenAI client
                import openai
                client = openai.OpenAI(
                    api_key=config['OPENAI_API_KEY'],
                    base_url=config.get('OPENAI_API_BASE', 'https://api.openai.com/v1')
                )
                # Test the connection
                models = client.models.list()
                print(f"✅ OpenAI client initialized successfully")
                return client, None
            except ImportError:
                print("⚠️ OpenAI library not installed, using placeholder mode")
                return "openai_client_placeholder", None
            except Exception as e:
                print(f"⚠️ OpenAI client initialization failed: {e}")
                return "openai_client_placeholder", None
        else:
            print("⚠️ OpenAI API key not configured")
            return None, None
            
    except Exception as e:
        print(f"Error initializing AI services: {e}")
        return None, None

def generate_ai_response(client, system_prompt, user_prompt, model="gpt-4o-mini"):
    """Generate AI response using configured service"""
    if client and client != "openai_client_placeholder":
        try:
            # Actual OpenAI API call
            import openai
            response = client.chat.completions.create(
                model=model,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.3,
                max_tokens=2000
            )
            return response.choices[0].message.content
        except Exception as e:
            return f"[AI Error]\nFailed to generate response: {str(e)}\nFalling back to demo mode."
    elif client == "openai_client_placeholder":
        return f"[AI Ready] OpenAI library detected but no valid API key. Configure API key for processing."
    else:
        return f"[AI Unavailable] No AI service configured. Add API key to config/config.json for AI processing."

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