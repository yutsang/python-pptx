#!/usr/bin/env python3
"""
Launcher script for the Financial Data Processor Streamlit app
"""

import subprocess
import sys
import os

def main():
    """Launch the Streamlit app"""
    try:
        # Check if streamlit is installed
        import streamlit
        print("🚀 Starting Financial Data Processor...")
        print("📊 Streamlit app will open in your browser")
        print("📍 URL: http://localhost:8501")
        print("⏹️  Press Ctrl+C to stop the server")
        print("-" * 50)
        
        # Run streamlit app
        subprocess.run([sys.executable, "-m", "streamlit", "run", "app.py", "--server.port", "8501"])
        
    except ImportError:
        print("❌ Streamlit is not installed!")
        print("📦 Please install dependencies first:")
        print("   pip install -r requirements.txt")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n👋 App stopped by user")
    except Exception as e:
        print(f"❌ Error starting app: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 