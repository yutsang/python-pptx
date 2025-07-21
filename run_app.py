#!/usr/bin/env python3
"""
Launcher script for the Financial Data Processor Streamlit app
"""

import subprocess
import sys
import os
import socket

def find_available_port(start_port=8501, max_attempts=50):
    """Find an available port starting from start_port"""
    for port in range(start_port, start_port + max_attempts):
        try:
            # Try to bind to the port to check if it's available
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
                sock.bind(('localhost', port))
                return port
        except OSError:
            continue  # Port is in use, try next one
    
    # If no port found in range, fall back to letting Streamlit choose
    return None

def main():
    """Launch the Streamlit app"""
    try:
        # Check if streamlit is installed
        import streamlit
        
        # Find an available port
        available_port = find_available_port()
        
        if available_port:
            print("🚀 Starting Financial Data Processor...")
            print("📊 Streamlit app will open in your browser")
            print(f"📍 URL: http://localhost:{available_port}")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Run streamlit app with the available port
            subprocess.run([
                sys.executable, "-m", "streamlit", "run", "app.py", 
                "--server.port", str(available_port)
            ])
        else:
            print("🚀 Starting Financial Data Processor...")
            print("📊 Streamlit app will open in your browser")
            print("📍 URL: Streamlit will choose an available port")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Let Streamlit choose the port automatically
            subprocess.run([sys.executable, "-m", "streamlit", "run", "app.py"])
        
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