#!/usr/bin/env python3
"""
Launcher script for the Financial Data Processor Streamlit app
Simple and reliable port detection
"""

import subprocess
import sys
import socket
import random

def is_port_available(port, host='localhost'):
    """Check if a port is available"""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(1)
            result = sock.connect_ex((host, port))
            return result != 0  # True if port is available (connection failed)
    except Exception:
        return True  # Assume available if check fails

def find_available_port():
    """Find an available port starting from 8501"""
    # Try ports 8501-8520 first
    for port in range(8501, 8521):
        if is_port_available(port):
            return port
    
    # If all busy, try some random alternatives
    alternative_ranges = [8601, 8701, 9001, 7501]
    for start in alternative_ranges:
        for i in range(10):
            port = start + i
            if is_port_available(port):
                return port
    
    return None

def main():
    """Launch the Streamlit app"""
    try:
        # Check if streamlit is installed
        import streamlit
        
        print("🚀 Starting Financial Data Processor...")
        
        # Find an available port
        available_port = find_available_port()
        
        if available_port:
            print(f"📍 Using port: {available_port}")
            print("📊 Streamlit app will open in your browser")
            print(f"🌐 Local URL: http://localhost:{available_port}")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Simple launch with the available port
            subprocess.run([
                sys.executable, "-m", "streamlit", "run", "app.py",
                "--server.port", str(available_port)
            ])
        else:
            print("🔄 Using Streamlit's automatic port selection...")
            print("📊 Streamlit app will open in your browser")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Let Streamlit handle port selection
            subprocess.run([sys.executable, "-m", "streamlit", "run", "app.py"])
        
    except ImportError:
        print("❌ Streamlit is not installed!")
        print("📦 Please install dependencies: pip install -r requirements.txt")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n👋 App stopped by user")
    except Exception as e:
        print(f"❌ Error starting app: {e}")
        print("💡 Try using: python run_app_simple.py")
        sys.exit(1)

if __name__ == "__main__":
    main() 