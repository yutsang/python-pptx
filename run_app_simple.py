#!/usr/bin/env python3
"""
Simple backup launcher for the Financial Data Processor Streamlit app
Minimal dependencies, robust port detection
"""

import subprocess
import sys
import socket
import random
import time

def simple_port_check(port, host='localhost'):
    """Simple port availability check"""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(2)
            result = sock.connect_ex((host, port))
            return result != 0  # True if port is available (connection failed)
    except Exception:
        return True  # Assume available if check fails

def find_simple_port():
    """Find an available port using simple method"""
    # Try multiple random ports to avoid conflicts
    port_ranges = [
        range(8501, 8600),
        range(8701, 8800), 
        range(9001, 9100),
        range(7501, 7600)
    ]
    
    for port_range in port_ranges:
        # Try a few random ports from each range
        ports = random.sample(list(port_range), min(10, len(port_range)))
        for port in ports:
            if simple_port_check(port):
                return port
    
    return None

def main():
    """Simple launcher with minimal dependencies"""
    try:
        # Check if streamlit is installed
        import streamlit
        
        print("🚀 Starting Financial Data Processor (Simple Mode)...")
        
        # Find an available port
        available_port = find_simple_port()
        
        if available_port:
            print(f"📍 Using port: {available_port}")
            print("📊 Streamlit app will open in your browser")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Run with found port
            subprocess.run([
                sys.executable, "-m", "streamlit", "run", "app.py",
                "--server.port", str(available_port)
            ])
        else:
            print("🔄 Using automatic port selection...")
            print("📊 Streamlit app will open in your browser")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Let Streamlit choose
            subprocess.run([sys.executable, "-m", "streamlit", "run", "app.py"])
        
    except ImportError:
        print("❌ Streamlit is not installed!")
        print("📦 Please run: pip install -r requirements.txt")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n👋 App stopped by user")
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main() 