#!/usr/bin/env python3
"""
Launcher script for the Financial Data Processor Streamlit app
Enhanced with robust port detection and process management
"""

import subprocess
import sys
import os
import socket
import time
import psutil
import random

def kill_streamlit_processes():
    """Kill any existing Streamlit processes to free up ports"""
    try:
        killed_count = 0
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                # Check if process is a Streamlit process
                if proc.info['name'] and 'python' in proc.info['name'].lower():
                    cmdline = proc.info['cmdline']
                    if cmdline and any('streamlit' in str(arg).lower() for arg in cmdline):
                        print(f"🔄 Killing existing Streamlit process (PID: {proc.info['pid']})")
                        proc.kill()
                        killed_count += 1
                        time.sleep(0.5)  # Give time for cleanup
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
        
        if killed_count > 0:
            print(f"✅ Killed {killed_count} existing Streamlit process(es)")
            time.sleep(2)  # Extra time for port cleanup
        return killed_count > 0
    except Exception as e:
        print(f"⚠️ Could not check for existing processes: {e}")
        return False

def is_port_available(port, host='localhost'):
    """Simple but reliable port availability check"""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.settimeout(3)  # Longer timeout for reliability
            result = sock.bind((host, port))
            return True  # Port is available
    except (OSError, socket.error):
        return False  # Port is in use or not available

def find_available_port(start_port=8501, max_attempts=50):
    """Find an available port with simple but reliable detection"""
    print(f"🔍 Searching for available port starting from {start_port}...")
    
    # Check ports sequentially first
    for i in range(max_attempts):
        port = start_port + i
        if is_port_available(port):
            print(f"✅ Found available port: {port}")
            return port
        elif i < 5:  # Only show details for first few attempts
            print(f"❌ Port {port} is in use")
    
    print(f"⚠️ No available ports found in range {start_port}-{start_port + max_attempts}")
    return None

def get_random_port_range():
    """Get a random port range to avoid conflicts"""
    ranges = [
        8601,   # Above standard Streamlit range
        8701,   # Alternative range 1
        9001,   # Alternative range 2
        7501,   # Alternative range 3
    ]
    return random.choice(ranges)

def main():
    """Launch the Streamlit app with enhanced port management"""
    try:
        # Check if streamlit is installed
        import streamlit
        
        print("🚀 Starting Financial Data Processor...")
        print("🔧 Enhanced port detection enabled")
        
        # Kill any existing Streamlit processes (optional, can be disabled)
        try:
            kill_streamlit_processes()
        except Exception as e:
            print(f"⚠️ Could not clean up processes: {e}")
        
        # Try to find an available port
        available_port = None
        
        # Approach 1: Try standard port range
        available_port = find_available_port(8501, 20)
        
        # Approach 2: If no port found, try alternative ranges
        if not available_port:
            print("🔄 Trying alternative port ranges...")
            for _ in range(3):  # Try 3 different ranges
                random_start = get_random_port_range()
                available_port = find_available_port(random_start, 20)
                if available_port:
                    break
        
        # Launch Streamlit
        if available_port:
            print("📊 Streamlit app will open in your browser")
            print(f"📍 URL: http://localhost:{available_port}")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Simple launch with just the port - no extra flags that might cause issues
            subprocess.run([
                sys.executable, "-m", "streamlit", "run", "app.py", 
                "--server.port", str(available_port)
            ])
        else:
            print("🔄 No specific port found, using Streamlit default...")
            print("📊 Streamlit app will open in your browser")
            print("📍 URL: Streamlit will choose an available port")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Default launch - let Streamlit handle everything
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
        print("🔄 Trying simple fallback launch...")
        try:
            # Simplest possible launch
            subprocess.run([sys.executable, "-m", "streamlit", "run", "app.py"])
        except Exception as fallback_error:
            print(f"❌ Fallback also failed: {fallback_error}")
            print("💡 Try using: python run_app_simple.py")
            sys.exit(1)

if __name__ == "__main__":
    main() 