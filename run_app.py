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

def is_port_in_use(port, host='localhost'):
    """Check if a port is in use with multiple methods"""
    # Method 1: Try to bind to the port
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.bind((host, port))
            return False  # Port is available
    except OSError:
        pass  # Port might be in use, check further
    
    # Method 2: Try to connect to the port
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(1)
            result = sock.connect_ex((host, port))
            return result == 0  # 0 means connection successful, port is in use
    except Exception:
        pass
    
    # Method 3: Check with netstat if available (Unix systems)
    try:
        if sys.platform != 'win32':
            result = subprocess.run(['netstat', '-an'], capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                lines = result.stdout.split('\n')
                for line in lines:
                    if f':{port} ' in line and ('LISTEN' in line or 'ESTABLISHED' in line):
                        return True
    except Exception:
        pass
    
    return False  # Assume available if all checks pass

def find_available_port(start_port=8501, max_attempts=100):
    """Find an available port with enhanced detection"""
    print(f"🔍 Searching for available port starting from {start_port}...")
    
    # Check a wider range of ports
    ports_to_try = list(range(start_port, start_port + max_attempts))
    
    # Add some randomization to avoid conflicts on shared systems
    random.shuffle(ports_to_try[10:])  # Keep first 10 in order, randomize rest
    
    for i, port in enumerate(ports_to_try):
        if not is_port_in_use(port):
            print(f"✅ Found available port: {port} (checked {i+1} ports)")
            return port
        else:
            if i < 10:  # Only show details for first few attempts
                print(f"❌ Port {port} is in use")
    
    print(f"⚠️ No available ports found in range {start_port}-{start_port + max_attempts}")
    return None

def get_random_port_range():
    """Get a random port range to avoid conflicts"""
    # Choose a random starting point in the valid range
    ranges = [
        (8501, 8600),   # Standard Streamlit range
        (8701, 8800),   # Alternative range 1
        (9001, 9100),   # Alternative range 2
        (7501, 7600),   # Alternative range 3
    ]
    
    start_range, end_range = random.choice(ranges)
    start_port = random.randint(start_range, end_range - 50)
    return start_port

def main():
    """Launch the Streamlit app with enhanced port management"""
    try:
        # Check if streamlit is installed
        import streamlit
        
        print("🚀 Starting Financial Data Processor...")
        print("🔧 Enhanced port detection and process management enabled")
        
        # Kill any existing Streamlit processes
        processes_killed = kill_streamlit_processes()
        
        # Try different approaches for port finding
        available_port = None
        
        # Approach 1: Try standard port range
        available_port = find_available_port(8501, 50)
        
        # Approach 2: If no port found, try a random range
        if not available_port:
            print("🔄 Trying alternative port ranges...")
            for _ in range(3):  # Try 3 different random ranges
                random_start = get_random_port_range()
                available_port = find_available_port(random_start, 30)
                if available_port:
                    break
        
        # Approach 3: Let Streamlit choose automatically
        if available_port:
            print("📊 Streamlit app will open in your browser")
            print(f"📍 URL: http://localhost:{available_port}")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Run streamlit app with the available port
            cmd = [
                sys.executable, "-m", "streamlit", "run", "app.py", 
                "--server.port", str(available_port),
                "--server.headless", "true",
                "--server.enableCORS", "false",
                "--server.enableXsrfProtection", "false"
            ]
            
            subprocess.run(cmd)
        else:
            print("🔄 No specific port found, letting Streamlit auto-select...")
            print("📊 Streamlit app will open in your browser")
            print("📍 URL: Streamlit will choose an available port automatically")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Let Streamlit choose the port automatically with enhanced settings
            cmd = [
                sys.executable, "-m", "streamlit", "run", "app.py",
                "--server.headless", "true",
                "--server.enableCORS", "false",
                "--server.enableXsrfProtection", "false"
            ]
            
            subprocess.run(cmd)
        
    except ImportError:
        print("❌ Streamlit is not installed!")
        print("📦 Please install dependencies first:")
        print("   pip install -r requirements.txt")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n👋 App stopped by user")
    except Exception as e:
        print(f"❌ Error starting app: {e}")
        print("🔄 Trying fallback launch method...")
        try:
            # Fallback: basic launch
            subprocess.run([sys.executable, "-m", "streamlit", "run", "app.py"])
        except Exception as fallback_error:
            print(f"❌ Fallback also failed: {fallback_error}")
            sys.exit(1)

if __name__ == "__main__":
    main() 