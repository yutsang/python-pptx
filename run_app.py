#!/usr/bin/env python3
"""
Launcher script for the Financial Data Processor Streamlit app
Enhanced with robust port detection and conservative process management
"""

import subprocess
import sys
import os
import socket
import time
import psutil
import random

def find_streamlit_processes():
    """Find actual Streamlit processes running this specific app"""
    streamlit_processes = []
    try:
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                cmdline = proc.info['cmdline']
                if cmdline and len(cmdline) > 0:
                    # Look for more specific Streamlit patterns
                    cmdline_str = ' '.join(cmdline).lower()
                    if ('streamlit' in cmdline_str and 
                        ('run' in cmdline_str or 'app.py' in cmdline_str)):
                        streamlit_processes.append({
                            'pid': proc.info['pid'],
                            'cmdline': cmdline_str
                        })
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
    except Exception as e:
        print(f"⚠️ Error checking processes: {e}")
    
    return streamlit_processes

def kill_streamlit_processes(force=False):
    """Kill existing Streamlit processes - only if explicitly requested or if conflicts detected"""
    if not force:
        # Only kill if there are actual port conflicts
        streamlit_processes = find_streamlit_processes()
        if not streamlit_processes:
            print("ℹ️ No Streamlit processes found")
            return False
        
        print(f"🔍 Found {len(streamlit_processes)} Streamlit process(es):")
        for proc in streamlit_processes:
            print(f"   PID {proc['pid']}: {proc['cmdline'][:80]}...")
        
        # Check if any of these processes are using common ports
        ports_in_use = []
        for port in range(8501, 8510):
            if not is_port_available(port):
                ports_in_use.append(port)
        
        if not ports_in_use:
            print("ℹ️ No port conflicts detected, keeping existing processes")
            return False
        
        print(f"⚠️ Ports in use: {ports_in_use}")
        print("🔄 Will clean up conflicting processes...")
    
    try:
        killed_count = 0
        streamlit_processes = find_streamlit_processes()
        
        for proc_info in streamlit_processes:
            try:
                proc = psutil.Process(proc_info['pid'])
                print(f"🔄 Stopping Streamlit process (PID: {proc_info['pid']})")
                proc.terminate()  # Use terminate instead of kill for graceful shutdown
                killed_count += 1
                time.sleep(0.5)  # Give time for cleanup
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        if killed_count > 0:
            print(f"✅ Stopped {killed_count} Streamlit process(es)")
            time.sleep(2)  # Extra time for port cleanup
        
        return killed_count > 0
    except Exception as e:
        print(f"⚠️ Could not clean up processes: {e}")
        return False

def is_port_available(port, host='localhost'):
    """Simple but reliable port availability check"""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.settimeout(2)  # Quick timeout
            result = sock.bind((host, port))
            return True  # Port is available
    except (OSError, socket.error):
        return False  # Port is in use or not available

def find_available_port(start_port=8501, max_attempts=20):
    """Find an available port with simple but reliable detection"""
    print(f"🔍 Searching for available port starting from {start_port}...")
    
    # Check ports sequentially
    for i in range(max_attempts):
        port = start_port + i
        if is_port_available(port):
            print(f"✅ Found available port: {port}")
            return port
        elif i < 3:  # Only show details for first few attempts
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
        print("🔧 Smart port detection enabled")
        
        # Optional: Clean up only if there are actual conflicts
        # You can disable this by setting CLEANUP_PROCESSES=False
        CLEANUP_PROCESSES = True  # Set to False to disable process cleanup
        
        if CLEANUP_PROCESSES:
            try:
                kill_streamlit_processes(force=False)  # Only kill if conflicts detected
            except Exception as e:
                print(f"⚠️ Process cleanup skipped: {e}")
        else:
            print("ℹ️ Process cleanup disabled")
        
        # Try to find an available port
        available_port = None
        
        # Approach 1: Try standard port range
        available_port = find_available_port(8501, 15)
        
        # Approach 2: If no port found, try alternative ranges
        if not available_port:
            print("🔄 Trying alternative port ranges...")
            for _ in range(2):  # Try 2 different ranges
                random_start = get_random_port_range()
                available_port = find_available_port(random_start, 15)
                if available_port:
                    break
        
        # Launch Streamlit
        if available_port:
            print("📊 Streamlit app will open in your browser")
            print(f"📍 URL: http://localhost:{available_port}")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Simple launch with just the port
            subprocess.run([
                sys.executable, "-m", "streamlit", "run", "app.py", 
                "--server.port", str(available_port)
            ])
        else:
            print("🔄 Using Streamlit's automatic port selection...")
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