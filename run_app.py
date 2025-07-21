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
        # Use basic process info that's compatible across psutil versions
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            try:
                cmdline = proc.info['cmdline']
                if cmdline and len(cmdline) > 0:
                    # Look for more specific Streamlit patterns
                    cmdline_str = ' '.join(cmdline).lower()
                    if ('streamlit' in cmdline_str and 
                        ('run' in cmdline_str or 'app.py' in cmdline_str)):
                        
                        # Try to get port information with compatibility handling
                        used_ports = []
                        try:
                            # Try different methods to get connections based on psutil version
                            process = psutil.Process(proc.info['pid'])
                            
                            # Method 1: Try the net_connections() method (newer psutil)
                            if hasattr(process, 'net_connections'):
                                try:
                                    connections = process.net_connections()
                                    for conn in connections:
                                        if hasattr(conn, 'laddr') and conn.laddr and conn.status == 'LISTEN':
                                            used_ports.append(conn.laddr.port)
                                except (AttributeError, psutil.AccessDenied):
                                    pass
                            
                            # Method 2: Try get_connections() method (older psutil)
                            elif hasattr(process, 'get_connections'):
                                try:
                                    connections = process.get_connections()
                                    for conn in connections:
                                        if hasattr(conn, 'local_address') and conn.local_address and conn.status == 'LISTEN':
                                            used_ports.append(conn.local_address[1])
                                        elif hasattr(conn, 'laddr') and conn.laddr and conn.status == 'LISTEN':
                                            used_ports.append(conn.laddr[1])
                                except (AttributeError, psutil.AccessDenied):
                                    pass
                                    
                        except (psutil.NoSuchProcess, psutil.AccessDenied, AttributeError):
                            # If we can't get connections, still track the process
                            pass
                        
                        streamlit_processes.append({
                            'pid': proc.info['pid'],
                            'cmdline': cmdline_str,
                            'ports': used_ports
                        })
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
    except Exception as e:
        print(f"⚠️ Error checking processes: {e}")
    
    return streamlit_processes

def is_port_available(port, host='localhost'):
    """Enhanced port availability check with multiple methods"""
    # Method 1: Try to bind to the port
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.settimeout(1)
            sock.bind((host, port))
            # If we can bind, the port should be available
            return True
    except (OSError, socket.error):
        # Port is definitely in use
        return False

def check_ports_in_use_fallback():
    """Fallback method to check ports using netstat or system commands"""
    ports_in_use = []
    
    try:
        # Try using netstat command as fallback
        if sys.platform != 'win32':  # Unix-like systems
            result = subprocess.run(['netstat', '-an'], capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                lines = result.stdout.split('\n')
                for line in lines:
                    # Look for listening ports in the Streamlit range
                    for port in range(8501, 8520):
                        if f':{port} ' in line and ('LISTEN' in line or 'LISTENING' in line):
                            ports_in_use.append(port)
        else:  # Windows
            result = subprocess.run(['netstat', '-an'], capture_output=True, text=True, timeout=5, shell=True)
            if result.returncode == 0:
                lines = result.stdout.split('\n')
                for line in lines:
                    for port in range(8501, 8520):
                        if f':{port} ' in line and 'LISTENING' in line:
                            ports_in_use.append(port)
    except Exception:
        pass
    
    return list(set(ports_in_use))

def check_ports_in_use():
    """Check which ports in the Streamlit range are actually in use"""
    ports_in_use = []
    
    # Method 1: Direct socket binding test (most reliable)
    for port in range(8501, 8520):
        if not is_port_available(port):
            ports_in_use.append(port)
    
    # Method 2: Try to get ports from processes (if psutil supports it)
    streamlit_processes = find_streamlit_processes()
    for proc in streamlit_processes:
        if proc['ports']:
            ports_in_use.extend(proc['ports'])
    
    # Method 3: Fallback using system commands
    fallback_ports = check_ports_in_use_fallback()
    ports_in_use.extend(fallback_ports)
    
    return list(set(ports_in_use))  # Remove duplicates

def kill_streamlit_processes(force=False):
    """Kill existing Streamlit processes - only if explicitly requested or if conflicts detected"""
    streamlit_processes = find_streamlit_processes()
    
    if not streamlit_processes:
        print("ℹ️ No Streamlit processes found")
        return False
    
    print(f"🔍 Found {len(streamlit_processes)} Streamlit process(es):")
    for proc in streamlit_processes:
        ports_info = f" (ports: {proc['ports']})" if proc['ports'] else " (no port info)"
        print(f"   PID {proc['pid']}: {proc['cmdline'][:60]}...{ports_info}")
    
    if not force:
        # Check if any of these processes are actually using ports we want
        ports_in_use = check_ports_in_use()
        
        if not ports_in_use:
            print("ℹ️ No port conflicts detected, keeping existing processes")
            return False
        
        print(f"⚠️ Ports currently in use: {ports_in_use}")
        
        # Check if any Streamlit processes are using these ports
        streamlit_using_ports = []
        for proc in streamlit_processes:
            if proc['ports']:
                streamlit_using_ports.extend(proc['ports'])
        
        # If we can't determine port usage, assume conflict and clean up
        if streamlit_using_ports or len(streamlit_processes) > 0:
            if streamlit_using_ports:
                print(f"🔄 Streamlit processes using ports: {streamlit_using_ports}")
            else:
                print("🔄 Cannot determine port usage, cleaning up Streamlit processes for safety")
            print("🔄 Will clean up conflicting processes...")
        else:
            print("ℹ️ Streamlit processes not using target ports, keeping them")
            return False
    
    try:
        killed_count = 0
        
        for proc_info in streamlit_processes:
            try:
                proc = psutil.Process(proc_info['pid'])
                print(f"🔄 Stopping Streamlit process (PID: {proc_info['pid']})")
                proc.terminate()  # Use terminate for graceful shutdown
                killed_count += 1
                time.sleep(0.5)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        
        if killed_count > 0:
            print(f"✅ Stopped {killed_count} Streamlit process(es)")
            time.sleep(3)  # More time for port cleanup
        
        return killed_count > 0
    except Exception as e:
        print(f"⚠️ Could not clean up processes: {e}")
        return False

def find_available_port(start_port=8501, max_attempts=20):
    """Find an available port with enhanced detection"""
    print(f"🔍 Searching for available port starting from {start_port}...")
    
    # First, get a list of currently used ports
    ports_in_use = check_ports_in_use()
    if ports_in_use:
        print(f"⚠️ Ports currently in use: {ports_in_use}")
    
    # Check ports sequentially, skipping known used ports
    for i in range(max_attempts):
        port = start_port + i
        
        # Skip ports we know are in use
        if port in ports_in_use:
            if i < 5:
                print(f"❌ Port {port} is in use (detected)")
            continue
        
        # Double-check port availability
        if is_port_available(port):
            print(f"✅ Found available port: {port}")
            return port
        elif i < 5:
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
        
        # Check for process conflicts more intelligently
        CLEANUP_PROCESSES = False  # Only cleanup if there are actual conflicts
        
        # First, check what ports are actually in use
        ports_in_use = check_ports_in_use()
        streamlit_processes = find_streamlit_processes()
        
        if streamlit_processes:
            print(f"🔍 Found {len(streamlit_processes)} existing Streamlit process(es)")
            
            # Only cleanup if there are actual port conflicts
            needs_cleanup = False
            if ports_in_use:
                print(f"⚠️ Ports currently in use: {ports_in_use}")
                # Check if any used ports overlap with what we want (8501-8520)
                target_ports = list(range(8501, 8521))
                if any(port in target_ports for port in ports_in_use):
                    needs_cleanup = True
                    print("🔄 Port conflict detected, will cleanup processes...")
            
            if needs_cleanup or CLEANUP_PROCESSES:
                try:
                    kill_streamlit_processes(force=True)
                except Exception as e:
                    print(f"⚠️ Process cleanup failed: {e}")
            else:
                print("ℹ️ No port conflicts detected, keeping existing processes")
        else:
            print("ℹ️ No existing Streamlit processes found")
        
        # Try to find an available port
        available_port = None
        
        # Approach 1: Try standard port range, but start from 8502 if 8501 is busy
        start_port = 8501
        if 8501 in ports_in_use:
            print("ℹ️ Port 8501 is busy, trying alternative ports...")
            start_port = 8502
        
        available_port = find_available_port(start_port, 15)
        
        # Approach 2: If no port found, try alternative ranges
        if not available_port:
            print("🔄 Trying alternative port ranges...")
            for _ in range(2):
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
            cmd = [sys.executable, "-m", "streamlit", "run", "app.py", 
                   "--server.port", str(available_port),
                   "--server.headless", "true",
                   "--server.enableXsrfProtection", "false"]
            subprocess.run(cmd)
        else:
            print("🔄 Using Streamlit's automatic port selection...")
            print("📊 Streamlit app will open in your browser")
            print("📍 URL: Streamlit will choose an available port")
            print("⏹️  Press Ctrl+C to stop the server")
            print("-" * 50)
            
            # Default launch - let Streamlit handle everything
            cmd = [sys.executable, "-m", "streamlit", "run", "app.py",
                   "--server.headless", "true",
                   "--server.enableXsrfProtection", "false"]
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