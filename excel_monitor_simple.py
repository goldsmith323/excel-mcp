#!/usr/bin/env python3
"""
Excel Claude Integration Monitor - Simple Version

A lightweight desktop application that monitors when you open Excel files
and offers to connect them to Claude Desktop for AI assistance.

FEATURES:
- Automatically detects when Excel files are opened
- Prompts user to connect files to Claude Desktop
- Configures Claude Desktop MCP server automatically
- Simple command-line interface (no GUI dependencies)
- Works with any Excel file (.xlsx, .xls, .xlsm, .xlsb)

USAGE:
    python3 excel_monitor_simple.py
    or
    ./start_simple_monitor.sh

WORKFLOW:
1. Start this monitor in Terminal
2. Open any Excel file in Excel application
3. Monitor detects the file and prompts you
4. Choose 'y' to connect the file to Claude
5. Claude Desktop opens and is connected to that specific file
6. Ask Claude questions about the Excel file

DEVELOPMENT:
    To modify the monitor:
    - Edit detection logic in get_open_excel_files()
    - Change prompts in prompt_user_for_integration()
    - Modify Claude config in update_claude_config()
    - Adjust monitoring frequency in monitor_excel_files()

DEPENDENCIES:
    - psutil: For detecting Excel processes and open files
    - subprocess: For launching Claude Desktop
    - json: For updating Claude Desktop configuration
"""

import os
import sys
import json
import time
import psutil
import subprocess
import threading
from pathlib import Path
from datetime import datetime


# ============================================================================
# EXCEL MONITOR CLASS
# ============================================================================

class SimpleExcelMonitor:
    """
    Main monitor class that detects Excel files and manages Claude integration.
    
    This class runs in the background and:
    1. Monitors running Excel processes
    2. Detects when new Excel files are opened
    3. Prompts user for Claude integration
    4. Configures Claude Desktop MCP server
    5. Tracks connected files
    """
    
    def __init__(self):
        self.monitoring = False
        self.connected_files = {}
        self.last_excel_files = set()
        
        # Paths
        self.project_root = Path(__file__).parent
        self.mcp_server_path = self.project_root / "src" / "excel_mcp" / "simple_server.py" 
        self.claude_config_path = Path.home() / "Library" / "Application Support" / "Claude" / "claude_desktop_config.json"
        
    def get_open_excel_files(self):
        """Get list of currently open Excel files"""
        excel_files = set()
        
        try:
            for proc in psutil.process_iter(['pid', 'name', 'open_files']):
                if proc.info['name'] and 'excel' in proc.info['name'].lower():
                    try:
                        if proc.info['open_files']:
                            for file_info in proc.info['open_files']:
                                file_path = file_info.path
                                if file_path.endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')):
                                    if not os.path.basename(file_path).startswith('~'):
                                        excel_files.add(file_path)
                    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                        continue
        except Exception as e:
            print(f"Error getting Excel files: {e}")
            
        return excel_files
    
    def prompt_user_for_integration(self, file_path):
        """Simple command-line prompt for integration"""
        file_name = os.path.basename(file_path)
        
        print("\n" + "="*60)
        print("🤖 CLAUDE EXCEL INTEGRATION")
        print("="*60)
        print(f"📁 Excel file detected: {file_name}")
        print(f"🗂️  Path: {file_path}")
        print()
        print("Would you like Claude to assist with this Excel file?")
        print()
        print("Options:")
        print("  [y] Yes, connect to Claude Desktop")
        print("  [n] No, skip this file")
        print("  [q] Quit monitoring")
        print()
        
        while True:
            try:
                choice = input("Your choice [y/n/q]: ").lower().strip()
                
                if choice in ['y', 'yes']:
                    return 'connect'
                elif choice in ['n', 'no']:
                    return 'skip'
                elif choice in ['q', 'quit']:
                    return 'quit'
                else:
                    print("Please enter 'y', 'n', or 'q'")
                    
            except KeyboardInterrupt:
                return 'quit'
    
    def connect_file_to_claude(self, file_path):
        """Connect Excel file to Claude Desktop"""
        try:
            file_name = os.path.basename(file_path)
            print(f"\n🔄 Connecting {file_name} to Claude Desktop...")
            
            # Update Claude Desktop configuration
            self.update_claude_config(file_path)
            
            # Track connection
            self.connected_files[file_path] = {
                'connected_at': datetime.now().isoformat(),
                'file_name': file_name
            }
            
            print(f"✅ Successfully connected {file_name} to Claude Desktop!")
            print()
            print("Next steps:")
            print("1. Open Claude Desktop")
            print("2. Ask: 'What Excel file am I connected to?'")
            print("3. Try: 'Show me information about this Excel file'")
            print()
            
            # Try to open Claude Desktop
            try:
                subprocess.run(['open', '-a', 'Claude Desktop'], check=False, 
                             stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                print("🚀 Attempting to open Claude Desktop...")
            except:
                print("💡 Please open Claude Desktop manually")
                
        except Exception as e:
            print(f"❌ Failed to connect: {str(e)}")
    
    def update_claude_config(self, file_path):
        """Update Claude Desktop configuration"""
        try:
            # Read existing config
            config = {}
            if self.claude_config_path.exists():
                with open(self.claude_config_path, 'r') as f:
                    config = json.load(f)
            
            # Ensure mcpServers exists
            if 'mcpServers' not in config:
                config['mcpServers'] = {}
            
            # Add Excel MCP configuration
            config['mcpServers']['excel-mcp'] = {
                'command': 'python3',
                'args': [str(self.mcp_server_path)],
                'cwd': str(self.project_root),
                'env': {
                    'EXCEL_FILE_PATH': file_path
                }
            }
            
            # Save config
            self.claude_config_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.claude_config_path, 'w') as f:
                json.dump(config, f, indent=2)
                
        except Exception as e:
            raise Exception(f"Failed to update Claude config: {e}")
    
    def monitor_excel_files(self):
        """Main monitoring loop"""
        print("🔍 Monitoring for Excel files...")
        print("💡 Open any Excel file to see integration prompt")
        print("⏹️  Press Ctrl+C to stop monitoring")
        print()
        
        while self.monitoring:
            try:
                current_files = self.get_open_excel_files()
                new_files = current_files - self.last_excel_files
                
                for file_path in new_files:
                    if file_path not in self.connected_files:
                        choice = self.prompt_user_for_integration(file_path)
                        
                        if choice == 'connect':
                            self.connect_file_to_claude(file_path)
                        elif choice == 'skip':
                            print(f"⏸️ Skipped: {os.path.basename(file_path)}")
                        elif choice == 'quit':
                            print("👋 Stopping monitor...")
                            self.monitoring = False
                            return
                
                # Check for closed files
                closed_files = self.last_excel_files - current_files
                for file_path in closed_files:
                    if file_path in self.connected_files:
                        file_name = self.connected_files[file_path]['file_name']
                        print(f"📝 Excel file closed: {file_name}")
                        del self.connected_files[file_path]
                
                self.last_excel_files = current_files
                
            except Exception as e:
                print(f"❌ Error in monitoring: {e}")
            
            time.sleep(3)  # Check every 3 seconds
    
    def start_monitoring(self):
        """Start monitoring"""
        self.monitoring = True
        self.monitor_excel_files()
    
    def show_status(self):
        """Show current status"""
        print("\n📊 Current Status:")
        print(f"   Monitoring: {'✅ Active' if self.monitoring else '❌ Stopped'}")
        print(f"   Connected files: {len(self.connected_files)}")
        
        if self.connected_files:
            print("\n🔗 Connected files:")
            for file_path, info in self.connected_files.items():
                print(f"   • {info['file_name']}")


# ============================================================================
# MAIN ENTRY POINT
# ============================================================================

def main():
    """
    Main entry point for the Excel monitor.
    
    Starts the monitor and handles keyboard interrupts gracefully.
    The monitor runs continuously until stopped with Ctrl+C.
    """
    print("🤖 Excel Claude Integration Monitor")
    print("=" * 50)
    print()
    
    monitor = SimpleExcelMonitor()
    
    try:
        monitor.start_monitoring()
    except KeyboardInterrupt:
        print("\n👋 Shutting down monitor...")
    except Exception as e:
        print(f"❌ Error: {e}")
    finally:
        print("✅ Monitor stopped")


# ============================================================================
# SCRIPT STARTUP
# ============================================================================

if __name__ == "__main__":
    """
    Start the Excel monitor when this script is run directly.
    
    This is the main way to start the monitor:
        python3 excel_monitor_simple.py
    or
        ./start_simple_monitor.sh
    """
    main()