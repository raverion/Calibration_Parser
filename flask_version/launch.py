#!/usr/bin/env python3
"""
Cross-platform launcher for Measurement Data Analyzer
Double-click this file to start the application
"""

import subprocess
import sys
import os
import webbrowser
import time
import threading

def check_and_install_dependencies():
    """Check if required packages are installed, install if missing."""
    required = ['flask', 'pandas', 'numpy', 'openpyxl', 'plotly']
    missing = []
    
    for package in required:
        try:
            __import__(package)
        except ImportError:
            missing.append(package)
    
    if missing:
        print(f"Installing missing packages: {', '.join(missing)}")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + missing)
        print("Packages installed successfully!\n")

def open_browser_delayed():
    """Open browser after a short delay to allow server to start."""
    time.sleep(2)
    webbrowser.open('http://localhost:5000')

def main():
    print("=" * 50)
    print("  Measurement Data Analyzer")
    print("=" * 50)
    print()
    
    # Change to script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    # Check dependencies
    print("Checking dependencies...")
    check_and_install_dependencies()
    
    # Start browser in background thread
    print("Starting server and opening browser...")
    browser_thread = threading.Thread(target=open_browser_delayed)
    browser_thread.daemon = True
    browser_thread.start()
    
    print()
    print("Server running at http://localhost:5000")
    print("Press Ctrl+C to stop the server")
    print()
    
    # Import and run the Flask app
    from app import app
    app.run(debug=False, host='127.0.0.1', port=5000)

if __name__ == '__main__':
    main()
