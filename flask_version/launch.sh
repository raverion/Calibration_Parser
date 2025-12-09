#!/bin/bash

echo "========================================"
echo "  Measurement Data Analyzer"
echo "========================================"
echo ""
echo "Starting server..."
echo ""

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check if Python is available
if ! command -v python3 &> /dev/null; then
    if ! command -v python &> /dev/null; then
        echo "ERROR: Python is not installed"
        echo "Please install Python from https://www.python.org/"
        exit 1
    fi
    PYTHON_CMD="python"
else
    PYTHON_CMD="python3"
fi

# Check if required packages are installed, install if missing
if ! $PYTHON_CMD -c "import flask" &> /dev/null; then
    echo "Installing required packages..."
    $PYTHON_CMD -m pip install -r requirements.txt
fi

# Function to open browser (works on macOS and Linux)
open_browser() {
    sleep 2
    if command -v open &> /dev/null; then
        # macOS
        open "http://localhost:5000"
    elif command -v xdg-open &> /dev/null; then
        # Linux
        xdg-open "http://localhost:5000"
    elif command -v gnome-open &> /dev/null; then
        # Linux (GNOME)
        gnome-open "http://localhost:5000"
    fi
}

# Start browser in background
open_browser &

# Start the Flask server
echo ""
echo "Server running at http://localhost:5000"
echo "Press Ctrl+C to stop the server"
echo ""
$PYTHON_CMD app.py
