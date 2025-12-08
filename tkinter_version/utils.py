import os
from pathlib import Path
import re

# Optional dependency check
try:
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False

# Shared color palette used by both excel_charts.py and html_report.py
CHANNEL_COLORS = [
    '4472C4',  # Muted blue
    'C45B5B',  # Muted red
    '70AD47',  # Muted green
    'ED7D31',  # Muted orange
    '7B7B7B',  # Gray
    '9E5ECE',  # Muted purple
    '43A6A2',  # Muted teal
    'C4A24E',  # Muted gold
    '5B9BC4',  # Steel blue
    'A85B5B',  # Dusty rose
    '5BAF7B',  # Sea green
    'C47B4E',  # Terracotta
    '6B6BAF',  # Muted indigo
    '8B6BAF',  # Muted violet
    '4EAFAF',  # Turquoise
    'AF8B4E',  # Bronze
]

# HTML version with # prefix
CHANNEL_COLORS_HEX = [f'#{c}' for c in CHANNEL_COLORS]

def get_versioned_filename(base_path):
    """
    Generate a versioned filename if the file already exists.
    Returns the next available filename with _v2, _v3, etc.
    """
    if not os.path.exists(base_path):
        return base_path
    
    # Split the path into directory, name, and extension
    directory = os.path.dirname(base_path)
    filename = os.path.basename(base_path)
    name, ext = os.path.splitext(filename)
    
    # Check if filename already has a version suffix
    import re
    version_match = re.match(r'^(.+)_v(\d+)$', name)
    if version_match:
        base_name = version_match.group(1)
        current_version = int(version_match.group(2))
    else:
        base_name = name
        current_version = 1
    
    # Find the next available version
    version = current_version + 1 if current_version > 1 else 2
    while True:
        new_filename = f"{base_name}_v{version}{ext}"
        new_path = os.path.join(directory, new_filename)
        if not os.path.exists(new_path):
            return new_path
        version += 1