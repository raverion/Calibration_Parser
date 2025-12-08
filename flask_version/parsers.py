import re
from pathlib import Path

def parse_filename(filename):
    """
    Extract test value, unit, channel number (if present), and range setting from filename.
    Supports formats like:
    - VT2816A_m2V5_R10V_CH3.csv (voltage: -2.5V, range: 10V, channel: 3)
    - VIO2004_3mA_R10mA_CH1.txt (current: 3mA, range: 10mA, channel: 1)
    - VT2816A_10V_R10V_1000x.txt (voltage: 10V, range: 10V, no channel - multi-channel file)
    - VT2516A_25V_1000x.txt (voltage: 25V, no range, no channel - multi-channel file)
    """
    name = Path(filename).stem
    
    # Extract channel pattern (e.g., CH1, CH2, CH3, CH4) - may not be present
    channel_pattern = r'_CH(\d+)'
    channel_match = re.search(channel_pattern, name, re.IGNORECASE)
    channel_num = int(channel_match.group(1)) if channel_match else None
    
    # Extract range setting pattern (e.g., R10V, R10mA, R100ohm)
    range_pattern = r'_R(\d+(?:\.\d+)?)(V|mV|mA|uA|A|ohm|Ohm|kOhm|MOhm)(?:_|$)'
    range_match = re.search(range_pattern, name, re.IGNORECASE)
    
    range_setting = None
    if range_match:
        range_value = range_match.group(1)
        range_unit = range_match.group(2)
        range_setting = f"{range_value}{range_unit}"
    
    # Try voltage pattern first (m2V5, p7V5, 0V, 10V, 25V) - but not matching the R prefix range
    voltage_pattern = r'(?<!R)_([mp]?\d+V\d*)(?:_|$)'
    voltage_match = re.search(voltage_pattern, name, re.IGNORECASE)
    
    if voltage_match:
        voltage_str = voltage_match.group(1).lower()
        sign = -1 if voltage_str.startswith('m') else 1
        voltage_str = voltage_str.lstrip('mp').replace('v', '.')
        # Handle cases like "10V" -> "10." -> need to strip trailing dot
        if voltage_str.endswith('.'):
            voltage_str = voltage_str[:-1]
        try:
            value = sign * float(voltage_str)
            return value, 'V', channel_num, range_setting
        except ValueError:
            pass
    
    # Try milliampere pattern (e.g., 3mA, m5mA, p10mA)
    ma_pattern = r'(?<!R)_([mp]?\d+(?:\.\d+)?)\s*mA(?:_|$)'
    ma_match = re.search(ma_pattern, name, re.IGNORECASE)
    
    if ma_match:
        value_str = ma_match.group(1)
        sign = -1 if value_str.startswith('m') else 1
        value_str = value_str.lstrip('mp')
        try:
            value = sign * float(value_str)
            return value, 'mA', channel_num, range_setting
        except ValueError:
            pass
    
    # Try microampere pattern (e.g., 100uA, m50uA)
    ua_pattern = r'(?<!R)_([mp]?\d+(?:\.\d+)?)\s*uA(?:_|$)'
    ua_match = re.search(ua_pattern, name, re.IGNORECASE)
    
    if ua_match:
        value_str = ua_match.group(1)
        sign = -1 if value_str.startswith('m') else 1
        value_str = value_str.lstrip('mp')
        try:
            value = sign * float(value_str)
            return value, 'uA', channel_num, range_setting
        except ValueError:
            pass
    
    # Try ampere pattern (e.g., 1A, 2A)
    a_pattern = r'(?<!R|m|u)_([mp]?\d+(?:\.\d+)?)\s*A(?:_|$)'
    a_match = re.search(a_pattern, name, re.IGNORECASE)
    
    if a_match:
        value_str = a_match.group(1)
        sign = -1 if value_str.startswith('m') else 1
        value_str = value_str.lstrip('mp')
        try:
            value = sign * float(value_str)
            return value, 'A', channel_num, range_setting
        except ValueError:
            pass
    
    # Try ohms pattern (10_ohms, 100ohms, etc.)
    ohms_pattern = r'_(\d+(?:\.\d+)?)[_\s]?ohms?(?:_|$)'
    ohms_match = re.search(ohms_pattern, name, re.IGNORECASE)
    
    if ohms_match:
        try:
            value = float(ohms_match.group(1))
            return value, 'Ohm', channel_num, range_setting
        except ValueError:
            pass
    
    # Try generic numeric pattern with underscore
    generic_pattern = r'_([mp]?\d+(?:\.\d+)?)_'
    generic_match = re.search(generic_pattern, name)
    
    if generic_match:
        value_str = generic_match.group(1)
        sign = -1 if value_str.startswith('m') else 1
        value_str = value_str.lstrip('mp')
        try:
            value = sign * float(value_str)
            return value, 'unknown', channel_num, range_setting
        except ValueError:
            pass
    
    return None, None, channel_num, range_setting

def get_unit_from_files(input_dir):
    """
    Determine the measurement unit from filenames.
    """
    all_files = list(Path(input_dir).glob('*.csv')) + list(Path(input_dir).glob('*.txt'))
    
    for file in all_files:
        _, unit, _, _ = parse_filename(file.name)
        if unit and unit != 'unknown':
            return unit
    
    return 'V'  # Default to volts

def scan_text_file_for_measurement_types(file_path):
    """
    Scan a text file to find all unique measurement types per channel.
    Returns a set of measurement type names found (e.g., {'Voltage', 'MeanVoltage'} or {'CurVoltage'} or {'Avg'})
    """
    measurement_types = set()
    
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
    except Exception:
        return measurement_types
    
    lines = content.strip().split('\n')
    
    # Check for hierarchical format (VIO1008 style)
    # Pattern: |  MeasurementType_Chxx   value   unit   ...
    hierarchical_pattern = r'\|\s+(\w+)_Ch\d+'
    for line in lines[:500]:  # Check first 500 lines
        match = re.search(hierarchical_pattern, line)
        if match:
            measurement_types.add(match.group(1))
    
    if measurement_types:
        return measurement_types
    
    # Check for flat format (VT2816A/VT2516A style)
    # Pattern: Time  Name::MeasurementType  Data
    flat_pattern = r'_Ch\d+::(\w+)'
    for line in lines[:500]:
        match = re.search(flat_pattern, line, re.IGNORECASE)
        if match:
            measurement_types.add(match.group(1))
    
    return measurement_types

def parse_text_file(file_path, selected_measurement_type=None, channel_from_filename=None):
    """
    Parse text files containing measurement data from multiple channels.
    
    Supports multiple formats:
    1. Hierarchical format (VIO1008 style):
       [-] timestamp   TaskName
             |  Voltage_Ch01       -2.498169   V   ...
             |  MeanVoltage_Ch01   -2.498347   V   ...
    
    2. Flat format with channel in name (VT2816A style):
       Time        Name                        Data
       66.001210   VT2816_1_Ch1::CurVoltage    10.011883
    
    3. Flat format with channel in name (VT2516A style):
       Time        Name                 Data
       30.000132   VT2516_1_Ch1::Avg    24.976000
    
    4. Simple flat format without channel in data (VN1630A style):
       Time        Name            Data
       15.001821   VN1600_1::AIN   0.686400
       (Channel comes from filename, e.g., VN1630A_0V7_CH1_100x.txt)
    
    Parameters:
    - file_path: Path to the text file
    - selected_measurement_type: If file has multiple measurement types per channel,
                                 use this one (e.g., 'Voltage' or 'MeanVoltage')
    - channel_from_filename: Channel number parsed from filename (used for format 4)
    
    Returns: Dictionary mapping channel numbers to lists of measurements
    """
    channel_data = {}
    
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return channel_data
    
    lines = content.strip().split('\n')
    if not lines:
        return channel_data
    
    # Try hierarchical format first (VIO1008 style)
    # Pattern: |  MeasurementType_Chxx   value   unit   value   description
    hierarchical_pattern = r'\|\s+(\w+)_Ch(\d+)\s+(-?\d+\.?\d*)\s+(\w+)'
    
    hierarchical_matches = []
    for line in lines:
        match = re.search(hierarchical_pattern, line)
        if match:
            hierarchical_matches.append({
                'type': match.group(1),
                'channel': int(match.group(2)),
                'value': float(match.group(3))
            })
    
    if hierarchical_matches:
        # Filter by selected measurement type if specified
        for match in hierarchical_matches:
            if selected_measurement_type and match['type'] != selected_measurement_type:
                continue
            
            channel = match['channel']
            value = match['value']
            
            if channel not in channel_data:
                channel_data[channel] = []
            channel_data[channel].append(value)
        
        if channel_data:
            return channel_data
    
    # Try flat format with channel in name (VT2816A/VT2516A style)
    # Pattern: Time  DeviceName_Chxx::MeasurementType  Data
    flat_pattern = r'^\s*[\d.]+\s+\S+_Ch(\d+)::(\w+)\s+(-?\d+\.?\d*)'
    
    for line in lines:
        match = re.search(flat_pattern, line)
        if match:
            channel = int(match.group(1))
            meas_type = match.group(2)
            value = float(match.group(3))
            
            # Filter by selected measurement type if specified
            if selected_measurement_type and meas_type != selected_measurement_type:
                continue
            
            if channel not in channel_data:
                channel_data[channel] = []
            channel_data[channel].append(value)
    
    if channel_data:
        return channel_data
    
    # Try simple flat format without channel in data (VN1630A style)
    # Pattern: Time  Name::Type  Data  OR  Time  Name  Data
    # Channel must come from filename
    simple_flat_pattern = r'^\s*[\d.]+\s+\S+\s+(-?\d+\.?\d*)\s*$'
    
    values_found = []
    for line in lines:
        match = re.search(simple_flat_pattern, line)
        if match:
            try:
                value = float(match.group(1))
                values_found.append(value)
            except ValueError:
                continue
    
    if values_found:
        # Use channel from filename if provided, otherwise default to channel 1
        channel = channel_from_filename if channel_from_filename is not None else 1
        channel_data[channel] = values_found
        return channel_data
    
    return channel_data
