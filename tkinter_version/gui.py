import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

def select_measurement_type(file_measurement_types):
    """
    Show a dialog to let user select which measurement type to process
    when files have multiple measurement types per channel.
    
    Parameters:
    - file_measurement_types: Dict mapping filename to set of measurement types
    
    Returns:
    - Dict mapping filename to selected measurement type, or None if cancelled
    """
    # Check if any file has multiple measurement types
    files_with_multiple = {f: types for f, types in file_measurement_types.items() if len(types) > 1}
    
    if not files_with_multiple:
        # No selection needed - return the single type for each file
        return {f: list(types)[0] if types else None for f, types in file_measurement_types.items()}
    
    # Create selection dialog
    root = tk.Tk()
    root.title("Measurement Type Selection")
    root.geometry("700x500")
    root.resizable(True, True)
    
    style = ttk.Style()
    style.theme_use('clam')
    
    # Header
    header_frame = tk.Frame(root, bg='#1F4E78', height=80)
    header_frame.pack(fill='x')
    header_frame.pack_propagate(False)
    
    title_label = tk.Label(header_frame, text="Measurement Type Selection", 
                          font=('Segoe UI', 14, 'bold'), 
                          bg='#1F4E78', fg='white')
    title_label.pack(pady=10)
    
    subtitle_label = tk.Label(header_frame, 
                             text="Some files have multiple measurement types per channel. Select which to process:", 
                             font=('Segoe UI', 10), 
                             bg='#1F4E78', fg='white')
    subtitle_label.pack()
    
    # Scrollable frame
    canvas = tk.Canvas(root, bg='white')
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg='white')
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Selection widgets
    selection_vars = {}
    
    for i, (filename, types) in enumerate(files_with_multiple.items()):
        bg_color = '#F8F8F8' if i % 2 == 0 else 'white'
        frame = tk.Frame(scrollable_frame, bg=bg_color, pady=10)
        frame.pack(fill='x', padx=20, pady=5)
        
        # File label
        file_label = tk.Label(frame, text=Path(filename).name, 
                             font=('Segoe UI', 10, 'bold'), bg=bg_color, 
                             anchor='w', width=40)
        file_label.pack(side='left', padx=10)
        
        # Dropdown for measurement type
        types_list = sorted(list(types))
        var = tk.StringVar(value=types_list[0])
        selection_vars[filename] = var
        
        dropdown = ttk.Combobox(frame, textvariable=var, values=types_list, 
                               state='readonly', width=20)
        dropdown.pack(side='left', padx=10)
        
        # Description of types
        desc_label = tk.Label(frame, text=f"Available: {', '.join(types_list)}", 
                             font=('Segoe UI', 9), bg=bg_color, fg='gray')
        desc_label.pack(side='left', padx=10)
    
    canvas.pack(side="left", fill="both", expand=True, padx=10, pady=(10, 90))
    scrollbar.pack(side="right", fill="y", pady=(10, 90))
    
    # Button frame
    button_frame = tk.Frame(root, bg='#F0F0F0', height=80)
    button_frame.pack(fill='x', side='bottom', pady=0, before=canvas)
    button_frame.pack_propagate(False)
    button_frame.lift()
    
    result = {'cancelled': False}
    
    def on_submit():
        result['cancelled'] = False
        root.quit()
        root.destroy()
    
    def on_cancel():
        result['cancelled'] = True
        root.quit()
        root.destroy()
    
    submit_btn = tk.Button(button_frame, text="Continue", command=on_submit,
                          font=('Segoe UI', 11, 'bold'), bg='#0070C0', fg='white',
                          width=12, height=2, cursor='hand2', relief='raised', bd=2)
    submit_btn.pack(side='right', padx=20, pady=15)
    
    cancel_btn = tk.Button(button_frame, text="Cancel", command=on_cancel,
                          font=('Segoe UI', 11), bg='#E0E0E0', fg='black',
                          width=12, height=2, cursor='hand2', relief='raised', bd=2)
    cancel_btn.pack(side='right', padx=5, pady=15)
    
    # Center window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f'+{x}+{y}')
    
    root.mainloop()
    
    if result['cancelled']:
        return None
    
    # Build result dictionary
    selections = {}
    for filename, types in file_measurement_types.items():
        if filename in selection_vars:
            selections[filename] = selection_vars[filename].get()
        elif len(types) == 1:
            selections[filename] = list(types)[0]
        else:
            selections[filename] = None
    
    return selections


def get_user_inputs(test_value_range_io_tuples, unit, input_dir=None):
    """
    Ask user to input range setting, reference value, and tolerance for each 
    test value/range/IO-type combination using a GUI.
    
    Parameters:
    - test_value_range_io_tuples: Set of (test_value, range_setting, io_type) tuples
    - unit: Measurement unit
    - input_dir: Directory path for saving/loading config files
    
    Returns:
    - Dictionary mapping (test_value, range_setting, io_type) to config dict
    """
    user_inputs = {}
    
    root = tk.Tk()
    root.title("Test Configuration Input")
    root.geometry("950x750")
    root.resizable(True, True)
    
    style = ttk.Style()
    style.theme_use('clam')
    
    # Header
    header_frame = tk.Frame(root, bg='#1F4E78', height=100)
    header_frame.pack(fill='x')
    header_frame.pack_propagate(False)
    
    title_label = tk.Label(header_frame, text="Test Configuration Input", 
                          font=('Segoe UI', 16, 'bold'), 
                          bg='#1F4E78', fg='white')
    title_label.pack(pady=10)
    
    subtitle_label = tk.Label(header_frame, 
                             text=f"Configure range, reference value, and tolerance for each test ({unit})", 
                             font=('Segoe UI', 10), 
                             bg='#1F4E78', fg='white')
    subtitle_label.pack()
    
    # Scrollable frame
    canvas = tk.Canvas(root, bg='white')
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas, bg='white')
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Column headers
    header_row = tk.Frame(scrollable_frame, bg='#E8E8E8', pady=8)
    header_row.pack(fill='x', padx=10, pady=(10, 5))
    
    tk.Label(header_row, text="Test Value", font=('Segoe UI', 10, 'bold'), 
             bg='#E8E8E8', width=14, anchor='w').pack(side='left', padx=5)
    tk.Label(header_row, text="I/O Type", font=('Segoe UI', 10, 'bold'), 
             bg='#E8E8E8', width=10, anchor='w').pack(side='left', padx=5)
    tk.Label(header_row, text="Range Setting", font=('Segoe UI', 10, 'bold'), 
             bg='#E8E8E8', width=12, anchor='w').pack(side='left', padx=5)
    tk.Label(header_row, text="Reference Value", font=('Segoe UI', 10, 'bold'), 
             bg='#E8E8E8', width=14, anchor='w').pack(side='left', padx=5)
    tk.Label(header_row, text="Tolerance (±)", font=('Segoe UI', 10, 'bold'), 
             bg='#E8E8E8', width=12, anchor='w').pack(side='left', padx=5)
    
    # Entry fields
    entry_widgets = {}
    # Sort by test value, then by I/O type (Input first, then Output), then by range
    sorted_tuples = sorted(test_value_range_io_tuples, key=lambda x: (x[0], x[2], x[1] or ''))
    
    for i, (test_value, range_setting, io_type) in enumerate(sorted_tuples):
        # Use different background colors for Input vs Output
        if io_type == 'Input':
            bg_color = '#E8F4E8' if i % 2 == 0 else '#F0FAF0'  # Light green tint
        else:
            bg_color = '#E8E8F4' if i % 2 == 0 else '#F0F0FA'  # Light blue tint
        
        frame = tk.Frame(scrollable_frame, bg=bg_color, pady=8)
        frame.pack(fill='x', padx=10, pady=2)
        
        # Test value label
        label = tk.Label(frame, text=f"{test_value:+.4g} {unit}", 
                        font=('Segoe UI', 10), bg=bg_color, width=14, anchor='w')
        label.pack(side='left', padx=5)
        
        # I/O Type label with color coding
        io_color = '#006400' if io_type == 'Input' else '#00008B'  # Dark green for Input, Dark blue for Output
        io_label = tk.Label(frame, text=io_type, 
                           font=('Segoe UI', 10, 'bold'), bg=bg_color, fg=io_color, width=10, anchor='w')
        io_label.pack(side='left', padx=5)
        
        # Range setting entry
        range_entry = ttk.Entry(frame, font=('Segoe UI', 10), width=12)
        range_entry.pack(side='left', padx=5)
        range_entry.insert(0, range_setting if range_setting else "N/A")
        
        # Reference value entry
        ref_entry = ttk.Entry(frame, font=('Segoe UI', 10), width=14)
        ref_entry.pack(side='left', padx=5)
        ref_entry.insert(0, f"{test_value:.6g}")
        
        # Tolerance entry
        tol_entry = ttk.Entry(frame, font=('Segoe UI', 10), width=12)
        tol_entry.pack(side='left', padx=5)
        tol_entry.insert(0, "0.015")
        
        # Unit label
        unit_label = tk.Label(frame, text=unit, font=('Segoe UI', 10), bg=bg_color, width=5)
        unit_label.pack(side='left')
        
        entry_widgets[(test_value, range_setting, io_type)] = {
            'range': range_entry,
            'reference': ref_entry,
            'tolerance': tol_entry
        }
    
    canvas.pack(side="left", fill="both", expand=True, padx=10, pady=(10, 140))
    scrollbar.pack(side="right", fill="y", pady=(10, 140))
    
    # Legend frame
    legend_frame = tk.Frame(scrollable_frame, bg='white', pady=10)
    legend_frame.pack(fill='x', padx=10, pady=(20, 5))
    
    tk.Label(legend_frame, text="Legend: ", font=('Segoe UI', 9, 'bold'), bg='white').pack(side='left', padx=5)
    tk.Label(legend_frame, text="■ Input (TXT files - e.g., Voltmeter readings)", 
             font=('Segoe UI', 9), bg='white', fg='#006400').pack(side='left', padx=10)
    tk.Label(legend_frame, text="■ Output (CSV files - e.g., Power supply output)", 
             font=('Segoe UI', 9), bg='white', fg='#00008B').pack(side='left', padx=10)
    
    # Button frame with two rows
    button_frame = tk.Frame(root, bg='#F0F0F0', height=130)
    button_frame.pack(fill='x', side='bottom', pady=0, before=canvas)
    button_frame.pack_propagate(False)
    button_frame.lift()
    
    # Config buttons row (Save/Load)
    config_row = tk.Frame(button_frame, bg='#F0F0F0')
    config_row.pack(fill='x', pady=(10, 5))
    
    # Status label for showing load/save messages
    status_var = tk.StringVar(value="")
    status_label = tk.Label(config_row, textvariable=status_var, 
                           font=('Segoe UI', 9, 'italic'), bg='#F0F0F0', fg='#666666')
    status_label.pack(side='left', padx=20)
    
    def save_config():
        """Save current configuration to a JSON file"""
        try:
            config_data = {
                'unit': unit,
                'configurations': []
            }
            
            for (test_value, range_setting, io_type), entries in entry_widgets.items():
                config_entry = {
                    'test_value': test_value,
                    'range_setting': range_setting,
                    'io_type': io_type,
                    'range_input': entries['range'].get().strip(),
                    'reference': entries['reference'].get().strip(),
                    'tolerance': entries['tolerance'].get().strip()
                }
                config_data['configurations'].append(config_entry)
            
            # Ask user for save location
            default_filename = "test_config.json"
            if input_dir:
                default_path = os.path.join(input_dir, default_filename)
            else:
                default_path = default_filename
            
            file_path = filedialog.asksaveasfilename(
                title="Save Configuration",
                initialdir=input_dir or os.getcwd(),
                initialfile=default_filename,
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(config_data, f, indent=2)
                status_var.set(f"✓ Configuration saved to {Path(file_path).name}")
                root.after(5000, lambda: status_var.set(""))  # Clear after 5 seconds
        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save configuration:\n{str(e)}")
    
    def load_config():
        """Load configuration from a JSON file"""
        try:
            file_path = filedialog.askopenfilename(
                title="Load Configuration",
                initialdir=input_dir or os.getcwd(),
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            
            if file_path:
                apply_config_file(file_path)
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to load configuration:\n{str(e)}")
    
    def apply_config_file(file_path):
        """Apply configuration from a file to the entry widgets"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            # Check unit compatibility
            if config_data.get('unit') != unit:
                result = messagebox.askyesno(
                    "Unit Mismatch", 
                    f"Config file unit ({config_data.get('unit')}) differs from current unit ({unit}).\n"
                    "Do you want to load it anyway?"
                )
                if not result:
                    return
            
            loaded_count = 0
            for config_entry in config_data.get('configurations', []):
                key = (
                    config_entry['test_value'],
                    config_entry['range_setting'],
                    config_entry['io_type']
                )
                
                if key in entry_widgets:
                    entries = entry_widgets[key]
                    
                    # Clear and set range
                    entries['range'].delete(0, tk.END)
                    entries['range'].insert(0, config_entry.get('range_input', 'N/A'))
                    
                    # Clear and set reference
                    entries['reference'].delete(0, tk.END)
                    entries['reference'].insert(0, config_entry.get('reference', str(key[0])))
                    
                    # Clear and set tolerance
                    entries['tolerance'].delete(0, tk.END)
                    entries['tolerance'].insert(0, config_entry.get('tolerance', '0.015'))
                    
                    loaded_count += 1
            
            status_var.set(f"✓ Loaded {loaded_count} configurations from {Path(file_path).name}")
            root.after(5000, lambda: status_var.set(""))  # Clear after 5 seconds
        except Exception as e:
            messagebox.showerror("Load Error", f"Failed to apply configuration:\n{str(e)}")
    
    save_btn = tk.Button(config_row, text="💾 Save Config", command=save_config,
                        font=('Segoe UI', 10), bg='#5B9BD5', fg='white',
                        width=14, height=1, cursor='hand2', relief='raised', bd=2)
    save_btn.pack(side='right', padx=5)
    
    load_btn = tk.Button(config_row, text="📂 Load Config", command=load_config,
                        font=('Segoe UI', 10), bg='#5B9BD5', fg='white',
                        width=14, height=1, cursor='hand2', relief='raised', bd=2)
    load_btn.pack(side='right', padx=5)
    
    # Main action buttons row (Cancel/Submit)
    action_row = tk.Frame(button_frame, bg='#F0F0F0')
    action_row.pack(fill='x', pady=(5, 15))
    
    result = {'cancelled': False}
    
    def on_submit():
        try:
            for (test_value, range_setting, io_type), entries in entry_widgets.items():
                range_str = entries['range'].get().strip()
                final_range = None if range_str.upper() == 'N/A' or range_str == '' else range_str
                
                ref_str = entries['reference'].get().strip().replace(',', '.')
                reference = float(ref_str)
                
                tol_str = entries['tolerance'].get().strip().replace(',', '.')
                tolerance = float(tol_str)
                
                if tolerance < 0:
                    messagebox.showerror("Error", f"Tolerance for {test_value} {unit} ({io_type}) must be positive!")
                    return
                
                user_inputs[(test_value, range_setting, io_type)] = {
                    'range': final_range,
                    'reference': reference,
                    'tolerance': tolerance
                }
            
            result['cancelled'] = False
            root.quit()
            root.destroy()
        except ValueError as e:
            messagebox.showerror("Error", f"Please enter valid numbers!\n{str(e)}")
    
    def on_cancel():
        result['cancelled'] = True
        root.quit()
        root.destroy()
    
    submit_btn = tk.Button(action_row, text="Submit", command=on_submit,
                          font=('Segoe UI', 11, 'bold'), bg='#0070C0', fg='white',
                          width=12, height=2, cursor='hand2', relief='raised', bd=2)
    submit_btn.pack(side='right', padx=20)
    
    cancel_btn = tk.Button(action_row, text="Cancel", command=on_cancel,
                          font=('Segoe UI', 11), bg='#E0E0E0', fg='black',
                          width=12, height=2, cursor='hand2', relief='raised', bd=2)
    cancel_btn.pack(side='right', padx=5)
    
    # Center window
    root.update_idletasks()
    x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
    y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
    root.geometry(f'+{x}+{y}')
    
    def on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    canvas.bind_all("<MouseWheel>", on_mousewheel)
    
    # Auto-load config if exists in input directory
    if input_dir:
        default_config_path = os.path.join(input_dir, "test_config.json")
        if os.path.exists(default_config_path):
            root.after(100, lambda: apply_config_file(default_config_path))
    
    root.mainloop()
    
    if result['cancelled']:
        return None
    
    return user_inputs
