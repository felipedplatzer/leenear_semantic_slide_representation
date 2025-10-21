import tkinter as tk
from tkinter import messagebox
import threading
import time
import win32com.client
import group_shapes
import basic
import structure_shapes
import create_shape_boxes
import resources
import pythoncom
# Global variables
root = None
invisible_root = None
start_button = None
status_label = None
test_index_entry = None
listener_active = False
ppt_app = None
selection_monitor_thread = None
group_list = []
individual_shape_labels = []  # Store individual shape labels
text_section_list = []
table_labels_list = []
last_selection_ref = None
last_selection_ref_shape_ids = []

ppt_app = resources.get_powerpoint_app()

x = create_shape_boxes.create_shape_boxes(ppt_app)
shape_data = x['shape_data']
slide_dimensions = x['slide_dimensions']

def setup_gui():
    """Create the main GUI window"""
    global root, start_button, status_label, test_index_entry
    
    root = tk.Tk()
    root.title("PowerPoint Shape Labeling Tool")
    root.geometry("350x200")
    
    # Bring window to foreground
    root.lift()
    root.attributes('-topmost', True)
    root.after_idle(lambda: root.attributes('-topmost', False))
    root.focus_force()
    
    # Main frame
    main_frame = tk.Frame(root, padx=20, pady=20)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Title
    title_label = tk.Label(main_frame, text="PowerPoint Shape Labeling", 
                          font=("Arial", 14, "bold"))
    title_label.pack(pady=(0, 15))
    
    # Test index input
    test_index_label = tk.Label(main_frame, text="Enter test index:", 
                               font=("Arial", 10))
    test_index_label.pack(anchor="w", pady=(0, 5))
    
    test_index_entry = tk.Entry(main_frame, font=("Arial", 11), width=20)
    test_index_entry.pack(pady=(0, 15))
    test_index_entry.focus()
    
    # Start labeling button
    start_button = tk.Button(main_frame, text="Start Labeling", 
                           command=start_labeling,
                           font=("Arial", 12),
                           bg="#4CAF50", fg="white",
                           width=15, height=2)
    start_button.pack(pady=10)
    
    # Status label
    status_label = tk.Label(main_frame, text="Ready to start", 
                          fg="gray")
    status_label.pack(pady=5)

def start_labeling():
    """Initialize PowerPoint connection and start the listener"""
    global listener_active, ppt_app, selection_monitor_thread, test_index
    
    try:
        # Get test index from input field
        test_index_str = test_index_entry.get().strip()
        if not test_index_str:
            messagebox.showerror("Error", "Please enter a test index")
            return
        
        test_index = int(test_index_str)
        status_label.config(text="Connected to PowerPoint", fg="green")
        
        # Start the selection monitor in a separate thread
        listener_active = True
        selection_monitor_thread = threading.Thread(target=monitor_selection)
        selection_monitor_thread.daemon = True
        selection_monitor_thread.start()
        
        # Update button
        start_button.config(text="Labeling Active", state="disabled", bg="#FF9800")
        
        # Hide the main window but keep the program running
        root.withdraw()
        
        # Create a new invisible root window to keep the program running
        global invisible_root
        invisible_root = tk.Tk()
        invisible_root.withdraw()  # Keep it invisible
        
    except ValueError:
        messagebox.showerror("Error", "Test index must be a number")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to start labeling: {str(e)}")
        status_label.config(text="Connection failed", fg="red")

def monitor_selection():
    """Monitor PowerPoint selection changes in background thread"""
    global listener_active, ppt_app
    
    while listener_active:
        try:
            # Use root.after to check selection from main thread
            root.after(0, check_selection)
            time.sleep(0.5)  # Check every 500ms
            
        except Exception as e:
            print(f"Selection monitor error: {e}")
            time.sleep(1)

def check_selection():
    """Check PowerPoint selection from main thread"""
    global listener_active, ppt_app, last_selection_ref, last_selection_ref_shape_ids

    try:
        if listener_active and ppt_app and ppt_app.ActiveWindow:
            selection = ppt_app.ActiveWindow.Selection
            
            # Check if selection has changed and is a group
            if selection.Type ==2:  # ppSelectionShapes
                current_selection_shape_ids = [shape.Id for shape in selection.ShapeRange]
                # current_selection = str(selection.ShapeRange.Count)
                
                if set(current_selection_shape_ids) != set(last_selection_ref_shape_ids) and selection.ShapeRange.Count >= 1:
                    print("Selection changed")
                    last_selection_ref = selection
                    last_selection_ref_shape_ids = current_selection_shape_ids
                    # Show group naming dialog
                    show_group_naming_dialog(selection)
                    
    except Exception as e:
        print(f"Selection check error: {e}")

def show_group_naming_dialog(selection):
    """Show dialog for naming the selected group"""
    # Create a new root window for the dialog
    dialog_root = tk.Tk()
    dialog_root.title("Name This Group")
    dialog_root.geometry("400x250")
    dialog_root.resizable(False, False)
    
    # Center on screen
    dialog_root.update_idletasks()
    x = (dialog_root.winfo_screenwidth() // 2) - (400 // 2)
    y = (dialog_root.winfo_screenheight() // 2) - (250 // 2)
    dialog_root.geometry(f"400x250+{x}+{y}")
    
    # Bring dialog to foreground
    dialog_root.lift()
    dialog_root.attributes('-topmost', True)
    dialog_root.after_idle(lambda: dialog_root.attributes('-topmost', False))
    dialog_root.focus_force()
    
    # Main frame
    main_frame = tk.Frame(dialog_root, padx=20, pady=20)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Title
    title_label = tk.Label(main_frame, text="Group Selected Shapes", 
                          font=("Arial", 12, "bold"))
    title_label.pack(pady=(0, 10))
    
    # Selection info
    info_text = f"Selected {selection.ShapeRange.Count} shapes"
    info_label = tk.Label(main_frame, text=info_text, fg="gray")
    info_label.pack(pady=(0, 15))
    
    # Group name input with updated header
    name_label = tk.Label(main_frame, text="Leave blank to not save group, e.g., if you made a mistake with the selection:", 
                         font=("Arial", 10), wraplength=350, justify="left")
    name_label.pack(anchor="w", pady=(0, 5))
    
    name_entry = tk.Entry(main_frame, font=("Arial", 11), width=40)
    name_entry.pack(pady=(0, 20))
    
    # Buttons frame
    buttons_frame = tk.Frame(main_frame)
    buttons_frame.pack(fill=tk.X, pady=(0, 10))
    
    # Save button (save and listen for next group)
    save_btn = tk.Button(buttons_frame, text="Save", 
                         command=lambda: go_to_next_group(name_entry.get().strip(), selection, dialog_root),
                         bg="#4CAF50", fg="white",
                         width=15, height=3)
    save_btn.pack(side=tk.LEFT, padx=(0, 10))
    
    # Go to Next Section button (with save confirmation)
    next_section_btn = tk.Button(buttons_frame, text="Go to Next Section\n(Individual Shapes)", 
                                 command=lambda: go_to_next_section_individual(dialog_root, name_entry, selection),
                            bg="#2196F3", fg="white",
                                 width=15, height=3)
    next_section_btn.pack(side=tk.LEFT)
    
    # Focus on textbox after dialog is created
    dialog_root.after(100, lambda: name_entry.focus())
    
    # Bind Enter key to Save
    name_entry.bind('<Return>', lambda e: go_to_next_group(name_entry.get().strip(), selection, dialog_root))
    
    # Bind Escape key to skip (go to next group without saving)
    dialog_root.bind('<Escape>', lambda e: go_to_next_group("", selection, dialog_root))
    
    # Bind Enter key to focused button when buttons are focused
    def on_button_focus(event):
        widget = event.widget
        if isinstance(widget, tk.Button):
            widget.bind('<Return>', lambda e: widget.invoke())
    
    # Apply focus binding to all buttons
    for button in [save_btn, next_section_btn]:
        button.bind('<FocusIn>', on_button_focus)
    
    # Handle window close (X button) - save and quit
    def on_closing():
        # Save current group if name is provided
        group_name = name_entry.get().strip()
        if group_name:
            save(group_name, selection)
        # Stop the listener and finish the process
        global listener_active
        listener_active = False
        dialog_root.destroy()
        finish_text_labeling_process()
    
    dialog_root.protocol("WM_DELETE_WINDOW", on_closing)

def go_to_next_group(group_name, selection, dialog_root):
    """Go to next group, saving current group if name is provided"""
    if group_name and group_name.strip():
        save(group_name, selection)
    else:
        print("Group name is blank - not saving group")
    dialog_root.destroy()
    print("Listening for next group")

def go_to_next_section_individual(dialog_root, name_entry, selection):
    """Go to individual shapes section with confirmation"""
    # Check if there's unsaved content
    group_name = name_entry.get().strip()
    if group_name:
        # Show confirmation popup
        from tkinter import messagebox
        result = messagebox.askyesnocancel("Save Changes?", 
                                          "Do you want to save changes before going to the next section?\n\n" +
                                          "Yes = Save and continue\n" +
                                          "No = Don't save and continue\n" +
                                          "Cancel = Stay on this page")
        if result is None:  # Cancel
            return
        elif result:  # Yes - save
            print("Saving group before going to next section")
            save(group_name, selection)
        else:  # No - don't save
            print("Going to next section - not saving group")
    else:
        print("Going to next section - no content to save")
    
    dialog_root.destroy()
    show_individual_shapes_form()

def go_to_text_sections(group_name, selection, dialog_root):
    """Go to individual shapes labeling, saving current group if name is provided"""
    if group_name and group_name.strip():
        save(group_name, selection)
    else:
        print("Group name is blank - not saving group")
    dialog_root.destroy()
    show_individual_shapes_form()

def skip_labeling(dialog_root):
    """Skip this selection and continue labeling"""
    dialog_root.destroy()
    # Status updates are not needed since main window is hidden
    print("Skipped selection - listening for next group")

def get_unlabeled_shapes():
    """Get list of shape IDs that have not been labeled yet"""
    # Get all labeled shape IDs
    labeled_ids = set()
    
    # Add shapes from groups (including single-shape groups)
    for group in group_list:
        if isinstance(group.get('shape_id'), list):
            labeled_ids.update(group['shape_id'])
        else:
            labeled_ids.add(group['shape_id'])
    
    # Add individually labeled shapes
    for shape_label in individual_shape_labels:
        labeled_ids.add(shape_label['shape_id'])
    
    # Get all shape IDs from shape_data
    all_shape_ids = [shape['shape_id'] for shape in shape_data]
    
    # Find unlabeled shapes
    unlabeled = [shape_id for shape_id in all_shape_ids if shape_id not in labeled_ids]
    
    return unlabeled

def show_individual_shapes_form():
    """Show form for labeling individual shapes"""
    global listener_active
    listener_active = False  # Stop listening for selections
    
    # Create dialog
    dialog_root = tk.Toplevel()
    dialog_root.title("Label Individual Shapes")
    dialog_root.geometry("500x400")
    
    # Bring to front
    dialog_root.lift()
    dialog_root.attributes('-topmost', True)
    dialog_root.after_idle(lambda: dialog_root.attributes('-topmost', False))
    dialog_root.focus_force()
    
    # Main frame
    main_frame = tk.Frame(dialog_root, padx=20, pady=20)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Title
    title_label = tk.Label(main_frame, text="Label Individual Shapes", 
                           font=("Arial", 14, "bold"))
    title_label.pack(pady=(0, 10))
    
    # Instructions
    instructions = tk.Label(main_frame, 
                           text="Select shapes in PowerPoint and enter a label.\nEach selected shape will be labeled individually with the same label.",
                           font=("Arial", 10), wraplength=450, justify="left")
    instructions.pack(pady=(0, 15))
    
    # Unlabeled shapes info
    unlabeled_shapes = get_unlabeled_shapes()
    unlabeled_frame = tk.LabelFrame(main_frame, text="Unlabeled Shapes", font=("Arial", 10, "bold"))
    unlabeled_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
    
    # Scrollable text widget for unlabeled shapes
    unlabeled_text = tk.Text(unlabeled_frame, height=8, width=50, font=("Arial", 9), wrap=tk.WORD)
    unlabeled_scroll = tk.Scrollbar(unlabeled_frame, command=unlabeled_text.yview)
    unlabeled_text.configure(yscrollcommand=unlabeled_scroll.set)
    unlabeled_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    unlabeled_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    if unlabeled_shapes:
        unlabeled_text.insert('1.0', f"Remaining shapes ({len(unlabeled_shapes)}): {', '.join(unlabeled_shapes)}")
    else:
        unlabeled_text.insert('1.0', "All shapes have been labeled!")
    unlabeled_text.config(state='disabled')
    
    # Label input
    label_frame = tk.Frame(main_frame)
    label_frame.pack(fill=tk.X, pady=(0, 20))
    
    tk.Label(label_frame, text="Label for selected shapes (NOT IN PLURAL - this is a label for individual shapes to be repeated across elements, not for a group):", 
             font=("Arial", 10), wraplength=450, justify="left").pack(anchor="w", pady=(0, 5))
    label_entry = tk.Entry(label_frame, font=("Arial", 10), width=40)
    label_entry.pack(fill=tk.X)
    label_entry.focus()
    
    # Buttons
    buttons_frame = tk.Frame(main_frame)
    buttons_frame.pack(fill=tk.X)
    
    def save_individual_shapes():
        label = label_entry.get().strip()
        if not label:
            print("Label is blank - not saving")
            return
        
        try:
            selection = ppt_app.ActiveWindow.Selection
            if selection.Type != 2:  # ppSelectionShapes
                print("No shapes selected")
                return
            
            # Save each shape individually with the same label
            for i in range(1, selection.ShapeRange.Count + 1):
                shape = selection.ShapeRange.Item(i)
                shape_id = str(shape.Id)
                
                individual_shape_labels.append({
                    'shape_id': shape_id,
                    'label': label
                })
                print(f"Labeled shape {shape_id} as '{label}'")
            
            # Update unlabeled shapes display
            unlabeled_shapes = get_unlabeled_shapes()
            unlabeled_text.config(state='normal')
            unlabeled_text.delete('1.0', tk.END)
            if unlabeled_shapes:
                unlabeled_text.insert('1.0', f"Remaining shapes ({len(unlabeled_shapes)}): {', '.join(unlabeled_shapes)}")
            else:
                unlabeled_text.insert('1.0', "All shapes have been labeled!")
            unlabeled_text.config(state='disabled')
            
            # Clear label entry for next input
            label_entry.delete(0, tk.END)
            label_entry.focus()
            
        except Exception as e:
            print(f"Error saving individual shapes: {e}")
    
    def go_to_previous_section_from_individual():
        """Go back to group naming (listening mode)"""
        dialog_root.destroy()
        # Restart group listening - just close the form and the listener will continue
        print("Going back to group selection")
    
    def go_to_text_sections_from_individual():
        """Go to text sections with confirmation if there's unsaved work"""
        # Check if label entry has unsaved content
        label_text = label_entry.get().strip()
        if label_text:
            from tkinter import messagebox
            result = messagebox.askyesno("Unsaved Label", 
                                        f"You have an unsaved label: '{label_text}'\n\n" +
                                        "Do you want to continue to the next section without saving?")
            if not result:  # User chose No - stay on page
                return
        
        dialog_root.destroy()
        show_text_labeling_form(None)
    
    # Go to Previous Section button
    prev_section_btn = tk.Button(buttons_frame, text="Go to Previous Section\n(Groups)", 
                                 command=go_to_previous_section_from_individual,
                                 bg="#FF9800", fg="white", width=20, height=2)
    prev_section_btn.pack(side=tk.LEFT, padx=(0, 10))
    
    save_btn = tk.Button(buttons_frame, text="Save", 
                        command=save_individual_shapes,
                        bg="#4CAF50", fg="white", width=20, height=2)
    save_btn.pack(side=tk.LEFT, padx=(0, 10))
    
    next_section_btn = tk.Button(buttons_frame, text="Go to Next Section\n(Text Sections)", 
                          command=go_to_text_sections_from_individual,
                          bg="#2196F3", fg="white", width=20, height=2)
    next_section_btn.pack(side=tk.LEFT)
    
    # Bind Enter key
    label_entry.bind('<Return>', lambda e: save_individual_shapes())
    
    # Handle window close
    def on_closing():
        dialog_root.destroy()
        show_text_labeling_form(None)
    
    dialog_root.protocol("WM_DELETE_WINDOW", on_closing)

def save(group_name, selection):
    if group_name is not None and group_name != '':
        x = group_shapes.group_selected_shapes(group_name, selection)
        group_list.append(x)
        print(f"Group named: {group_name}")
        print(f"Selection count: {selection.ShapeRange.Count}")
        print(f"Group '{group_name}' saved - listening for next group")
    else:
        print("Group name cannot be empty. Saving skipped")

def finish(dialog_root):
    global listener_active
    # Stop the listener
    dialog_root.destroy()
    listener_active = False
    print("Labeling finished - stopping listener")
    group_dl = basic.process_groups(shape_data, group_list, test_index, slide_dimensions)
    group_dl = structure_shapes.generate_structure_main(group_dl)
    group_df = basic.save_to_csv(group_dl, test_index)  # Saves as CSV
    print("Done")   
    print("Finished labeling process")
    exit()


def continue_labeling(group_name, selection, dialog_root):
    save(group_name, selection)
    dialog_root.destroy()


def save_and_finish(group_name, selection, dialog_root):
    """Save group and open text labeling form"""
    save(group_name, selection)
    dialog_root.destroy()
    show_text_labeling_form(selection)

def update_text_preview(preview_label, shape_id_entry, start_char_entry, end_char_entry, selection):
    """Update the text preview based on input values"""
    try:
        shape_id = shape_id_entry.get().strip()
        start_char = start_char_entry.get().strip()
        end_char = end_char_entry.get().strip()
        
        # Check if all required fields are filled
        if not shape_id or not start_char or not end_char:
            preview_label.config(text="Please fill in Shape ID, Start Char, and End Char", fg="gray")
            return
        
        # Convert to integers
        try:
            start_pos = int(start_char)
            end_pos = int(end_char)
        except ValueError:
            preview_label.config(text="Error: Start Char and End Char must be numbers", fg="red")
            return
        
        # Find the shape by ID
        target_shape = None
        for shape in ppt_app.ActiveWindow.View.Slide.Shapes:
            if str(shape.Id) == str(shape_id):
                target_shape = shape
                break
        
        if not target_shape:
            preview_label.config(text=f"Error: Shape ID '{shape_id}' not found in selection", fg="red")
            return
        
        # Check if shape has text
        if not hasattr(target_shape, 'TextFrame'):
            preview_label.config(text=f"Error: Shape {shape_id} does not contain text", fg="red")
            return
        
        # Get the text content
        full_text = target_shape.TextFrame.TextRange.Text
        
        # Validate character positions
        if start_pos < 0 or end_pos < 0:
            preview_label.config(text="Error: Character positions must be positive numbers", fg="red")
            return
        
        if start_pos >= len(full_text):
            preview_label.config(text=f"Error: Start position {start_pos} is beyond text length ({len(full_text)})", fg="red")
            return
        
        if end_pos > len(full_text):
            preview_label.config(text=f"Error: End position {end_pos} is beyond text length ({len(full_text)})", fg="red")
            return
        
        if start_pos >= end_pos:
            preview_label.config(text="Error: Start position must be less than end position", fg="red")
            return
        
        # Extract the text segment
        text_segment = full_text[start_pos:end_pos]
        
        # Display the preview
        preview_text = f"Text Preview (Chars {start_pos}-{end_pos}):\n\n{text_segment}"
        preview_label.config(text=preview_text, fg="black")
        
    except Exception as e:
        preview_label.config(text=f"Error: {str(e)}", fg="red")

def show_text_labeling_form(selection):
    """Show text labeling form for text sections"""
    # Create new text labeling dialog
    text_dialog = tk.Tk()
    text_dialog.title("Label Text Sections")
    
    # Get screen dimensions and make fullscreen
    screen_width = text_dialog.winfo_screenwidth()
    screen_height = text_dialog.winfo_screenheight()
    text_dialog.geometry(f"{screen_width}x{screen_height}+0+0")
    text_dialog.state('zoomed')  # Maximize window on Windows
    text_dialog.resizable(True, True)
    
    # Bring dialog to foreground
    text_dialog.lift()
    text_dialog.attributes('-topmost', True)
    text_dialog.after_idle(lambda: text_dialog.attributes('-topmost', False))
    text_dialog.focus_force()
    
    # Create main scrollable frame
    canvas = tk.Canvas(text_dialog, highlightthickness=0)
    scrollbar = tk.Scrollbar(text_dialog, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)
    
    # Bind to update scroll region
    def on_main_frame_configure(event):
        canvas.configure(scrollregion=canvas.bbox("all"))
    
    scrollable_frame.bind("<Configure>", on_main_frame_configure)
    
    # Create window that expands to canvas width
    main_canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    
    # Bind canvas resize to update frame width
    def on_main_canvas_configure(event):
        # Make the scrollable_frame fill the canvas width
        canvas.itemconfig(main_canvas_window, width=event.width)
    
    canvas.bind("<Configure>", on_main_canvas_configure)
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Main frame - reduce horizontal padding to maximize table width
    main_frame = tk.Frame(scrollable_frame, padx=20, pady=30)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Title
    title_label = tk.Label(main_frame, text="Label Text Sections", 
                          font=("Arial", 18, "bold"))
    title_label.pack(pady=(0, 20))
    
    # 1. Shape Index textbox
    shape_index_frame = tk.Frame(main_frame)
    shape_index_frame.pack(fill=tk.X, pady=(0, 10))
    tk.Label(shape_index_frame, text="Shape ID:", font=("Arial", 12, "bold")).pack(side=tk.LEFT, padx=(0, 15))
    shape_id_entry = tk.Entry(shape_index_frame, font=("Arial", 12), width=25)
    shape_id_entry.pack(side=tk.LEFT)
    
    # Shape text preview (below shape ID)
    shape_text_frame = tk.Frame(main_frame)
    shape_text_frame.pack(fill=tk.X, pady=(0, 20))
    tk.Label(shape_text_frame, text="Shape Text:", font=("Arial", 11, "bold")).pack(side=tk.LEFT, padx=(0, 15))
    shape_text_label = tk.Label(shape_text_frame, text="", font=("Arial", 11), fg="gray", anchor="w", justify="left")
    shape_text_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    # Store for text section rows (needed by functions below)
    text_section_rows = []
    
    # 2. Break by panel
    break_by_frame = tk.LabelFrame(main_frame, text="Break By", font=("Arial", 13, "bold"))
    break_by_frame.pack(fill=tk.X, pady=(0, 20))
    
    break_by_inner = tk.Frame(break_by_frame, padx=15, pady=15)
    break_by_inner.pack(fill=tk.X)
    
    # Checkboxes for break by options
    linebreaks_var = tk.BooleanVar(master=text_dialog, value=False)
    format_changes_var = tk.BooleanVar(master=text_dialog, value=False)
    custom_char_var = tk.BooleanVar(master=text_dialog, value=False)
    
    # Custom char entry (create early so it can be referenced)
    custom_char_frame = tk.Frame(break_by_inner)
    custom_char_entry = tk.Entry(custom_char_frame, font=("Arial", 11), width=8, state='disabled')
    
    # 3. Text sections panel
    text_sections_frame = tk.LabelFrame(main_frame, text="Text Sections", font=("Arial", 13, "bold"))
    text_sections_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
    
    # Table frame with scrollbar - minimize padding to maximize width
    table_container = tk.Frame(text_sections_frame)
    table_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=10)
    
    table_canvas = tk.Canvas(table_container, highlightthickness=0)
    table_scrollbar = tk.Scrollbar(table_container, orient="vertical", command=table_canvas.yview)
    table_frame = tk.Frame(table_canvas)
    
    # Bind to update scroll region
    def on_frame_configure(event):
        table_canvas.configure(scrollregion=table_canvas.bbox("all"))
    
    table_frame.bind("<Configure>", on_frame_configure)
    
    # Create window that expands to canvas width
    canvas_window = table_canvas.create_window((0, 0), window=table_frame, anchor="nw")
    
    # Bind canvas resize to update frame width
    def on_canvas_configure(event):
        # Make the frame fill the canvas width
        table_canvas.itemconfig(canvas_window, width=event.width)
    
    table_canvas.bind("<Configure>", on_canvas_configure)
    table_canvas.configure(yscrollcommand=table_scrollbar.set)
    
    table_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    table_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    
    # Table headers - using grid for better width control
    headers_frame = tk.Frame(table_frame)
    headers_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
    
    tk.Label(headers_frame, text="Start Char", font=("Arial", 10, "bold"), width=12, anchor="w").grid(row=0, column=0, padx=5, sticky="ew")
    tk.Label(headers_frame, text="End Char", font=("Arial", 10, "bold"), width=12, anchor="w").grid(row=0, column=1, padx=5, sticky="ew")
    tk.Label(headers_frame, text="Text", font=("Arial", 10, "bold"), anchor="w").grid(row=0, column=2, padx=5, sticky="ew")
    tk.Label(headers_frame, text="Name", font=("Arial", 10, "bold"), anchor="w").grid(row=0, column=3, padx=5, sticky="ew")
    tk.Label(headers_frame, text="Action", font=("Arial", 10, "bold"), width=10, anchor="w").grid(row=0, column=4, padx=5, sticky="ew")
    
    # Configure column weights to make them expand to fill available width
    headers_frame.columnconfigure(0, minsize=100)  # Start char minimum
    headers_frame.columnconfigure(1, minsize=100)  # End char minimum
    headers_frame.columnconfigure(2, weight=4, minsize=400)  # Text column - largest
    headers_frame.columnconfigure(3, weight=3, minsize=300)  # Name column - large
    headers_frame.columnconfigure(4, minsize=100)  # Action minimum
    
    def format_text_with_ellipsis(text, max_len=50):
        """Format text with ellipsis in the middle if too long"""
        if len(text) <= max_len:
            return text
        # Show beginning and end with [...] in middle
        chars_per_side = (max_len - 5) // 2  # 5 chars for " [...] "
        return text[:chars_per_side] + " [...] " + text[-chars_per_side:]
    
    def add_text_section_row(start_val="", end_val="", text_val="", name_val=""):
        """Add a new row to the text sections table"""
        row_frame = tk.Frame(table_frame)
        row_frame.pack(fill=tk.BOTH, expand=True, pady=3)
        
        # Start char entry (int only)
        start_entry = tk.Entry(row_frame, font=("Arial", 10), width=12)
        start_entry.grid(row=0, column=0, padx=5, sticky="ew")
        if start_val:
            start_entry.insert(0, str(start_val))
        
        # End char entry (int only)
        end_entry = tk.Entry(row_frame, font=("Arial", 10), width=12)
        end_entry.grid(row=0, column=1, padx=5, sticky="ew")
        if end_val:
            end_entry.insert(0, str(end_val))
        
        # Text label (capped to 50 chars with ellipsis in middle)
        text_display = format_text_with_ellipsis(text_val, 50)
        text_label = tk.Label(row_frame, text=text_display, font=("Arial", 10), anchor="w", bg="white", relief="sunken")
        text_label.grid(row=0, column=2, padx=5, sticky="ew")
        
        # Name entry (pre-populated with text, capped at 50 chars)
        name_entry = tk.Entry(row_frame, font=("Arial", 10))
        name_entry.grid(row=0, column=3, padx=5, sticky="ew")
        name_text = name_val if name_val else text_val
        if len(name_text) > 50:
            name_text = name_text[:50]
        name_entry.insert(0, name_text)
        
        # Remove button
        def remove_row():
            row_frame.destroy()
            text_section_rows.remove(row_data)
        
        remove_btn = tk.Button(row_frame, text="Remove", font=("Arial", 9), command=remove_row, width=10)
        remove_btn.grid(row=0, column=4, padx=5, sticky="ew")
        
        # Configure column weights to match headers
        row_frame.columnconfigure(0, minsize=100)  # Start char minimum
        row_frame.columnconfigure(1, minsize=100)  # End char minimum
        row_frame.columnconfigure(2, weight=4, minsize=400)  # Text column - largest
        row_frame.columnconfigure(3, weight=3, minsize=300)  # Name column - large
        row_frame.columnconfigure(4, minsize=100)  # Action minimum
        
        # Store row data
        row_data = {
            'frame': row_frame,
            'start_entry': start_entry,
            'end_entry': end_entry,
            'text_label': text_label,
            'name_entry': name_entry,
            'text_val': text_val
        }
        text_section_rows.append(row_data)
        
        # Update text display when start/end change
        def update_text_display(*args):
            try:
                shape_id = shape_id_entry.get().strip()
                start = start_entry.get().strip()
                end = end_entry.get().strip()
                
                if shape_id and start and end:
                    start_int = int(start)
                    end_int = int(end)
                    
                    # Find shape and get text
                    for shape in ppt_app.ActiveWindow.View.Slide.Shapes:
                        if str(shape.Id) == shape_id:
                            if hasattr(shape, 'TextFrame') and shape.HasTextFrame:
                                full_text = shape.TextFrame.TextRange.Text
                                if start_int >= 0 and end_int <= len(full_text) and start_int < end_int:
                                    text_segment = full_text[start_int:end_int]
                                    row_data['text_val'] = text_segment
                                    display_text = format_text_with_ellipsis(text_segment, 50)
                                    text_label.config(text=display_text)
                                    # Update name if it's empty
                                    if not name_entry.get().strip():
                                        name_entry.delete(0, tk.END)
                                        name_text = text_segment[:50] if len(text_segment) > 50 else text_segment
                                        name_entry.insert(0, name_text)
                            break
            except:
                pass
        
        start_entry.bind('<KeyRelease>', update_text_display)
        end_entry.bind('<KeyRelease>', update_text_display)
        
        return row_data
    
    # Add section button
    add_section_btn = tk.Button(text_sections_frame, text="Add Section", font=("Arial", 11, "bold"), 
                                command=lambda: add_text_section_row(), bg="#4CAF50", fg="white",
                                width=15, height=2)
    add_section_btn.pack(pady=(10, 15))
    
    def update_shape_text_preview():
        """Update the shape text preview label"""
        shape_id = shape_id_entry.get().strip()
        if not shape_id:
            shape_text_label.config(text="Enter a Shape ID to see text preview", fg="gray")
            return
        
        # Find the shape
        target_shape = None
        for shape in ppt_app.ActiveWindow.View.Slide.Shapes:
            if str(shape.Id) == shape_id:
                target_shape = shape
                break
        
        if not target_shape or not hasattr(target_shape, 'TextFrame') or not target_shape.HasTextFrame:
            shape_text_label.config(text="Shape not found or has no text", fg="red")
            return
        
        full_text = target_shape.TextFrame.TextRange.Text
        display_text = format_text_with_ellipsis(full_text, 50)
        shape_text_label.config(text=display_text, fg="black")
    
    # Define populate_sections_by_breaks with actual implementation
    def populate_sections_by_breaks():
        """Auto-populate text sections based on selected break options"""
        shape_id = shape_id_entry.get().strip()
        if not shape_id:
            # Clear existing rows
            for row in text_section_rows[:]:
                row['frame'].destroy()
            text_section_rows.clear()
            return
        
        # Find the shape
        target_shape = None
        for shape in ppt_app.ActiveWindow.View.Slide.Shapes:
            if str(shape.Id) == shape_id:
                target_shape = shape
                break
        
        if not target_shape or not hasattr(target_shape, 'TextFrame') or not target_shape.HasTextFrame:
            # Clear existing rows
            for row in text_section_rows[:]:
                row['frame'].destroy()
            text_section_rows.clear()
            return
        
        full_text = target_shape.TextFrame.TextRange.Text
        
        # Clear existing rows ALWAYS when this function is called
        for row in text_section_rows[:]:
            row['frame'].destroy()
        text_section_rows.clear()
        
        # Check if any break option is selected
        if not (linebreaks_var.get() or custom_char_var.get() or format_changes_var.get()):
            # No break options selected, don't populate
            return
        
        # Determine break points
        break_points = [0]
        
        if linebreaks_var.get():
            # Break by linebreaks (\n, \r, \r\n)
            for i, char in enumerate(full_text):
                if char in ['\n', '\r']:
                    if i + 1 not in break_points:
                        break_points.append(i + 1)
        
        if custom_char_var.get():
            custom_char = custom_char_entry.get()
            if custom_char:
                for i, char in enumerate(full_text):
                    if char == custom_char:
                        if i + 1 not in break_points:
                            break_points.append(i + 1)
        
        if format_changes_var.get():
            # Break by format changes (requires checking TextRange.Runs)
            try:
                runs = target_shape.TextFrame.TextRange.Runs()
                current_pos = 0
                for i in range(1, runs.Count + 1):
                    run = runs(i)
                    run_length = len(run.Text)
                    if current_pos > 0 and current_pos not in break_points:
                        break_points.append(current_pos)
                    current_pos += run_length
            except:
                pass
        
        # Sort and deduplicate break points
        break_points = sorted(set(break_points))
        if len(full_text) not in break_points:
            break_points.append(len(full_text))
        
        # Create sections from break points
        for i in range(len(break_points) - 1):
            start = break_points[i]
            end = break_points[i + 1]
            text_segment = full_text[start:end]
            if text_segment.strip():  # Only add non-empty sections
                add_text_section_row(start, end, text_segment, text_segment.strip())
    
    # Now create the checkbox command functions and wire them up
    def on_linebreaks_toggle():
        populate_sections_by_breaks()
    
    def on_format_changes_toggle():
        populate_sections_by_breaks()
    
    def on_custom_char_toggle():
        if custom_char_var.get():
            custom_char_entry.config(state='normal')
        else:
            custom_char_entry.config(state='disabled')
        populate_sections_by_breaks()
    
    # Create and pack the checkboxes
    linebreaks_check = tk.Checkbutton(break_by_inner, text="Linebreaks", variable=linebreaks_var, 
                                     font=("Arial", 11), command=on_linebreaks_toggle)
    linebreaks_check.pack(anchor="w", pady=5)
    
    format_changes_check = tk.Checkbutton(break_by_inner, text="Format changes", variable=format_changes_var, 
                                         font=("Arial", 11), command=on_format_changes_toggle)
    format_changes_check.pack(anchor="w", pady=5)
    
    # Pack custom char frame and checkbox
    custom_char_frame.pack(anchor="w", pady=5)
    custom_char_check = tk.Checkbutton(custom_char_frame, text="Custom character:", variable=custom_char_var, 
                                      font=("Arial", 11), command=on_custom_char_toggle)
    custom_char_check.pack(side=tk.LEFT)
    custom_char_entry.pack(side=tk.LEFT, padx=(10, 0))
    
    # Bind shape ID changes to update preview and populate
    def on_shape_id_change(event):
        update_shape_text_preview()
        populate_sections_by_breaks()
    
    shape_id_entry.bind('<FocusOut>', on_shape_id_change)
    
    # Bind custom char entry to trigger populate when typing
    custom_char_entry.bind('<KeyRelease>', lambda e: populate_sections_by_breaks())
    
    # 4. Buttons
    buttons_frame = tk.Frame(main_frame)
    buttons_frame.pack(fill=tk.X, pady=(10, 0))
    
    def go_to_previous_section():
        """Go back to individual shapes"""
        text_dialog.destroy()
        show_individual_shapes_form()
    
    def save_and_continue():
        """Save current sections and stay on form"""
        shape_id = shape_id_entry.get().strip()
        if not shape_id:
            print("No shape ID specified")
            return
        
        # Save all sections
        for row in text_section_rows:
            try:
                start = int(row['start_entry'].get().strip())
                end = int(row['end_entry'].get().strip())
                name = row['name_entry'].get().strip()
                text = row['text_val']
                
                if start >= 0 and end > start and name:
                    text_section_list.append({
                        'shape_id': shape_id,
                        'start_char': start,
                        'end_char': end,
                        'label': name,
                        'text': text
                    })
                    print(f"Saved text section: {name} ({start}-{end})")
            except:
                pass
        
        # Clear form for next shape
        shape_id_entry.delete(0, tk.END)
        linebreaks_var.set(False)
        format_changes_var.set(False)
        custom_char_var.set(False)
        custom_char_entry.delete(0, tk.END)
        for row in text_section_rows[:]:
            row['frame'].destroy()
        text_section_rows.clear()
        shape_id_entry.focus()
    
    def go_to_next_section():
        """Save and go to tables"""
        # Save current sections first
        shape_id = shape_id_entry.get().strip()
        if shape_id:
            for row in text_section_rows:
                try:
                    start = int(row['start_entry'].get().strip())
                    end = int(row['end_entry'].get().strip())
                    name = row['name_entry'].get().strip()
                    text = row['text_val']
                    
                    if start >= 0 and end > start and name:
                        text_section_list.append({
                            'shape_id': shape_id,
                            'start_char': start,
                            'end_char': end,
                            'label': name,
                            'text': text
                        })
                        print(f"Saved text section: {name} ({start}-{end})")
                except:
                    pass
        
        text_dialog.destroy()
        show_table_labeling_form()
    
    prev_btn = tk.Button(buttons_frame, text="Go to Previous Section", command=go_to_previous_section,
                        bg="#FF9800", fg="white", font=("Arial", 11, "bold"), width=22, height=2)
    prev_btn.pack(side=tk.LEFT, padx=(0, 15))
    
    save_btn = tk.Button(buttons_frame, text="Save", command=save_and_continue,
                        bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), width=22, height=2)
    save_btn.pack(side=tk.LEFT, padx=(0, 15))
    
    next_btn = tk.Button(buttons_frame, text="Go to Next Section", command=go_to_next_section,
                        bg="#2196F3", fg="white", font=("Arial", 11, "bold"), width=22, height=2)
    next_btn.pack(side=tk.LEFT)
    
    # Pack canvas and scrollbar
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    # Focus on shape ID entry
    text_dialog.after(100, lambda: shape_id_entry.focus())
    
    # Handle window close (X button) - save and quit
    def on_text_closing():
        # Save current sections
        shape_id = shape_id_entry.get().strip()
        if shape_id:
            for row in text_section_rows:
                try:
                    start = int(row['start_entry'].get().strip())
                    end = int(row['end_entry'].get().strip())
                    name = row['name_entry'].get().strip()
                    text = row['text_val']
                    
                    if start >= 0 and end > start and name:
                        text_section_list.append({
                            'shape_id': shape_id,
                            'start_char': start,
                            'end_char': end,
                            'label': name,
                            'text': text
                        })
                except:
                    pass
        
        text_dialog.destroy()
        finish_text_labeling_process()
    
    text_dialog.protocol("WM_DELETE_WINDOW", on_text_closing)

def get_text_dict(shape_id, start_char, end_char, name):
    for x in ppt_app.ActiveWindow.View.Slide.Shapes:
        if str(x.ID) == shape_id:
            if x.HasTextFrame:
                text = x.TextFrame.TextRange.Text[int(start_char):int(end_char)]
                return {'shape_id': shape_id, 'start_char': int(start_char), 'end_char': int(end_char), 'label': name, 'text': text}
    return {'shape_id': shape_id, 'start_char': int(start_char), 'end_char': int(end_char), 'label': name, 'text': ''}

def save_and_next_text(shape_id, start_char, end_char, name, dialog):
    """Save text section and continue to next"""
    # TODO: Implement save and next functionality
    text_section_list.append(get_text_dict(shape_id, start_char, end_char, name))
    print(f"Save & Next: Shape ID={shape_id}, Start={start_char}, End={end_char}, Name={name}")
    
    # Find the input fields and update them
    for widget in dialog.winfo_children():
        if isinstance(widget, tk.Frame):
            for child in widget.winfo_children():
                if isinstance(child, tk.Frame):
                    entries = []
                    for entry in child.winfo_children():
                        if isinstance(entry, tk.Entry):
                            entries.append(entry)
                    
                    if len(entries) >= 4:  # shape_id, start_char, end_char, name
                        # Keep shape_id the same (entries[0])
                        # Set start_char to previous end_char (entries[1])
                        entries[1].delete(0, tk.END)
                        entries[1].insert(0, end_char)
                        # Clear end_char (entries[2])
                        entries[2].delete(0, tk.END)
                        # Clear name (entries[3])
                        entries[3].delete(0, tk.END)
                        # Focus on end_char field
                        entries[2].focus()
                        break

def go_to_previous_section_from_text(shape_id_entry, start_char_entry, end_char_entry, name_entry, dialog):
    """Go back to individual shapes section"""
    # Check if there's unsaved content
    shape_id = shape_id_entry.get().strip()
    start_char = start_char_entry.get().strip()
    end_char = end_char_entry.get().strip()
    name = name_entry.get().strip()
    
    if shape_id or start_char or end_char or name:
        from tkinter import messagebox
        result = messagebox.askyesno("Unsaved Changes", 
                                    "You have unsaved changes.\n\n" +
                                    "Do you want to continue to the previous section without saving?")
        if not result:  # User chose No - stay on page
            return
    
    dialog.destroy()
    show_individual_shapes_form()

def go_to_tables_with_confirmation(shape_id_entry, start_char_entry, end_char_entry, name_entry, dialog):
    """Go to tables section with save confirmation"""
    # Check if there's unsaved content
    shape_id = shape_id_entry.get().strip()
    start_char = start_char_entry.get().strip()
    end_char = end_char_entry.get().strip()
    name = name_entry.get().strip()
    
    if shape_id and start_char and end_char and name:
        from tkinter import messagebox
        result = messagebox.askyesnocancel("Save Changes?", 
                                          "Do you want to save changes before going to the next section?\n\n" +
                                          "Yes = Save and continue\n" +
                                          "No = Don't save and continue\n" +
                                          "Cancel = Stay on this page")
        if result is None:  # Cancel
            return
        elif result:  # Yes - save
            go_to_tables(shape_id, start_char, end_char, name, dialog)
            return
        # else: No - don't save, just continue
    
    # No unsaved content or user chose not to save
    dialog.destroy()
    show_table_labeling_form()

def save_and_finish_text(shape_id, start_char, end_char, name, dialog):
    """Save text section and finish"""
    # TODO: Implement save and finish functionality
    text_section_list.append(get_text_dict(shape_id, start_char, end_char, name))
    print(f"Save & Finish: Shape ID={shape_id}, Start={start_char}, End={end_char}, Name={name}")
    dialog.destroy()
    finish_text_labeling_process()

def go_to_next_text_section(shape_id, start_char, end_char, name, dialog):
    """Go to next text section, saving current section if shape ID is provided"""
    if shape_id and shape_id.strip():
        text_section_list.append(get_text_dict(shape_id, start_char, end_char, name))
        print(f"Saved text section: Shape ID={shape_id}, Start={start_char}, End={end_char}, Name={name}")
        
        # Find the input fields and update them for next section
        for widget in dialog.winfo_children():
            if isinstance(widget, tk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, tk.Frame):
                        entries = []
                        for entry in child.winfo_children():
                            if isinstance(entry, tk.Entry):
                                entries.append(entry)
                        
                        if len(entries) >= 4:  # shape_id, start_char, end_char, name
                            # Keep shape_id the same (entries[0])
                            # Set start_char to previous end_char (entries[1])
                            entries[1].delete(0, tk.END)
                            entries[1].insert(0, end_char)
                            # Clear end_char (entries[2])
                            entries[2].delete(0, tk.END)
                            # Clear name (entries[3])
                            entries[3].delete(0, tk.END)
                            # Focus on end_char field
                            entries[2].focus()
                            break
    else:
        print("Shape ID is blank - not saving text section")
        # Clear all fields for next section
        for widget in dialog.winfo_children():
            if isinstance(widget, tk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, tk.Frame):
                        entries = []
                        for entry in child.winfo_children():
                            if isinstance(entry, tk.Entry):
                                entries.append(entry)
                        
                        if len(entries) >= 4:  # shape_id, start_char, end_char, name
                            # Clear all fields
                            for entry in entries:
                                entry.delete(0, tk.END)
                            # Focus on shape_id field
                            entries[0].focus()
                            break

def go_to_tables(shape_id, start_char, end_char, name, dialog):
    """Go to tables, saving current text section if shape ID is provided"""
    if shape_id and shape_id.strip():
        text_section_list.append(get_text_dict(shape_id, start_char, end_char, name))
        print(f"Saved text section: Shape ID={shape_id}, Start={start_char}, End={end_char}, Name={name}")
    else:
        print("Shape ID is blank - not saving text section")
    
    dialog.destroy()
    show_table_labeling_form()

def finish_text_labeling(dialog):
    """Finish text labeling without saving"""
    dialog.destroy()
    show_table_labeling_form()

def get_tables_from_slide():
    """Get all tables from the current slide"""
    tables = []
    try:
        if ppt_app and ppt_app.ActiveWindow and ppt_app.ActiveWindow.View.Slide:
            slide = ppt_app.ActiveWindow.View.Slide
            for shape in slide.Shapes:
                if shape.HasTable:
                    tables.append(shape)
        return tables
    except Exception as e:
        print(f"Error getting tables: {e}")
        return []

def show_table_labeling_form():
    """Show table labeling form for table rows/columns"""
    # Get all tables from the slide
    tables = get_tables_from_slide()
    
    if not tables:
        messagebox.showinfo("No Tables", "No tables found in the current slide.")
        finish_text_labeling_process()
        return
    
    # Start with the first table
    show_table_form_with_data(tables, 0)

def show_table_form_with_data(tables, table_index):
    """Show table labeling form with actual table data"""
    if table_index >= len(tables):
        # No more tables, finish the process
        finish_text_labeling_process()
        return
    
    current_table = tables[table_index]
    
    # Create new table labeling dialog
    table_dialog = tk.Tk()
    table_dialog.title(f"Label Table Rows / Cols ({table_index + 1}/{len(tables)})")
    table_dialog.geometry("1000x800")
    table_dialog.resizable(True, True)
    
    # Center on screen
    table_dialog.update_idletasks()
    x = (table_dialog.winfo_screenwidth() // 2) - (1000 // 2)
    y = (table_dialog.winfo_screenheight() // 2) - (800 // 2)
    table_dialog.geometry(f"1000x800+{x}+{y}")
    
    # Bring dialog to foreground
    table_dialog.lift()
    table_dialog.attributes('-topmost', True)
    table_dialog.after_idle(lambda: table_dialog.attributes('-topmost', False))
    table_dialog.focus_force()
    
    # Create main scrollable frame with regular scrollbar
    canvas = tk.Canvas(table_dialog)
    scrollbar = tk.Scrollbar(table_dialog, orient="vertical", command=canvas.yview)
    scrollable_frame = tk.Frame(canvas)
    
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Main frame
    main_frame = tk.Frame(scrollable_frame, padx=20, pady=20)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Title with actual table ID
    title_label = tk.Label(main_frame, text=f"Table #{current_table.Id}", 
                          font=("Arial", 16, "bold"))
    title_label.pack(pady=(0, 20))
    
    # Build cell coordinates map for this table (needed for overlay counts)
    cell_coords = get_cell_coordinates_map(current_table)
    
    # Row labels section
    row_frame = tk.LabelFrame(main_frame, text="Row Labels", font=("Arial", 12, "bold"))
    row_frame.pack(fill=tk.X, pady=(0, 20))
    
    # Save row labels checkbox
    save_row_labels_var = tk.BooleanVar(value=True)
    save_row_labels_checkbox = tk.Checkbutton(row_frame, text="Save row labels", 
                                             variable=save_row_labels_var,
                                             font=("Arial", 10))
    save_row_labels_checkbox.pack(anchor="w", padx=10, pady=(5, 5))
    
    # Add overlaid shapes checkbox and tolerance for rows
    row_overlay_frame = tk.Frame(row_frame)
    row_overlay_frame.pack(anchor="w", padx=10, pady=(0, 10))
    
    add_row_overlaid_shapes_var = tk.BooleanVar(master=table_dialog, value=False)
    add_row_overlaid_shapes_checkbox = tk.Checkbutton(row_overlay_frame, text="Add overlaid shapes", 
                                                      variable=add_row_overlaid_shapes_var,
                                                      font=("Arial", 10))
    add_row_overlaid_shapes_checkbox.pack(side=tk.LEFT)
    
    tk.Label(row_overlay_frame, text="  Tolerance for overlay:", font=("Arial", 10)).pack(side=tk.LEFT, padx=(10, 5))
    row_tolerance_entry = tk.Entry(row_overlay_frame, font=("Arial", 10), width=5, state='disabled')
    row_tolerance_entry.insert(0, "5")
    row_tolerance_entry.pack(side=tk.LEFT)
    
    def toggle_row_tolerance():
        if add_row_overlaid_shapes_var.get():
            row_tolerance_entry.config(state='normal')
            update_row_overlay_counts()
        else:
            row_tolerance_entry.config(state='disabled')
            update_row_overlay_counts()
    
    add_row_overlaid_shapes_checkbox.config(command=toggle_row_tolerance)
    
    # Row labels table
    row_table_frame = tk.Frame(row_frame)
    row_table_frame.pack(fill=tk.X, padx=10, pady=10)
    
    # Row table headers
    tk.Label(row_table_frame, text="Row Index", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    tk.Label(row_table_frame, text="Row Text (50 chars max)", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=5, sticky="w")
    tk.Label(row_table_frame, text="Row Name", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, pady=5, sticky="w")
    tk.Label(row_table_frame, text="# Overlaid", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5, pady=5, sticky="w")
    
    # Populate with actual row data
    row_entries = []
    row_overlay_labels = []  # Store labels for overlay counts
    for i in range(current_table.Table.Rows.Count):
        # Row index
        tk.Label(row_table_frame, text=f"{i}", font=("Arial", 9)).grid(row=i+1, column=0, padx=5, pady=2, sticky="w")
        
        # Get first non-merged cell text from each row
        # Check each cell's position to detect merging
        first_cell_text = ""
        row_name_text = f"Row {i}"
        try:
            for j in range(current_table.Table.Columns.Count):
                try:
                    cell = current_table.Table.Cell(i + 1, j + 1)
                    cell_text = cell.Shape.TextFrame.TextRange.Text.strip()
                    
                    if not cell_text:
                        continue
                    
                    # Check if merged vertically by comparing top position with adjacent cells
                    is_merged_vertically = False
                    cell_top = cell.Shape.Top
                    
                    # Check cell above (if exists)
                    if i > 0:
                        try:
                            cell_above = current_table.Table.Cell(i, j + 1)
                            if abs(cell_above.Shape.Top - cell_top) < 1:  # Same top = merged
                                is_merged_vertically = True
                        except:
                            pass
                    
                    # Check cell below (if exists)
                    if not is_merged_vertically and i < current_table.Table.Rows.Count - 1:
                        try:
                            cell_below = current_table.Table.Cell(i + 2, j + 1)
                            if abs(cell_below.Shape.Top - cell_top) < 1:  # Same top = merged
                                is_merged_vertically = True
                        except:
                            pass
                    
                    # Use this cell if not merged vertically
                    if not is_merged_vertically:
                        first_cell_text = cell_text
                        row_name_text = cell_text
                        break
                except:
                    continue
            
            # Truncate to 50 characters
            if len(first_cell_text) > 50:
                first_cell_text = first_cell_text[:47] + "..."
        except:
            first_cell_text = f"Row {i}"
        
        # Row text
        tk.Label(row_table_frame, text=first_cell_text, font=("Arial", 9), width=30, anchor="w").grid(row=i+1, column=1, padx=5, pady=2, sticky="w")
        
        # Row name entry
        row_name_entry = tk.Entry(row_table_frame, font=("Arial", 9), width=20)
        row_name_entry.grid(row=i+1, column=2, padx=5, pady=2, sticky="w")
        row_name_entry.insert(0, row_name_text)
        row_entries.append(row_name_entry)
        
        # Overlay count label
        overlay_count_label = tk.Label(row_table_frame, text="0", font=("Arial", 9), width=8, anchor="center")
        overlay_count_label.grid(row=i+1, column=3, padx=5, pady=2, sticky="w")
        row_overlay_labels.append(overlay_count_label)
    
    # Function to update row overlay counts
    def update_row_overlay_counts():
        if not add_row_overlaid_shapes_var.get():
            # Clear counts when disabled
            for label in row_overlay_labels:
                label.config(text="0")
            return
        
        try:
            tolerance = float(row_tolerance_entry.get())
        except:
            tolerance = 5
        
        num_cols = current_table.Table.Columns.Count
        for i in range(current_table.Table.Rows.Count):
            cells = [f"{i}.{j}" for j in range(num_cols)]
            bounds = calculate_section_bounds(cells, cell_coords)
            overlaid_shapes = get_overlaid_shapes_in_bounds(bounds, tolerance)
            row_overlay_labels[i].config(text=str(len(overlaid_shapes)))
    
    # Bind tolerance entry to update counts on change
    row_tolerance_entry.bind('<KeyRelease>', lambda e: update_row_overlay_counts())
    
    # Column labels section
    col_frame = tk.LabelFrame(main_frame, text="Column Labels", font=("Arial", 12, "bold"))
    col_frame.pack(fill=tk.X, pady=(0, 20))
    
    # Save column labels checkbox
    save_col_labels_var = tk.BooleanVar(value=True)
    save_col_labels_checkbox = tk.Checkbutton(col_frame, text="Save column labels", 
                                             variable=save_col_labels_var,
                                             font=("Arial", 10))
    save_col_labels_checkbox.pack(anchor="w", padx=10, pady=(5, 5))
    
    # Add overlaid shapes checkbox and tolerance for columns
    col_overlay_frame = tk.Frame(col_frame)
    col_overlay_frame.pack(anchor="w", padx=10, pady=(0, 10))
    
    add_col_overlaid_shapes_var = tk.BooleanVar(master=table_dialog, value=False)
    add_col_overlaid_shapes_checkbox = tk.Checkbutton(col_overlay_frame, text="Add overlaid shapes", 
                                                      variable=add_col_overlaid_shapes_var,
                                                      font=("Arial", 10))
    add_col_overlaid_shapes_checkbox.pack(side=tk.LEFT)
    
    tk.Label(col_overlay_frame, text="  Tolerance for overlay:", font=("Arial", 10)).pack(side=tk.LEFT, padx=(10, 5))
    col_tolerance_entry = tk.Entry(col_overlay_frame, font=("Arial", 10), width=5, state='disabled')
    col_tolerance_entry.insert(0, "5")
    col_tolerance_entry.pack(side=tk.LEFT)
    
    def toggle_col_tolerance():
        if add_col_overlaid_shapes_var.get():
            col_tolerance_entry.config(state='normal')
            update_col_overlay_counts()
        else:
            col_tolerance_entry.config(state='disabled')
            update_col_overlay_counts()
    
    add_col_overlaid_shapes_checkbox.config(command=toggle_col_tolerance)
    
    # Column labels table
    col_table_frame = tk.Frame(col_frame)
    col_table_frame.pack(fill=tk.X, padx=10, pady=10)
    
    # Column table headers
    tk.Label(col_table_frame, text="Col Index", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    tk.Label(col_table_frame, text="Col Text (50 chars max)", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=5, sticky="w")
    tk.Label(col_table_frame, text="Col Name", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, pady=5, sticky="w")
    tk.Label(col_table_frame, text="# Overlaid", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5, pady=5, sticky="w")
    
    # Populate with actual column data
    col_entries = []
    col_overlay_labels = []  # Store labels for overlay counts
    for i in range(current_table.Table.Columns.Count):
        # Column index
        tk.Label(col_table_frame, text=f"{i}", font=("Arial", 9)).grid(row=i+1, column=0, padx=5, pady=2, sticky="w")
        
        # Get first non-merged cell text from each column
        # Check each cell's position to detect merging
        first_cell_text = ""
        col_name_text = f"Col {i}"
        try:
            for j in range(current_table.Table.Rows.Count):
                try:
                    cell = current_table.Table.Cell(j + 1, i + 1)
                    cell_text = cell.Shape.TextFrame.TextRange.Text.strip()
                    
                    if not cell_text:
                        continue
                    
                    # Check if merged horizontally by comparing left position with adjacent cells
                    is_merged_horizontally = False
                    cell_left = cell.Shape.Left
                    
                    # Check cell to the left (if exists)
                    if i > 0:
                        try:
                            cell_left_neighbor = current_table.Table.Cell(j + 1, i)
                            if abs(cell_left_neighbor.Shape.Left - cell_left) < 1:  # Same left = merged
                                is_merged_horizontally = True
                        except:
                            pass
                    
                    # Check cell to the right (if exists)
                    if not is_merged_horizontally and i < current_table.Table.Columns.Count - 1:
                        try:
                            cell_right_neighbor = current_table.Table.Cell(j + 1, i + 2)
                            if abs(cell_right_neighbor.Shape.Left - cell_left) < 1:  # Same left = merged
                                is_merged_horizontally = True
                        except:
                            pass
                    
                    # Use this cell if not merged horizontally
                    if not is_merged_horizontally:
                        first_cell_text = cell_text
                        col_name_text = cell_text
                        break
                except:
                    continue
            
            # Truncate to 50 characters
            if len(first_cell_text) > 50:
                first_cell_text = first_cell_text[:47] + "..."
        except:
            first_cell_text = f"Col {i}"
        
        # Column text
        tk.Label(col_table_frame, text=first_cell_text, font=("Arial", 9), width=30, anchor="w").grid(row=i+1, column=1, padx=5, pady=2, sticky="w")
        
        # Column name entry
        col_name_entry = tk.Entry(col_table_frame, font=("Arial", 9), width=20)
        col_name_entry.grid(row=i+1, column=2, padx=5, pady=2, sticky="w")
        col_name_entry.insert(0, col_name_text)
        col_entries.append(col_name_entry)
        
        # Overlay count label
        overlay_count_label = tk.Label(col_table_frame, text="0", font=("Arial", 9), width=8, anchor="center")
        overlay_count_label.grid(row=i+1, column=3, padx=5, pady=2, sticky="w")
        col_overlay_labels.append(overlay_count_label)
    
    # Function to update column overlay counts
    def update_col_overlay_counts():
        if not add_col_overlaid_shapes_var.get():
            # Clear counts when disabled
            for label in col_overlay_labels:
                label.config(text="0")
            return
        
        try:
            tolerance = float(col_tolerance_entry.get())
        except:
            tolerance = 5
        
        num_rows = current_table.Table.Rows.Count
        for j in range(current_table.Table.Columns.Count):
            cells = [f"{i}.{j}" for i in range(num_rows)]
            bounds = calculate_section_bounds(cells, cell_coords)
            overlaid_shapes = get_overlaid_shapes_in_bounds(bounds, tolerance)
            col_overlay_labels[j].config(text=str(len(overlaid_shapes)))
    
    # Bind tolerance entry to update counts on change
    col_tolerance_entry.bind('<KeyRelease>', lambda e: update_col_overlay_counts())
    
    # Custom groups section
    custom_frame = tk.LabelFrame(main_frame, text="Custom Groups", font=("Arial", 12, "bold"))
    custom_frame.pack(fill=tk.X, pady=(0, 20))
    
    # Note about indexing
    note_label = tk.Label(custom_frame, text="Note: Use index 0 for first row/column", 
                         font=("Arial", 9, "italic"), fg="gray")
    note_label.pack(anchor="w", padx=10, pady=(5, 0))
    
    # Custom groups table
    custom_table_frame = tk.Frame(custom_frame)
    custom_table_frame.pack(fill=tk.X, padx=10, pady=10)
    
    # Custom groups headers
    tk.Label(custom_table_frame, text="Rows/Cols", font=("Arial", 10, "bold")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    tk.Label(custom_table_frame, text="Start Index", font=("Arial", 10, "bold")).grid(row=0, column=1, padx=5, pady=5, sticky="w")
    tk.Label(custom_table_frame, text="End Index", font=("Arial", 10, "bold")).grid(row=0, column=2, padx=5, pady=5, sticky="w")
    tk.Label(custom_table_frame, text="Group Name", font=("Arial", 10, "bold")).grid(row=0, column=3, padx=5, pady=5, sticky="w")
    tk.Label(custom_table_frame, text="Action", font=("Arial", 10, "bold")).grid(row=0, column=4, padx=5, pady=5, sticky="w")
    
    # Store custom group rows for management (each item is a dict with frame and data)
    custom_group_rows = []
    
    def add_custom_group():
        """Add a new custom group row"""
        row_num = len(custom_group_rows) + 1
        
        # Create new row frame
        custom_row_frame = tk.Frame(custom_table_frame)
        custom_row_frame.grid(row=row_num, column=0, columnspan=5, sticky="ew", padx=5, pady=2)
        
        # Radio buttons for rows/cols
        row_col_var = tk.StringVar(master=table_dialog, value="rows")
        tk.Radiobutton(custom_row_frame, text="Rows", variable=row_col_var, value="rows").pack(side=tk.LEFT, padx=(0, 10))
        tk.Radiobutton(custom_row_frame, text="Cols", variable=row_col_var, value="cols").pack(side=tk.LEFT, padx=(0, 10))
        
        # Start and end index entries
        start_entry = tk.Entry(custom_row_frame, font=("Arial", 9), width=10)
        start_entry.pack(side=tk.LEFT, padx=(0, 10))
        start_entry.insert(0, "0")
        
        end_entry = tk.Entry(custom_row_frame, font=("Arial", 9), width=10)
        end_entry.pack(side=tk.LEFT, padx=(0, 10))
        end_entry.insert(0, "2")
        
        # Group name entry
        group_name_entry = tk.Entry(custom_row_frame, font=("Arial", 9), width=20)
        group_name_entry.pack(side=tk.LEFT, padx=(0, 10))
        group_name_entry.insert(0, f"Group {row_num}")
        
        # Remove button
        def remove_this_group():
            custom_row_frame.destroy()
            for group_data in custom_group_rows:
                if group_data['frame'] == custom_row_frame:
                    custom_group_rows.remove(group_data)
                    break
        
        remove_btn = tk.Button(custom_row_frame, text="Remove", font=("Arial", 9), 
                              bg="#F44336", fg="white", width=8, command=remove_this_group)
        remove_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Store the row frame and its data for management
        custom_group_rows.append({
            'frame': custom_row_frame,
            'row_col_var': row_col_var,
            'start_entry': start_entry,
            'end_entry': end_entry,
            'group_name_entry': group_name_entry
        })
    
    # Add button
    add_btn = tk.Button(custom_frame, text="Add Group", font=("Arial", 10), 
                       bg="#4CAF50", fg="white", width=12, command=add_custom_group)
    add_btn.pack(pady=10)
    
    # Pack canvas and scrollbar
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    # Buttons frame
    buttons_frame = tk.Frame(table_dialog)
    buttons_frame.pack(fill=tk.X, pady=10)
    
    # Left side buttons
    left_buttons_frame = tk.Frame(buttons_frame)
    left_buttons_frame.pack(side=tk.LEFT, padx=(20, 10))
    
    # Go to Previous Section button (only show on first table)
    if table_index == 0:
        prev_section_btn = tk.Button(left_buttons_frame, text="Go to Previous Section\n(Text Sections)", 
                                     command=lambda: go_to_previous_section_from_tables(table_dialog),
                                     bg="#FF9800", fg="white",
                                     width=18, height=2)
        prev_section_btn.pack(side=tk.LEFT, padx=(0, 5))
    
    # Navigation buttons
    nav_frame = tk.Frame(left_buttons_frame)
    nav_frame.pack(side=tk.LEFT, padx=(0, 10))
    
    # Previous button (only show if not first table)
    if table_index > 0:
        prev_btn = tk.Button(nav_frame, text="Previous", 
                            command=lambda: (table_dialog.destroy(), show_table_form_with_data(tables, table_index - 1)),
                            bg="#FF9800", fg="white",
                            width=10, height=2)
        prev_btn.pack(side=tk.LEFT, padx=(0, 5))
    
    # Next button (only show if not last table)
    if table_index < len(tables) - 1:
        next_btn = tk.Button(nav_frame, text="Next", 
                            command=lambda: (table_dialog.destroy(), show_table_form_with_data(tables, table_index + 1)),
                            bg="#2196F3", fg="white",
                            width=10, height=2)
        next_btn.pack(side=tk.LEFT, padx=(0, 5))
    
    # Save button (save and go to next table)
    save_btn = tk.Button(buttons_frame, text="Save", 
                         command=lambda: save_and_go_to_next_table(table_dialog, tables, table_index, current_table,
                                                                  save_row_labels_var, row_entries, add_row_overlaid_shapes_var, row_tolerance_entry,
                                                                  save_col_labels_var, col_entries, add_col_overlaid_shapes_var, col_tolerance_entry,
                                                                  custom_group_rows),
                         bg="#4CAF50", fg="white",
                         width=15, height=2)
    save_btn.pack(side=tk.RIGHT, padx=(10, 20))
    
    # Skip button (skip and go to next table)
    skip_btn = tk.Button(buttons_frame, text="Skip", 
                        command=lambda: skip_table_and_go_to_next(table_dialog, tables, table_index),
                        bg="#FF9800", fg="white",
                        width=15, height=2)
    skip_btn.pack(side=tk.RIGHT, padx=(0, 10))
    
    # Focus on Save button
    table_dialog.after(100, lambda: save_btn.focus())
    
    # Bind Enter key to Save
    table_dialog.bind('<Return>', lambda e: save_and_go_to_next_table(table_dialog, tables, table_index, current_table,
                                                                      save_row_labels_var, row_entries, add_row_overlaid_shapes_var, row_tolerance_entry,
                                                                      save_col_labels_var, col_entries, add_col_overlaid_shapes_var, col_tolerance_entry,
                                                                      custom_group_rows))
    
    # Bind Escape key to Skip
    table_dialog.bind('<Escape>', lambda e: skip_table_and_go_to_next(table_dialog, tables, table_index))
    
    # Handle window close (X button) - save and quit
    def on_table_closing():
        # Finish the table labeling process
        table_dialog.destroy()
        finish_text_labeling_process()
    
    table_dialog.protocol("WM_DELETE_WINDOW", on_table_closing)

def get_cell_coordinates_map(table_shape):
    """
    Build a map of cell coordinates for all cells in a table.
    Returns a dictionary where keys are "row.col" strings and values are coordinate dicts.
    """
    cell_coords = {}
    num_rows = table_shape.Table.Rows.Count
    num_cols = table_shape.Table.Columns.Count
    
    for i in range(num_rows):
        for j in range(num_cols):
            try:
                cell_shape = table_shape.Table.Cell(i+1, j+1).Shape
                top = cell_shape.Top
                left = cell_shape.Left
                width = cell_shape.Width
                height = cell_shape.Height
                bottom = top + height
                right = left + width
                
                cell_coords[f"{i}.{j}"] = {
                    'top': top,
                    'left': left,
                    'right': right,
                    'bottom': bottom,
                    'width': width,
                    'height': height
                }
            except:
                pass
    
    return cell_coords

def calculate_section_bounds(cells, cell_coords):
    """
    Calculate the bounding box for a section based on its cells.
    cells: list of cell coordinate strings like ["0.0", "0.1"]
    cell_coords: map from cell strings to coordinate dicts
    Returns: dict with top, left, right, bottom, width, height
    """
    if not cells:
        return {'top': 0, 'left': 0, 'right': 0, 'bottom': 0, 'width': 0, 'height': 0}
    
    # Get coordinates for all cells in this section
    valid_cells = [cell_coords[c] for c in cells if c in cell_coords]
    
    if not valid_cells:
        return {'top': 0, 'left': 0, 'right': 0, 'bottom': 0, 'width': 0, 'height': 0}
    
    # Calculate bounds
    left = min(c['left'] for c in valid_cells)
    right = max(c['right'] for c in valid_cells)
    top = min(c['top'] for c in valid_cells)
    bottom = max(c['bottom'] for c in valid_cells)
    width = right - left
    height = bottom - top
    
    return {
        'top': top,
        'left': left,
        'right': right,
        'bottom': bottom,
        'width': width,
        'height': height
    }

def get_overlaid_shapes_in_bounds(bounds, tolerance=0):
    """
    Get all shapes and groups that are fully contained within the given bounds (with tolerance).
    bounds: dict with 'top', 'left', 'right', 'bottom'
    tolerance: expand the bounds by this amount in all directions
    Returns: list of shape_ids that are fully contained
    """
    overlaid_shape_ids = []
    
    # Apply tolerance to bounds (expand bounds)
    expanded_top = bounds['top'] - tolerance
    expanded_left = bounds['left'] - tolerance
    expanded_right = bounds['right'] + tolerance
    expanded_bottom = bounds['bottom'] + tolerance
    
    # Check shape_data (individual shapes)
    for shape in shape_data:
        if (shape['top'] >= expanded_top and 
            shape['left'] >= expanded_left and 
            shape['right'] <= expanded_right and 
            shape['bottom'] <= expanded_bottom):
            overlaid_shape_ids.append(shape['shape_id'])
    
    # Check group_list (shape groups)
    for group in group_list:
        # Groups have their bounds already calculated
        if (group.get('top', 0) >= expanded_top and 
            group.get('left', 0) >= expanded_left and 
            group.get('right', 0) <= expanded_right and 
            group.get('bottom', 0) <= expanded_bottom):
            # For groups, get all shape_ids
            if isinstance(group.get('shape_id'), list):
                overlaid_shape_ids.extend(group['shape_id'])
            else:
                overlaid_shape_ids.append(group['shape_id'])
    
    return overlaid_shape_ids

def save_and_go_to_next_table(dialog, tables, current_index, current_table, 
                             save_row_labels_var, row_entries, add_row_overlaid_shapes_var, row_tolerance_entry,
                             save_col_labels_var, col_entries, add_col_overlaid_shapes_var, col_tolerance_entry,
                             custom_group_rows):
    """Save current table and go to next table"""
    table_shape_id = str(current_table.Id)
    
    # Build cell coordinates map for this table
    cell_coords = get_cell_coordinates_map(current_table)
    
    # 1. Save row labels if checked
    if save_row_labels_var.get():
        num_rows = current_table.Table.Rows.Count
        num_cols = current_table.Table.Columns.Count
        
        for i in range(num_rows):
            # Get text from all cells in the row
            row_text_parts = []
            for j in range(1, num_cols + 1):
                try:
                    cell_text = current_table.Table.Rows(i+1).Cells(j).Shape.TextFrame.TextRange.Text.strip()
                    row_text_parts.append(cell_text)
                except:
                    pass
            
            # Concatenate with spaces
            row_text = " ".join(row_text_parts)
            
            # Create cells array
            cells = [f"{i}.{j}" for j in range(num_cols)]
            
            # Get label from textbox
            label = row_entries[i].get() if i < len(row_entries) else f"Row {i}"
            
            # Calculate bounds for this row
            bounds = calculate_section_bounds(cells, cell_coords)
            
            # Get overlaid shapes if checkbox is checked
            overlaid_shapes = []
            if add_row_overlaid_shapes_var.get():
                try:
                    tolerance = float(row_tolerance_entry.get())
                except:
                    tolerance = 5
                overlaid_shapes = get_overlaid_shapes_in_bounds(bounds, tolerance)
            
            # Build shape_id array: cell IDs + overlaid shape IDs
            # Cell ID format: table_shape_id.row_index.col_index
            cell_ids = [f"{table_shape_id}.{i}.{j}" for j in range(num_cols)]
            shape_id_array = cell_ids + overlaid_shapes
            
            # Add to data model
            row_data = {
                'shape_name': table_shape_id,
                'shape_id': shape_id_array,
                'text': row_text,
                'cells': cells,
                'label': label,
                'section_type': 'row',
                'top': bounds['top'],
                'left': bounds['left'],
                'right': bounds['right'],
                'bottom': bounds['bottom'],
                'width': bounds['width'],
                'height': bounds['height']
            }
            if overlaid_shapes:
                row_data['overlaid_shapes'] = overlaid_shapes
            
            table_labels_list.append(row_data)
            print(f"Saved row {i}: {label}" + (f" (with {len(overlaid_shapes)} overlaid shapes)" if overlaid_shapes else ""))
    
    # 2. Save column labels if checked
    if save_col_labels_var.get():
        num_rows = current_table.Table.Rows.Count
        num_cols = current_table.Table.Columns.Count
        
        for j in range(num_cols):
            # Get text from all cells in the column
            col_text_parts = []
            for i in range(1, num_rows + 1):
                try:
                    cell_text = current_table.Table.Columns(j+1).Cells(i).Shape.TextFrame.TextRange.Text.strip()
                    col_text_parts.append(cell_text)
                except:
                    pass
            
            # Concatenate with spaces
            col_text = " ".join(col_text_parts)
            
            # Create cells array
            cells = [f"{i}.{j}" for i in range(num_rows)]
            
            # Get label from textbox
            label = col_entries[j].get() if j < len(col_entries) else f"Col {j}"
            
            # Calculate bounds for this column
            bounds = calculate_section_bounds(cells, cell_coords)
            
            # Get overlaid shapes if checkbox is checked
            overlaid_shapes = []
            if add_col_overlaid_shapes_var.get():
                try:
                    tolerance = float(col_tolerance_entry.get())
                except:
                    tolerance = 5
                overlaid_shapes = get_overlaid_shapes_in_bounds(bounds, tolerance)
            
            # Build shape_id array: cell IDs + overlaid shape IDs
            # Cell ID format: table_shape_id.row_index.col_index
            cell_ids = [f"{table_shape_id}.{i}.{j}" for i in range(num_rows)]
            shape_id_array = cell_ids + overlaid_shapes
            
            # Add to data model
            col_data = {
                'shape_name': table_shape_id,
                'shape_id': shape_id_array,
                'text': col_text,
                'cells': cells,
                'label': label,
                'section_type': 'col',
                'top': bounds['top'],
                'left': bounds['left'],
                'right': bounds['right'],
                'bottom': bounds['bottom'],
                'width': bounds['width'],
                'height': bounds['height']
            }
            if overlaid_shapes:
                col_data['overlaid_shapes'] = overlaid_shapes
            
            table_labels_list.append(col_data)
            print(f"Saved column {j}: {label}" + (f" (with {len(overlaid_shapes)} overlaid shapes)" if overlaid_shapes else ""))
    
    # 3. Save custom groups
    num_rows = current_table.Table.Rows.Count
    num_cols = current_table.Table.Columns.Count
    
    for group_data in custom_group_rows:
        # Get values from stored widgets
        row_col_type = group_data['row_col_var'].get()  # "rows" or "cols"
        start_index = int(group_data['start_entry'].get())
        end_index = int(group_data['end_entry'].get())
        group_name = group_data['group_name_entry'].get()
        
        print(f"DEBUG: Processing group '{group_name}': type={row_col_type}, start={start_index}, end={end_index}")
        
        # Get text from all cells in the group
        text_parts = []
        cells = []
        
        if row_col_type == "rows":
            # Multiple rows
            for i in range(start_index, end_index + 1):
                if i < num_rows:
                    for j in range(num_cols):
                        try:
                            cell_text = current_table.Table.Rows(i+1).Cells(j+1).Shape.TextFrame.TextRange.Text.strip()
                            text_parts.append(cell_text)
                            cells.append(f"{i}.{j}")
                        except:
                            pass
        else:  # "cols"
            # Multiple columns
            for j in range(start_index, end_index + 1):
                if j < num_cols:
                    for i in range(num_rows):
                        try:
                            cell_text = current_table.Table.Columns(j+1).Cells(i+1).Shape.TextFrame.TextRange.Text.strip()
                            text_parts.append(cell_text)
                            cells.append(f"{i}.{j}")
                        except:
                            pass
        
        # Concatenate with spaces
        group_text = " ".join(text_parts)
        
        # Calculate bounds for this custom group
        bounds = calculate_section_bounds(cells, cell_coords)
        
        # Determine section_type based on row_col_type
        section_type = 'group_of_rows' if row_col_type == 'rows' else 'group_of_cols'
        
        # Build shape_id array: cell IDs for all cells in this group
        # Cell ID format: table_shape_id.row_index.col_index
        cell_ids = [f"{table_shape_id}.{cell}" for cell in cells]
        
        # Add to data model
        table_labels_list.append({
            'shape_name': table_shape_id,
            'shape_id': cell_ids,
            'text': group_text,
            'cells': cells,
            'label': group_name,
            'section_type': section_type,
            'top': bounds['top'],
            'left': bounds['left'],
            'right': bounds['right'],
            'bottom': bounds['bottom'],
            'width': bounds['width'],
            'height': bounds['height']
        })
        print(f"Saved custom group: {group_name} ({row_col_type} {start_index}-{end_index})")
    
    print(f"Saved table {current_index + 1}")
    dialog.destroy()
    
    # Go to next table or finish if this was the last one
    if current_index + 1 < len(tables):
        show_table_form_with_data(tables, current_index + 1)
    else:
        finish_text_labeling_process()

def skip_table_and_go_to_next(dialog, tables, current_index):
    """Skip current table and go to next table"""
    print(f"Skipped table {current_index + 1}")
    dialog.destroy()
    
    # Go to next table or finish if this was the last one
    if current_index + 1 < len(tables):
        show_table_form_with_data(tables, current_index + 1)
    else:
        finish_text_labeling_process()

def go_to_previous_section_from_tables(dialog):
    """Go back to text sections"""
    dialog.destroy()
    show_text_labeling_form(None)

def finish_table_labeling(dialog):
    """Complete the table labeling process"""
    dialog.destroy()
    finish_text_labeling_process()

def cancel_table_labeling(dialog):
    """Cancel table labeling and go back to text labeling"""
    dialog.destroy()
    # For now, just finish the process
    finish_text_labeling_process()

def finish_text_labeling_process():
    """Complete the text labeling process"""
    global listener_active
    listener_active = False
    print("Text labeling finished - stopping listener")
    
    # Step 1 & 2: Add section_type to all lists
    # Add section_type to group_list
    for x in group_list:
        x['section_type'] = 'shape_group'
    
    # Add section_type to text_section_list
    for x in text_section_list:
        x['section_type'] = 'text_section'
    
    # Add section_type to shape_data and set labels from individual_shape_labels
    # Create a mapping of shape_id -> label from individual_shape_labels
    individual_shape_label_map = {}
    for shape_label in individual_shape_labels:
        individual_shape_label_map[shape_label['shape_id']] = shape_label['label']
    
    for x in shape_data:
        x['section_type'] = 'orphan_shape'
        # Set label if found in individual_shape_labels, otherwise leave blank
        if x['shape_id'] in individual_shape_label_map:
            x['label'] = individual_shape_label_map[x['shape_id']]
        else:
            x['label'] = ''
    
    # table_labels_list already has section_type added in save_and_go_to_next_table
    
    # Step 3: Remove orphan shapes that have been labeled in groups
    # Collect all labeled shape IDs from groups only
    """labeled_shape_ids = set()
    for group in group_list:
        if isinstance(group.get('shape_id'), list):
            labeled_shape_ids.update(group['shape_id'])
        else:
            labeled_shape_ids.add(group['shape_id'])
    
    # Filter shape_data to remove shapes that are in groups, rename remaining as individual_shape
    filtered_shape_data = []
    for shape in shape_data:
        if shape['shape_id'] not in labeled_shape_ids:
            shape_copy = shape.copy()
            shape_copy['section_type'] = 'individual_shape'  # Rename from orphan_shape
            filtered_shape_data.append(shape_copy)"""
    
    # Combine filtered_shape_data and group_list
    # (text_section_list and table_labels_list will be inserted later)
    dl = shape_data + group_list.copy()
    
    # Ensure all items have shape_id as a list
    for x in dl:
        if 'shape_id' in x and not isinstance(x['shape_id'], list):
            x['shape_id'] = [x['shape_id']]
    
    # Remove duplicates based on shape_id (comparing as sets, regardless of order)
    # Keep the instance with longer label
    seen_shape_ids = {}  # Maps frozenset of shape_ids to item
    filtered_dl = []
    for item in dl:
        shape_id_set = frozenset(item.get('shape_id', []))
        
        if shape_id_set in seen_shape_ids:
            # Duplicate found - keep the one with longer label
            existing_item = seen_shape_ids[shape_id_set]
            existing_label = existing_item.get('label', '')
            current_label = item.get('label', '')
            
            if len(current_label) > len(existing_label):
                # Replace with current item (has longer label)
                seen_shape_ids[shape_id_set] = item
                # Update in filtered_dl
                for i, x in enumerate(filtered_dl):
                    if frozenset(x.get('shape_id', [])) == shape_id_set:
                        filtered_dl[i] = item
                        break
        else:
            # New shape_id set
            seen_shape_ids[shape_id_set] = item
            filtered_dl.append(item)
    
    dl = filtered_dl
    
    # Insert text sections as tree structures based on char indices
    # This re-arranges text sections within their parent shapes
    dl = basic.add_text_sections(dl, text_section_list)
    
    # Add table sections to the data list
    # Overlaid shapes are already merged into shape_id
    dl = basic.add_table_sections(dl, table_labels_list)
    
    # Step 4: Run generate_structure_main to arrange into tree based on spatial coordinates
    # This arranges shapes and groups based on top, left, bottom, right
    dl = structure_shapes.generate_structure_main(dl)
    
    # Add unlabeled individual shapes to groups where they are contained
    # dl = basic.add_unlabeled_shapes_to_groups(dl)
    
    # Remove individual shapes that don't belong to any group
    # (They've been copied into the tree structure where needed)
    # dl = basic.remove_ungrouped_individual_shapes(dl)
    
    # Add metadata
    for x in dl:
        x['slide_height'] = slide_dimensions['height']
        x['slide_width'] = slide_dimensions['width']
        x['test_index'] = test_index  
    
    df = basic.save_to_csv(dl, test_index)  # Saves as CSV
    print("Done")   
    print("Finished labeling process")
    exit()



def main():
    """Main function to run the GUI"""
    setup_gui()
    root.mainloop()
    
    # After the main window is closed, keep the invisible root running
    if invisible_root:
        invisible_root.mainloop()


if __name__ == "__main__":
    main()
