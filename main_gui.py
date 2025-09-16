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
    dialog_root.geometry("350x200")
    dialog_root.resizable(False, False)
    
    # Center on screen
    dialog_root.update_idletasks()
    x = (dialog_root.winfo_screenwidth() // 2) - (350 // 2)
    y = (dialog_root.winfo_screenheight() // 2) - (200 // 2)
    dialog_root.geometry(f"350x200+{x}+{y}")
    
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
    
    # Group name input
    name_label = tk.Label(main_frame, text="Name this group:", 
                         font=("Arial", 10))
    name_label.pack(anchor="w", pady=(0, 5))
    
    name_entry = tk.Entry(main_frame, font=("Arial", 11), width=30)
    name_entry.pack(pady=(0, 20))
    
    # Buttons frame
    buttons_frame = tk.Frame(main_frame)
    buttons_frame.pack(fill=tk.X, pady=(0, 10))
    
    # Continue button
    continue_btn = tk.Button(buttons_frame, text="Continue", 
                            command=lambda: continue_labeling(name_entry.get().strip(), selection, dialog_root),
                            bg="#2196F3", fg="white",
                            width=10, height=5)
    continue_btn.pack(side=tk.LEFT, padx=(0, 5))
    
    # Finish button
    finish_btn = tk.Button(buttons_frame, text="Save & Finish", 
                          command=lambda: save_and_finish(name_entry.get().strip(), selection, dialog_root),
                          bg="#F44336", fg="white",
                          width=10, height=5)
    finish_btn.pack(side=tk.LEFT, padx=(0, 5))
    
    # Skip button
    skip_btn = tk.Button(buttons_frame, text="Skip", 
                        command=lambda: skip_labeling(dialog_root),
                        bg="#FFC107", fg="black",
                        width=10, height=5)
    skip_btn.pack(side=tk.LEFT, padx=(0, 5))
    
    # Finish button (no save)
    finish_btn_no_save = tk.Button(buttons_frame, text="Finish", 
                                  command=lambda: finish(dialog_root),
                                  bg="#9C27B0", fg="white",
                                  width=10, height=5)
    finish_btn_no_save.pack(side=tk.LEFT)
    
    # Focus on textbox after dialog is created
    dialog_root.after(100, lambda: name_entry.focus())
    
    # Bind Enter key to Continue (only when textbox is focused)
    name_entry.bind('<Return>', lambda e: continue_labeling(name_entry.get().strip(), selection, dialog_root))
    
    # Bind Escape key to Skip
    dialog_root.bind('<Escape>', lambda e: skip_labeling(dialog_root))
    
    # Bind Enter key to focused button when buttons are focused
    def on_button_focus(event):
        widget = event.widget
        if isinstance(widget, tk.Button):
            widget.bind('<Return>', lambda e: widget.invoke())
    
    # Apply focus binding to all buttons
    for button in [continue_btn, finish_btn, skip_btn, finish_btn_no_save]:
        button.bind('<FocusIn>', on_button_focus)

def skip_labeling(dialog_root):
    """Skip this selection and continue labeling"""
    dialog_root.destroy()
    # Status updates are not needed since main window is hidden
    print("Skipped selection - listening for next group")

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
    """Finish the labeling process"""
    global listener_active
    save(group_name, selection)
    finish(dialog_root)



def main():
    """Main function to run the GUI"""
    setup_gui()
    root.mainloop()
    
    # After the main window is closed, keep the invisible root running
    if invisible_root:
        invisible_root.mainloop()


if __name__ == "__main__":
    main()
