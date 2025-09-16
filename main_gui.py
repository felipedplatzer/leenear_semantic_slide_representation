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
text_section_list = []
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
    text_dialog.geometry("500x400")
    text_dialog.resizable(False, False)
    
    # Center on screen
    text_dialog.update_idletasks()
    x = (text_dialog.winfo_screenwidth() // 2) - (500 // 2)
    y = (text_dialog.winfo_screenheight() // 2) - (400 // 2)
    text_dialog.geometry(f"500x400+{x}+{y}")
    
    # Bring dialog to foreground
    text_dialog.lift()
    text_dialog.attributes('-topmost', True)
    text_dialog.after_idle(lambda: text_dialog.attributes('-topmost', False))
    text_dialog.focus_force()
    
    # Main frame
    main_frame = tk.Frame(text_dialog, padx=20, pady=20)
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Title
    title_label = tk.Label(main_frame, text="Label Text Sections", 
                          font=("Arial", 14, "bold"))
    title_label.pack(pady=(0, 20))
    
    # Text preview panel
    preview_frame = tk.Frame(main_frame, relief="sunken", bd=2, bg="white")
    preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
    
    preview_label = tk.Label(preview_frame, text="Text preview will appear here", 
                            fg="gray", wraplength=450, justify="left", 
                            bg="white", anchor="nw")
    preview_label.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
    
    # Input fields frame
    inputs_frame = tk.Frame(main_frame)
    inputs_frame.pack(fill=tk.X, pady=(0, 20))
    
    # Shape ID input
    shape_id_label = tk.Label(inputs_frame, text="Shape ID:", font=("Arial", 10))
    shape_id_label.grid(row=0, column=0, sticky="w", pady=(0, 5))
    shape_id_entry = tk.Entry(inputs_frame, font=("Arial", 11), width=30)
    shape_id_entry.grid(row=0, column=1, sticky="ew", pady=(0, 5))
    shape_id_entry.bind('<FocusOut>', lambda e: update_text_preview(preview_label, shape_id_entry, start_char_entry, end_char_entry, selection))
    
    # Start char input
    start_char_label = tk.Label(inputs_frame, text="Start Char (0-indexed):", font=("Arial", 10))
    start_char_label.grid(row=1, column=0, sticky="w", pady=(0, 5))
    start_char_entry = tk.Entry(inputs_frame, font=("Arial", 11), width=30)
    start_char_entry.grid(row=1, column=1, sticky="ew", pady=(0, 5))
    start_char_entry.bind('<FocusOut>', lambda e: update_text_preview(preview_label, shape_id_entry, start_char_entry, end_char_entry, selection))
    
    # End char input
    end_char_label = tk.Label(inputs_frame, text="End Char (0-indexed):", font=("Arial", 10))
    end_char_label.grid(row=2, column=0, sticky="w", pady=(0, 5))
    end_char_entry = tk.Entry(inputs_frame, font=("Arial", 11), width=30)
    end_char_entry.grid(row=2, column=1, sticky="ew", pady=(0, 5))
    end_char_entry.bind('<FocusOut>', lambda e: update_text_preview(preview_label, shape_id_entry, start_char_entry, end_char_entry, selection))
    
    # Name of text section input
    name_label = tk.Label(inputs_frame, text="Name of Text Section:", font=("Arial", 10))
    name_label.grid(row=3, column=0, sticky="w", pady=(0, 5))
    name_entry = tk.Entry(inputs_frame, font=("Arial", 11), width=30)
    name_entry.grid(row=3, column=1, sticky="ew", pady=(0, 5))
    
    # Configure grid weights
    inputs_frame.columnconfigure(1, weight=1)
    
    # Buttons frame
    buttons_frame = tk.Frame(main_frame)
    buttons_frame.pack(fill=tk.X, pady=(20, 0))
    
    # Save & Next button
    save_next_btn = tk.Button(buttons_frame, text="Save & Next", 
                             command=lambda: save_and_next_text(shape_id_entry.get().strip(), 
                                                              start_char_entry.get().strip(),
                                                              end_char_entry.get().strip(),
                                                              name_entry.get().strip(),
                                                              text_dialog),
                             bg="#2196F3", fg="white",
                             width=12, height=2)
    save_next_btn.pack(side=tk.LEFT, padx=(0, 10))
    
    # Save & Finish button
    save_finish_btn = tk.Button(buttons_frame, text="Save & Finish", 
                               command=lambda: save_and_finish_text(shape_id_entry.get().strip(), 
                                                                  start_char_entry.get().strip(),
                                                                  end_char_entry.get().strip(),
                                                                  name_entry.get().strip(),
                                                                  text_dialog),
                               bg="#4CAF50", fg="white",
                               width=12, height=2)
    save_finish_btn.pack(side=tk.LEFT, padx=(0, 10))
    
    # Finish button
    finish_btn = tk.Button(buttons_frame, text="Finish", 
                          command=lambda: finish_text_labeling(text_dialog),
                          bg="#F44336", fg="white",
                          width=12, height=2)
    finish_btn.pack(side=tk.LEFT)
    
    # Focus on first input
    text_dialog.after(100, lambda: shape_id_entry.focus())
    
    # Bind Enter key to Save & Next
    text_dialog.bind('<Return>', lambda e: save_and_next_text(shape_id_entry.get().strip(), 
                                                             start_char_entry.get().strip(),
                                                             end_char_entry.get().strip(),
                                                             name_entry.get().strip(),
                                                             text_dialog))

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

def save_and_finish_text(shape_id, start_char, end_char, name, dialog):
    """Save text section and finish"""
    # TODO: Implement save and finish functionality
    text_section_list.append(get_text_dict(shape_id, start_char, end_char, name))
    print(f"Save & Finish: Shape ID={shape_id}, Start={start_char}, End={end_char}, Name={name}")
    dialog.destroy()
    finish_text_labeling_process()

def finish_text_labeling(dialog):
    """Finish text labeling without saving"""
    dialog.destroy()
    finish_text_labeling_process()

def finish_text_labeling_process():
    """Complete the text labeling process"""
    global listener_active
    listener_active = False
    print("Text labeling finished - stopping listener")
    dl = basic.add_unnamed_shapes(shape_data, group_list)
    dl = structure_shapes.generate_structure_main(dl)
    dl = basic.add_text_sections(dl, text_section_list)
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
