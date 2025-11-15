import tkinter as tk
from tkinter import ttk, messagebox
import win32com.client
import os
import json
import tempfile
from datetime import datetime
from PIL import ImageGrab
import pythoncom


class TableSectionWidget(ttk.Frame):
    """Widget for a single table section with radio buttons and textboxes"""
    
    def __init__(self, parent, on_remove_callback, powerpoint_app):
        super().__init__(parent, relief=tk.RIDGE, borderwidth=2, padding=10)
        self.on_remove_callback = on_remove_callback
        self.powerpoint_app = powerpoint_app
        
        # Table ID
        ttk.Label(self, text="Table ID:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.table_id_entry = ttk.Entry(self, width=15)
        self.table_id_entry.grid(row=0, column=1, padx=5, pady=5)
        
        # Radio button variable
        self.selection_type = tk.StringVar(value="cells")
        
        # Create radio buttons and textboxes
        row_start = 1
        
        # Cells option
        self.cells_radio = ttk.Radiobutton(self, text="Cells", variable=self.selection_type, 
                                           value="cells", command=self.on_radio_change)
        self.cells_radio.grid(row=row_start, column=0, sticky=tk.W, padx=5, pady=3)
        
        self.cells_entry = ttk.Entry(self, width=30)
        self.cells_entry.grid(row=row_start, column=1, padx=5, pady=3)
        self.cells_entry.bind('<KeyRelease>', self.on_cells_change)
        
        self.cells_label = ttk.Label(self, text="", foreground="blue", cursor="hand2")
        self.cells_label.grid(row=row_start, column=2, sticky=tk.W, padx=5, pady=3)
        self.cells_label.bind('<Button-1>', lambda e: self.toggle_label_text(self.cells_label))
        
        # Rows option
        self.rows_radio = ttk.Radiobutton(self, text="Rows", variable=self.selection_type, 
                                          value="rows", command=self.on_radio_change)
        self.rows_radio.grid(row=row_start+1, column=0, sticky=tk.W, padx=5, pady=3)
        
        self.rows_entry = ttk.Entry(self, width=30, state=tk.DISABLED)
        self.rows_entry.grid(row=row_start+1, column=1, padx=5, pady=3)
        self.rows_entry.bind('<KeyRelease>', self.on_rows_change)
        
        self.rows_label = ttk.Label(self, text="", foreground="blue", cursor="hand2")
        self.rows_label.grid(row=row_start+1, column=2, sticky=tk.W, padx=5, pady=3)
        self.rows_label.bind('<Button-1>', lambda e: self.toggle_label_text(self.rows_label))
        
        # Cols option
        self.cols_radio = ttk.Radiobutton(self, text="Cols", variable=self.selection_type, 
                                          value="cols", command=self.on_radio_change)
        self.cols_radio.grid(row=row_start+2, column=0, sticky=tk.W, padx=5, pady=3)
        
        self.cols_entry = ttk.Entry(self, width=30, state=tk.DISABLED)
        self.cols_entry.grid(row=row_start+2, column=1, padx=5, pady=3)
        self.cols_entry.bind('<KeyRelease>', self.on_cols_change)
        
        self.cols_label = ttk.Label(self, text="", foreground="blue", cursor="hand2")
        self.cols_label.grid(row=row_start+2, column=2, sticky=tk.W, padx=5, pady=3)
        self.cols_label.bind('<Button-1>', lambda e: self.toggle_label_text(self.cols_label))
        
        # Remove button
        ttk.Button(self, text="Remove Section", command=self.on_remove).grid(
            row=row_start+3, column=1, pady=10, sticky=tk.E)
        
        # Store full text for labels
        self.cells_full_text = ""
        self.rows_full_text = ""
        self.cols_full_text = ""
        
    def on_radio_change(self):
        """Enable/disable textboxes based on radio button selection"""
        selection = self.selection_type.get()
        
        # Disable all
        self.cells_entry.config(state=tk.DISABLED)
        self.rows_entry.config(state=tk.DISABLED)
        self.cols_entry.config(state=tk.DISABLED)
        
        # Enable selected
        if selection == "cells":
            self.cells_entry.config(state=tk.NORMAL)
        elif selection == "rows":
            self.rows_entry.config(state=tk.NORMAL)
        elif selection == "cols":
            self.cols_entry.config(state=tk.NORMAL)
    
    def on_cells_change(self, event):
        """Validate and display cell text"""
        text = self.cells_entry.get().strip()
        if not text:
            self.cells_label.config(text="", foreground="blue")
            return
        
        table_id = self.table_id_entry.get().strip()
        if not table_id:
            self.cells_label.config(text="", foreground="blue")
            return
        
        # Validate format
        cells = [c.strip() for c in text.split(',')]
        for cell in cells:
            if not cell:
                continue
            parts = cell.split('.')
            if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
                self.cells_label.config(text="INVALID FORMAT", foreground="red")
                return
        
        # Get cell text from PowerPoint
        try:
            cell_texts = self.get_cell_texts(table_id, cells)
            if isinstance(cell_texts, str) and ("NOT AVAILABLE" in cell_texts or "not found" in cell_texts):
                self.cells_label.config(text=cell_texts, foreground="red")
            else:
                self.cells_full_text = " | ".join(cell_texts)
                display_text = self.truncate_text(self.cells_full_text)
                self.cells_label.config(text=display_text, foreground="blue")
        except Exception as e:
            self.cells_label.config(text=f"Error: {str(e)}", foreground="red")
    
    def on_rows_change(self, event):
        """Validate and display row text"""
        text = self.rows_entry.get().strip()
        if not text:
            self.rows_label.config(text="", foreground="blue")
            return
        
        table_id = self.table_id_entry.get().strip()
        if not table_id:
            self.rows_label.config(text="", foreground="blue")
            return
        
        # Validate format
        rows = [r.strip() for r in text.split(',')]
        for row in rows:
            if not row:
                continue
            if not row.isdigit():
                self.rows_label.config(text="INVALID FORMAT", foreground="red")
                return
        
        # Get row text from PowerPoint
        try:
            row_texts = self.get_row_texts(table_id, [int(r) for r in rows if r])
            if isinstance(row_texts, str) and ("NOT AVAILABLE" in row_texts or "not found" in row_texts):
                self.rows_label.config(text=row_texts, foreground="red")
            else:
                self.rows_full_text = " | ".join(row_texts)
                display_text = self.truncate_text(self.rows_full_text)
                self.rows_label.config(text=display_text, foreground="blue")
        except Exception as e:
            self.rows_label.config(text=f"Error: {str(e)}", foreground="red")
    
    def on_cols_change(self, event):
        """Validate and display col text"""
        text = self.cols_entry.get().strip()
        if not text:
            self.cols_label.config(text="", foreground="blue")
            return
        
        table_id = self.table_id_entry.get().strip()
        if not table_id:
            self.cols_label.config(text="", foreground="blue")
            return
        
        # Validate format
        cols = [c.strip() for c in text.split(',')]
        for col in cols:
            if not col:
                continue
            if not col.isdigit():
                self.cols_label.config(text="INVALID FORMAT", foreground="red")
                return
        
        # Get col text from PowerPoint
        try:
            col_texts = self.get_col_texts(table_id, [int(c) for c in cols if c])
            if isinstance(col_texts, str) and ("NOT AVAILABLE" in col_texts or "not found" in col_texts):
                self.cols_label.config(text=col_texts, foreground="red")
            else:
                self.cols_full_text = " | ".join(col_texts)
                display_text = self.truncate_text(self.cols_full_text)
                self.cols_label.config(text=display_text, foreground="blue")
        except Exception as e:
            self.cols_label.config(text=f"Error: {str(e)}", foreground="red")
    
    def get_cell_texts(self, table_id, cells):
        """Get text from specific cells in the table"""
        try:
            table_id_int = int(table_id)
            slide = self.powerpoint_app.ActiveWindow.View.Slide
            
            # Find the table shape
            table_shape = None
            for shape in slide.Shapes:
                if shape.Id == table_id_int and shape.HasTable:
                    table_shape = shape
                    break
            
            if not table_shape:
                return f"Table ID {table_id} not found"
            
            table = table_shape.Table
            texts = []
            
            for cell in cells:
                if not cell:
                    continue
                parts = cell.split('.')
                row_idx = int(parts[0])
                col_idx = int(parts[1])
                
                if row_idx < 1 or row_idx > table.Rows.Count or col_idx < 1 or col_idx > table.Columns.Count:
                    return f"CELL {cell} NOT AVAILABLE"
                
                cell_text = table.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Text.strip()
                texts.append(cell_text if cell_text else "[empty]")
            
            return texts
        except Exception as e:
            return f"Error: {str(e)}"
    
    def get_row_texts(self, table_id, rows):
        """Get text from specific rows in the table"""
        try:
            table_id_int = int(table_id)
            slide = self.powerpoint_app.ActiveWindow.View.Slide
            
            # Find the table shape
            table_shape = None
            for shape in slide.Shapes:
                if shape.Id == table_id_int and shape.HasTable:
                    table_shape = shape
                    break
            
            if not table_shape:
                return f"Table ID {table_id} not found"
            
            table = table_shape.Table
            texts = []
            
            for row_idx in rows:
                if row_idx < 1 or row_idx > table.Rows.Count:
                    return f"ROW {row_idx} NOT AVAILABLE"
                
                row_text = []
                for col_idx in range(1, table.Columns.Count + 1):
                    cell_text = table.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Text.strip()
                    row_text.append(cell_text if cell_text else "[empty]")
                
                texts.append(" ".join(row_text))
            
            return texts
        except Exception as e:
            return f"Error: {str(e)}"
    
    def get_col_texts(self, table_id, cols):
        """Get text from specific cols in the table"""
        try:
            table_id_int = int(table_id)
            slide = self.powerpoint_app.ActiveWindow.View.Slide
            
            # Find the table shape
            table_shape = None
            for shape in slide.Shapes:
                if shape.Id == table_id_int and shape.HasTable:
                    table_shape = shape
                    break
            
            if not table_shape:
                return f"Table ID {table_id} not found"
            
            table = table_shape.Table
            texts = []
            
            for col_idx in cols:
                if col_idx < 1 or col_idx > table.Columns.Count:
                    return f"COL {col_idx} NOT AVAILABLE"
                
                col_text = []
                for row_idx in range(1, table.Rows.Count + 1):
                    cell_text = table.Cell(row_idx, col_idx).Shape.TextFrame.TextRange.Text.strip()
                    col_text.append(cell_text if cell_text else "[empty]")
                
                texts.append(" ".join(col_text))
            
            return texts
        except Exception as e:
            return f"Error: {str(e)}"
    
    def truncate_text(self, text, max_length=20):
        """Truncate text to max_length and add ellipsis"""
        if len(text) <= max_length:
            return text
        return text[:max_length] + "…"
    
    def toggle_label_text(self, label):
        """Toggle between truncated and full text"""
        current_text = label.cget("text")
        
        # Determine which label was clicked
        if label == self.cells_label:
            full_text = self.cells_full_text
        elif label == self.rows_label:
            full_text = self.rows_full_text
        elif label == self.cols_label:
            full_text = self.cols_full_text
        else:
            return
        
        if not full_text or "INVALID" in current_text or "NOT AVAILABLE" in current_text or "Error" in current_text:
            return
        
        if current_text.endswith("…"):
            label.config(text=full_text)
        else:
            label.config(text=self.truncate_text(full_text))
    
    def on_remove(self):
        """Call the remove callback"""
        self.on_remove_callback(self)
    
    def get_data(self):
        """Get the table section data"""
        table_id = self.table_id_entry.get().strip()
        if not table_id:
            return None
        
        selection = self.selection_type.get()
        
        if selection == "cells":
            text = self.cells_entry.get().strip()
            if not text:
                return None
            cells = [c.strip() for c in text.split(',') if c.strip()]
            return {
                "table_id": table_id,
                "type": "cells",
                "values": cells
            }
        elif selection == "rows":
            text = self.rows_entry.get().strip()
            if not text:
                return None
            rows = [int(r.strip()) for r in text.split(',') if r.strip()]
            return {
                "table_id": table_id,
                "type": "rows",
                "values": rows
            }
        elif selection == "cols":
            text = self.cols_entry.get().strip()
            if not text:
                return None
            cols = [int(c.strip()) for c in text.split(',') if c.strip()]
            return {
                "table_id": table_id,
                "type": "cols",
                "values": cols
            }


class PowerPointShapeCaptureApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Shape Capture")
        self.root.geometry("800x700")
        
        # Initialize PowerPoint connection
        try:
            pythoncom.CoInitialize()
            self.ppt = win32com.client.Dispatch("PowerPoint.Application")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to connect to PowerPoint: {str(e)}")
            self.root.destroy()
            return
        
        # Main container
        main_frame = ttk.Frame(root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File label
        self.file_label = ttk.Label(main_frame, text="File: ", foreground="gray", wraplength=780, justify=tk.LEFT)
        self.file_label.pack(fill=tk.X, pady=(0, 5))
        
        # Slide number label
        self.slide_label = ttk.Label(main_frame, text="Slide number: ", foreground="gray", justify=tk.LEFT)
        self.slide_label.pack(fill=tk.X, pady=(0, 10))
        
        # Name field
        name_frame = ttk.Frame(main_frame)
        name_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(name_frame, text="Name:").pack(side=tk.LEFT, padx=5)
        self.name_entry = ttk.Entry(name_frame, width=40)
        self.name_entry.pack(side=tk.LEFT, padx=5)
        
        # Radio button for selection mode
        radio_frame = ttk.Frame(main_frame)
        radio_frame.pack(fill=tk.X, pady=10)
        
        self.selection_mode = tk.StringVar(value="shapes")
        
        # Select shapes radio button with label on same line
        shapes_frame = ttk.Frame(radio_frame)
        shapes_frame.pack(fill=tk.X, pady=2)
        
        self.shapes_radio = ttk.Radiobutton(shapes_frame, text="Select shapes", 
                                           variable=self.selection_mode, 
                                           value="shapes", 
                                           command=self.on_mode_change)
        self.shapes_radio.pack(side=tk.LEFT, padx=5)
        
        # Label for selected shapes (shown next to "Select shapes" radio button)
        self.selected_shapes_label = ttk.Label(shapes_frame, text="", foreground="blue")
        self.selected_shapes_label.pack(side=tk.LEFT, padx=5)
        
        # Drag rectangle radio button
        self.rectangle_radio = ttk.Radiobutton(radio_frame, text="Drag rectangle", 
                                              variable=self.selection_mode, 
                                              value="rectangle", 
                                              command=self.on_mode_change)
        self.rectangle_radio.pack(fill=tk.X, padx=5, pady=2)
        
        # Table radio button with label on same line
        table_frame = ttk.Frame(radio_frame)
        table_frame.pack(fill=tk.X, pady=2)
        
        self.table_radio = ttk.Radiobutton(table_frame, text="Select table rows/cols/cells", 
                                          variable=self.selection_mode, 
                                          value="table", 
                                          command=self.on_mode_change)
        self.table_radio.pack(side=tk.LEFT, padx=5)
        
        # Label for selected table (shown next to "Select table" radio button)
        self.selected_table_label = ttk.Label(table_frame, text="", foreground="blue")
        self.selected_table_label.pack(side=tk.LEFT, padx=5)
        
        # Container for mode-specific content
        self.mode_content_frame = ttk.Frame(main_frame)
        self.mode_content_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Table cells section (hidden by default)
        self.table_section_frame = ttk.LabelFrame(self.mode_content_frame, text="Table Cells", padding=10)
        
        # Scrollable frame for table sections
        canvas = tk.Canvas(self.table_section_frame, height=300)
        scrollbar = ttk.Scrollbar(self.table_section_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Add section button
        self.add_section_btn = ttk.Button(self.table_section_frame, text="Add Section", 
                                          command=self.add_table_section)
        self.add_section_btn.pack(pady=5)
        
        # List to store table section widgets
        self.table_sections = []
        
        # Add initial section
        self.add_table_section()
        
        # Drag rectangle section (hidden by default)
        self.rectangle_section_frame = ttk.LabelFrame(self.mode_content_frame, text="Drag Rectangle", padding=10)
        
        # Instructions label
        instruction_text = "Drag rectangles on top of the image. Once you're done with one rectangle, drag the next one"
        ttk.Label(self.rectangle_section_frame, text=instruction_text, wraplength=760, justify=tk.LEFT).pack(fill=tk.X, pady=(0, 10))
        
        # Canvas for slide image with rectangles
        self.rect_canvas = tk.Canvas(self.rectangle_section_frame, bg="white", height=400)
        self.rect_canvas.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Bind mouse events for drawing rectangles
        self.rect_canvas.bind("<ButtonPress-1>", self.on_rect_press)
        self.rect_canvas.bind("<B1-Motion>", self.on_rect_drag)
        self.rect_canvas.bind("<ButtonRelease-1>", self.on_rect_release)
        
        # Variables for rectangle drawing
        self.rect_start_x = None
        self.rect_start_y = None
        self.current_rect = None
        self.rectangles = []  # List of drawn rectangle IDs
        self.rectangle_bboxes = []  # List of bounding boxes in relative coordinates
        self.slide_image = None
        self.canvas_image_id = None
        
        # Clear selections button
        ttk.Button(self.rectangle_section_frame, text="Clear Selections", 
                  command=self.clear_rectangles).pack(pady=5)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.ok_button = ttk.Button(button_frame, text="OK", command=self.on_ok)
        self.ok_button.pack(side=tk.LEFT, padx=5)
        
        self.exit_button = ttk.Button(button_frame, text="Exit", command=self.on_exit)
        self.exit_button.pack(side=tk.LEFT, padx=5)
        
        # Status label below buttons
        self.status_label = ttk.Label(main_frame, text="", foreground="green")
        self.status_label.pack(fill=tk.X, pady=(5, 0))
        
        # Key bindings
        self.root.bind('<Return>', lambda e: self.on_ok())
        self.root.bind('<Escape>', lambda e: self.on_exit())
        
        # Start polling for selection changes
        self.last_selection_info = ""
        self.update_info_label()
        self.poll_selection_changes()
        
        # Initialize mode display
        self.on_mode_change()
    
    def on_mode_change(self):
        """Handle radio button mode change"""
        mode = self.selection_mode.get()
        
        # Hide all sections first
        self.table_section_frame.pack_forget()
        self.rectangle_section_frame.pack_forget()
        
        # Clear labels for non-active modes
        if mode != "shapes":
            self.selected_shapes_label.config(text="")
        if mode != "table":
            self.selected_table_label.config(text="")
        
        # Show appropriate section
        if mode == "shapes":
            # Update selected shapes label
            self.update_selected_shapes_label()
        elif mode == "table":
            self.table_section_frame.pack(fill=tk.BOTH, expand=True)
            # Update selected table label
            self.update_selected_table_label()
        elif mode == "rectangle":
            self.rectangle_section_frame.pack(fill=tk.BOTH, expand=True)
            self.load_slide_image()
    
    def update_selected_shapes_label(self):
        """Update the label showing selected shape IDs"""
        try:
            if self.ppt.Presentations.Count == 0:
                self.selected_shapes_label.config(text="")
                return
            
            selection = self.ppt.ActiveWindow.Selection
            if selection.Type == 2:  # ppSelectionShapes = 2
                shapes = selection.ShapeRange
                shape_ids = [str(shape.Id) for shape in shapes]
                if shape_ids:
                    self.selected_shapes_label.config(text=f"Selected shapes: {', '.join(shape_ids)}")
                else:
                    self.selected_shapes_label.config(text="")
            else:
                self.selected_shapes_label.config(text="")
        except:
            self.selected_shapes_label.config(text="")
    
    def update_selected_table_label(self):
        """Update the label showing selected table ID or error message"""
        try:
            if self.ppt.Presentations.Count == 0:
                self.selected_table_label.config(text="ERROR: PLEASE SELECT A TABLE", foreground="red")
                return
            
            selection = self.ppt.ActiveWindow.Selection
            if selection.Type == 2:  # ppSelectionShapes = 2
                shapes = selection.ShapeRange
                # Check if any selected shape is a table
                for shape in shapes:
                    if shape.HasTable:
                        self.selected_table_label.config(text=f"Table ID: {shape.Id}", foreground="blue")
                        return
                
                # No table found in selection
                self.selected_table_label.config(text="ERROR: PLEASE SELECT A TABLE", foreground="red")
            else:
                self.selected_table_label.config(text="ERROR: PLEASE SELECT A TABLE", foreground="red")
        except:
            self.selected_table_label.config(text="ERROR: PLEASE SELECT A TABLE", foreground="red")
    
    def load_slide_image(self):
        """Load the current slide image into the canvas"""
        try:
            if self.ppt.Presentations.Count == 0:
                return
            
            presentation = self.ppt.ActivePresentation
            slide = self.ppt.ActiveWindow.View.Slide
            
            # Get slide dimensions
            self.slide_width_pts = presentation.PageSetup.SlideWidth
            self.slide_height_pts = presentation.PageSetup.SlideHeight
            
            # Export slide to temporary image
            temp_img = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
            temp_path = temp_img.name
            temp_img.close()
            
            slide.Export(temp_path, "PNG")
            
            # Load image
            from PIL import Image, ImageTk
            img = Image.open(temp_path)
            
            # Calculate scaling to fit canvas
            canvas_width = self.rect_canvas.winfo_width()
            canvas_height = self.rect_canvas.winfo_height()
            
            # Use canvas size or default if not yet rendered
            if canvas_width <= 1:
                canvas_width = 760
            if canvas_height <= 1:
                canvas_height = 400
            
            # Calculate scale to fit
            scale_x = canvas_width / img.width
            scale_y = canvas_height / img.height
            self.image_scale = min(scale_x, scale_y, 1.0)  # Don't scale up
            
            new_width = int(img.width * self.image_scale)
            new_height = int(img.height * self.image_scale)
            
            img = img.resize((new_width, new_height), Image.Resampling.LANCZOS)
            self.slide_image = ImageTk.PhotoImage(img)
            
            # Clear canvas and display image
            self.rect_canvas.delete("all")
            self.canvas_image_id = self.rect_canvas.create_image(0, 0, anchor=tk.NW, image=self.slide_image)
            
            # Store image dimensions for coordinate calculations
            self.image_width = new_width
            self.image_height = new_height
            
            # Clean up temp file
            try:
                os.unlink(temp_path)
            except:
                pass
                
        except Exception as e:
            print(f"Error loading slide image: {str(e)}")
    
    def on_rect_press(self, event):
        """Handle mouse press for rectangle drawing"""
        self.rect_start_x = event.x
        self.rect_start_y = event.y
        
        # Create a new rectangle
        self.current_rect = self.rect_canvas.create_rectangle(
            self.rect_start_x, self.rect_start_y, 
            self.rect_start_x, self.rect_start_y,
            outline='red', width=2, fill='yellow', stipple='gray25'
        )
    
    def on_rect_drag(self, event):
        """Handle mouse drag for rectangle drawing"""
        if self.current_rect:
            # Update rectangle coordinates
            self.rect_canvas.coords(self.current_rect, 
                                   self.rect_start_x, self.rect_start_y,
                                   event.x, event.y)
    
    def on_rect_release(self, event):
        """Handle mouse release for rectangle drawing"""
        if self.current_rect:
            # Get final coordinates
            x1 = min(self.rect_start_x, event.x)
            y1 = min(self.rect_start_y, event.y)
            x2 = max(self.rect_start_x, event.x)
            y2 = max(self.rect_start_y, event.y)
            
            # Update rectangle to final position
            self.rect_canvas.coords(self.current_rect, x1, y1, x2, y2)
            
            # Store rectangle
            self.rectangles.append(self.current_rect)
            
            # Calculate relative coordinates (relative to the image on canvas)
            if hasattr(self, 'image_width') and hasattr(self, 'image_height'):
                rel_x = x1 / self.image_width
                rel_y = y1 / self.image_height
                rel_w = (x2 - x1) / self.image_width
                rel_h = (y2 - y1) / self.image_height
                
                # Round to 4 decimal places
                bbox = [round(rel_x, 4), round(rel_y, 4), round(rel_w, 4), round(rel_h, 4)]
                self.rectangle_bboxes.append(bbox)
            
            # Reset for next rectangle
            self.current_rect = None
    
    def clear_rectangles(self):
        """Clear all drawn rectangles"""
        for rect in self.rectangles:
            self.rect_canvas.delete(rect)
        self.rectangles.clear()
        self.rectangle_bboxes.clear()
    
    def add_table_section(self):
        """Add a new table section"""
        section = TableSectionWidget(self.scrollable_frame, self.remove_table_section, self.ppt)
        section.pack(fill=tk.X, pady=5)
        self.table_sections.append(section)
    
    def remove_table_section(self, section):
        """Remove a table section"""
        section.destroy()
        self.table_sections.remove(section)
        
        # If no sections left, keep the add button visible
        if len(self.table_sections) == 0:
            pass  # Add button is always visible
    
    def update_info_label(self):
        """Update the file and slide labels with current presentation info"""
        try:
            if self.ppt.Presentations.Count == 0:
                self.file_label.config(text="File: No presentation open")
                self.slide_label.config(text="Slide number: ")
                return
            
            presentation = self.ppt.ActivePresentation
            slide = self.ppt.ActiveWindow.View.Slide
            slide_number = slide.SlideIndex
            
            # Get presentation path (truncate if too long)
            pres_path = presentation.FullName
            if len(pres_path) > 100:
                pres_path = "..." + pres_path[-97:]
            
            self.file_label.config(text=f"File: {pres_path}")
            self.slide_label.config(text=f"Slide number: {slide_number}")
            
            # Update mode-specific labels
            mode = self.selection_mode.get()
            if mode == "shapes":
                self.update_selected_shapes_label()
            elif mode == "table":
                self.update_selected_table_label()
            
        except Exception as e:
            self.file_label.config(text=f"File: Error: {str(e)}")
            self.slide_label.config(text="Slide number: ")
    
    def poll_selection_changes(self):
        """Poll for selection changes every 500ms"""
        try:
            if self.ppt.Presentations.Count > 0:
                presentation = self.ppt.ActivePresentation
                slide = self.ppt.ActiveWindow.View.Slide
                slide_number = slide.SlideIndex
                
                selection = self.ppt.ActiveWindow.Selection
                shape_ids = []
                if selection.Type == 2:  # ppSelectionShapes = 2
                    shapes = selection.ShapeRange
                    shape_ids = [shape.Id for shape in shapes]
                
                # Create a unique string representing current state
                current_info = f"{presentation.FullName}|{slide_number}|{shape_ids}"
                
                # Update label if changed
                if current_info != self.last_selection_info:
                    self.last_selection_info = current_info
                    self.update_info_label()
        except:
            pass  # Ignore errors during polling
        
        # Schedule next poll
        self.root.after(500, self.poll_selection_changes)
    
    def clear_form(self):
        """Clear all form inputs"""
        self.name_entry.delete(0, tk.END)
        
        # Reset to shapes mode
        self.selection_mode.set("shapes")
        
        # Remove all table sections
        for section in list(self.table_sections):
            section.destroy()
        self.table_sections.clear()
        
        # Add one empty section back
        self.add_table_section()
        
        # Clear rectangles
        self.clear_rectangles()
        
        # Update mode display
        self.on_mode_change()
    
    def show_status_message(self, message, duration=1000):
        """Show a temporary status message"""
        self.status_label.config(text=message)
        self.root.after(duration, lambda: self.status_label.config(text=""))
    
    def bbox_to_relative(self, bbox, slide_width, slide_height):
        """Convert absolute bbox to relative coordinates (4 decimal points)"""
        left, top, width, height = bbox
        rel_left = round(left / slide_width, 4)
        rel_top = round(top / slide_height, 4)
        rel_width = round(width / slide_width, 4)
        rel_height = round(height / slide_height, 4)
        return [rel_left, rel_top, rel_width, rel_height]
    
    def is_color_white_or_transparent(self, color_obj):
        """Check if a color is white or transparent"""
        try:
            # Check if transparent (no fill)
            if color_obj.Type == 0:  # msoColorTypeBackground = 0 (transparent)
                return True
            
            if color_obj.Type == 1:  # msoColorTypeRGB = 1
                rgb = color_obj.RGB
                r = rgb & 0xFF
                g = (rgb >> 8) & 0xFF
                b = (rgb >> 16) & 0xFF
                # Check if white (all RGB values are 255)
                if r >= 250 and g >= 250 and b >= 250:
                    return True
            
            return False
        except:
            return False
    
    def is_shape_invisible(self, shape):
        """Check if shape fill and outline are both white or transparent"""
        try:
            # Check fill
            fill_invisible = False
            if shape.Fill.Visible == 0:  # Fill not visible
                fill_invisible = True
            elif shape.Fill.Transparency >= 0.99:  # Nearly fully transparent
                fill_invisible = True
            elif self.is_color_white_or_transparent(shape.Fill.ForeColor):
                fill_invisible = True
            
            # Check outline
            outline_invisible = False
            if shape.Line.Visible == 0:  # Line not visible
                outline_invisible = True
            elif shape.Line.Transparency >= 0.99:  # Nearly fully transparent
                outline_invisible = True
            elif self.is_color_white_or_transparent(shape.Line.ForeColor):
                outline_invisible = True
            
            return fill_invisible and outline_invisible
        except:
            return False
    
    def is_text_visible(self, shape):
        """Check if shape has visible text (not white/transparent)"""
        try:
            if not shape.HasTextFrame:
                return False
            
            text_range = shape.TextFrame.TextRange
            text = text_range.Text.strip()
            
            if not text:
                return False
            
            # Check text color
            font_color = text_range.Font.Color
            if self.is_color_white_or_transparent(font_color):
                return False
            
            return True
        except:
            return False
    
    def get_next_file_id(self):
        """Get the next sequential file ID by checking existing files"""
        # Create directories if they don't exist
        img_dir = os.path.join("resources", "img")
        json_dir = os.path.join("resources", "json")
        os.makedirs(img_dir, exist_ok=True)
        os.makedirs(json_dir, exist_ok=True)
        
        # Check for existing files in both directories
        max_id = 0
        
        # Check img directory
        if os.path.exists(img_dir):
            for filename in os.listdir(img_dir):
                if filename.endswith('.png'):
                    try:
                        file_id = int(filename.replace('.png', ''))
                        max_id = max(max_id, file_id)
                    except ValueError:
                        pass
        
        # Check json directory
        if os.path.exists(json_dir):
            for filename in os.listdir(json_dir):
                if filename.endswith('.json'):
                    try:
                        file_id = int(filename.replace('.json', ''))
                        max_id = max(max_id, file_id)
                    except ValueError:
                        pass
        
        return max_id + 1
    
    def on_ok(self):
        """Handle OK button click"""
        try:
            # Get next sequential file ID
            file_id = self.get_next_file_id()
            
            # Get active presentation
            if self.ppt.Presentations.Count == 0:
                messagebox.showerror("Error", "No PowerPoint presentation is open")
                return
            
            presentation = self.ppt.ActivePresentation
            slide = self.ppt.ActiveWindow.View.Slide
            slide_number = slide.SlideIndex
            
            # Get slide dimensions
            slide_width = presentation.PageSetup.SlideWidth
            slide_height = presentation.PageSetup.SlideHeight
            
            # Check if presentation is in the cloud and download if necessary
            presentation_path = presentation.FullName
            is_cloud = presentation_path.startswith("http://") or presentation_path.startswith("https://")
            
            if is_cloud:
                # Create cloud_presentations directory
                cloud_dir = os.path.join("resources", "cloud_presentations")
                os.makedirs(cloud_dir, exist_ok=True)
                
                # Extract filename from presentation name
                pres_name = presentation.Name
                if not pres_name.endswith('.pptx') and not pres_name.endswith('.ppt'):
                    pres_name += '.pptx'
                
                # Create local path for downloaded presentation (use timestamp for cloud files)
                timestamp = datetime.now().strftime("%Y-%m-%d-%H-%M-%S")
                local_pres_path = os.path.join(cloud_dir, f"{timestamp}_{pres_name}")
                
                # Download presentation
                try:
                    presentation.SaveCopyAs(os.path.abspath(local_pres_path))
                    presentation_path = os.path.abspath(local_pres_path)
                    print(f"Cloud presentation downloaded to: {presentation_path}")
                except Exception as e:
                    messagebox.showwarning("Warning", 
                        f"Could not download cloud presentation: {str(e)}\nUsing cloud URL in JSON.")
            
            # Create resources directories
            img_dir = os.path.join("resources", "img")
            json_dir = os.path.join("resources", "json")
            os.makedirs(img_dir, exist_ok=True)
            os.makedirs(json_dir, exist_ok=True)
            
            # Take screenshot of current slide
            img_path = os.path.join(img_dir, f"{file_id}.png")
            self.capture_slide_screenshot(slide, img_path)
            
            # Collect data based on selection mode
            json_data = []
            mode = self.selection_mode.get()
            
            if mode == "shapes":
                # Get selected shapes
                selection = self.ppt.ActiveWindow.Selection
                if selection.Type == 2:  # ppSelectionShapes = 2
                    shapes = selection.ShapeRange
                    
                    # Extract shape data
                    shape_ids = []
                    bboxes = []
                    colors = []
                    texts = []
                    
                    for shape in shapes:
                        # Check if shape is invisible (white/transparent fill and outline)
                        shape_invisible = self.is_shape_invisible(shape)
                        
                        if shape_invisible:
                            # Check if text is visible
                            text_visible = self.is_text_visible(shape)
                            
                            if not text_visible:
                                # Skip this shape entirely
                                continue
                            
                            # Use text bounding box instead of shape bounding box
                            try:
                                text_range = shape.TextFrame.TextRange
                                text_left = text_range.BoundLeft
                                text_top = text_range.BoundTop
                                text_width = text_range.BoundWidth
                                text_height = text_range.BoundHeight
                                abs_bbox = [text_left, text_top, text_width, text_height]
                            except:
                                # Fallback to shape bounds if text bounds fail
                                abs_bbox = [shape.Left, shape.Top, shape.Width, shape.Height]
                        else:
                            # Use normal shape bounding box
                            abs_bbox = [shape.Left, shape.Top, shape.Width, shape.Height]
                        
                        shape_ids.append(shape.Id)
                        
                        # Convert to relative coordinates
                        rel_bbox = self.bbox_to_relative(abs_bbox, slide_width, slide_height)
                        bboxes.append(rel_bbox)
                        
                        # Color (RGB)
                        try:
                            if shape.Fill.ForeColor.Type == 1:  # msoColorTypeRGB = 1
                                rgb = shape.Fill.ForeColor.RGB
                                # RGB is stored as BGR in COM, need to convert
                                r = rgb & 0xFF
                                g = (rgb >> 8) & 0xFF
                                b = (rgb >> 16) & 0xFF
                                colors.append([r, g, b])
                            else:
                                colors.append([0, 0, 0])
                        except:
                            colors.append([0, 0, 0])
                        
                        # Text
                        try:
                            if shape.HasTextFrame:
                                text = shape.TextFrame.TextRange.Text.strip()
                                texts.append(text)
                            else:
                                texts.append("")
                        except:
                            texts.append("")
                    
                    # Add shapes data to JSON
                    shapes_data = {
                        "name": self.name_entry.get().strip(),
                        "path": presentation_path,
                        "slide_number": slide_number,
                        "slide_width": slide_width,
                        "slide_height": slide_height,
                        "selection_type": "shape",
                        "shape_ids": shape_ids,
                        "table_rows": "",
                        "table_cols": "",
                        "table_cells": "",
                        "bbox": bboxes,
                        "color_rgb": colors,
                        "text": texts
                    }
                    json_data.append(shapes_data)
                else:
                    messagebox.showwarning("Warning", "No shapes selected")
                    return
            
            elif mode == "table":
                # Get table data from sections
                for section in self.table_sections:
                    section_data = section.get_data()
                    if not section_data:
                        continue
                    
                    table_id = int(section_data["table_id"])
                    values = section_data["values"]
                    section_type = section_data["type"]
                    
                    # Find the table shape
                    table_shape = None
                    for shape in slide.Shapes:
                        if shape.Id == table_id and shape.HasTable:
                            table_shape = shape
                            break
                    
                    if not table_shape:
                        messagebox.showwarning("Warning", f"Table ID {table_id} not found")
                        continue
                    
                    table = table_shape.Table
                    
                    # Extract bboxes
                    bboxes = []
                    
                    if section_type == "cells":
                        for cell_ref in values:
                            parts = cell_ref.split('.')
                            row_idx = int(parts[0])
                            col_idx = int(parts[1])
                            
                            cell = table.Cell(row_idx, col_idx)
                            abs_bbox = [cell.Shape.Left, cell.Shape.Top, 
                                   cell.Shape.Width, cell.Shape.Height]
                            rel_bbox = self.bbox_to_relative(abs_bbox, slide_width, slide_height)
                            bboxes.append(rel_bbox)
                        
                        table_data = {
                            "name": self.name_entry.get().strip(),
                            "path": presentation_path,
                            "slide_number": slide_number,
                            "slide_width": slide_width,
                            "slide_height": slide_height,
                            "selection_type": "table_cells",
                            "shape_ids": table_id,
                            "table_rows": "",
                            "table_cols": "",
                            "table_cells": ",".join(values),
                            "bbox": bboxes,
                            "color_rgb": [],
                            "text": []
                        }
                    
                    elif section_type == "rows":
                        for row_idx in values:
                            # Get bounding box of entire row
                            first_cell = table.Cell(row_idx, 1)
                            last_cell = table.Cell(row_idx, table.Columns.Count)
                            
                            left = first_cell.Shape.Left
                            top = first_cell.Shape.Top
                            width = (last_cell.Shape.Left + last_cell.Shape.Width) - left
                            height = first_cell.Shape.Height
                            
                            abs_bbox = [left, top, width, height]
                            rel_bbox = self.bbox_to_relative(abs_bbox, slide_width, slide_height)
                            bboxes.append(rel_bbox)
                        
                        table_data = {
                            "name": self.name_entry.get().strip(),
                            "path": presentation_path,
                            "slide_number": slide_number,
                            "slide_width": slide_width,
                            "slide_height": slide_height,
                            "selection_type": "table_rows",
                            "shape_ids": table_id,
                            "table_rows": ",".join(map(str, values)),
                            "table_cols": "",
                            "table_cells": "",
                            "bbox": bboxes,
                            "color_rgb": [],
                            "text": []
                        }
                    
                    elif section_type == "cols":
                        for col_idx in values:
                            # Get bounding box of entire column
                            first_cell = table.Cell(1, col_idx)
                            last_cell = table.Cell(table.Rows.Count, col_idx)
                            
                            left = first_cell.Shape.Left
                            top = first_cell.Shape.Top
                            width = first_cell.Shape.Width
                            height = (last_cell.Shape.Top + last_cell.Shape.Height) - top
                            
                            abs_bbox = [left, top, width, height]
                            rel_bbox = self.bbox_to_relative(abs_bbox, slide_width, slide_height)
                            bboxes.append(rel_bbox)
                        
                        table_data = {
                            "name": self.name_entry.get().strip(),
                            "path": presentation_path,
                            "slide_number": slide_number,
                            "slide_width": slide_width,
                            "slide_height": slide_height,
                            "selection_type": "table_cols",
                            "shape_ids": table_id,
                            "table_rows": "",
                            "table_cols": ",".join(map(str, values)),
                            "table_cells": "",
                            "bbox": bboxes,
                            "color_rgb": [],
                            "text": []
                        }
                    
                    json_data.append(table_data)
            
            elif mode == "rectangle":
                # Handle drag rectangle mode
                if not self.rectangle_bboxes:
                    messagebox.showwarning("Warning", "No rectangles drawn")
                    return
                
                rectangle_data = {
                    "name": self.name_entry.get().strip(),
                    "path": presentation_path,
                    "slide_number": slide_number,
                    "slide_width": slide_width,
                    "slide_height": slide_height,
                    "selection_type": "rectangle",
                    "bbox": self.rectangle_bboxes,
                    "table_rows": "",
                    "table_cols": "",
                    "table_cells": "",
                    "shape_ids": []
                }
                json_data.append(rectangle_data)
            
            # Save JSON
            json_path = os.path.join(json_dir, f"{file_id}.json")
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, indent=2, ensure_ascii=False)
            
            # Show success message and clear form
            self.show_status_message("Data saved")
            self.clear_form()
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    
    def capture_slide_screenshot(self, slide, output_path):
        """Capture screenshot of the current slide"""
        try:
            # Convert to absolute path for PowerPoint Export
            abs_output_path = os.path.abspath(output_path)
            
            # Export slide as image directly
            slide.Export(abs_output_path, "PNG")
            
        except Exception as e:
            print(f"Screenshot error: {str(e)}")
            # Fallback: just create a placeholder
            from PIL import Image
            img = Image.new('RGB', (800, 600), color='white')
            img.save(output_path)
    
    def on_exit(self):
        """Handle exit button click"""
        self.root.destroy()


def main():
    root = tk.Tk()
    app = PowerPointShapeCaptureApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

