import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from PIL import Image, ImageDraw, ImageTk


class TestViewerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerPoint Capture Test Viewer")
        self.root.geometry("1000x800")
        
        # Current data
        self.current_data = None
        self.current_image = None
        self.box_thickness = 5
        
        # Main container
        main_frame = ttk.Frame(root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Top section - Test ID input
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(top_frame, text="Enter Test ID:").pack(side=tk.LEFT, padx=5)
        self.test_id_entry = ttk.Entry(top_frame, width=40)
        self.test_id_entry.pack(side=tk.LEFT, padx=5)
        
        self.ok_button = ttk.Button(top_frame, text="OK", command=self.load_test_data)
        self.ok_button.pack(side=tk.LEFT, padx=5)
        
        self.exit_button = ttk.Button(top_frame, text="Exit", command=self.on_exit)
        self.exit_button.pack(side=tk.LEFT, padx=5)
        
        # Box thickness controls
        thickness_frame = ttk.Frame(top_frame)
        thickness_frame.pack(side=tk.LEFT, padx=20)
        
        ttk.Label(thickness_frame, text="Box Thickness:").pack(side=tk.LEFT, padx=5)
        
        self.thickness_var = tk.IntVar(value=5)
        self.thickness_slider = ttk.Scale(
            thickness_frame, 
            from_=1, 
            to=20, 
            orient=tk.HORIZONTAL,
            variable=self.thickness_var,
            length=150
        )
        self.thickness_slider.pack(side=tk.LEFT, padx=5)
        
        self.thickness_label = ttk.Label(thickness_frame, text="5")
        self.thickness_label.pack(side=tk.LEFT, padx=2)
        
        # Update label when slider moves
        self.thickness_slider.config(command=self.update_thickness_label)
        
        self.apply_button = ttk.Button(thickness_frame, text="Apply", command=self.apply_thickness)
        self.apply_button.pack(side=tk.LEFT, padx=5)
        
        # Display section
        display_frame = ttk.LabelFrame(main_frame, text="Test Data Viewer", padding=10)
        display_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Prompt label
        prompt_label_frame = ttk.Frame(display_frame)
        prompt_label_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(prompt_label_frame, text="Prompt:", font=('Arial', 10, 'bold')).pack(anchor=tk.W)
        self.prompt_text = tk.Text(prompt_label_frame, height=4, wrap=tk.WORD, state=tk.DISABLED)
        self.prompt_text.pack(fill=tk.X, pady=(5, 0))
        
        # Image display with scrollbar
        image_frame = ttk.Frame(display_frame)
        image_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas for image with scrollbars
        self.canvas = tk.Canvas(image_frame, bg='gray')
        v_scrollbar = ttk.Scrollbar(image_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        h_scrollbar = ttk.Scrollbar(image_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.canvas.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        image_frame.grid_rowconfigure(0, weight=1)
        image_frame.grid_columnconfigure(0, weight=1)
        
        self.image_on_canvas = None
        
        # Key bindings
        self.root.bind('<Return>', lambda e: self.load_test_data())
        self.root.bind('<Escape>', lambda e: self.on_exit())
        
        # Focus on test ID entry
        self.test_id_entry.focus()
    
    def update_thickness_label(self, value):
        """Update the thickness label when slider moves"""
        self.thickness_label.config(text=str(int(float(value))))
    
    def apply_thickness(self):
        """Apply the new thickness and redraw"""
        self.box_thickness = int(self.thickness_var.get())
        if self.current_data and self.current_image:
            self.draw_and_display_image()
    
    def load_test_data(self):
        """Load test data from JSON file"""
        test_id = self.test_id_entry.get().strip()
        if not test_id:
            messagebox.showwarning("Warning", "Please enter a test ID")
            return
        
        # Look for JSON file
        json_path = os.path.join("resources", "json", f"{test_id}.json")
        if not os.path.exists(json_path):
            messagebox.showerror("Error", f"Test ID '{test_id}' not found in resources/json/")
            return
        
        # Load JSON data
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                self.current_data = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load JSON: {str(e)}")
            return
        
        # Look for image file
        img_path = os.path.join("resources", "img", f"{test_id}.png")
        if not os.path.exists(img_path):
            messagebox.showerror("Error", f"Image file '{test_id}.png' not found in resources/img/")
            return
        
        # Load image
        try:
            self.current_image = Image.open(img_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load image: {str(e)}")
            return
        
        # Update prompt
        self.update_prompt()
        
        # Draw and display image with bounding boxes
        self.draw_and_display_image()
        
        # Update thickness
        self.box_thickness = int(self.thickness_var.get())
    
    def update_prompt(self):
        """Update the prompt text area"""
        if not self.current_data:
            return
        
        # Build prompt from data
        prompt_lines = []
        for idx, item in enumerate(self.current_data):
            name = item.get('name', 'N/A')
            selection_type = item.get('selection_type', 'N/A')
            path = item.get('path', 'N/A')
            slide_num = item.get('slide_number', 'N/A')
            
            prompt_lines.append(f"[{idx+1}] Name: {name} | Type: {selection_type} | Slide: {slide_num}")
            
            if selection_type == "shape":
                shape_ids = item.get('shape_ids', [])
                texts = item.get('text', [])
                prompt_lines.append(f"    Shape IDs: {shape_ids}")
                if texts and any(texts):
                    prompt_lines.append(f"    Texts: {texts}")
            elif selection_type == "table_cells":
                cells = item.get('table_cells', '')
                prompt_lines.append(f"    Table ID: {item.get('shape_ids', 'N/A')} | Cells: {cells}")
            elif selection_type == "table_rows":
                rows = item.get('table_rows', '')
                prompt_lines.append(f"    Table ID: {item.get('shape_ids', 'N/A')} | Rows: {rows}")
            elif selection_type == "table_cols":
                cols = item.get('table_cols', '')
                prompt_lines.append(f"    Table ID: {item.get('shape_ids', 'N/A')} | Cols: {cols}")
            
            prompt_lines.append("")  # Empty line between items
        
        prompt_text = "\n".join(prompt_lines)
        
        self.prompt_text.config(state=tk.NORMAL)
        self.prompt_text.delete(1.0, tk.END)
        self.prompt_text.insert(1.0, prompt_text)
        self.prompt_text.config(state=tk.DISABLED)
    
    def draw_and_display_image(self):
        """Draw bounding boxes on image and display"""
        if not self.current_image or not self.current_data:
            return
        
        # Create a copy of the image to draw on
        img_with_boxes = self.current_image.copy()
        draw = ImageDraw.Draw(img_with_boxes, 'RGBA')
        
        # Get image dimensions
        img_width, img_height = img_with_boxes.size
        
        # Draw bounding boxes for all items
        for item in self.current_data:
            bboxes = item.get('bbox', [])
            
            # Get slide dimensions (for converting relative to absolute coords)
            slide_width = item.get('slide_width', img_width)
            slide_height = item.get('slide_height', img_height)
            
            # Calculate scale factors (in case image was resized during export)
            scale_x = img_width / slide_width
            scale_y = img_height / slide_height
            
            for bbox in bboxes:
                if len(bbox) != 4:
                    continue
                
                # bbox is now in relative coordinates (0-1 range)
                rel_left, rel_top, rel_width, rel_height = bbox
                
                # Convert to absolute coordinates using slide dimensions and scale
                left = rel_left * slide_width * scale_x
                top = rel_top * slide_height * scale_y
                width = rel_width * slide_width * scale_x
                height = rel_height * slide_height * scale_y
                
                # Calculate rectangle coordinates
                x1 = int(left)
                y1 = int(top)
                x2 = int(left + width)
                y2 = int(top + height)
                
                # Draw transparent red rectangle with red outline
                # First draw the transparent fill
                overlay = Image.new('RGBA', img_with_boxes.size, (0, 0, 0, 0))
                overlay_draw = ImageDraw.Draw(overlay)
                overlay_draw.rectangle([x1, y1, x2, y2], fill=(255, 0, 0, 30))
                img_with_boxes = Image.alpha_composite(img_with_boxes.convert('RGBA'), overlay)
                
                # Then draw the red outline
                draw = ImageDraw.Draw(img_with_boxes)
                for i in range(self.box_thickness):
                    draw.rectangle([x1-i, y1-i, x2+i, y2+i], outline=(255, 0, 0, 255))
        
        # Convert to PhotoImage and display
        img_with_boxes = img_with_boxes.convert('RGB')
        self.photo_image = ImageTk.PhotoImage(img_with_boxes)
        
        # Update canvas
        self.canvas.delete("all")
        self.image_on_canvas = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo_image)
        
        # Update scroll region
        self.canvas.config(scrollregion=self.canvas.bbox(tk.ALL))
    
    def on_exit(self):
        """Handle exit button click"""
        self.root.destroy()


def main():
    root = tk.Tk()
    app = TestViewerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

