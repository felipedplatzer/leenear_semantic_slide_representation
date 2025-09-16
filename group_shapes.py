import win32com.client
import resources
import re
from datetime import datetime
import os


def get_selected_shapes_info(selection):
    """
    Get information about currently selected shapes in PowerPoint.
    Returns list of shape information including ID and coordinates.
    """
    try:

        if selection.Type != 2:  # ppSelectionShapes
            print("No shapes are currently selected.")
            return []
        
        selected_shapes = []
        for i in range(1, selection.ShapeRange.Count + 1):
            shape = selection.ShapeRange.Item(i)
            
            # Get shape properties
            shape_id = shape.Id
            left = shape.Left
            top = shape.Top
            width = shape.Width
            height = shape.Height
            right = left + width
            bottom = top + height
            
            # Extract shape ID from text
            shape_info = {
                'shape_id': shape_id,  # ID of the red box
                'left': left,
                'top': top,
                'right': right,
                'bottom': bottom,
                'width': width,
                'height': height,
            }
            
            selected_shapes.append(shape_info)
        
        return selected_shapes
        
    except Exception as e:
        print(f"Error getting selected shapes: {e}")
        return []

def calculate_bounds(shapes):
    """
    Calculate the bounding box for a group of shapes.
    Returns min left, min top, max right, max bottom.
    """
    if not shapes:
        return None
    
    min_left = min(shape['left'] for shape in shapes)
    min_top = min(shape['top'] for shape in shapes)
    max_right = max(shape['right'] for shape in shapes)
    max_bottom = max(shape['bottom'] for shape in shapes)
    
    return {
        'left': min_left,
        'top': min_top,
        'right': max_right,
        'bottom': max_bottom,
        'width': max_right - min_left,
        'height': max_bottom - min_top
    }

def group_selected_shapes(group_name, selection):
    """
    Main function to group selected shapes based on user input.
    """    
    print()
    
    # Get currently selected shapes
    selected_shapes = get_selected_shapes_info(selection)
    
    # Calculate group bounds
    bounds = calculate_bounds(selected_shapes)
    
    # Create group entry
    group_entry = {
        'label': group_name,
        'shape_count': len(selected_shapes),
        'shape_id': [shape['shape_id'] for shape in selected_shapes],
        'top': bounds['top'],
        'left': bounds['left'],
        'right': bounds['right'],
        'bottom': bounds['bottom'],
        'width': bounds['width'],
        'height': bounds['height'],
        'timestamp': datetime.now()
    }
    
    return group_entry



"""
def get_groups_from_file(filename=None):
    # LOAD GROUPS FROM FILE
    try:
        if filename is None:
            # Find the most recent groups file
            import glob
            pattern = os.path.join(os.path.dirname(__file__), "shape_groups_*.txt")
            files = glob.glob(pattern)
            if not files:
                print("No groups files found.")
                return []
            filename = max(files, key=os.path.getctime)
        
        print(f"Loading groups from: {filename}")
        # This is a simple implementation - you might want to parse the file more robustly
        with open(filename, 'r', encoding='utf-8') as f:
            content = f.read()
            print(content)
        
        return content
            except Exception as e:
        print(f"Error loading groups from file: {e}")
        return []"""