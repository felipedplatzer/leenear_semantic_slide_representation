import json
from PIL import Image
import base64
import os
import win32com.client

# Get the directory where this script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

slide_folder = os.path.join(script_dir, 'resources', 'input_slides')
picture_folder = os.path.join(script_dir, 'resources', 'input_slide_pictures')
training_data_folder = os.path.join(script_dir, 'resources', 'manual_sections_json_trees')

def get_script_dir():
    """Get the directory where the current script is located."""
    return os.path.dirname(os.path.abspath(__file__))

def get_project_path(*path_parts):
    """Get a path relative to the project root (where this script is located)."""
    return os.path.join(script_dir, *path_parts)

def open_slide(rump):
    slide_file = f'{slide_folder}/{rump}.pptx'
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.DisplayAlerts = False
    # ppt_app.Visible = False
    presentation = ppt_app.Presentations.Open(slide_file, WithWindow=False)
    slide = presentation.Slides(1)
    return presentation, slide

def get_active_slide():
    """Get the currently active slide from PowerPoint"""
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    presentation = ppt_app.ActivePresentation
    slide = ppt_app.ActiveWindow.View.Slide  # Get the currently active/selected slide
    return presentation, slide


def get_json_from_llm_response(text):
    if '```json' in text:
        json_start = text.find('```json') + 7
        json_end = text.find('```', json_start)
        json_text = text[json_start:json_end].strip()
    else:
        json_text = text
    # Parse the JSON response
    data = json.loads(json_text)
    return data

def get_bounds_from_shape_ids(dl, elements_data):
    for x in dl:
        shape_ids = x['shape_id']
        if len(shape_ids) == 0:
            print(f"No shape ids found for {x['label']}")
            continue
        shapes = [y for y in elements_data if y['shape_id'] in shape_ids]
        x['left'] = min([x['left'] for x in shapes])
        x['top'] = min([x['top'] for x in shapes])
        x['right'] = max([x['right'] for x in shapes])
        x['bottom'] = max([x['bottom'] for x in shapes])
    return dl



def get_picture(rump, active_bool):
    if os.path.exists(f'{picture_folder}/{rump}.png'):
        return Image.open(f'{picture_folder}/{rump}.png')
    else:
        if active_bool:
            presentation, slide = get_active_slide()
        else:
            presentation, slide = open_slide(rump)
        png_file = f'{picture_folder}/{rump}.png'
        slide.Export(png_file, "PNG")
        if active_bool == False:
            presentation.Close()
        return Image.open(png_file)

def get_sections(rump):
    return json.load(open(os.path.join(training_data_folder, f'{rump}.json')))


def encode_image(image_path):
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def get_shape_info(shape):
    shape_info = {
        'shape_id': str(shape.Id),
        'shape_type': shape.Type,
        'top': shape.Top,
        'left': shape.Left,
        'width': shape.Width,
        'height': shape.Height,
        'right': shape.Left + shape.Width,
        'bottom': shape.Top + shape.Height,
        'name': shape.Name,
    }
    return shape_info

def get_text_pos(shape):
    return {
        'top': shape.TextFrame2.TextRange.BoundTop,
        'left': shape.TextFrame2.TextRange.BoundLeft,
        'width': shape.TextFrame2.TextRange.BoundWidth,
        'height': shape.TextFrame2.TextRange.BoundHeight,
        'right': shape.TextFrame2.TextRange.BoundLeft + shape.TextFrame2.TextRange.BoundWidth,
        'bottom': shape.TextFrame2.TextRange.BoundTop + shape.TextFrame2.TextRange.BoundHeight,
    }


def classify_shapes(slide):
    textboxes = []
    tables = []
    charts = []
    pictures = []
    for i,x in enumerate(slide.Shapes):
        if x.Type == 17 and hasattr(x, 'TextFrame') and x.TextFrame.HasText and x.Fill.ForeColor.RGB == 0xFFFFFF:
            textboxes.append(x)
        elif x.Type == 19:
            tables.append(x)
        elif x.Type == 3:
            charts.append(x)
        else:
            pictures.append(x)
    return textboxes, tables, charts, pictures


def flatten_tree(tree, parent_index=""):
    """
    Convert a tree structure back to a flat list with hierarchical indices.
    
    Input: [{'label': 'Introduction', 'sections': [{'label': 'Overview', 'sections': []}]}]
    Output: [{'index': '1', 'label': 'Introduction'}, {'index': '1.1', 'label': 'Overview'}]
    """
    flat_list = []
    
    for i, node in enumerate(tree, 1):
        # Create current index
        current_index = f"{parent_index}.{i}" if parent_index else str(i)
        
        # Add current node to flat list
        flat_list.append({
            'index': current_index,
            'label': node['label']
        })
        
        # Recursively process sections
        if 'sections' in node and node['sections']:
            flat_list.extend(flatten_tree(node['sections'], current_index))
    
    return flat_list


def get_cell_info(shape, i_row, i_cell):
    x =  {
        'shape_id': str(shape.Id) + '.' + str(i_row + 1) + '.' + str(i_cell + 1), # Table ID.Row Index.Cell Index
        'shape_type': 'cell',
        'top': shape.Table.Rows[i_row].Cells[i_cell].Shape.Top,
        'left': shape.Table.Rows[i_row].Cells[i_cell].Shape.Left,
        'width': shape.Table.Rows[i_row].Cells[i_cell].Shape.Width,
        'height': shape.Table.Rows[i_row].Cells[i_cell].Shape.Height,
        'right': shape.Table.Rows[i_row].Cells[i_cell].Shape.Left + shape.Table.Rows[i_row].Cells[i_cell].Shape.Width,
        'bottom': shape.Table.Rows[i_row].Cells[i_cell].Shape.Top + shape.Table.Rows[i_row].Cells[i_cell].Shape.Height,
        'name': shape.Name  +'.'+ str(i_row + 1) + '.' + str(i_cell + 1),
        'i_row': i_row + 1,
        'i_cell': i_cell + 1,
    }
    if hasattr(shape.Table.Rows[i_row].Cells[i_cell].Shape, 'TextFrame'):
        x['text'] = shape.Table.Rows[i_row].Cells[i_cell].Shape.TextFrame.TextRange.Text.strip()
    else:
        x['text'] = ''
    return x

def table_to_cells(table_ppt):
    cells = []
    for row in table_ppt.Table.Rows:
        for cell in row.Cells:
            cells.append(cell)
    return cells

def get_shape_images(slide):
    """
    Extract images of each shape in the slide and return as a dictionary.
    
    Args:
        slide: PowerPoint slide object
        
    Returns:
        dict: Dictionary with shape_id as key and PIL Image as value
    """
    import tempfile
    import os
    from io import BytesIO
    
    shape_images = {}
    
    # Classify shapes into different types
    textboxes, tables, charts, pictures = classify_shapes(slide)
    
    # Process each shape type
    all_shapes = []
    all_shapes.extend(textboxes)
    all_shapes.extend(tables)
    all_shapes.extend(charts)
    all_shapes.extend(pictures)
    
    for shape in all_shapes:
        try:
            # Get shape info to extract ID and bounds
            shape_info = get_shape_info(shape)
            shape_id = shape_info['shape_id']
            
            # Create a temporary file for the shape export
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                temp_path = temp_file.name
            
            try:
                # Export the shape as PNG
                shape.Export(temp_path, 2)
                
                # Open with PIL and convert to bytes
                with Image.open(temp_path) as pil_image:
                    # Convert to RGB if necessary
                    if pil_image.mode != 'RGB':
                        pil_image = pil_image.convert('RGB')
                    
                    # Convert to bytes
                    buffer = BytesIO()
                    pil_image.save(buffer, format='PNG')
                    shape_images[shape_id] = buffer.getvalue()
                    
            finally:
                # Clean up temporary file
                if os.path.exists(temp_path):
                    os.unlink(temp_path)
                    
        except Exception as e:
            print(f"Error processing shape {shape.Id}: {e}")
            continue
    
    # Also process table cells
    for table in tables:
        try:
            for i_row, row in enumerate(table.Table.Rows):
                for i_cell, cell in enumerate(row.Cells):
                    try:
                        # Get cell info
                        cell_info = get_cell_info(table, i_row, i_cell)
                        cell_id = cell_info['shape_id']
                        
                        # Create a temporary file for the cell export
                        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                            temp_path = temp_file.name
                        
                        try:
                            # Export the cell as PNG
                            cell.Shape.Export(temp_path, 2)
                            
                            # Open with PIL and convert to bytes
                            with Image.open(temp_path) as pil_image:
                                # Convert to RGB if necessary
                                if pil_image.mode != 'RGB':
                                    pil_image = pil_image.convert('RGB')
                                
                                # Convert to bytes
                                buffer = BytesIO()
                                pil_image.save(buffer, format='PNG')
                                shape_images[cell_id] = buffer.getvalue()
                                
                        finally:
                            # Clean up temporary file
                            if os.path.exists(temp_path):
                                os.unlink(temp_path)
                                
                    except Exception as e:
                        print(f"Error processing cell {i_row}.{i_cell} in table {table.Id}: {e}")
                        continue
                        
        except Exception as e:
            print(f"Error processing table {table.Id}: {e}")
            continue
    
    return shape_images