import pandas as pd
import json
import csv
import os
from datetime import datetime
import structure_text
import structure_table

def is_orphan(shape_dict, group_list):
    shape_id = shape_dict['shape_id']
    for y in group_list:
        if len(y['shape_id']) == 1 and y['shape_id'][0] == shape_id:
            return False
    return True

def add_unnamed_shapes(shape_data, group_list):
    unnamed_shapes = [x for x in shape_data if is_orphan(x, group_list)] #shapes that don't have a name, i.e. there's no group with 1 shape (only containint that shape) wit ha name 
    for x in unnamed_shapes:
        x['shape_id'] = [x['shape_id']]
        group_list.append(x)
    return group_list


def process_groups(shape_data, group_list, test_index, slide_dimensions):
    """
    Process groups by adding unnamed shapes and adding required metadata fields.
    
    Args:
        shape_data: List of shape data dictionaries
        group_list: List of grouped shapes
        test_index: Test index for the current session
        slide_dimensions: Dictionary with 'height' and 'width' of the slide
    
    Returns:
        List of processed group data with metadata fields added
    """
    # Add unnamed shapes to the group list
    dl = add_unnamed_shapes(shape_data, group_list)
    
    # Add required metadata fields to each item
    for x in dl:
        x['slide_height'] = slide_dimensions['height']
        x['slide_width'] = slide_dimensions['width']
        x['test_index'] = test_index
    
    return dl


def get_parent_shape(this_id, dl):
    for i, x in enumerate(dl):
        if len(x['shape_id']) == 1 and x['shape_id'][0] == this_id:
            return i, x
    return None


def add_text_sections(dl, text_section_list):
    # Group into shape id's
    shape_ids_with_text = list(set([x['shape_id'] for x in text_section_list]))
    dl_new = dl.copy()
    # For each shape id, get tree structure
    for x in shape_ids_with_text:
        # Get all sections in this shape
        sections_in_shape = [y for y in text_section_list if y['shape_id'] == x]
        # Get tree structure
        sections_in_shape = structure_text.structure_text_sections(sections_in_shape)
        # Get parent shape
        result = get_parent_shape(x, dl)
        if result is None:
            print(f"Warning: Could not find parent shape for text section with shape_id={x}")
            continue
        parent_shape_i, parent_shape = result
        if 'index' not in parent_shape:
            print(f"Warning: Parent shape {x} does not have an index")
            continue
        # Add prefix 
        for y in sections_in_shape: 
            y['index'] = parent_shape['index'] + '.' + y['index']
            y['shape_id'] = [y['shape_id']] # fix
        # Insert into array
        dl_new[parent_shape_i:parent_shape_i] = sections_in_shape
    return dl_new


def get_next_child_index_for_parent(sections_list, parent_index):
    """
    Get the next available child index for a given parent.
    E.g., if parent is "01.01" and existing children are "01.01.01", "01.01.02",
    this returns "03".
    """
    # Find all children of this parent
    existing_children = []
    for section in sections_list:
        if 'index' in section and section['index'].startswith(parent_index + '.'):
            # Extract the immediate child index
            child_part = section['index'][len(parent_index) + 1:]
            # Get only the first level (in case of deeper nesting)
            if '.' in child_part:
                child_part = child_part.split('.')[0]
            existing_children.append(child_part)
    
    # Find the next available number
    if not existing_children:
        return "01"
    
    # Convert to integers, find max, and increment
    max_child = max([int(c) for c in existing_children if c.isdigit()])
    return str(max_child + 1).zfill(2)

def add_table_sections(dl, table_labels_list, individual_shape_label_map=None):
    # Group into shape names (table shape IDs)
    shape_names_with_tables = list(set([x['shape_name'] for x in table_labels_list]))
    dl_new = dl.copy()
    # For each table shape ID, get tree structure
    for x in shape_names_with_tables:
        # Get all sections in this table
        sections_in_table = [y for y in table_labels_list if y['shape_name'] == x]
        # Get tree structure
        sections_in_table = structure_table.structure_table_sections(sections_in_table)
        # Get parent shape
        result = get_parent_shape(x, dl)
        if result is None:
            print(f"Warning: Could not find parent table for table section with shape_name={x}")
            continue
        parent_shape_i, parent_shape = result
        if 'index' not in parent_shape:
            print(f"Warning: Parent table {x} does not have an index")
            continue
        # Add prefix and handle overlaid shapes
        sections_with_overlaid = []
        for y in sections_in_table: 
            y['index'] = parent_shape['index'] + '.' + y['index']
            # Note: shape_id is already set as an array in main_gui.py
            sections_with_overlaid.append(y)
            
            # If this section has overlaid shapes, add them as children
            if 'overlaid_shapes' in y and y['overlaid_shapes']:
                for overlaid_shape_id in y['overlaid_shapes']:
                    # Find the original shape in dl_new to copy its attributes
                    original_shape = None
                    for shape in dl_new:
                        if isinstance(shape.get('shape_id'), list) and overlaid_shape_id in shape['shape_id']:
                            original_shape = shape
                            break
                        elif shape.get('shape_id') == overlaid_shape_id:
                            original_shape = shape
                            break
                    
                    if original_shape:
                        # Create a copy of the shape as a child of this section
                        child_shape = original_shape.copy()
                        # Get next child index for this parent
                        child_index = get_next_child_index_for_parent(sections_with_overlaid, y['index'])
                        child_shape['index'] = y['index'] + '.' + child_index
                        
                        # Apply label from individual_shape_label_map if available
                        if individual_shape_label_map:
                            # Check if this shape_id has a label
                            if overlaid_shape_id in individual_shape_label_map:
                                child_shape['label'] = individual_shape_label_map[overlaid_shape_id]
                        
                        sections_with_overlaid.append(child_shape)
                
                # Remove overlaid_shapes attribute so it doesn't appear in CSV
                del y['overlaid_shapes']
        
        # Insert into array
        dl_new[parent_shape_i:parent_shape_i] = sections_with_overlaid
    return dl_new   


def get_exclude_bool(x, dl):
    """
    Check if an individual shape should be excluded.
    Returns True if the shape should be excluded (is referenced elsewhere).
    Only applies to root-level items (index without a dot).
    """
    # If not an individual_shape, don't exclude
    if x.get('section_type') != 'individual_shape':
        return False
    
    # Only exclude root-level items (index without a dot)
    # If index has a dot, it's a nested item - don't remove it
    index = x.get('index', '')
    if '.' in index:
        return False
    
    # Get the first shape_id from this item
    shape_ids = x.get('shape_id', [])
    if not shape_ids:
        return False
    
    if isinstance(shape_ids, list):
        target_shape_id = shape_ids[0]
    else:
        target_shape_id = shape_ids
    
    # Check if this shape_id appears in any other item in dl
    for y in dl:
        # Skip comparing with itself
        if y is x:
            continue
        
        y_shape_ids = y.get('shape_id', [])
        if isinstance(y_shape_ids, list):
            if target_shape_id in y_shape_ids:
                return True
        else:
            if target_shape_id == y_shape_ids:
                return True
    
    return False

def remove_ungrouped_individual_shapes(dl):
    """
    Remove individual shapes that don't belong to any group.
    Individual shapes that appear in other groups/sections will be excluded.
    """
    indices_to_remove = [i for i, x in enumerate(dl) if get_exclude_bool(x, dl) == True]
    dl_new = [x for i, x in enumerate(dl) if i not in indices_to_remove]
    return dl_new

def save_to_csv(dl, test_index):
    df = pd.DataFrame(dl)
    # add missing cols
    col_list = ["test_index", "index",  "label", "shape_id", "section_type", "start_char", "end_char", "cells", "text", "top", "left", "right", "bottom", "width", "height", "slide_height", "slide_width"]
    for x in col_list:
        if x not in df.columns:
            df[x] = ''
    # fill na
    df = df.fillna('')
    # Handle commas and arrays
    df["shape_id"] = df["shape_id"].apply(json.dumps)
    # Handle cells array (convert to JSON string if it's a list)
    if 'cells' in df.columns:
        df["cells"] = df["cells"].apply(lambda x: json.dumps(x) if isinstance(x, list) else x)
    for x in ["shape_id", "cells", "text", "label"]:
        if x in df.columns:
            df[x] = df[x].str.replace('"', '""').apply(lambda y: f'"{y}"')
    
    # round to 2 decimal places
    for x in ["top", "left", "right", "bottom", "width", "height"]:
        if x in df.columns:
            df[x] = df[x].round(2)
    #reorder cols
    df = df[["test_index", "index",  "label", "shape_id", "section_type", "start_char", "end_char", "cells", "text", "top", "left", "right", "bottom", "width", "height", "slide_height", "slide_width"]]
    # Get filename and save
    test_index_str = str(test_index).zfill(3)
    filename = f"test_{str(test_index_str)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    # Get the directory where this script is located and navigate to resources/manual_sections_csv
    script_dir = os.path.dirname(os.path.abspath(__file__))
    # Go up one level from manual_labeling to the project root, then into resources/manual_sections_csv
    project_root = os.path.dirname(script_dir)
    csv_dir = os.path.join(project_root, 'resources', 'manual_sections_csv')
    filepath = os.path.join(csv_dir, filename)
    # Sort
    df = df.sort_values(by="index", ascending=True)
    df.to_csv(filepath, index=False, encoding="utf-8", quoting=csv.QUOTE_NONE, escapechar='\\', na_rep='')
    return df
    