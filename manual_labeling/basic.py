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


def remove_ungrouped_individual_shapes(dl):
    """
    Remove individual shapes that don't belong to any group.
    Individual shapes are those with section_type='individual_shape' that:
    - Have only 1 shape_id in their shape_id array
    - Are at the root level (index has only 2 digits like "01", not "01.01")
    
    These shapes should be removed if they haven't been used elsewhere in the tree.
    """
    # Collect all shape_ids that are referenced as children (non-root level)
    referenced_shape_ids = set()
    for item in dl:
        if 'index' in item and item['index'].count('.') > 0:  # Not a root-level item
            if 'shape_id' in item and isinstance(item['shape_id'], list):
                referenced_shape_ids.update(item['shape_id'])
    
    # Filter out individual shapes that are root-level and not referenced elsewhere
    filtered_dl = []
    for item in dl:
        # Check if this is a root-level individual shape
        is_root_individual = (
            item.get('section_type') == 'individual_shape' and
            'index' in item and 
            item['index'].count('.') == 0  # Root level (e.g., "01" not "01.01")
        )
        
        if is_root_individual:
            # Check if any of its shape_ids are referenced elsewhere
            item_shape_ids = item.get('shape_id', [])
            if isinstance(item_shape_ids, list):
                # If any shape_id is referenced elsewhere, keep it
                if any(sid in referenced_shape_ids for sid in item_shape_ids):
                    filtered_dl.append(item)
                # Otherwise, skip it (don't add to filtered_dl)
            else:
                # Single shape_id (not a list)
                if item_shape_ids in referenced_shape_ids:
                    filtered_dl.append(item)
        else:
            # Not a root-level individual shape, keep it
            filtered_dl.append(item)
    
    return filtered_dl

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
    