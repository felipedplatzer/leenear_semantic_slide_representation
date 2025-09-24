import pandas as pd
import json
import csv
import os
from datetime import datetime
import structure_text

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
        parent_shape_i, parent_shape = get_parent_shape(x, dl)
        # Add prefix 
        for y in sections_in_shape: 
            y['index'] = parent_shape['index'] + '.' + y['index']
            y['shape_id'] = [y['shape_id']] # fix
        # Insert into array
        dl_new[parent_shape_i:parent_shape_i] = sections_in_shape
    return dl_new



def save_to_csv(dl, test_index):
    df = pd.DataFrame(dl)
    # add missing cols
    col_list = ["test_index", "index",  "label", "shape_id", "start_char", "end_char", "text", "top", "left", "right", "bottom", "width", "height", "slide_height", "slide_width"]
    for x in col_list:
        if x not in df.columns:
            df[x] = ''
    # fill na
    df = df.fillna('')
    # Handle commas 
    df["shape_id"] = df["shape_id"].apply(json.dumps)
    for x in ["shape_id", "text", "label"]:
        df[x] = df[x].str.replace('"', '""').apply(lambda y: f'"{y}"')
    
    # round to 2 decimal places
    for x in ["top", "left", "right", "bottom", "width", "height"]:
        df[x] = df[x].round(2)
    #reorder cols
    df = df[["test_index", "index",  "label", "shape_id", "start_char", "end_char", "text", "top", "left", "right", "bottom", "width", "height", "slide_height", "slide_width"]]
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
    