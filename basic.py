import pandas as pd
import json
import csv
from datetime import datetime

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
    # concat shape data and group list
    data_list = add_unnamed_shapes(shape_data, group_list)

    # add slide dimensions and test index to each dict
    for x in data_list:
        x['slide_height'] = slide_dimensions['height']
        x['slide_width'] = slide_dimensions['width']
        x['test_index'] = test_index  
    # Create dataframe
    return data_list



def save_to_csv(dl, test_index):
    df = pd.DataFrame(dl)
    # Handle commas 
    df["shape_id"] = df["shape_id"].apply(json.dumps)
    for x in ["shape_id", "text", "label"]:
        df[x] = df[x].str.replace('"', '""').apply(lambda y: f'"{y}"')
    
    # round to 2 decimal places
    for x in ["top", "left", "right", "bottom", "width", "height"]:
        df[x] = df[x].round(2)
    #reorder cols
    df = df[["test_index", "index",  "label", "shape_id", "text", "top", "left", "right", "bottom", "width", "height", "slide_height", "slide_width"]]
    # Get filename and save
    test_index_str = str(test_index).zfill(3)
    filename = f"test_{str(test_index_str)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    filepath = f'./csv_files/{filename}'
    # Sort
    df = df.sort_values(by="index", ascending=True)
    df.to_csv(filepath, index=False, encoding="utf-8", quoting=csv.QUOTE_NONE, escapechar='\\')
    return df
    