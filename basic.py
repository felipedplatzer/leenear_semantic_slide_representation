import pandas as pd
import json
import csv
from datetime import datetime

def process_groups(shape_data, group_list, test_index, slide_dimensions):
    # concat shape data and group list
    data_list = shape_data + group_list

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
    