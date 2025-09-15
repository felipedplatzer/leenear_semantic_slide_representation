"""
Complete workflow for shape analysis and grouping:
1. Create red boxes for shapes on active slide
2. Group selected red boxes interactively
"""
import structure_shapes
import create_shape_boxes
import group_shapes
import os
from datetime import datetime
import resources
import pandas as pd
import csv
import json

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
    df["text"] = df["text"].str.replace('"', '""').apply(lambda x: f'"{x}"')
    df["shape_id"] = df["shape_id"].apply(json.dumps)
    df["shape_id"] = df["shape_id"].str.replace('"', '""').apply(lambda x: f'"{x}"')

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
    

def main(ppt_app):
    """Main menu for the workflow."""
    test_index = input("Enter test index: ")
    test_index = int(test_index)
    print("PowerPoint Shape Analysis Tool")
    
    x = create_shape_boxes.create_shape_boxes(ppt_app)
    shape_data = x['shape_data']
    slide_dimensions = x['slide_dimensions']
    print("Extracted shape data and created red boxes")
    
    group_list = []

    while True:
        group_name = input("Select shapes to group and enter a group name. Press X for exit: ")

        if group_name.lower().strip() == 'x':
            #save_groups_to_file.save_groups_to_file(test_index, output, slide_dimensions)  # Saves as JSON
            group_dl = process_groups(shape_data, group_list, test_index, slide_dimensions)
            group_dl = structure_shapes.generate_structure_main(group_dl)
            group_df = save_to_csv(group_dl, test_index)  # Saves as CSV
            print("Done")   
            return group_df
        else:
            selection = ppt_app.ActiveWindow.Selection
            #level = input("Enter level: ")
            #level = int(level)
            x = group_shapes.group_selected_shapes(group_name, selection)
            group_list.append(x)

    

if __name__ == "__main__":
    ppt_app = resources.get_powerpoint_app()
    main(ppt_app)
