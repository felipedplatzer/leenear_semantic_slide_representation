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

import basic

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
            group_dl = basic.process_groups(shape_data, group_list, test_index, slide_dimensions)
            group_dl = structure_shapes.generate_structure_main(group_dl)
            group_df = basic.save_to_csv(group_dl, test_index)  # Saves as CSV
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
