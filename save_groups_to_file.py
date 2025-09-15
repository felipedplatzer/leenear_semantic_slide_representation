import json
from datetime import datetime
import os
import csv
import pandas as pd

def save_groups_to_file(test_index, data_list, slide_dimensions):
    """
    Save a list of dictionaries to a JSON file.
    
    Args:
        data_list: List of dictionaries to save
        filename: Optional custom filename. If None, uses timestamp.
    
    Returns:
        str: Path to the saved file
    """
    try:
        test_index_str = str(test_index).zfill(3)
        filename = f"test_{str(test_index_str)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"

        # Get the directory of the current script
        filepath = f'./json_files/{filename}'
        
        # Save to JSON file
        with open(filepath, 'w', encoding='utf-8') as f:
            data_dict = {"test_index": test_index, "slide_dimensions": slide_dimensions, "data": data_list}
            json.dump(data_dict, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Data saved to: {filepath}")
        return filepath
        
    except Exception as e:
        print(f"❌ Error saving to JSON: {e}")
        return None


def save_groups_to_csv(test_index, data_list, slide_dimensions):
    for x in data_list:
        x['slide_height'] = slide_dimensions['height']
        x['slide_width'] = slide_dimensions['width']
        x['test_index'] = test_index
        
    test_index_str = str(test_index).zfill(3)
    filename = f"test_{str(test_index_str)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"

    # Get the directory of the current script
    filepath = f'./csv_files/{filename}'

    df = pd.DataFrame(data_list)

    df["text"] = df["text"].str.replace('"', '""').apply(lambda x: f'"{x}"')
    df["shape_id"] = df["shape_id"].apply(json.dumps)
    df["shape_id"] = df["shape_id"].str.replace('"', '""').apply(lambda x: f'"{x}"')

# round to 2 decimal places
    for x in ["top", "left", "right", "bottom", "width", "height"]:
        df[x] = df[x].round(2)
    #reorder cols
    df = df[["test_index", "label", "level", "shape_id", "text", "top", "left", "right", "bottom", "width", "height", "slide_height", "slide_width"]]
    df.to_csv(filepath, index=False, encoding="utf-8", quoting=csv.QUOTE_NONE, escapechar='\\')
    
    """

    with open(filepath, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        # write header
        writer.writerow(["id", "category", "text"])

        for row in data_list:
            # manually quote only the text field
            text_value = row["text"].replace('"', '""')   # escape inner quotes
            text_value = f'"{text_value}"'                # wrap in quotes
            writer.writerow([row["id"], row["category"], text_value])
    """