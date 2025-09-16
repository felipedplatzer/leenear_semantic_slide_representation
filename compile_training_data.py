import os
import pandas as pd
import win32com.client
import json
import numpy as np

ppt_app = win32com.client.Dispatch("PowerPoint.Application")
ppt_app.DisplayAlerts = False


def save_to_png(slide_file, png_file):
    presentation = ppt_app.Presentations.Open(slide_file, WithWindow=False)
    slide = presentation.Slides(1)
    slide.Export(png_file, "PNG")
    return png_file

def normalize_positions(csv_data):
    csv_data['top'] = np.round(csv_data['top'] / csv_data['slide_height'], 3)
    csv_data['left'] = np.round(csv_data['left'] / csv_data['slide_width'], 3)
    csv_data['right'] = np.round(csv_data['right'] / csv_data['slide_width'], 3)
    csv_data['bottom'] = np.round(csv_data['bottom'] / csv_data['slide_height'], 3)
    return csv_data

def convert_to_json(csv_data):
    csv_data = csv_data[['index','label','top','left','right','bottom']]
    csv_data = csv_data.to_dict(orient='records')
    return csv_data

def save_to_json(csv_data,  json_file):
    with open(json_file, 'w') as f:
        json.dump(csv_data, f, indent=4)

if __name__ == "__main__":
    # Get current working directory
    current_dir = os.getcwd()
    
    # Use absolute paths
    slides_dir = os.path.join(current_dir, 'slides_for_tests')
    csv_dir = os.path.join(current_dir, 'csv_files', 'after_processing')
    training_data_dir = os.path.join(current_dir, 'training_data')

    rump_slides = [x.replace('.pptx', '') for x in os.listdir(slides_dir) if x.endswith('.pptx')]
    rump_csv_files = [x.replace('.csv', '') for x in os.listdir(csv_dir) if x.endswith('.csv')]

    rump_both = [x for x in rump_slides if x in rump_csv_files]
    
    for x in rump_both:
        slide_file = os.path.join(slides_dir, x + '.pptx')
        
        csv_file = os.path.join(csv_dir, x + '.csv')
        png_file = os.path.join(training_data_dir, x + '.png')
        json_file = os.path.join(training_data_dir, x + '.json')
        img = save_to_png(slide_file, png_file)
        csv_data = pd.read_csv(csv_file)
        # normalize positoins (index x and y positoins to 0-1)
        csv_data = normalize_positions(csv_data)
        # convert to json
        csv_data = convert_to_json(csv_data)
        save_to_json(csv_data, json_file)
    ppt_app.Quit()

