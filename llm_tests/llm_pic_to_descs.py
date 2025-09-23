import os
import json
import tempfile
import win32com.client
from openai import OpenAI
from dotenv import load_dotenv
from PIL import Image
import base64
from io import BytesIO
import math
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import base


def describe_image(client, shape_img_filepath):
    """Use OpenAI to describe an image"""
    try:
        # Open and convert image to ensure proper format
        with Image.open(shape_img_filepath) as img:
            # Convert to RGB if necessary (for JPEG compatibility)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Save to BytesIO buffer as JPEG
            buffer = BytesIO()
            img.save(buffer, format='JPEG')
            buffer.seek(0)
            
            # Encode to base64
            base64_image = base64.b64encode(buffer.getvalue()).decode('utf-8')
        
        # Call OpenAI Vision API
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {   
                            "type": "text",
                            "text": "Describe this image in a few words."
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=10,
            temperature=0.1
        )
        
        return response.choices[0].message.content
    except Exception as e:
        print(f"Error describing image: {e}")
        return "Unable to describe image"

def describe_image_from_bytes(client, shape_img_bytes):
    """Use OpenAI to describe an image from byte data"""
    try:
        # Open and convert image to ensure proper format
        with Image.open(BytesIO(shape_img_bytes)) as img:
            # Convert to RGB if necessary (for JPEG compatibility)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Save to BytesIO buffer as JPEG
            buffer = BytesIO()
            img.save(buffer, format='JPEG')
            buffer.seek(0)
            
            # Encode to base64
            base64_image = base64.b64encode(buffer.getvalue()).decode('utf-8')
        
        # Call OpenAI Vision API
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {   
                            "type": "text",
                            "text": "Describe this image in a few words."
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=10,
            temperature=0.1
        )
        
        return response.choices[0].message.content
    except Exception as e:
        print(f"Error describing image: {e}")
        return "Unable to describe image"

def describe_image_from_bytes(client, shape_img_bytes):
    """Use OpenAI to describe an image from byte data"""
    try:
        # Open and convert image to ensure proper format
        with Image.open(BytesIO(shape_img_bytes)) as img:
            # Convert to RGB if necessary (for JPEG compatibility)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Save to BytesIO buffer as JPEG
            buffer = BytesIO()
            img.save(buffer, format='JPEG')
            buffer.seek(0)
            
            # Encode to base64
            base64_image = base64.b64encode(buffer.getvalue()).decode('utf-8')
        
        # Call OpenAI Vision API
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {
                    "role": "user",
                    "content": [
                        {   
                            "type": "text",
                            "text": "Describe this image in a few words."
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/jpeg;base64,{base64_image}"
                            }
                        }
                    ]
                }
            ],
            max_tokens=10,
            temperature=0.1
        )
        
        return response.choices[0].message.content
    except Exception as e:
        print(f"Error describing image: {e}")
        return "Unable to describe image"





def get_description_table(client, shape_img_filepath):
    try:
        # table = shape.Table
        table_desc = f"Table " # with {table.Rows.Count} rows and {table.Columns.Count} columns"
        table_desc += '. ' + describe_image(client, shape_img_filepath)
        return table_desc
    except:
        return "Table"

def get_description_chart(client, shape_img_filepath):
    try:
        # chart_type = shape.Chart.ChartType
        chart_desc = f"Chart "
        chart_desc += '. ' + describe_image(client, shape_img_filepath)
        return chart_desc
    except:
        return "Chart"  


def get_description_cell(text, table_id, i_row, i_cell):
    cell_desc = f"Cell of table {table_id} at row {i_row} and column {i_cell}"
    try:
        if text.strip() != "":
            cell_desc += f" with text = {text}"
        return cell_desc
    except:
        return cell_desc

# Byte-based versions of the description functions
def get_description_table_from_bytes(client, shape_img_bytes):
    try:
        table_desc = f"Table "
        table_desc += '. ' + describe_image_from_bytes(client, shape_img_bytes)
        return table_desc
    except:
        return "Table"

def get_description_chart_from_bytes(client, shape_img_bytes):
    try:
        chart_desc = f"Chart "
        chart_desc += '. ' + describe_image_from_bytes(client, shape_img_bytes)
        return chart_desc
    except:
        return "Chart"

def process_slide_shapes(slide):
    """Process slide shapes and extract basic info without descriptions"""
    # Process each shape in the slide
    slide_height = slide.Parent.PageSetup.SlideHeight
    slide_width = slide.Parent.PageSetup.SlideWidth

    # Classify shapes into textboxes, tables, charts, and pictures
    textboxes, tables, charts, pictures = base.classify_shapes(slide)
    
    all_elements_dl = []
    
    for x in textboxes:
        d = base.get_shape_info(x)
        pos_dict = base.get_text_pos(x)
        d['top'] = pos_dict['top']
        d['left'] = pos_dict['left']
        d['width'] = pos_dict['width']
        d['height'] = pos_dict['height']
        d['right'] = pos_dict['right']
        d['bottom'] = pos_dict['bottom']
        d['shape_type_category'] = 'textbox'
        all_elements_dl.append(d)

    for x in tables:
        d = base.get_shape_info(x)
        d['shape_type_category'] = 'table'
        all_elements_dl.append(d)

    for x in charts:
        d = base.get_shape_info(x)
        d['shape_type_category'] = 'chart'
        all_elements_dl.append(d)

    for x in pictures:
        d = base.get_shape_info(x)
        d['shape_type_category'] = 'picture'
        all_elements_dl.append(d)

    # Postprocess tables - break into cells
    for table in tables:
        for i_row, row in enumerate(table.Table.Rows):
            for i_cell, cell in enumerate(row.Cells):
                d = base.get_cell_info(table, i_row, i_cell)
                d['shape_type_category'] = 'cell'
                d['table_id'] = table.Id
                d['i_row'] = i_row
                d['i_cell'] = i_cell
                all_elements_dl.append(d)
    
    # Postprocessing - convert to relative coordinates and round
    for x in all_elements_dl:
        x['top'] = round(x['top'] / slide_height, 3)  
        x['left'] = round(x['left'] / slide_width, 3)
        x['width'] = round(x['width'] / slide_width, 3)
        x['height'] = round(x['height'] / slide_height, 3)
        x['right'] = round(x['right'] / slide_width, 3)
        x['bottom'] = round(x['bottom'] / slide_height, 3)
 
    return all_elements_dl

def extract_shape_descriptions(shape_dict_list, client, shape_images):
    """Add descriptions to pre-processed shape dictionaries"""
    for shape_dict in shape_dict_list:
        if shape_dict['shape_type'] == 'textbox':
            shape_dict['description'] = "Textbox with text: " + shape_dict['text']
        elif shape_dict['shape_type'] == 'cell':
            shape_dict['description'] = get_description_cell(
                shape_dict['text'], 
                shape_dict['table_id'], 
                shape_dict['i_row'], 
                shape_dict['i_cell']
        )
        else:
            if shape_dict['shape_id'] in shape_images:
                if isinstance(shape_images[shape_dict['shape_id']], bytes):
                    # Byte data - use byte-based functions
                    shape_bytes = shape_images[shape_dict['shape_id']]
                else:
                    shape_dict['description'] = ''
                if shape_dict['shape_type'] == 'table':
                    shape_dict['description'] = get_description_table_from_bytes(client, shape_bytes)
                elif shape_dict['shape_type'] == 'chart':
                    shape_dict['description'] = get_description_chart_from_bytes(client, shape_bytes)
                else:
                    shape_dict['description'] = describe_image_from_bytes(client, shape_bytes)
            else:
                shape_dict['description'] = ''
    return shape_dict_list


def main(rump, shape_images, client, shape_dict_list):
    output_file = base.get_project_path('resources', 'llm_pic_to_descs_results', f'{rump}.json')    
    shape_descriptions = extract_shape_descriptions(shape_dict_list, client, shape_images)
    with open(output_file, 'w') as f:
        json.dump(shape_descriptions, f, indent=2)
    print(f"Saved {len(shape_descriptions)} shape descriptions to {output_file}")
    return shape_descriptions



# Main execution
if __name__ == "__main__":
    rump = input('Enter slide index: ')
    # Load environment variables
    load_dotenv('env.env')

    # Initialize OpenAI
    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

    active_bool = False # get slide from resources folder, not from active slide
    target_image = base.get_picture(rump, active_bool)
    presentation, slide = base.open_slide(rump)
    shape_images_bytes = base.get_shape_images(slide)
    
    # Process slide shapes before main
    shape_dict_list = process_slide_shapes(slide)
    
    # Call main with bytes

    main(rump, shape_images_bytes, client, shape_dict_list)
    presentation.Close()