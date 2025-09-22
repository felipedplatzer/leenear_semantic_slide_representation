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


def describe_image(image_path, client):
    """Use OpenAI to describe an image"""
    try:
        # Open and convert image to ensure proper format
        with Image.open(image_path) as img:
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


def get_description_textbox(shape):
    try:
        text_content = shape.TextFrame.TextRange.Text.strip()
        if text_content:
            return f"Text content: {text_content}"
    except:
        return "Textbox (unable to read content)"


def get_description_picture(shape, client):
    temp_path = None
    try:
        # Create temp file with explicit path
        temp_fd, temp_path = tempfile.mkstemp(suffix='.jpg')
        os.close(temp_fd)  # Close the file descriptor
        # Export shape as image
        shape.Export(temp_path, Filter=3)  # 2 = png, 3 = jpg, 4 = bmp
        # Get description
        description = describe_image(temp_path, client)
        
    except Exception as e:
        print(f"Error processing image in shape {shape.Id}: {e}")
        description = "Error processing image"
    finally:
        # Clean up temp file
        if temp_path and os.path.exists(temp_path):
            try:
                os.unlink(temp_path)
            except:
                pass
    return description


def get_description_table(shape, client):
    try:
        table = shape.Table
        table_desc = f"Table with {table.Rows.Count} rows and {table.Columns.Count} columns"
        table_desc += '. ' + get_description_picture(shape, client)
        return table_desc
    except:
        return "Table"

def get_description_chart(shape):
    try:
        chart_type = shape.Chart.ChartType
        return f"Chart: {chart_type}"
    except:
        return "Chart"  


def get_description_cell(cell, table_id, i_row, i_cell):
    cell_desc = f"Cell of table {table_id} at row {i_row} and column {i_cell}"
    try:
        text = cell.Shape.TextFrame.TextRange.Text.strip()
        if text.strip() != "":
            cell_desc += f" with text = {text}"
        return cell_desc
    except:
        return cell_desc

def extract_shape_descriptions(slide, client):
    """Extract descriptions of all shapes in a PowerPoint presentation using win32com"""
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
        d['description'] = get_description_textbox(x)
        all_elements_dl.append(d)

    for x in tables:
        d = base.get_shape_info(x)
        d['description'] = get_description_table(x, client)
        all_elements_dl.append(d)

    for x in charts:
        d = base.get_shape_info(x)
        d['description'] = get_description_chart(x)
        all_elements_dl.append(d)

    for x in pictures:
        d = base.get_shape_info(x)
        d['description'] = get_description_picture(x, client)
        all_elements_dl.append(d)

    # Postprocess tables - break into cells
    
    for table in tables:
        for i_row, row in enumerate(table.Table.Rows):
            for i_cell, cell in enumerate(row.Cells):
                d = base.get_cell_info(table, i_row, i_cell)
                d['description'] = get_description_cell(cell, table.Id,i_row, i_cell)
                all_elements_dl.append(d)
    
    # Postprocessing - convert to relative coordinates and round

    for x in all_elements_dl:
        x['top'] = round(x['top'] / slide_height, 3)  
        x['left'] = round(x['left'] / slide_width, 3)
        x['width'] = round(x['width'] / slide_width, 3)
        x['height'] = round(x['height'] / slide_height, 3)
        x['right'] = round(x['right'] / slide_width, 3)
        x['bottom'] = round(x['bottom'] / slide_height, 3)
 
    # Return
    return all_elements_dl


def main(rump, active_bool, client):
    if active_bool == False:
        presentation, slide = base.open_slide(rump)
    else:
        presentation, slide = base.get_active_slide()
    output_file = base.get_project_path('resources', 'llm_pic_to_descs_results', f'{rump}.json')    
    shape_descriptions = extract_shape_descriptions(slide, client)
    with open(output_file, 'w') as f:
        json.dump(shape_descriptions, f, indent=2)
    print(f"Saved {len(shape_descriptions)} shape descriptions to {output_file}")
    if active_bool == False:
        presentation.Close()
    return shape_descriptions


# Main execution
if __name__ == "__main__":
    rump = input('Enter slide index: ')
    # Load environment variables
    load_dotenv('env.env')

    # Initialize OpenAI
    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

    active_bool = False # get slide from resources folder, not from active slide
    main(rump, active_bool, client)
