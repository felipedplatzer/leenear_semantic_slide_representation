import json
from PIL import Image, ImageDraw, ImageFont
import base

def draw_section(draw, section, slide_width, slide_height, relative_coordinates):
    try:
        if relative_coordinates:
            x1 = int(section['left'] * slide_width)
            y1 = int(section['top'] * slide_height)
            x2 = int(section['right'] * slide_width)
            y2 = int(section['bottom'] * slide_height)
        else:
            x1 = int(section['left'])
            y1 = int(section['top'])
            x2 = int(section['right'])
            y2 = int(section['bottom'])
    except:
        return
    # Draw rectangle border
    draw.rectangle([x1, y1, x2, y2], outline='red', width=3)
    # Draw label
    label = section['label']
    draw.text((x1, y1-20), label, fill='red')


def annotate_image(image, data, one_image_bool, relative_coordinates=True, rump=None):
    # Create a copy of the image to draw on
    if one_image_bool:
        annotated_image = image.copy()
        draw = ImageDraw.Draw(annotated_image)
        # Get image dimensions
        width, height = image.size
        # Draw each section
        for section in data:
            draw_section(draw, section, width, height, relative_coordinates)
        # Save the annotated image
        output_path = base.get_project_path('resources', 'output_annotated_slides_temp', f'{rump}.png')
        annotated_image.save(output_path)
        print(f"Annotated image saved as '{output_path}'")
    else:
        for i,section in enumerate(data):
            annotated_image = image.copy()
            draw = ImageDraw.Draw(annotated_image)
            # Get image dimensions
            width, height = image.size
            # Draw each section
            draw_section(draw, section, width, height, relative_coordinates)
            # Save the annotated image
            output_path = base.get_project_path('resources', 'output_annotated_slides_temp', f'{rump}_{i}.png')
            annotated_image.save(output_path)
            print(f"Annotated image saved as '{output_path}'")