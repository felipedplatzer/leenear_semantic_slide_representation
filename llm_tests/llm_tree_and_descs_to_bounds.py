import google.generativeai as genai
import os
from PIL import Image
from dotenv import load_dotenv
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import base
import annotate_image
import openai
import base64
import json
load_dotenv('env.env')



def encode_image(image_path):
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def get_prompt(section_tree, relative_coordinates_bool):
    if relative_coordinates_bool:
        coordinates_string = 'relative coordinates between 0 and 1'
        x_min = 0.12
        y_min = 0.1
        x_max = 0.2
        y_max = 0.2
    else:
        coordinates_string = 'bounding box coordinates in pixels'
        x_min = 120
        y_min = 80
        x_max = 560
        y_max = 140
    
    prompt = f"""

    You are given an image of a slide. The slide contains several sections

    Your task is to:
    1. Identify the location of each section from the list provided.
    2. For each matched section, return:
    - The matched section name
    - 'x_min', 'y_min', 'x_max', 'y_max': {coordinates_string}
    - 'confidence': a confidence score from 0 to 1"

    The list of sections is a tree structure. Some sections can contain subsections. These subsections can contain subsections themselves.
    This is represented as a JSON array, where each section has a label and a list of subsections.
    Use this tree structure to guide your detection. For example, a subsection should be enclosed by its parent section (larger or equal x_min and y_min and smaller or equal x_max and y_max).

    Here is the list of sections (label and subsections) to detect:
    {section_tree}

    Some sections can intersect.

    Return a JSON array like this:

    [
    {{
        "label": "section_name",
        "x_min": {str(x_min)},
        "y_min": {str(y_min)},
        "x_max": {str(x_max)},
        "y_max": {str(y_max)},
        "confidence": 0.97
    }},
    ...
    ]
    """
    return prompt


def call_openai_model(prompt, base64_image):
    response = openai.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{base64_image}",
                            "detail": "high"
                        }
                    }
                ]
            }
        ],
        max_tokens=1000,
        temperature=0
    )
    return response.choices[0].message.content


if __name__ == "__main__":
    ONE_IMAGE_BOOL = False
    RELATIVE_COORDINATES_BOOL = True
    MODEL = 'google' #toggle between google and openai


    target_image_path = base.get_project_path('resources', 'input_slide_pictures', '12.png')
    image = Image.open(target_image_path)

    # Section headers you want to find
    load_dotenv('env.env')

    with open(base.get_project_path('resources', 'manual_sections_json_trees', '12.json'), 'r') as f:
        section_tree = json.load(f)


    prompt = get_prompt(section_tree, RELATIVE_COORDINATES_BOOL)


    # MAIN MODEL
    if MODEL == 'google':
        genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
        model = genai.GenerativeModel("gemini-1.5-pro")
        response = model.generate_content([prompt, image], stream=False)
        text = response.text
        elements_json = base.get_json_from_llm_response(text)
    elif MODEL == 'openai':
        openai.api_key = os.getenv("OPENAI_API_KEY")
        base64_image = encode_image(target_image_path)
        text = call_openai_model(prompt, base64_image)
        elements_json = base.get_json_from_llm_response(text)
    if ONE_IMAGE_BOOL:  
        annotate_image.annotate_in_one_image(image, elements_json, RELATIVE_COORDINATES_BOOL)
    else:
        annotate_image.annotate_in_separate_images(image, elements_json, RELATIVE_COORDINATES_BOOL)