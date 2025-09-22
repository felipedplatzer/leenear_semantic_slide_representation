import google.generativeai as genai
import os
from PIL import Image
from dotenv import load_dotenv
import base
import annotate_image
import openai
import base64
import json
load_dotenv('env.env')




def get_google_prompt(section_titles):      
    prompt = f"""

    You are given an image of a slide. The slide contains several sections. Some sections can contain subsections. Some sections can intersect

    Your task is to:
    1. Identify the location of each section from the list provided.
    2. For each matched section, return:
    - The matched section name
    - 'x_min', 'y_min', 'x_max', 'y_max': bounding box coordinates in pixels
    - 'confidence': a confidence score from 0 to 1"

    Here is the list of sections to detect:
    {section_titles}

    Return a JSON array like this:

    [
    {{
        "label": "section_name",
        "x_min": 0.12,
        "y_min": 0.1,
        "x_max": 0.2,
        "y_max": 0.2,
        "confidence": 0.97
    }},
    ...
    ]
    """
    return prompt


def get_openai_prompt(section_titles):
    prompt = f"""
    Here is an image of a slide.

    Your task is to find the bounding boxes for specific section headers. The section headers to detect are:
    {section_titles}

    Please do the following:
    1. Search the image for any text that matches or closely resembles one of the section titles.
    2. For each match, return:
    - The section title it matches
    - The exact text as it appears
    - x_min, y_min, x_max, y_max: bounding box coordinates in pixels
    - A confidence score between 0 and 1

    Return the result as a JSON array:

    [
    {{
        "label": "Introduction",
        "x_min": 120,
        "y_min": 80,
        "x_max": 560,
        "y_max": 140,
        "confidence": 0.95
    }},
    ...
    ]"""
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
    RELATIVE_COORDINATES_BOOL = False
    MODEL = 'openai' #toggle between google and openai


    target_image_path = base.get_project_path('resources', 'input_slide_pictures', '12.png')
    image = Image.open(target_image_path)


    # Section headers you want to find
    section_titles = [
        "Title", 
        "Footnote", 
        "Legend", 
        "Page number", 
        "Manifold logo", 
        "Impact of semiconductor trends on industry and KPMG's business lines",
        "List of trends",
        "Impact on semiconductor industry",
        "Impact on KPMG's business lines",
        "Callout",
        "Table rows",
        "Table columns"
    ]

    load_dotenv('env.env')

    # MAIN MODEL
    if MODEL == 'google':
        genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
        model = genai.GenerativeModel("gemini-1.5-pro")
        prompt = get_google_prompt(section_titles)
        response = model.generate_content([prompt, image], stream=False)
        text = response.text
        elements_json = base.get_json_from_llm_response(text)
    elif MODEL == 'openai':
        openai.api_key = os.getenv("OPENAI_API_KEY")
        base64_image = base.encode_image(target_image_path)
        prompt = get_openai_prompt(section_titles)
        text = call_openai_model(prompt, base64_image)
        elements_json = base.get_json_from_llm_response(text)
    if ONE_IMAGE_BOOL:  
        annotate_image.annotate_in_one_image(image, elements_json, RELATIVE_COORDINATES_BOOL)
    else:
        annotate_image.annotate_in_separate_images(image, elements_json, RELATIVE_COORDINATES_BOOL)