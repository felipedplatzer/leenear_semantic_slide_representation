import google.generativeai as genai
import os
from dotenv import load_dotenv
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import base
import openai
import json
import annotate_image
from PIL import Image

load_dotenv('env.env')


def get_prompt(section_tree, elements_data, relative_coordinates_bool):
    if relative_coordinates_bool:
        coordinates_string = 'relative coordinates between 0 and 1'
    else:
        coordinates_string = 'bounding box coordinates in pixels'
    
    # Format elements data for the prompt
    elements_text = "Elements on the page:\n"
    for element in elements_data:
        elements_text += f"- ID: {element['shape_id']}, Description: {element['description']}, "
        elements_text += f"Coordinates: left={element['left']}, top={element['top']}, "
        elements_text += f"right={element['right']}, bottom={element['bottom']}\n"
    
    prompt = f"""
    You are given a list of elements from a page with their descriptions, IDs, and {coordinates_string}.
    You are also given a hierarchical list of sections that should be found on this page.

    Your task is to:
    1. Analyze the element descriptions and coordinates to determine which elements belong to each section.
    2. For each section, return the list of element IDs that belong to that section.
    3. Elements can belong to multiple sections.
    4. Consider the hierarchical structure - subsections should contain elements that are more specific to that subsection.

    {elements_text}

    The list of sections is a tree structure. Some sections can contain subsections. These subsections can contain subsections themselves.
    This is represented as a JSON array, where each section has a label and a list of subsections.
    Use this tree structure to guide your classification. For example, elements that belong to a subsection should also be considered as belonging to its parent section.

    Here is the list of sections (label and subsections) to classify elements into:
    {json.dumps(section_tree, indent=2)}

    Some sections can overlap or intersect, and elements can belong to multiple sections.

    Return a JSON object with a label and the list of element IDs that belong to that section:
    Please enclose each element ID in quotes - i.e. treat it as a string
     [
         {{"label": "..." , "elements": ["...", "...", "..."]}},
         {{"label": "..." , "elements": ["...", "...", "..."]}},
         ...
     ]

    Make sure to include all sections from the tree structure, even if they have no elements assigned to them (use an empty array in that case).
    """
    return prompt


def call_openai_model(prompt):
    response = openai.chat.completions.create(
        model="gpt-4o",
        messages=[
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt}
                ]
            }
        ],
        max_tokens=10000,
        temperature=0
    )
    return response.choices[0].message.content


def call_google_model(prompt):
    genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
    model = genai.GenerativeModel("gemini-1.5-pro")
    response = model.generate_content(prompt, stream=False)
    return response.text


def get_elements_by_section(section_tree, elements_data, relative_coordinates_bool=True, model='google'):
    """
    Takes element data and section tree, returns mapping of sections to element IDs.
    
    Args:
        section_tree: JSON structure with hierarchical sections
        elements_data: List of dicts with 'id', 'description', 'left', 'top', 'right', 'bottom'
        relative_coordinates_bool: Whether coordinates are relative (0-1) or absolute pixels
        model: 'google' or 'openai'
    
    Returns:
        Dict mapping section names to lists of element IDs
    """
    prompt = get_prompt(section_tree, elements_data, relative_coordinates_bool)
    
    if model == 'google':
        text = call_google_model(prompt)
    elif model == 'openai':
        openai.api_key = os.getenv("OPENAI_API_KEY")
        text = call_openai_model(prompt)
    else:
        raise ValueError("Model must be 'google' or 'openai'")
    
    elements_json = base.get_json_from_llm_response(text)
    return elements_json



def main(descs, tree, slide_image_path=None):
    """Main function that takes descs and tree dictionaries directly"""
    RELATIVE_COORDINATES_BOOL = False
    MODEL = 'openai'  # toggle between google and openai
    ONE_IMAGE_BOOL = False
    
    # Use the provided descs and tree data
    elements_data = descs
    section_tree = tree

    # Get element assignments
    result = get_elements_by_section(section_tree, elements_data, RELATIVE_COORDINATES_BOOL, MODEL)
    
    # Rename "elements" to "shape_id" - "elements" makes sense for LLM but shape_id is used elasewhere
    for x in result:
        x['shape_id'] = x['elements']
        del x['elements']
    
    print("Element assignments by section:")
    print(json.dumps(result, indent=2))
    
    # Annotate the image if slide_image_path is provided
    if slide_image_path:
        result_with_pos = base.get_bounds_from_shape_ids(result, elements_data)
        slide_image = Image.open(slide_image_path)
        annotate_image.annotate_image(slide_image, result_with_pos, ONE_IMAGE_BOOL, RELATIVE_COORDINATES_BOOL)
        return result_with_pos
    
    return result


if __name__ == "__main__":
    rump = input('Enter slide index: ')
    
    # Read the descs and tree files
    with open(base.get_project_path('resources', 'llm_pic_to_descs_results', f'{rump}.json'), 'r') as f:
        descs = json.load(f)
    
    with open(base.get_project_path('resources', 'llm_pic_to_tree_results', f'{rump}.json'), 'r') as f:
        tree = json.load(f)
    
    # Get slide image path for annotation
    slide_image_path = base.get_project_path('resources', 'input_slide_pictures', f'{rump}.png')
    
    # Call main with the loaded data
    result = main(descs, tree, slide_image_path)
    with open(base.get_project_path('resources', 'llm_tree_and_descs_to_final_results', f'{rump}.json'), 'w') as f:
        json.dump(result, f, indent=2)
    