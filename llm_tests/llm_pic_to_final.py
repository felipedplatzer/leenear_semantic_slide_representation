import os
import json
from PIL import Image
import google.generativeai as genai
from dotenv import load_dotenv
import annotate_image
import base

# Get the directory where this script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Paths - using base module for consistent path handling
training_dir = base.training_data_folder
target_image_path = base.get_project_path('resources', 'input_slide_pictures', '12.png')  # Replace with your test slide

# One image or separate images - toggle this to True or False
ONE_IMAGE_BOOL = False


def build_example_list(training_dir):
    example_prompts = []
    for filename in os.listdir(training_dir):
        if filename.endswith('.png'):
            base_name = filename[:-4]  # Remove '.png'
            json_path = os.path.join(training_dir, f'{base_name}.json')
            img_path = os.path.join(training_dir, filename)

            if not os.path.exists(json_path):
                continue  # Skip if no matching JSON

            # Load image and annotation
            image = Image.open(img_path)
            with open(json_path, 'r') as f:
                annotations = json.load(f)

            # Append to prompt
            example_prompts.extend([
                f"Here is an example slide with annotated sections ({base_name}):",
                image,
                "The sections for this slide are:",
                json.dumps(annotations, indent=2)
            ])
    return example_prompts


def build_prompt(example_prompts, target_image):
    prompt = [
        "I want to detect visual sections in a slide. A 'section' is a rectangular area that groups related content. Most sections can be broken down into smaller subsections.",
        "Please return a list of sections in JSON format, where each section has the keys:",
        "- 'label': a short description of what's in this section",
        "- 'x_min', 'y_min', 'x_max', 'y_max': relative coordinates between 0 and 1",
        *example_prompts,
        "Now analyze this new slide and return the section coordinates in JSON format:",
        target_image
    ]
    return prompt







if __name__ == "__main__":
    load_dotenv('env.env')

    # Initialize Gemini
    genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
    model = genai.GenerativeModel("gemini-1.5-pro")


    # Step 1: Load training examples
    example_prompts = build_example_list(training_dir)
    
    # Step 2: Target slide to analyze
    target_image = Image.open(target_image_path)

    # Step 3: Final prompt
    prompt = build_prompt(example_prompts, target_image)

    # Step 4: Send to Gemini
    response = model.generate_content(prompt)
    elements_json = base.get_json_from_llm_response(response.text)

    if ONE_IMAGE_BOOL:  
        annotate_image.annotate_in_one_image(target_image, elements_json)
    else:
        annotate_image.annotate_in_separate_images(target_image, elements_json)
    
    # Step 5: Output the result
    print(response.text)
