import os
import json
from PIL import Image
import google.generativeai as genai
from dotenv import load_dotenv
import sys
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import base

# Paths



def build_example_list(training_dir):
    example_prompts = []
    for filename in os.listdir(training_dir):
        base_name = filename[:-5]  # Remove '.png'
        json_path = os.path.join(training_dir, f'{base_name}.json')
        image = base.get_picture(base_name)

        with open(json_path, 'r') as f:
            tree = json.load(f)

        # Append to prompt
        example_prompts.extend([
            f"Here is an example slide with annotated sections ({base_name}):",
            image,
            "The sections for this slide are:",
            json.dumps(tree, indent=2)
        ])
    return example_prompts


def build_prompt(example_prompts, target_image):
    prompt = [
        "I want to detect visual sections in a slide. A 'section' is a rectangular area that groups related content. Most sections can be broken down into smaller subsections.",
        "Please return a list of sections in JSON format, where each section has the keys:",
        "- 'label': a short description of what's in this section",
        "- 'sections': list of the subsections. Please note that sections can be nested indefinitely. Also, please note that some sections may overlap but not necessarily be contained within each other.",
        *example_prompts,
        "Now analyze this new slide and return the sections in JSON format:",
        target_image
    ]
    return prompt


def write_sections(elements_json, folder, base_name):
    if not os.path.exists(folder):
        os.makedirs(folder)
    with open(f'{folder}/{base_name}.json', 'w') as f:
        json.dump(elements_json, f, indent=2)




if __name__ == "__main__":
    rump = input('Enter slide index: ')


    load_dotenv('env.env')

    # Initialize Gemini
    genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
    model = genai.GenerativeModel("gemini-1.5-pro")


    # Step 1: Load training examples
    example_prompts = build_example_list(base.training_data_folder)
    
    # Step 2: Target slide to analyze
    target_image = base.get_picture(rump)

    # Step 3: Final prompt
    prompt = build_prompt(example_prompts, target_image)

    # Step 4: Send to Gemini
    response = model.generate_content(prompt)
    text = response.text
    elements_json = base.get_json_from_llm_response(text)

    write_sections(elements_json, base.get_project_path('resources', 'llm_pic_to_tree_results'), rump)
    # Step 5: Output the result
    print(response.text)
