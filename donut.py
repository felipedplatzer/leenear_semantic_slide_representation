from transformers import DonutProcessor, VisionEncoderDecoderModel
from PIL import Image
import torch
import json
import base
from get_list_from_tree import flatten_tree
# 1. Load pretrained Donut model
processor = DonutProcessor.from_pretrained("naver-clova-ix/donut-base")
model = VisionEncoderDecoderModel.from_pretrained("naver-clova-ix/donut-base")

# 2. Load your slide image
image_path = base.get_project_path('resources', 'input_slide_pictures', '12.png')
image = Image.open(image_path).convert("RGB")

# 3. Define the prompt (task instruction)
try:
    with open(base.get_project_path('resources', 'manual_sections_json_trees', '12.json'), 'r') as f:
        section_tree = json.load(f)
    section_titles = flatten_tree(section_tree)
    # Extract just the labels for the prompt
    section_labels = [item['label'] for item in section_titles]
except FileNotFoundError:
    print("Warning: sections/12.json not found, using default sections")
    section_labels = ["Introduction", "Methodology", "Results", "Conclusion"]

# Try a simpler approach - Donut works better with document understanding tasks
prompt = f"<s_docvqa><s_question>What are the main section headers in this document? Look for: {', '.join(section_labels)}</s_question><s_answer>"

# 4. Preprocess input
pixel_values = processor(image, return_tensors="pt").pixel_values

# 5. Generate output
task_prompt = processor.tokenizer(prompt, add_special_tokens=False, return_tensors="pt").input_ids

# Use the tokenizer's start token instead of the model's config
decoder_start_token_id = processor.tokenizer.bos_token_id if processor.tokenizer.bos_token_id is not None else processor.tokenizer.pad_token_id
if decoder_start_token_id is None:
    decoder_start_token_id = 0  # fallback to token 0
    print("Warning: Using token 0 as decoder start token")

decoder_input_ids = torch.ones((1, 1), dtype=torch.long) * decoder_start_token_id

print(f"Decoder start token ID: {decoder_start_token_id}")
print(f"EOS token ID: {processor.tokenizer.eos_token_id}")
print(f"UNK token ID: {processor.tokenizer.unk_token_id}")

outputs = model.generate(
    pixel_values,
    decoder_input_ids=decoder_input_ids,
    max_length=512,
    early_stopping=True,
    num_beams=1,
    do_sample=False,
    eos_token_id=processor.tokenizer.eos_token_id,
    pad_token_id=processor.tokenizer.pad_token_id,
    return_dict_in_generate=True,
)

# 6. Decode
sequence = processor.batch_decode(outputs.sequences)[0]
print(f"Raw generated sequence: {sequence}")

try:
    parsed_output = processor.token2json(sequence)
    print(f"Parsed JSON output: {parsed_output}")
except Exception as e:
    print(f"Error parsing JSON: {e}")
    print(f"Raw sequence: {sequence}")
