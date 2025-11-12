import os
import shutil

# Directories
img_dir = os.path.join("resources", "img")
json_dir = os.path.join("resources", "json")

# Collect all datetime stamps from both directories
datetime_stamps = set()

# Scan img directory
if os.path.exists(img_dir):
    for filename in os.listdir(img_dir):
        if filename.endswith('.png'):
            stamp = filename.replace('.png', '')
            datetime_stamps.add(stamp)

# Scan json directory
if os.path.exists(json_dir):
    for filename in os.listdir(json_dir):
        if filename.endswith('.json'):
            stamp = filename.replace('.json', '')
            datetime_stamps.add(stamp)

# Sort datetime stamps and create mapping
sorted_stamps = sorted(datetime_stamps)
stamp_to_id = {stamp: idx + 1 for idx, stamp in enumerate(sorted_stamps)}

print(f"Found {len(sorted_stamps)} unique datetime stamps")
print("\nMapping:")
for stamp, file_id in stamp_to_id.items():
    print(f"  {stamp} -> {file_id}")

# Rename files in img directory
if os.path.exists(img_dir):
    for filename in os.listdir(img_dir):
        if filename.endswith('.png'):
            stamp = filename.replace('.png', '')
            if stamp in stamp_to_id:
                new_id = stamp_to_id[stamp]
                old_path = os.path.join(img_dir, filename)
                new_path = os.path.join(img_dir, f"{new_id}.png")
                shutil.move(old_path, new_path)
                print(f"Renamed: {filename} -> {new_id}.png")

# Rename files in json directory
if os.path.exists(json_dir):
    for filename in os.listdir(json_dir):
        if filename.endswith('.json'):
            stamp = filename.replace('.json', '')
            if stamp in stamp_to_id:
                new_id = stamp_to_id[stamp]
                old_path = os.path.join(json_dir, filename)
                new_path = os.path.join(json_dir, f"{new_id}.json")
                shutil.move(old_path, new_path)
                print(f"Renamed: {filename} -> {new_id}.json")

print("\nRenaming complete!")