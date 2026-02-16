import os

folder_path = r"./"

for filename in os.listdir(folder_path):
    if filename.lower().endswith(".pdf") and " " in filename:
        old_path = os.path.join(folder_path, filename)
        new_name = filename.replace(" ", "_")
        new_path = os.path.join(folder_path, new_name)

        if not os.path.exists(new_path):
            os.rename(old_path, new_path)
            print(f"Renamed: {filename} -> {new_name}")
        else:
            print(f"Skipped (already exists): {new_name}")