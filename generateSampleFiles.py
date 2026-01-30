import os
from datetime import datetime

# Pick your target parent folder
target_root = r"C:\Users\corri\Documents"

# Create a unique output folder every run to avoid locked/read-only leftovers
stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
base_dir = os.path.join(target_root, f"sample_files_{stamp}")

structure = {
    "Projects": {
        "Docs": {
            "readme.txt": "This is a sample README file.\n" * 200,
            "notes.txt": "Some notes about the project.\n" * 500,
        },
        "Code": {
            "app.py": "print('Hello, world!')\n" * 300,
            "utils.py": "# Utility functions\n" * 400,
        },
    },
    "Images": {
        "image_list.txt": "image1.jpg\nimage2.jpg\nimage3.jpg\n",
    },
    "EmptyFolder": {}
}

def create_structure(base_path, tree):
    for name, content in tree.items():
        path = os.path.join(base_path, name)

        if isinstance(content, dict):
            os.makedirs(path, exist_ok=True)
            create_structure(path, content)
        else:
            # Ensure parent folder exists
            os.makedirs(os.path.dirname(path), exist_ok=True)
            try:
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)
                print(f"✅ wrote {path}")
            except PermissionError as e:
                print(f"❌ permission denied writing {path}")
                raise

os.makedirs(base_dir, exist_ok=True)
create_structure(base_dir, structure)
print(f"\nDone. Created at: {base_dir}")
