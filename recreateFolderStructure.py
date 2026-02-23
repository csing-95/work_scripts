import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

class FolderRebuilderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Folder Structure Rebuilder")
        self.root.geometry("950x650")

        self.source_path = tk.StringVar()
        self.destination_path = tk.StringVar()
        self.safe_mode = tk.BooleanVar(value=True)
        self.include_files = tk.BooleanVar(value=False)
        self.default_filename = tk.StringVar(value="folder_structure_template.txt")

        self.current_structure = []

        self.create_tabs()

    # ---------------------------
    # Tabs
    # ---------------------------

    def create_tabs(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill="both", expand=True)

        self.main_tab = ttk.Frame(notebook)
        self.help_tab = ttk.Frame(notebook)

        notebook.add(self.main_tab, text="Tool")
        notebook.add(self.help_tab, text="Guide")

        self.build_main_tab()
        self.build_help_tab()

    # ---------------------------
    # Main UI
    # ---------------------------

    def build_main_tab(self):
        frame = self.main_tab

        ttk.Label(frame, text="Source Folder (optional):").pack(anchor="w", padx=10, pady=(10,0))
        source_frame = ttk.Frame(frame)
        source_frame.pack(fill="x", padx=10)
        ttk.Entry(source_frame, textvariable=self.source_path).pack(side="left", fill="x", expand=True)
        ttk.Button(source_frame, text="Browse", command=self.select_source).pack(side="left", padx=5)

        ttk.Label(frame, text="Or Paste Manual Folder Structure:").pack(anchor="w", padx=10, pady=(15,0))
        self.manual_text = tk.Text(frame, height=10)
        self.manual_text.pack(fill="both", padx=10)

        ttk.Label(frame, text="Destination Folder:").pack(anchor="w", padx=10, pady=(15,0))
        dest_frame = ttk.Frame(frame)
        dest_frame.pack(fill="x", padx=10)
        ttk.Entry(dest_frame, textvariable=self.destination_path).pack(side="left", fill="x", expand=True)
        ttk.Button(dest_frame, text="Browse", command=self.select_destination).pack(side="left", padx=5)

        options_frame = ttk.Frame(frame)
        options_frame.pack(fill="x", padx=10, pady=10)

        ttk.Checkbutton(options_frame, text="Safe Mode (Don't overwrite existing folders)", variable=self.safe_mode).pack(anchor="w")
        ttk.Checkbutton(options_frame, text="Include files (create empty files too)", variable=self.include_files).pack(anchor="w")

        button_frame = ttk.Frame(frame)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Preview Structure", command=self.preview_structure).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Create Structure", command=self.create_structure).pack(side="left", padx=5)

        # Save Template Section
        save_frame = ttk.LabelFrame(frame, text="Save Structure as Template")
        save_frame.pack(fill="x", padx=10, pady=10)

        ttk.Label(save_frame, text="Default File Name:").pack(anchor="w", padx=5)
        ttk.Entry(save_frame, textvariable=self.default_filename).pack(fill="x", padx=5, pady=5)

        ttk.Button(save_frame, text="Save Structure to TXT", command=self.save_structure_to_txt).pack(pady=5)

        ttk.Label(frame, text="Log:").pack(anchor="w", padx=10)
        self.log_box = tk.Text(frame, height=8, bg="#f4f4f4")
        self.log_box.pack(fill="both", padx=10, pady=(0,10))

    # ---------------------------
    # Help Tab
    # ---------------------------

    def build_help_tab(self):
        help_text = """
FOLDER STRUCTURE REBUILDER - USER GUIDE

NEW FEATURE:
You can now save any generated structure as a reusable .txt template.

HOW TEMPLATE SAVING WORKS:
1. Click Preview Structure
2. Adjust default file name if desired
3. Click Save Structure to TXT
4. Choose where to save the template

You can later paste that template back into the manual input area.

Template format:
Each line represents a folder path.
Example:
FolderA
FolderA/SubFolder1
FolderB
        """

        text_widget = tk.Text(self.help_tab)
        text_widget.insert("1.0", help_text)
        text_widget.config(state="disabled")
        text_widget.pack(fill="both", expand=True)

    # ---------------------------
    # Folder Selection
    # ---------------------------

    def select_source(self):
        path = filedialog.askdirectory()
        if path:
            self.source_path.set(path)

    def select_destination(self):
        path = filedialog.askdirectory()
        if path:
            self.destination_path.set(path)

    # ---------------------------
    # Core Logic
    # ---------------------------

    def get_structure(self):
        structure = []

        if self.source_path.get():
            for root, dirs, files in os.walk(self.source_path.get()):
                rel_path = os.path.relpath(root, self.source_path.get())
                if rel_path == ".":
                    rel_path = ""
                structure.append(rel_path)

                if self.include_files.get():
                    for file in files:
                        structure.append(os.path.join(rel_path, file))
        else:
            manual_input = self.manual_text.get("1.0", "end").strip().split("\n")
            structure = [line.strip() for line in manual_input if line.strip()]

        return structure

    def preview_structure(self):
        self.log_box.delete("1.0", "end")
        self.current_structure = self.get_structure()

        if not self.current_structure:
            messagebox.showwarning("Warning", "No structure found.")
            return

        for item in self.current_structure:
            self.log_box.insert("end", f"{item}\n")

    def create_structure(self):
        dest = self.destination_path.get()
        if not dest:
            messagebox.showerror("Error", "Please select a destination folder.")
            return

        if not self.current_structure:
            self.current_structure = self.get_structure()

        for item in self.current_structure:
            full_path = Path(dest) / item

            try:
                if "." in Path(item).name and self.include_files.get():
                    full_path.parent.mkdir(parents=True, exist_ok=True)
                    if not full_path.exists() or not self.safe_mode.get():
                        full_path.touch()
                else:
                    full_path.mkdir(parents=True, exist_ok=not self.safe_mode.get())

                self.log_box.insert("end", f"Created: {full_path}\n")

            except Exception as e:
                self.log_box.insert("end", f"Error: {e}\n")

        messagebox.showinfo("Done", "Folder structure creation complete.")

    def save_structure_to_txt(self):
        if not self.current_structure:
            messagebox.showwarning("Warning", "Please preview structure first.")
            return

        filename = self.default_filename.get().strip()
        if not filename.endswith(".txt"):
            filename += ".txt"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=filename,
            filetypes=[("Text Files", "*.txt")]
        )

        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    for item in self.current_structure:
                        f.write(f"{item}\n")

                self.log_box.insert("end", f"Template saved to: {file_path}\n")
                messagebox.showinfo("Success", "Template saved successfully.")

            except Exception as e:
                messagebox.showerror("Error", f"Failed to save template: {e}")

# ---------------------------
# Run
# ---------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = FolderRebuilderApp(root)
    root.mainloop()