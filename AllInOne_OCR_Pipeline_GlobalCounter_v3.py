
import os
import shutil
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

# =========================================================
# Helpers
# =========================================================

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def parse_extensions(ext_text: str):
    parts = [p.strip().lower() for p in str(ext_text).split(",") if p.strip()]
    out = []
    for p in parts:
        if not p.startswith("."):
            p = "." + p
        out.append(p)
    return tuple(dict.fromkeys(out))


# =========================================================
# PRE-OCR RENAME (Size OR Global Counter)
# =========================================================

def pre_ocr_rename(
    source_root: str,
    mode: str,  # copy or rename
    staging_root: str,
    rename_mode: str,  # "size" or "counter"
    separator: str,
    padding: int,
    report_path: str,
    log_fn=print,
):
    if not os.path.isdir(source_root):
        raise ValueError("Source folder not found.")

    if mode == "copy":
        if not staging_root:
            raise ValueError("Staging root required in copy mode.")
        ensure_dir(staging_root)

    counter = 1
    rows = []

    for root, _, files in os.walk(source_root):
        files.sort(key=lambda s: s.lower())

        for name in files:
            src_full = os.path.join(root, name)

            base, ext = os.path.splitext(name)

            if rename_mode == "size":
                size = os.path.getsize(src_full)
                new_name = f"{base}{separator}{size}{ext}"
            else:
                new_name = f"{base}{separator}{str(counter).zfill(padding)}{ext}"
                counter += 1

            if mode == "rename":
                dest_full = os.path.join(root, new_name)
            else:
                rel = os.path.relpath(root, source_root)
                rel = "" if rel == "." else rel
                dest_dir = os.path.join(staging_root, rel)
                ensure_dir(dest_dir)
                dest_full = os.path.join(dest_dir, new_name)

            if os.path.exists(dest_full):
                base2, ext2 = os.path.splitext(dest_full)
                i = 1
                while os.path.exists(f"{base2}_{i}{ext2}"):
                    i += 1
                dest_full = f"{base2}_{i}{ext2}"

            if mode == "rename":
                os.rename(src_full, dest_full)
                action = "RENAMED"
            else:
                shutil.copy2(src_full, dest_full)
                action = "COPIED"

            rows.append({
                "original": src_full,
                "new_path": dest_full,
                "action": action
            })

            log_fn(f"✅ {action}: {name} -> {os.path.basename(dest_full)}")

    if report_path:
        ensure_dir(os.path.dirname(report_path))
        with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
            pd.DataFrame(rows).to_excel(writer, index=False, sheet_name="Renames")

    log_fn("Finished pre-OCR rename.")


# =========================================================
# SIMPLE REBUILD (base + suffix + ext)
# =========================================================

def rebuild_structure(
    original_root,
    flat_folder,
    dest_root,
    output_ext,
    output_suffix,
    move_files,
    log_fn=print,
):
    ensure_dir(dest_root)

    flat_files = {f.lower(): os.path.join(flat_folder, f)
                  for f in os.listdir(flat_folder)
                  if f.lower().endswith(output_ext.lower())}

    for root, _, files in os.walk(original_root):
        files.sort()
        for name in files:
            base = os.path.splitext(name)[0]
            expected = f"{base}{output_suffix}{output_ext}"
            expected_key = expected.lower()

            if expected_key not in flat_files:
                log_fn(f"⚠ Missing: {expected}")
                continue

            src = flat_files[expected_key]

            rel = os.path.relpath(root, original_root)
            rel = "" if rel == "." else rel
            dest_dir = os.path.join(dest_root, rel)
            ensure_dir(dest_dir)

            dest_path = os.path.join(dest_dir, expected)

            if move_files:
                shutil.move(src, dest_path)
            else:
                shutil.copy2(src, dest_path)

            log_fn(f"✅ Placed: {expected}")


# =========================================================
# GUI
# =========================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("All-in-One OCR Pipeline (Counter + Size)")
        self.geometry("1000x800")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        self.tab1 = ttk.Frame(nb)
        self.tab2 = ttk.Frame(nb)

        nb.add(self.tab1, text="Pre-OCR Rename")
        nb.add(self.tab2, text="Rebuild Structure")

        self.build_tab1()
        self.build_tab2()

    def build_tab1(self):
        self.src = tk.StringVar()
        self.staging = tk.StringVar()
        self.mode = tk.StringVar(value="copy")
        self.rename_mode = tk.StringVar(value="counter")
        self.separator = tk.StringVar(value="__")
        self.padding = tk.IntVar(value=6)
        self.report = tk.StringVar()

        frm = ttk.Frame(self.tab1)
        frm.pack(fill="x", padx=10, pady=10)

        ttk.Label(frm, text="Source Root").pack(anchor="w")
        ttk.Entry(frm, textvariable=self.src, width=100).pack()
        ttk.Button(frm, text="Browse", command=lambda: self.pick_folder(self.src)).pack()

        ttk.Label(frm, text="Staging Root (copy mode)").pack(anchor="w", pady=(10,0))
        ttk.Entry(frm, textvariable=self.staging, width=100).pack()
        ttk.Button(frm, text="Browse", command=lambda: self.pick_folder(self.staging)).pack()

        ttk.Label(frm, text="Rename Mode").pack(anchor="w", pady=(10,0))
        ttk.Radiobutton(frm, text="Global Counter", variable=self.rename_mode, value="counter").pack(anchor="w")
        ttk.Radiobutton(frm, text="File Size", variable=self.rename_mode, value="size").pack(anchor="w")

        ttk.Label(frm, text="Counter Padding").pack(anchor="w")
        ttk.Entry(frm, textvariable=self.padding, width=10).pack(anchor="w")

        ttk.Label(frm, text="Separator").pack(anchor="w")
        ttk.Entry(frm, textvariable=self.separator, width=10).pack(anchor="w")

        ttk.Label(frm, text="Report Path (.xlsx)").pack(anchor="w", pady=(10,0))
        ttk.Entry(frm, textvariable=self.report, width=100).pack()
        ttk.Button(frm, text="Save As", command=lambda: self.pick_report(self.report)).pack()

        ttk.Button(frm, text="Run", command=self.run_tab1).pack(pady=10)

        self.log1 = tk.Text(self.tab1, height=20)
        self.log1.pack(fill="both", expand=True)

    def build_tab2(self):
        self.orig = tk.StringVar()
        self.flat = tk.StringVar()
        self.dest = tk.StringVar()
        self.output_ext = tk.StringVar(value=".pdf")
        self.output_suffix = tk.StringVar(value="_OCR")
        self.move = tk.BooleanVar(value=True)

        frm = ttk.Frame(self.tab2)
        frm.pack(fill="x", padx=10, pady=10)

        ttk.Label(frm, text="Original Root").pack(anchor="w")
        ttk.Entry(frm, textvariable=self.orig, width=100).pack()
        ttk.Button(frm, text="Browse", command=lambda: self.pick_folder(self.orig)).pack()

        ttk.Label(frm, text="Flat OCR Output Folder").pack(anchor="w")
        ttk.Entry(frm, textvariable=self.flat, width=100).pack()
        ttk.Button(frm, text="Browse", command=lambda: self.pick_folder(self.flat)).pack()

        ttk.Label(frm, text="Destination Root").pack(anchor="w")
        ttk.Entry(frm, textvariable=self.dest, width=100).pack()
        ttk.Button(frm, text="Browse", command=lambda: self.pick_folder(self.dest)).pack()

        ttk.Label(frm, text="Output Extension").pack(anchor="w")
        ttk.Entry(frm, textvariable=self.output_ext, width=10).pack(anchor="w")

        ttk.Label(frm, text="Output Suffix (e.g. _OCR)").pack(anchor="w")
        ttk.Entry(frm, textvariable=self.output_suffix, width=10).pack(anchor="w")

        ttk.Checkbutton(frm, text="Move Files (unchecked = copy)", variable=self.move).pack(anchor="w", pady=10)

        ttk.Button(frm, text="Run", command=self.run_tab2).pack()

        self.log2 = tk.Text(self.tab2, height=20)
        self.log2.pack(fill="both", expand=True)

    def run_tab1(self):
        pre_ocr_rename(
            self.src.get(),
            "copy" if self.mode.get()=="copy" else "rename",
            self.staging.get(),
            self.rename_mode.get(),
            self.separator.get(),
            self.padding.get(),
            self.report.get(),
            log_fn=lambda m: self.log1.insert("end", m + "\n")
        )

    def run_tab2(self):
        rebuild_structure(
            self.orig.get(),
            self.flat.get(),
            self.dest.get(),
            self.output_ext.get(),
            self.output_suffix.get(),
            self.move.get(),
            log_fn=lambda m: self.log2.insert("end", m + "\n")
        )

    def pick_folder(self, var):
        p = filedialog.askdirectory()
        if p:
            var.set(p)

    def pick_report(self, var):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if p:
            var.set(p)


if __name__ == "__main__":
    App().mainloop()
