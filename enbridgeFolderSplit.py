import os
import re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


# -----------------------------
# Config
# -----------------------------
# CATEGORIES = [
#     "AAA_GRM_VAULT",
#     "CONSTRUCTION_SET_DRAWINGS_CAD_FILES",
#     "DGNS_AND_MORE_ARCHIVE",
#     "FIELD_ASBUILT_SCANS",
#     "MERIDIAN_MIGRATION_UPLOADS_STILL_NEEDED",
#     "OFFSHORE",
#     "REVISION_DONOTUSE",
#     "SABALTRAIL",
#     "VECTOR",
# ]

CATEGORIES = [
    "MERIDIAN MIGRATION UPLOADS STILL NEEDED",
    "OFFSHORE",
]

PATH_COL_CANDIDATES = [
    "Original File Path",
    "Original_File_Path",
    "OriginalFilePath",
]

EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")


# -----------------------------
# Helpers
# -----------------------------
def is_excel_file(filename: str) -> bool:
    if filename.startswith("~$"):
        return False
    return filename.lower().endswith(EXCEL_EXTS)


def safe_sheet_name(name: str) -> str:
    r"""Excel sheet names max 31 chars, can't contain : \ / ? * [ ]"""
    bad = r'[:\\/?*\[\]]'
    name = re.sub(bad, "_", name)
    return name[:31]


def safe_filename(name: str) -> str:
    """Make a safe Windows filename."""
    name = name.strip()
    if not name:
        return "output"
    bad = r'[<>:"/\\|?*\x00-\x1F]'
    name = re.sub(bad, "_", name)
    name = name.rstrip(". ")
    return name[:150]


def build_category_regex(cat: str) -> re.Pattern:
    """
    Match category in path even if underscores become spaces.
    Example: AAA_GRM_VAULT matches AAA GRM VAULT or AAA_GRM_VAULT
    """
    esc = re.escape(cat)
    esc = esc.replace(r"\_", r"[_\s]+")
    return re.compile(esc, re.IGNORECASE)


CATEGORY_PATTERNS = {c: build_category_regex(c) for c in CATEGORIES}


def detect_path_column(columns) -> str | None:
    cols = list(columns)
    for c in PATH_COL_CANDIDATES:
        if c in cols:
            return c

    lowered = {c: str(c).lower() for c in cols}
    for c, lc in lowered.items():
        if "original" in lc and "path" in lc:
            return c
    return None


def pick_best_sheet(xls: pd.ExcelFile) -> str:
    """
    Prefer 'Documents' if it has Original File Path column.
    Otherwise return first sheet that contains a path column.
    """
    if "Documents" in xls.sheet_names:
        try:
            test = pd.read_excel(xls, sheet_name="Documents", nrows=5)
            if detect_path_column(test.columns):
                return "Documents"
        except Exception:
            pass

    for s in xls.sheet_names:
        try:
            test = pd.read_excel(xls, sheet_name=s, nrows=5)
            if detect_path_column(test.columns):
                return s
        except Exception:
            continue

    return xls.sheet_names[0]


def assign_category(path_value: str) -> str:
    if not isinstance(path_value, str):
        return "OTHER"
    for cat in CATEGORIES:
        if CATEGORY_PATTERNS[cat].search(path_value):
            return cat
    return "OTHER"


def read_one_excel(file_path: str) -> pd.DataFrame:
    xls = pd.ExcelFile(file_path)
    sheet = pick_best_sheet(xls)
    df = pd.read_excel(xls, sheet_name=sheet)

    path_col = detect_path_column(df.columns)
    if not path_col:
        raise ValueError(
            f"No 'Original File Path' column found in any sheet of: {os.path.basename(file_path)}"
        )

    if path_col != "Original File Path":
        df = df.rename(columns={path_col: "Original File Path"})

    df["Source File"] = os.path.basename(file_path)
    df["Source Sheet"] = sheet
    return df


def merge_and_split_to_files(input_folder: str, output_folder: str, output_prefix: str, log_fn):
    files = [
        os.path.join(input_folder, f)
        for f in os.listdir(input_folder)
        if is_excel_file(f)
    ]
    if not files:
        raise ValueError("No Excel files found in that folder.")

    log_fn(f"Found {len(files)} Excel files.")

    all_dfs = []
    for i, fp in enumerate(files, start=1):
        log_fn(f"[{i}/{len(files)}] Reading: {os.path.basename(fp)}")
        df = read_one_excel(fp)
        all_dfs.append(df)

    merged = pd.concat(all_dfs, ignore_index=True)

    if "Original File Path" not in merged.columns:
        raise ValueError("Merged data does not contain 'Original File Path' column (unexpected).")

    # Duplicate marker across ALL merged rows (based on filepath only)
    merged["Filepath Dupe?"] = merged["Original File Path"].duplicated(keep=False).map(
        {True: "YES", False: "NO"}
    )

    # Category assignment (first match wins)
    merged["Category"] = merged["Original File Path"].apply(assign_category)

    # Output files: one per category (plus OTHER)
    prefix = safe_filename(output_prefix.replace(".xlsx", ""))

    created = 0
    for cat in CATEGORIES + ["OTHER"]:
        part = merged[merged["Category"] == cat].copy()
        if part.empty:
            log_fn(f"Skipping {cat} (no rows).")
            continue

        out_name = f"{prefix}_{cat}.xlsx" if prefix else f"{cat}.xlsx"
        out_name = safe_filename(out_name)
        out_path = os.path.join(output_folder, out_name)

        log_fn(f"Writing: {out_name} ({len(part):,} rows)")
        # one sheet per file (nice + simple)
        part.to_excel(out_path, index=False)

        created += 1

    if created == 0:
        raise ValueError("No output files were created (no category matches found).")

    log_fn(f"Done ✅ Created {created} file(s) in: {output_folder}")


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Merger → Split to Separate Files")
        self.geometry("820x540")
        self.minsize(760, 520)

        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        # This is now a PREFIX (not a single output filename)
        self.output_prefix = tk.StringVar(value="ENB_SPLIT")

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True)

        # Input folder
        row1 = ttk.LabelFrame(frm, text="1) Input folder (Excel files)")
        row1.pack(fill="x", **pad)

        ttk.Entry(row1, textvariable=self.input_folder).pack(
            side="left", fill="x", expand=True, padx=8, pady=8
        )
        ttk.Button(row1, text="Browse…", command=self.browse_input).pack(
            side="left", padx=8, pady=8
        )

        # Output
        row2 = ttk.LabelFrame(frm, text="2) Output")
        row2.pack(fill="x", **pad)

        out_grid = ttk.Frame(row2)
        out_grid.pack(fill="x", padx=8, pady=8)

        ttk.Label(out_grid, text="Output folder:").grid(row=0, column=0, sticky="w")
        ttk.Entry(out_grid, textvariable=self.output_folder).grid(
            row=0, column=1, sticky="ew", padx=6
        )
        ttk.Button(out_grid, text="Browse…", command=self.browse_output).grid(
            row=0, column=2, padx=6
        )

        ttk.Label(out_grid, text="Output file prefix:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(out_grid, textvariable=self.output_prefix).grid(
            row=1, column=1, sticky="ew", padx=6, pady=(8, 0)
        )
        ttk.Label(out_grid, text="(e.g. ENB_RUN_01 → ENB_RUN_01_VECTOR.xlsx)").grid(
            row=2, column=1, sticky="w", padx=6, pady=(4, 0)
        )

        out_grid.columnconfigure(1, weight=1)

        # Run button + progress
        row3 = ttk.Frame(frm)
        row3.pack(fill="x", **pad)

        self.run_btn = ttk.Button(row3, text="Run merge + split", command=self.on_run)
        self.run_btn.pack(side="left")

        self.prog = ttk.Progressbar(row3, mode="indeterminate")
        self.prog.pack(side="left", fill="x", expand=True, padx=10)

        # Log box
        row4 = ttk.LabelFrame(frm, text="Log")
        row4.pack(fill="both", expand=True, **pad)

        self.log = tk.Text(row4, height=14, wrap="word")
        self.log.pack(fill="both", expand=True, padx=8, pady=8)

        tip = (
            "Output created:\n"
            "• One Excel file per category (AAA_GRM_VAULT, …, VECTOR)\n"
            "• OTHER.xlsx for anything that doesn't match\n\n"
            "Adds columns: Source File, Source Sheet, Filepath Dupe?, Category\n"
        )
        self.log.insert("end", tip + "\n")

    def browse_input(self):
        folder = filedialog.askdirectory(title="Select folder containing Excel files")
        if folder:
            self.input_folder.set(folder)
            if not self.output_folder.get():
                self.output_folder.set(folder)

    def browse_output(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_folder.set(folder)

    def log_line(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.update_idletasks()

    def on_run(self):
        in_dir = self.input_folder.get().strip()
        out_dir = self.output_folder.get().strip()
        prefix = self.output_prefix.get().strip()

        if not in_dir or not os.path.isdir(in_dir):
            messagebox.showerror("Missing input", "Please choose a valid input folder.")
            return

        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showerror("Missing output", "Please choose a valid output folder.")
            return

        if not prefix:
            # allow blank prefix (will just name by category)
            prefix = ""

        self.run_btn.config(state="disabled")
        self.prog.start(10)
        self.log_line("Starting…")

        def worker():
            try:
                merge_and_split_to_files(in_dir, out_dir, prefix, self.log_line)
                messagebox.showinfo("Done", f"Saved outputs in:\n{out_dir}")
            except Exception as e:
                self.log_line(f"ERROR: {e}")
                messagebox.showerror("Error", str(e))
            finally:
                self.prog.stop()
                self.run_btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    # ttk theme polish (optional)
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    app = App()
    app.mainloop()
