import os
import re
import pandas as pd
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# -----------------------------
# Helpers
# -----------------------------
def is_blank(x) -> bool:
    return x is None or (isinstance(x, float) and pd.isna(x)) or str(x).strip() == ""

def quote(s: str) -> str:
    return f'"{s}"'

def sanitize_relpath(p: str) -> str:
    """Turn 'C:\\a\\b\\c' into 'a\\b\\c' (remove drive) for mirroring."""
    p = os.path.normpath(str(p))
    drive, tail = os.path.splitdrive(p)
    tail = tail.lstrip("\\/")
    return tail

def safe_norm(s: str) -> str:
    return os.path.normpath(str(s).strip()).replace('"', '')

def safe_str(s: str) -> str:
    return str(s).strip().replace('"', '')

# -----------------------------
# GUI App
# -----------------------------
class RoboCopyGui(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Robocopy Command Generator")
        self.geometry("920x620")
        self.minsize(920, 620)

        self.df = None
        self.file_path = None
        self.sheets = []

        # Variables
        self.input_path_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.dest_root_var = tk.StringVar()
        self.output_basename_var = tk.StringVar(value="generated_robocopy")

        self.mode_var = tk.StringVar(value="fullpath")  # "fullpath" or "folderfile"
        self.fullpath_col_var = tk.StringVar()
        self.folder_col_var = tk.StringVar()
        self.filename_col_var = tk.StringVar()

        self.mirror_var = tk.BooleanVar(value=True)
        self.pause_var = tk.BooleanVar(value=True)
        self.echo_var = tk.BooleanVar(value=True)
        self.dryrun_var = tk.BooleanVar(value=False)

        self.flags_var = tk.StringVar(value='/R:2 /W:1 /NP /NFL /NDL /NJH /NJS')

        # Build UI
        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 6}

        # Top: file picker
        file_frame = ttk.LabelFrame(self, text="Input Spreadsheet")
        file_frame.pack(fill="x", **pad)

        ttk.Label(file_frame, text="File:").grid(row=0, column=0, sticky="w", padx=8, pady=8)
        ttk.Entry(file_frame, textvariable=self.input_path_var, width=90).grid(row=0, column=1, sticky="we", padx=8, pady=8)
        ttk.Button(file_frame, text="Browse…", command=self.browse_file).grid(row=0, column=2, sticky="e", padx=8, pady=8)

        file_frame.columnconfigure(1, weight=1)

        # Sheet selector (Excel only)
        sheet_frame = ttk.Frame(file_frame)
        sheet_frame.grid(row=1, column=0, columnspan=3, sticky="we", padx=8, pady=(0, 8))
        ttk.Label(sheet_frame, text="Sheet (Excel):").grid(row=0, column=0, sticky="w")
        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.sheet_var, state="disabled", width=40)
        self.sheet_combo.grid(row=0, column=1, sticky="w", padx=(8, 0))
        ttk.Button(sheet_frame, text="Load Columns", command=self.load_dataframe).grid(row=0, column=2, sticky="w", padx=12)

        # Destination + output settings
        out_frame = ttk.LabelFrame(self, text="Output Settings")
        out_frame.pack(fill="x", **pad)

        ttk.Label(out_frame, text="Destination root folder:").grid(row=0, column=0, sticky="w", padx=8, pady=8)
        ttk.Entry(out_frame, textvariable=self.dest_root_var, width=70).grid(row=0, column=1, sticky="we", padx=8, pady=8)
        ttk.Button(out_frame, text="Browse…", command=self.browse_dest).grid(row=0, column=2, sticky="e", padx=8, pady=8)

        ttk.Label(out_frame, text="Output base name:").grid(row=1, column=0, sticky="w", padx=8, pady=8)
        ttk.Entry(out_frame, textvariable=self.output_basename_var, width=30).grid(row=1, column=1, sticky="w", padx=8, pady=8)

        out_frame.columnconfigure(1, weight=1)

        # Mode selection + columns
        mode_frame = ttk.LabelFrame(self, text="How to Read Source Info")
        mode_frame.pack(fill="x", **pad)

        # Mode radios
        radios = ttk.Frame(mode_frame)
        radios.pack(fill="x", padx=8, pady=(8, 0))

        ttk.Radiobutton(radios, text="Use a FULL Source Path column (includes filename)",
                        variable=self.mode_var, value="fullpath", command=self._refresh_mode).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(radios, text="Use Folder Path + File Name columns",
                        variable=self.mode_var, value="folderfile", command=self._refresh_mode).grid(row=1, column=0, sticky="w", pady=(6, 0))

        # Column selectors
        cols = ttk.Frame(mode_frame)
        cols.pack(fill="x", padx=8, pady=10)

        ttk.Label(cols, text="Full Source Path column:").grid(row=0, column=0, sticky="w")
        self.fullpath_combo = ttk.Combobox(cols, textvariable=self.fullpath_col_var, state="disabled", width=45)
        self.fullpath_combo.grid(row=0, column=1, sticky="w", padx=8)

        ttk.Label(cols, text="Folder Path column:").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.folder_combo = ttk.Combobox(cols, textvariable=self.folder_col_var, state="disabled", width=45)
        self.folder_combo.grid(row=1, column=1, sticky="w", padx=8, pady=(10, 0))

        ttk.Label(cols, text="File Name column:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.filename_combo = ttk.Combobox(cols, textvariable=self.filename_col_var, state="disabled", width=45)
        self.filename_combo.grid(row=2, column=1, sticky="w", padx=8, pady=(10, 0))

        # Options
        opt_frame = ttk.LabelFrame(self, text="Robocopy Options")
        opt_frame.pack(fill="x", **pad)

        ttk.Checkbutton(opt_frame, text="Mirror source folder structure under Destination Root",
                        variable=self.mirror_var).grid(row=0, column=0, sticky="w", padx=8, pady=6)
        ttk.Checkbutton(opt_frame, text="Add echo lines", variable=self.echo_var).grid(row=0, column=1, sticky="w", padx=8, pady=6)
        ttk.Checkbutton(opt_frame, text="Add pause at end", variable=self.pause_var).grid(row=0, column=2, sticky="w", padx=8, pady=6)
        ttk.Checkbutton(opt_frame, text="Dry run (/L)", variable=self.dryrun_var).grid(row=0, column=3, sticky="w", padx=8, pady=6)

        ttk.Label(opt_frame, text="Extra flags:").grid(row=1, column=0, sticky="w", padx=8, pady=6)
        ttk.Entry(opt_frame, textvariable=self.flags_var, width=90).grid(row=1, column=1, columnspan=3, sticky="we", padx=8, pady=6)
        opt_frame.columnconfigure(3, weight=1)

        # Generate button + preview
        bottom = ttk.Frame(self)
        bottom.pack(fill="both", expand=True, **pad)

        ttk.Button(bottom, text="Generate .BAT + .TXT (+ errors.xlsx)", command=self.generate).pack(anchor="w")

        self.preview = tk.Text(bottom, height=18, wrap="none")
        self.preview.pack(fill="both", expand=True, pady=(10, 0))

        # scrollbars
        yscroll = ttk.Scrollbar(self.preview, orient="vertical", command=self.preview.yview)
        self.preview.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side="right", fill="y")

        self._refresh_mode()

    def _refresh_mode(self):
        mode = self.mode_var.get()
        if mode == "fullpath":
            self.fullpath_combo.configure(state="readonly" if self.df is not None else "disabled")
            self.folder_combo.configure(state="disabled")
            self.filename_combo.configure(state="disabled")
        else:
            self.fullpath_combo.configure(state="disabled")
            st = "readonly" if self.df is not None else "disabled"
            self.folder_combo.configure(state=st)
            self.filename_combo.configure(state=st)

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select spreadsheet",
            filetypes=[("Excel", "*.xlsx *.xlsm *.xls"), ("CSV", "*.csv"), ("All files", "*.*")]
        )
        if not path:
            return
        self.file_path = path
        self.input_path_var.set(path)
        self.df = None
        self.preview.delete("1.0", tk.END)

        # set up sheet combobox for Excel
        ext = Path(path).suffix.lower()
        if ext in [".xlsx", ".xlsm", ".xls"]:
            try:
                xl = pd.ExcelFile(path)
                self.sheets = xl.sheet_names
                self.sheet_combo.configure(state="readonly")
                self.sheet_combo["values"] = self.sheets
                self.sheet_var.set(self.sheets[0] if self.sheets else "")
            except Exception as e:
                messagebox.showerror("Error", f"Could not read Excel sheets:\n\n{e}")
                self.sheet_combo.configure(state="disabled")
                self.sheet_combo["values"] = []
                self.sheet_var.set("")
        else:
            self.sheet_combo.configure(state="disabled")
            self.sheet_combo["values"] = []
            self.sheet_var.set("")

    def browse_dest(self):
        folder = filedialog.askdirectory(title="Select destination root folder")
        if not folder:
            return

        folder = os.path.normpath(folder)

        self.dest_root_var.set(folder)

    # force refresh (Windows ttk quirk)
        self.update_idletasks()



    def load_dataframe(self):
        path = self.input_path_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Missing file", "Pick a valid spreadsheet first.")
            return

        try:
            ext = Path(path).suffix.lower()
            if ext == ".csv":
                df = pd.read_csv(path)
            else:
                sheet = self.sheet_var.get()
                if is_blank(sheet):
                    sheet = 0
                df = pd.read_excel(path, sheet_name=sheet)
        except Exception as e:
            messagebox.showerror("Load failed", f"Could not load spreadsheet:\n\n{e}")
            return

        self.df = df
        cols = list(df.columns)

        # Populate column selectors
        for combo in [self.fullpath_combo, self.folder_combo, self.filename_combo]:
            combo["values"] = cols

        # attempt some smart defaults
        self._smart_set_defaults(cols)

        self.preview.delete("1.0", tk.END)
        self.preview.insert(tk.END, f"Loaded {len(df)} rows and {len(cols)} columns.\n\n")
        self.preview.insert(tk.END, "Columns:\n- " + "\n- ".join(map(str, cols)) + "\n")

        self._refresh_mode()

    def _smart_set_defaults(self, cols):
        def find_col(candidates):
            lower_map = {str(c).lower(): c for c in cols}
            for cand in candidates:
                if cand.lower() in lower_map:
                    return lower_map[cand.lower()]
            # partial match
            for c in cols:
                cl = str(c).lower()
                for cand in candidates:
                    if cand.lower() in cl:
                        return c
            return ""

        self.fullpath_col_var.set(find_col(["File Path", "Full Path", "Source Path", "Path"]))
        self.folder_col_var.set(find_col(["Folder Path", "Source Folder", "Directory", "Folder"]))
        self.filename_col_var.set(find_col(["File Name", "Filename", "Name"]))

    def _build_dest_folder(self, source_folder: str) -> str:
        dest_root = self.dest_root_var.get().strip()
        if is_blank(dest_root):
            return ""
        if not self.mirror_var.get():
            return dest_root
        rel = sanitize_relpath(source_folder)
        return str(Path(dest_root) / rel)

    def generate(self):
        if self.df is None:
            messagebox.showwarning("Not loaded", "Click 'Load Columns' first.")
            return

        dest_root = self.dest_root_var.get().strip()
        if is_blank(dest_root):
            messagebox.showwarning("Missing destination", "Set a Destination root folder.")
            return

        out_base = self.output_basename_var.get().strip() or "generated_robocopy"
        input_path = self.input_path_var.get().strip()
        out_dir = str(Path(input_path).parent)

        bat_path = str(Path(out_dir) / f"{out_base}.bat")
        txt_path = str(Path(out_dir) / f"{out_base}.txt")
        err_path = str(Path(out_dir) / f"{out_base}_errors.xlsx")

        mode = self.mode_var.get()
        flags = self.flags_var.get().strip()
        if self.dryrun_var.get() and "/L" not in flags.upper():
            flags = (flags + " /L").strip()

        commands = []
        errors = []

        cols = set(self.df.columns)

        full_col = self.fullpath_col_var.get()
        folder_col = self.folder_col_var.get()
        name_col = self.filename_col_var.get()

        if mode == "fullpath":
            if is_blank(full_col) or full_col not in cols:
                messagebox.showwarning("Column needed", "Pick a valid 'Full Source Path' column.")
                return
        else:
            if is_blank(folder_col) or folder_col not in cols or is_blank(name_col) or name_col not in cols:
                messagebox.showwarning("Columns needed", "Pick valid 'Folder Path' and 'File Name' columns.")
                return

        for i, row in self.df.iterrows():
            excelish_row = i + 2  # header row counts as 1

            try:
                if mode == "fullpath":
                    v = row.get(full_col)
                    if is_blank(v):
                        errors.append({"row": excelish_row, "reason": "Blank full source path", "details": ""})
                        continue
                    full_source = safe_norm(v)
                    source_folder = os.path.dirname(full_source)
                    filename = os.path.basename(full_source)
                else:
                    fv = row.get(folder_col)
                    nv = row.get(name_col)
                    if is_blank(fv) or is_blank(nv):
                        errors.append({
                            "row": excelish_row,
                            "reason": "Blank folder or filename",
                            "details": f"folder={fv!r}, name={nv!r}"
                        })
                        continue
                    source_folder = safe_norm(fv)
                    filename = safe_str(nv)

                dest_folder = self._build_dest_folder(source_folder)
                if is_blank(dest_folder):
                    errors.append({"row": excelish_row, "reason": "Destination root not set", "details": ""})
                    continue

                cmd = f'robocopy {quote(source_folder)} {quote(dest_folder)} {quote(filename)} {flags}'
                commands.append(cmd)

            except Exception as e:
                errors.append({"row": excelish_row, "reason": "Exception building command", "details": str(e)})

        # Write outputs
        try:
            Path(out_dir).mkdir(parents=True, exist_ok=True)

            with open(txt_path, "w", encoding="utf-8") as f:
                f.write("\n".join(commands) + ("\n" if commands else ""))

            with open(bat_path, "w", encoding="utf-8") as f:
                f.write("@echo off\n")
                f.write("setlocal enableextensions\n")
                if self.echo_var.get():
                    f.write("echo Running robocopy commands...\n")
                    f.write("echo.\n")

                # mkdir each unique destination folder (nice-to-have)
                dests = sorted(set(re.findall(r'robocopy\s+"[^"]+"\s+"([^"]+)"', "\n".join(commands))))
                for d in dests:
                    f.write(f'if not exist "{d}" mkdir "{d}"\n')

                f.write("\n")
                for c in commands:
                    f.write(c + "\n")

                if self.echo_var.get():
                    f.write("\necho.\n")
                    f.write("echo Done.\n")
                if self.pause_var.get():
                    f.write("pause\n")
                f.write("endlocal\n")

            if errors:
                pd.DataFrame(errors).to_excel(err_path, index=False)

        except Exception as e:
            messagebox.showerror("Write failed", f"Could not write output files:\n\n{e}")
            return

        # Preview
        self.preview.delete("1.0", tk.END)
        self.preview.insert(tk.END, f"✅ Generated {len(commands)} commands\n")
        self.preview.insert(tk.END, f"- BAT: {bat_path}\n")
        self.preview.insert(tk.END, f"- TXT: {txt_path}\n")
        if errors:
            self.preview.insert(tk.END, f"- Errors: {err_path}  ({len(errors)} rows had issues)\n")
        else:
            self.preview.insert(tk.END, "- Errors: none\n")

        self.preview.insert(tk.END, "\n--- Preview (first 25 commands) ---\n")
        for c in commands[:25]:
            self.preview.insert(tk.END, c + "\n")

        messagebox.showinfo("Done", f"Generated {len(commands)} commands.\n\nSaved in:\n{out_dir}")

# -----------------------------
# Run
# -----------------------------
if __name__ == "__main__":

    app = RoboCopyGui()
    app.mainloop()
