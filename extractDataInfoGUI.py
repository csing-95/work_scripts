import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

# ----------------------------
# Helpers
# ----------------------------
def safe_upper_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.upper()

def is_blank_series(s: pd.Series) -> pd.Series:
    return s.isnull() | (s.astype(str).str.strip() == "")

def basename_from_path(p: pd.Series) -> pd.Series:
    # Handles NaN safely
    return p.fillna("").astype(str).apply(lambda x: os.path.basename(x))

def stem_from_filename(fn: pd.Series) -> pd.Series:
    # Remove extension
    return fn.fillna("").astype(str).apply(lambda x: os.path.splitext(x)[0])


# ----------------------------
# Core analysis (returns text + dfs)
# ----------------------------
def analyze_document_data(
    df: pd.DataFrame,
    col_name: str,
    col_number: str,
    col_size: str,
    col_type: str,
    col_path: str | None,
    col_rev: str | None,
    dupe_cols: list[str],
    derive_filename_from_path: bool,
    derive_stem_from_filename: bool,
):
    lines = []

    # Work on a copy so we can add derived cols without messing original view
    df2 = df.copy()

    # Derived fields (optional)
    derived_cols = []
    if col_path and derive_filename_from_path:
        df2["__FileName"] = basename_from_path(df2[col_path])
        derived_cols.append("__FileName")
        if derive_stem_from_filename:
            df2["__FileStem"] = stem_from_filename(df2["__FileName"])
            derived_cols.append("__FileStem")

    # 1. Total count
    lines.append(f"1. Total File Count (Rows): {len(df2)}")

    # 2. File type breakdown
    if col_type:
        type_breakdown = safe_upper_series(df2[col_type]).value_counts(dropna=False)
        lines.append("\n2. File Type Breakdown:")
        lines.append(type_breakdown.to_string())

    # 3. No filesize
    no_size = df2[is_blank_series(df2[col_size])] if col_size else pd.DataFrame()
    lines.append(f"\n3. Files with No Filesizes: {len(no_size)} found")
    if not no_size.empty:
        lines.append("   Sample records (first 5):")
        sample_cols = [c for c in [col_name, col_size, col_path] if c]
        lines.append(no_size[sample_cols].head(5).to_string(index=False))

    # 4. No filetype
    no_type = df2[is_blank_series(df2[col_type])] if col_type else pd.DataFrame()
    lines.append(f"\n4. Files with No Filetypes: {len(no_type)} found")
    if not no_type.empty:
        lines.append("   Sample records (first 5):")
        sample_cols = [c for c in [col_name, col_type, col_path] if c]
        lines.append(no_type[sample_cols].head(5).to_string(index=False))

    # 5. Duplicates (user-chosen columns)
    if not dupe_cols:
        raise ValueError("No duplicate columns selected.")

    duplicates_mask = df2.duplicated(subset=dupe_cols, keep="first")
    duplicates = df2[duplicates_mask]
    lines.append(f"\n5. Total Duplicates (based on {', '.join(dupe_cols)}): {len(duplicates)}")

    all_dupe_rows = pd.DataFrame()
    if len(duplicates) > 0:
        all_dupe_rows = df2[df2.duplicated(subset=dupe_cols, keep=False)].sort_values(by=dupe_cols)
        lines.append("   Sample records of Duplicates (first 10 of the full set):")
        show_cols = dupe_cols.copy()
        # add path/name columns for context if not already included
        for extra in [col_path, col_name, col_number, col_rev]:
            if extra and extra not in show_cols:
                show_cols.append(extra)
        # add derived cols if they exist and not already there
        for dc in derived_cols:
            if dc in df2.columns and dc not in show_cols:
                show_cols.append(dc)
        lines.append(all_dupe_rows[show_cols].head(10).to_string(index=False))

    # 6. Revision stacks by Document Number (if provided)
    revision_stacks = pd.Series(dtype=int)
    if col_number:
        stack_counts = df2[col_number].value_counts(dropna=False)
        revision_stacks = stack_counts[stack_counts > 1]

        lines.append("\n6. Document Revision Stacks Found:")
        lines.append(f"   Total unique Document Numbers: {len(stack_counts)}")
        lines.append(f"   Document Numbers with Revisions (Count > 1): {len(revision_stacks)}")
        lines.append(f"   Total number of documents in a revision stack: {int(revision_stacks.sum())}")

        if not revision_stacks.empty:
            lines.append("   Top 5 largest revision stacks (Document Number: Count):")
            lines.append(revision_stacks.head(5).to_string())

    # Return text + useful frames
    results_text = "\n".join(lines)

    # Strip derived columns in exported frames? Keep them—it’s helpful.
    return results_text, no_size, no_type, duplicates, all_dupe_rows, revision_stacks


# ----------------------------
# GUI
# ----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Document Data Analyzer (GUI)")
        self.geometry("1100x780")

        self.file_path = tk.StringVar(value="")
        self.sheet_name = tk.StringVar(value="")
        self.status = tk.StringVar(value="Select an Excel file to begin.")

        self.df = None
        self.columns = []

        # column mapping
        self.col_name = tk.StringVar()
        self.col_number = tk.StringVar()
        self.col_size = tk.StringVar()
        self.col_type = tk.StringVar()
        self.col_path = tk.StringVar()
        self.col_rev = tk.StringVar()

        # dupe options checkboxes
        self.use_doc_name = tk.BooleanVar(value=True)
        self.use_doc_number = tk.BooleanVar(value=False)
        self.use_revision = tk.BooleanVar(value=False)
        self.use_file_size = tk.BooleanVar(value=True)
        self.use_file_type = tk.BooleanVar(value=True)
        self.use_file_path = tk.BooleanVar(value=False)

        # derived options
        self.derive_filename_from_path = tk.BooleanVar(value=False)
        self.derive_stem_from_filename = tk.BooleanVar(value=False)
        self.use_derived_filename = tk.BooleanVar(value=False)  # __FileName
        self.use_derived_stem = tk.BooleanVar(value=False)      # __FileStem

        # outputs
        self.out_no_size = None
        self.out_no_type = None
        self.out_duplicates = None
        self.out_all_dupes = None
        self.out_revision_stacks = None

        self._build_ui()

    def _build_ui(self):
        # File picker row
        top = ttk.Frame(self, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="Excel file:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.file_path, width=90).grid(row=0, column=1, padx=8, sticky="we")
        ttk.Button(top, text="Browse…", command=self.browse_file).grid(row=0, column=2, sticky="e")
        top.columnconfigure(1, weight=1)

        # Sheet row
        mid = ttk.Frame(self, padding=10)
        mid.pack(fill="x")

        ttk.Label(mid, text="Sheet:").grid(row=0, column=0, sticky="w")
        self.sheet_combo = ttk.Combobox(mid, textvariable=self.sheet_name, state="readonly", width=40, values=[])
        self.sheet_combo.grid(row=0, column=1, padx=8, sticky="w")

        ttk.Button(mid, text="Load Sheet", command=self.load_sheet).grid(row=0, column=2, padx=8)
        ttk.Button(mid, text="Run Analysis", command=self.run_analysis).grid(row=0, column=3, padx=8)
        ttk.Button(mid, text="Export Results…", command=self.export_results).grid(row=0, column=4, padx=8)

        # Mapping frame
        mapf = ttk.LabelFrame(self, text="Column Mapping", padding=10)
        mapf.pack(fill="x", padx=10, pady=6)

        self.name_combo = ttk.Combobox(mapf, textvariable=self.col_name, state="readonly", width=55, values=[])
        self.num_combo  = ttk.Combobox(mapf, textvariable=self.col_number, state="readonly", width=55, values=[])
        self.size_combo = ttk.Combobox(mapf, textvariable=self.col_size, state="readonly", width=55, values=[])
        self.type_combo = ttk.Combobox(mapf, textvariable=self.col_type, state="readonly", width=55, values=[])
        self.path_combo = ttk.Combobox(mapf, textvariable=self.col_path, state="readonly", width=55, values=[])
        self.rev_combo  = ttk.Combobox(mapf, textvariable=self.col_rev, state="readonly", width=55, values=[])

        ttk.Label(mapf, text="Document Name:").grid(row=0, column=0, sticky="w")
        self.name_combo.grid(row=0, column=1, padx=8, pady=2, sticky="w")

        ttk.Label(mapf, text="Document Number:").grid(row=1, column=0, sticky="w")
        self.num_combo.grid(row=1, column=1, padx=8, pady=2, sticky="w")

        ttk.Label(mapf, text="Revision (optional):").grid(row=2, column=0, sticky="w")
        self.rev_combo.grid(row=2, column=1, padx=8, pady=2, sticky="w")

        ttk.Label(mapf, text="File Size:").grid(row=0, column=2, sticky="w")
        self.size_combo.grid(row=0, column=3, padx=8, pady=2, sticky="w")

        ttk.Label(mapf, text="File Type:").grid(row=1, column=2, sticky="w")
        self.type_combo.grid(row=1, column=3, padx=8, pady=2, sticky="w")

        ttk.Label(mapf, text="File Path (optional):").grid(row=2, column=2, sticky="w")
        self.path_combo.grid(row=2, column=3, padx=8, pady=2, sticky="w")

        # Duplicate options frame
        dupf = ttk.LabelFrame(self, text="Duplicate Detection Options (tick the fields to define duplicates)", padding=10)
        dupf.pack(fill="x", padx=10, pady=6)

        # left column of checkboxes
        ttk.Checkbutton(dupf, text="Document Name", variable=self.use_doc_name).grid(row=0, column=0, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(dupf, text="Document Number", variable=self.use_doc_number).grid(row=1, column=0, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(dupf, text="Revision", variable=self.use_revision).grid(row=2, column=0, sticky="w", padx=6, pady=2)

        # middle column
        ttk.Checkbutton(dupf, text="File Size", variable=self.use_file_size).grid(row=0, column=1, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(dupf, text="File Type / Extension", variable=self.use_file_type).grid(row=1, column=1, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(dupf, text="File Path", variable=self.use_file_path).grid(row=2, column=1, sticky="w", padx=6, pady=2)

        # derived controls
        derived = ttk.LabelFrame(dupf, text="Derived options (from File Path)", padding=8)
        derived.grid(row=0, column=2, rowspan=3, padx=10, sticky="nw")

        ttk.Checkbutton(
            derived,
            text="Derive FileName from File Path",
            variable=self.derive_filename_from_path,
            command=self._on_derive_toggle
        ).grid(row=0, column=0, sticky="w", pady=2)

        ttk.Checkbutton(
            derived,
            text="Also derive FileStem (no extension)",
            variable=self.derive_stem_from_filename,
        ).grid(row=1, column=0, sticky="w", pady=2)

        ttk.Checkbutton(
            derived,
            text="Use derived FileName in duplicate keys",
            variable=self.use_derived_filename,
        ).grid(row=2, column=0, sticky="w", pady=2)

        ttk.Checkbutton(
            derived,
            text="Use derived FileStem in duplicate keys",
            variable=self.use_derived_stem,
        ).grid(row=3, column=0, sticky="w", pady=2)

        # Presets
        presetf = ttk.Frame(self, padding=(10, 0))
        presetf.pack(fill="x")
        ttk.Label(presetf, text="Presets:").pack(side="left")
        ttk.Button(presetf, text="Strict (Name+Size+Type)", command=self.preset_strict).pack(side="left", padx=6)
        ttk.Button(presetf, text="Super strict (+Path)", command=self.preset_super_strict).pack(side="left", padx=6)
        ttk.Button(presetf, text="Revision-aware (DocNo+Rev+Type+Size)", command=self.preset_revision_aware).pack(side="left", padx=6)
        ttk.Button(presetf, text="DocNo duplicates (DocNo+Type+Size)", command=self.preset_docno).pack(side="left", padx=6)

        # Output area
        out = ttk.Frame(self, padding=10)
        out.pack(fill="both", expand=True)

        ttk.Label(out, text="Output:").pack(anchor="w")
        self.text = tk.Text(out, wrap="word")
        self.text.pack(fill="both", expand=True, side="left")

        scroll = ttk.Scrollbar(out, command=self.text.yview)
        scroll.pack(fill="y", side="right")
        self.text.configure(yscrollcommand=scroll.set)

        # Status bar
        status_bar = ttk.Frame(self, padding=(10, 6))
        status_bar.pack(fill="x")
        ttk.Label(status_bar, textvariable=self.status).pack(anchor="w")

        self._on_derive_toggle()

    def _on_derive_toggle(self):
        # if not deriving filename, disable the "use derived" toggles
        enabled = self.derive_filename_from_path.get()
        state = "normal" if enabled else "disabled"

        # ttk.Checkbutton doesn't have direct state var; set via widget state:
        # We'll find them by walking children of derived frame
        # (Simple: just enforce that these vars are false if disabled)
        if not enabled:
            self.use_derived_filename.set(False)
            self.use_derived_stem.set(False)
            self.derive_stem_from_filename.set(False)

    # ----------------------------
    # Presets
    # ----------------------------
    def preset_strict(self):
        self.use_doc_name.set(True)
        self.use_file_size.set(True)
        self.use_file_type.set(True)
        self.use_doc_number.set(False)
        self.use_revision.set(False)
        self.use_file_path.set(False)
        self.use_derived_filename.set(False)
        self.use_derived_stem.set(False)

    def preset_super_strict(self):
        self.preset_strict()
        self.use_file_path.set(True)

    def preset_revision_aware(self):
        self.use_doc_name.set(False)
        self.use_doc_number.set(True)
        self.use_revision.set(True)
        self.use_file_type.set(True)
        self.use_file_size.set(True)
        self.use_file_path.set(False)
        self.use_derived_filename.set(False)
        self.use_derived_stem.set(False)

    def preset_docno(self):
        self.use_doc_name.set(False)
        self.use_doc_number.set(True)
        self.use_revision.set(False)
        self.use_file_type.set(True)
        self.use_file_size.set(True)
        self.use_file_path.set(False)
        self.use_derived_filename.set(False)
        self.use_derived_stem.set(False)

    # ----------------------------
    # File + sheet loading
    # ----------------------------
    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path:
            return
        self.file_path.set(path)
        self.status.set("Loading sheet list…")
        self.populate_sheets(path)

    def populate_sheets(self, path):
        try:
            xl = pd.ExcelFile(path)
            sheets = xl.sheet_names
            self.sheet_combo["values"] = sheets
            if sheets:
                self.sheet_name.set(sheets[0])
            self.status.set(f"Found {len(sheets)} sheet(s). Select one and click Load Sheet.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not read Excel file.\n\n{e}")
            self.status.set("Failed to read Excel file.")

    def load_sheet(self):
        path = self.file_path.get().strip()
        sheet = self.sheet_name.get().strip()
        if not path:
            messagebox.showwarning("Missing file", "Choose an Excel file first.")
            return
        if not sheet:
            messagebox.showwarning("Missing sheet", "Choose a sheet first.")
            return

        try:
            self.df = pd.read_excel(path, sheet_name=sheet)
            self.columns = list(self.df.columns)

            for combo in (self.name_combo, self.num_combo, self.rev_combo, self.size_combo, self.type_combo, self.path_combo):
                combo["values"] = self.columns

            self._auto_map_columns()
            self.status.set(f"Loaded {len(self.df)} rows from '{sheet}'. Now run analysis.")
            self.text.delete("1.0", "end")
            self.text.insert("end", f"Loaded sheet '{sheet}'. Columns detected:\n- " + "\n- ".join(map(str, self.columns)) + "\n")
        except Exception as e:
            messagebox.showerror("Error", f"Could not load sheet.\n\n{e}")
            self.status.set("Failed to load sheet.")

    def _auto_map_columns(self):
        def pick(preferred, fallbacks):
            cols_lower = {str(c).lower(): c for c in self.columns}
            if preferred.lower() in cols_lower:
                return cols_lower[preferred.lower()]
            for fb in fallbacks:
                if fb.lower() in cols_lower:
                    return cols_lower[fb.lower()]
            return ""

        self.col_name.set(pick("Document Name", ["Name", "Document_Name", "Title", "DocumentName"]))
        self.col_number.set(pick("Document Number", ["Doc Number", "Document_Number", "Number", "DocumentNo"]))
        self.col_rev.set(pick("Revision", ["Rev", "Document Revision", "Document_Revision", "Revision Number"]))
        self.col_size.set(pick("File Size", ["Size", "File_Size", "Bytes", "FileSize"]))
        self.col_type.set(pick("File Type", ["Type", "Extension", "File Extension", "FileType"]))
        self.col_path.set(pick("File Path", ["Path", "Full Path", "FullPath", "Rendition Path", "RenditionPath"]))

    # ----------------------------
    # Analysis
    # ----------------------------
    def _build_dupe_cols(self):
        """
        Build the list of columns used to detect duplicates.
        Includes derived cols if selected.
        """
        if self.df is None:
            return []

        dupe_cols = []

        # mapped col names
        name = self.col_name.get().strip()
        num = self.col_number.get().strip()
        rev = self.col_rev.get().strip()
        size = self.col_size.get().strip()
        ftyp = self.col_type.get().strip()
        path = self.col_path.get().strip()

        # regular cols
        if self.use_doc_name.get():
            if not name:
                raise ValueError("Duplicate rule includes Document Name, but it's not mapped.")
            dupe_cols.append(name)

        if self.use_doc_number.get():
            if not num:
                raise ValueError("Duplicate rule includes Document Number, but it's not mapped.")
            dupe_cols.append(num)

        if self.use_revision.get():
            if not rev:
                raise ValueError("Duplicate rule includes Revision, but it's not mapped.")
            dupe_cols.append(rev)

        if self.use_file_size.get():
            if not size:
                raise ValueError("Duplicate rule includes File Size, but it's not mapped.")
            dupe_cols.append(size)

        if self.use_file_type.get():
            if not ftyp:
                raise ValueError("Duplicate rule includes File Type, but it's not mapped.")
            dupe_cols.append(ftyp)

        if self.use_file_path.get():
            if not path:
                raise ValueError("Duplicate rule includes File Path, but it's not mapped.")
            dupe_cols.append(path)

        # derived cols (must have path mapped)
        if self.derive_filename_from_path.get():
            if not path:
                raise ValueError("To derive FileName/FileStem, you must map File Path.")
            if self.use_derived_filename.get():
                dupe_cols.append("__FileName")
            if self.use_derived_stem.get():
                dupe_cols.append("__FileStem")

        # unique + preserve order
        seen = set()
        ordered = []
        for c in dupe_cols:
            if c not in seen:
                ordered.append(c)
                seen.add(c)

        return ordered

    def run_analysis(self):
        if self.df is None:
            messagebox.showwarning("Not loaded", "Load a sheet first.")
            return

        # required mappings (baseline analysis needs these)
        name = self.col_name.get().strip()
        num  = self.col_number.get().strip()
        size = self.col_size.get().strip()
        ftyp = self.col_type.get().strip()
        path = self.col_path.get().strip() or None
        rev  = self.col_rev.get().strip() or None

        for required_label, required_value in [
            ("Document Name", name),
            ("Document Number", num),
            ("File Size", size),
            ("File Type", ftyp),
        ]:
            if not required_value:
                messagebox.showwarning("Missing mapping", f"Please map: {required_label}")
                return

        for col in [name, num, size, ftyp]:
            if col not in self.df.columns:
                messagebox.showerror("Invalid mapping", f"Column not found in sheet: {col}")
                return
        if path and path not in self.df.columns:
            messagebox.showerror("Invalid mapping", f"File Path column not found in sheet: {path}")
            return
        if rev and rev not in self.df.columns:
            messagebox.showerror("Invalid mapping", f"Revision column not found in sheet: {rev}")
            return

        try:
            dupe_cols = self._build_dupe_cols()
            if not dupe_cols:
                messagebox.showwarning("Duplicate rules", "Tick at least one field for duplicate detection.")
                return

            text, no_size, no_type, dupes, all_dupes, stacks = analyze_document_data(
                self.df,
                col_name=name,
                col_number=num,
                col_size=size,
                col_type=ftyp,
                col_path=path,
                col_rev=rev,
                dupe_cols=dupe_cols,
                derive_filename_from_path=self.derive_filename_from_path.get(),
                derive_stem_from_filename=self.derive_stem_from_filename.get(),
            )

            self.out_no_size = no_size
            self.out_no_type = no_type
            self.out_duplicates = dupes
            self.out_all_dupes = all_dupes
            self.out_revision_stacks = stacks

            self.text.delete("1.0", "end")
            self.text.insert("end", text)
            self.status.set("Analysis complete.")
        except Exception as e:
            messagebox.showerror("Error", f"Analysis failed.\n\n{e}")
            self.status.set("Analysis failed.")

    # ----------------------------
    # Export
    # ----------------------------
    def export_results(self):
        if self.out_no_size is None and self.out_no_type is None and self.out_duplicates is None:
            messagebox.showinfo("Nothing to export", "Run the analysis first.")
            return

        default_name = "document_analysis_results.xlsx"
        start_dir = os.path.dirname(self.file_path.get()) if self.file_path.get() else os.getcwd()

        save_path = filedialog.asksaveasfilename(
            title="Save results workbook",
            initialdir=start_dir,
            initialfile=default_name,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not save_path:
            return

        try:
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                if self.out_no_size is not None:
                    self.out_no_size.to_excel(writer, index=False, sheet_name="No File Size")
                if self.out_no_type is not None:
                    self.out_no_type.to_excel(writer, index=False, sheet_name="No File Type")
                if self.out_duplicates is not None:
                    self.out_duplicates.to_excel(writer, index=False, sheet_name="Duplicates (Copies)")
                if self.out_all_dupes is not None and not self.out_all_dupes.empty:
                    self.out_all_dupes.to_excel(writer, index=False, sheet_name="Duplicates (Full Set)")
                if self.out_revision_stacks is not None and not self.out_revision_stacks.empty:
                    self.out_revision_stacks.rename("Count").to_frame().to_excel(
                        writer, sheet_name="Revision Stacks"
                    )

            messagebox.showinfo("Saved", f"Exported results to:\n{save_path}")
            self.status.set("Export complete.")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed.\n\n{e}")
            self.status.set("Export failed.")


if __name__ == "__main__":
    App().mainloop()
