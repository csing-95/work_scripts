import os
import re
import threading
import queue
from datetime import datetime

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# -----------------------------
# General helpers
# -----------------------------
EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")
ALL_INPUT_EXTS = (".xlsx", ".xlsm", ".xls", ".csv")


def norm(s: str) -> str:
    return (s or "").strip().lower()


def is_spreadsheet(filename: str) -> bool:
    if filename.startswith("~$"):
        return False
    return filename.lower().endswith(ALL_INPUT_EXTS)


def is_excel_file(filename: str) -> bool:
    if filename.startswith("~$"):
        return False
    return filename.lower().endswith(EXCEL_EXTS)


def safe_filename(name: str) -> str:
    """Make a safe Windows filename."""
    name = (name or "").strip()
    if not name:
        return "output"
    bad = r'[<>:"/\\|?*\x00-\x1F]'
    name = re.sub(bad, "_", name)
    name = name.rstrip(". ")
    return name[:150]


def safe_sheet_name(name: str) -> str:
    """Excel sheet names max 31 chars, can't contain : \ / ? * [ ]"""
    name = name or "Sheet"
    bad = r'[:\\/?*\[\]]'
    name = re.sub(bad, "_", name)
    return name[:31]


def list_files(folder: str, excel_only: bool = False):
    exts = EXCEL_EXTS if excel_only else ALL_INPUT_EXTS
    out = []
    for name in os.listdir(folder):
        if name.lower().endswith(exts) and not name.startswith("~$"):
            out.append(os.path.join(folder, name))
    return sorted(out)


# -----------------------------
# Matching logic (WHOLE FOLDER SEGMENTS)
# -----------------------------
def compile_keyword_patterns(keywords: list[str], allow_spaces_for_underscores: bool = True) -> dict[str, re.Pattern]:
    """
    Folder-segment matching:
      - keyword must match a whole folder segment, e.g. \\VECTOR\\ or /VECTOR/
      - supports nested keywords with separators, e.g. "A/B" means ".../A/B/..."
      - case-insensitive
      - optional underscore~space equivalence within a segment
    """
    patterns = {}
    for kw in keywords:
        kw_clean = kw.strip()
        if not kw_clean:
            continue

        # split keyword on / or \ so user can paste either
        parts = re.split(r"[\\/]+", kw_clean)
        parts = [p for p in parts if p.strip()]

        part_pats = []
        for p in parts:
            esc = re.escape(p.strip())
            if allow_spaces_for_underscores:
                # treat underscores as "_" OR whitespace within the segment
                esc = esc.replace(r"\_", r"[_\s]+")
            part_pats.append(esc)

        sep = r"[\\/]+"
        inner = sep.join(part_pats)
        full = rf"(^|{sep}){inner}({sep}|$)"
        patterns[kw_clean] = re.compile(full, re.IGNORECASE)

    return patterns


def assign_first_match(path_value: str, keywords: list[str], patterns: dict[str, re.Pattern]) -> str | None:
    """Return the first matching keyword (whole-folder-segment match) or None."""
    if not isinstance(path_value, str):
        return None
    for kw in keywords:
        pat = patterns.get(kw)
        if pat and pat.search(path_value):
            return kw
    return None


# -----------------------------
# Column detection (universal)
# -----------------------------
DEFAULT_PATH_COL_CANDIDATES = [
    "Original File Path",
    "Original_File_Path",
    "OriginalFilePath",
    "Rendition Path",
    "Rendition_Path",
    "Source Path",
    "Source_Path",
    "File Path",
    "FilePath",
    "Path",
]

LIKELY_PATH_TOKENS = [
    "file path", "filepath", " path", "path ", "rendition", "source path", "original", "full path"
]


def looks_like_path_col(col_name: str) -> bool:
    c = norm(col_name)
    return any(tok in c for tok in LIKELY_PATH_TOKENS)


def detect_path_column(columns) -> str | None:
    cols = list(columns)

    for c in DEFAULT_PATH_COL_CANDIDATES:
        if c in cols:
            return c

    lowered = {c: str(c).lower() for c in cols}
    best = None
    best_score = -1
    for c, lc in lowered.items():
        score = 0
        for tok in LIKELY_PATH_TOKENS:
            if tok.strip() and tok.strip() in lc:
                score += 1
        if score > best_score:
            best_score = score
            best = c

    return best if best_score > 0 else None


def get_column_candidates_from_samples(folder: str, sample_limit: int = 5):
    files = list_files(folder, excel_only=False)
    samples = files[:sample_limit]
    scores = {}

    for fp in samples:
        try:
            if fp.lower().endswith(".csv"):
                df = pd.read_csv(fp, nrows=5, dtype=str, low_memory=False)
                cols = list(df.columns)
            else:
                xl = pd.ExcelFile(fp)
                sheets = xl.sheet_names[:3]
                cols = []
                for sh in sheets:
                    df = pd.read_excel(fp, sheet_name=sh, nrows=5, dtype=str)
                    cols.extend(list(df.columns))

            for c in cols:
                scores[c] = scores.get(c, 0) + (3 if looks_like_path_col(c) else 1)
        except Exception:
            continue

    ranked = sorted(scores.items(), key=lambda x: x[1], reverse=True)
    return [c for c, _ in ranked][:60]


# -----------------------------
# Tab 1: Merge + Split
# -----------------------------
def pick_best_sheet_for_merge(xls: pd.ExcelFile) -> str:
    """
    Prefer 'Documents' if it has a path column; else first sheet containing a path column; else sheet 0.
    """
    if "Documents" in xls.sheet_names:
        try:
            test = pd.read_excel(xls, sheet_name="Documents", nrows=10, dtype=str)
            if detect_path_column(test.columns):
                return "Documents"
        except Exception:
            pass

    for s in xls.sheet_names:
        try:
            test = pd.read_excel(xls, sheet_name=s, nrows=10, dtype=str)
            if detect_path_column(test.columns):
                return s
        except Exception:
            continue

    return xls.sheet_names[0]


def read_one_excel_for_merge(file_path: str, chosen_sheet_mode: str) -> pd.DataFrame:
    """
    chosen_sheet_mode:
      - "Auto (prefer Documents)"
      - "Documents"
      - "First sheet"
      - or a literal sheet name (handled by caller)
    """
    xls = pd.ExcelFile(file_path)
    if chosen_sheet_mode == "Documents":
        sheet = "Documents"
        if sheet not in xls.sheet_names:
            raise ValueError(f"'Documents' sheet not found in {os.path.basename(file_path)}")
    elif chosen_sheet_mode == "First sheet":
        sheet = xls.sheet_names[0]
    elif chosen_sheet_mode == "Auto (prefer Documents)":
        sheet = pick_best_sheet_for_merge(xls)
    else:
        sheet = chosen_sheet_mode
        if sheet not in xls.sheet_names:
            raise ValueError(f"Sheet '{sheet}' not found in {os.path.basename(file_path)}")

    # IMPORTANT: keep types as-is (no dtype=str) so values aren't coerced
    df = pd.read_excel(xls, sheet_name=sheet)

    path_col = detect_path_column(df.columns)
    if not path_col:
        raise ValueError(f"No path-like column found in: {os.path.basename(file_path)} (sheet: {sheet})")

    if path_col != "File Path":
        df = df.rename(columns={path_col: "File Path"})

    df["Source File"] = os.path.basename(file_path)
    df["Source Sheet"] = sheet
    return df


def merge_and_split(input_folder: str, output_folder: str, output_prefix: str,
                    keywords: list[str], allow_spaces_for_underscores: bool,
                    sheet_mode: str, log_fn):
    files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if is_excel_file(f)]
    if not files:
        raise ValueError("No Excel files (.xlsx/.xlsm/.xls) found in that folder.")

    log_fn(f"Found {len(files)} Excel files.")

    all_dfs = []
    for i, fp in enumerate(files, start=1):
        log_fn(f"[{i}/{len(files)}] Reading: {os.path.basename(fp)}")
        df = read_one_excel_for_merge(fp, sheet_mode)
        all_dfs.append(df)

    merged = pd.concat(all_dfs, ignore_index=True)

    if "File Path" not in merged.columns:
        raise ValueError("Merged data does not contain a 'File Path' column (unexpected).")

    merged["Filepath Dupe?"] = merged["File Path"].duplicated(keep=False).map({True: "YES", False: "NO"})

    keywords = [k.strip() for k in keywords if k.strip()]
    patterns = compile_keyword_patterns(keywords, allow_spaces_for_underscores=allow_spaces_for_underscores)

    def cat_func(v):
        m = assign_first_match(v, keywords, patterns)
        return m if m else "OTHER"

    merged["Category"] = merged["File Path"].apply(cat_func)

    prefix = safe_filename((output_prefix or "").replace(".xlsx", ""))
    created = 0

    for cat in keywords + ["OTHER"]:
        part = merged[merged["Category"] == cat].copy()
        if part.empty:
            log_fn(f"Skipping {cat} (no rows).")
            continue

        out_name = f"{prefix}_{cat}.xlsx" if prefix else f"{cat}.xlsx"
        out_name = safe_filename(out_name)
        out_path = os.path.join(output_folder, out_name)

        log_fn(f"Writing: {out_name} ({len(part):,} rows)")
        part.to_excel(out_path, index=False)
        created += 1

    if created == 0:
        raise ValueError("No output files were created (no matches found).")

    log_fn(f"Done ✅ Created {created} file(s) in: {output_folder}")


# -----------------------------
# Tab 2: Folder Finder Report
# -----------------------------
def read_file_paths(file_path: str, sheet_choice: str, path_col: str):
    """
    Returns list of tuples: (batch_filename, sheet_name, path_value)
    """
    batch = os.path.basename(file_path)
    rows = []

    if file_path.lower().endswith(".csv"):
        df = pd.read_csv(file_path, dtype=str, low_memory=False)
        if path_col not in df.columns:
            raise KeyError(f"Column '{path_col}' not found in {batch}")
        for v in df[path_col].dropna().astype(str):
            rows.append((batch, "CSV", v))
        return rows

    xl = pd.ExcelFile(file_path)
    sheets = xl.sheet_names

    if sheet_choice == "First sheet":
        target_sheets = [sheets[0]]
    elif sheet_choice == "All sheets":
        target_sheets = sheets
    else:
        if sheet_choice not in sheets:
            raise ValueError(f"Sheet '{sheet_choice}' not found in {batch}")
        target_sheets = [sheet_choice]

    for sh in target_sheets:
        df = pd.read_excel(file_path, sheet_name=sh, dtype=str)
        if path_col not in df.columns:
            continue
        for v in df[path_col].dropna().astype(str):
            rows.append((batch, sh, v))
    return rows


def build_report(files: list[str], keywords: list[str], sheet_choice: str, path_col: str,
                 out_path: str, include_details: bool,
                 allow_spaces_for_underscores: bool, stop_flag: threading.Event,
                 msg_queue: queue.Queue):
    keywords = [k.strip() for k in keywords if k.strip()]
    patterns = compile_keyword_patterns(keywords, allow_spaces_for_underscores=allow_spaces_for_underscores)

    counts = {kw: {} for kw in keywords}     # counts[keyword][batch] = count
    total_counts = {kw: 0 for kw in keywords}
    details_rows = []

    for i, fp in enumerate(files, start=1):
        if stop_flag.is_set():
            msg_queue.put(("status", "Stopped."))
            msg_queue.put(("done", None))
            return

        batch = os.path.basename(fp)
        msg_queue.put(("log", f"Reading: {batch}"))

        try:
            path_rows = read_file_paths(fp, sheet_choice, path_col)
        except Exception as e:
            msg_queue.put(("log", f"  ⚠ Skipped (error): {e}"))
            msg_queue.put(("progress", i))
            continue

        for (b, sh, pval) in path_rows:
            m = assign_first_match(pval, keywords, patterns)
            if not m:
                continue

            total_counts[m] += 1
            counts[m][b] = counts[m].get(b, 0) + 1

            if include_details:
                details_rows.append({
                    "Batch Spreadsheet": b,
                    "Sheet": sh,
                    "Matched Keyword": m,
                    "File Path": pval
                })

        msg_queue.put(("progress", i))

    summary_rows = []
    for kw in keywords:
        batches = sorted(counts[kw].keys())
        summary_rows.append({
            "Keyword / Folder": kw,
            "Total File Count": total_counts[kw],
            "Batch Spreadsheet Count": len(batches),
            "Batch Spreadsheets": ", ".join(batches)
        })

    summary_df = pd.DataFrame(summary_rows)

    all_batches = sorted({os.path.basename(f) for f in files})
    matrix = []
    for kw in keywords:
        row = {"Keyword / Folder": kw}
        for b in all_batches:
            row[b] = counts[kw].get(b, 0)
        matrix.append(row)
    matrix_df = pd.DataFrame(matrix)

    details_df = pd.DataFrame(details_rows) if include_details else pd.DataFrame()

    msg_queue.put(("status", "Writing Excel…"))
    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            summary_df.to_excel(writer, index=False, sheet_name="Summary")
            matrix_df.to_excel(writer, index=False, sheet_name="BatchCounts")
            if include_details:
                details_df.to_excel(writer, index=False, sheet_name="Matches")

            readme = pd.DataFrame([{
                "Generated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Input folder": os.path.dirname(files[0]) if files else "",
                "Sheet option": sheet_choice,
                "Path column": path_col,
                "Keywords": len(keywords),
                "Batch files scanned": len(files),
                "Details included": include_details,
                "Underscore~Space matching": allow_spaces_for_underscores,
                "Matching mode": "Whole folder segments",
            }])
            readme.to_excel(writer, index=False, sheet_name="ReadMe")
    except Exception as e:
        msg_queue.put(("log", f"❌ Failed to write output: {e}"))
        msg_queue.put(("status", "Failed."))
        msg_queue.put(("done", None))
        return

    msg_queue.put(("log", f"✅ Done! Output saved to: {out_path}"))
    msg_queue.put(("status", "Complete."))
    msg_queue.put(("done", None))


# -----------------------------
# GUI (combined)
# -----------------------------
class CombinedApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Universal Batch Tool (Merge+Split + Folder Finder Report)")
        self.geometry("1040x760")
        self.minsize(940, 680)

        self.msg_queue = queue.Queue()
        self.stop_flag = threading.Event()
        self.worker_thread = None

        self._build_ui()
        self._poll_queue()

    def _build_ui(self):
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        self.tab_merge = ttk.Frame(nb)
        self.tab_report = ttk.Frame(nb)

        nb.add(self.tab_merge, text="Merge + Split")
        nb.add(self.tab_report, text="Folder Finder Report")

        self._build_merge_tab(self.tab_merge)
        self._build_report_tab(self.tab_report)

    # -------- Merge+Split tab --------
    def _build_merge_tab(self, parent):
        pad = {"padx": 10, "pady": 6}

        self.m_in_var = tk.StringVar()
        self.m_out_var = tk.StringVar()
        self.m_prefix_var = tk.StringVar(value="SPLIT")
        self.m_allow_us_var = tk.BooleanVar(value=True)
        self.m_sheet_var = tk.StringVar(value="Auto (prefer Documents)")

        frm1 = ttk.LabelFrame(parent, text="1) Input folder (Excel files)")
        frm1.pack(fill="x", **pad)

        row = ttk.Frame(frm1)
        row.pack(fill="x", padx=10, pady=8)
        ttk.Label(row, text="Folder:").pack(side="left")
        ttk.Entry(row, textvariable=self.m_in_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Browse…", command=self._browse_merge_input).pack(side="left")

        frm2 = ttk.LabelFrame(parent, text="2) Output")
        frm2.pack(fill="x", **pad)

        grid = ttk.Frame(frm2)
        grid.pack(fill="x", padx=10, pady=8)
        ttk.Label(grid, text="Output folder:").grid(row=0, column=0, sticky="w")
        ttk.Entry(grid, textvariable=self.m_out_var).grid(row=0, column=1, sticky="ew", padx=8)
        ttk.Button(grid, text="Browse…", command=self._browse_merge_output).grid(row=0, column=2)

        ttk.Label(grid, text="Output file prefix:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(grid, textvariable=self.m_prefix_var).grid(row=1, column=1, sticky="ew", padx=8, pady=(8, 0))

        ttk.Label(grid, text="Sheet mode:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.m_sheet_combo = ttk.Combobox(
            grid,
            textvariable=self.m_sheet_var,
            values=["Auto (prefer Documents)", "Documents", "First sheet"],
            state="readonly",
            width=28
        )
        self.m_sheet_combo.grid(row=2, column=1, sticky="w", padx=8, pady=(8, 0))

        ttk.Checkbutton(
            grid,
            text="Treat underscores like spaces when matching (AAA_GRM_VAULT matches 'AAA GRM VAULT')",
            variable=self.m_allow_us_var
        ).grid(row=3, column=1, sticky="w", padx=8, pady=(8, 0))

        grid.columnconfigure(1, weight=1)

        frm3 = ttk.LabelFrame(parent, text="3) Categories / folders / keywords to split by (one per line)")
        frm3.pack(fill="both", expand=True, **pad)

        self.m_keywords = tk.Text(frm3, height=10)
        self.m_keywords.pack(fill="both", expand=True, padx=10, pady=8)
        self.m_keywords.insert("1.0", "VECTOR\nOFFSHORE\nMERIDIAN MIGRATION UPLOADS STILL NEEDED\n")

        frm4 = ttk.Frame(parent)
        frm4.pack(fill="x", **pad)

        self.m_run_btn = ttk.Button(frm4, text="Run merge + split", command=self._run_merge_split)
        self.m_run_btn.pack(side="left")

        self.m_prog = ttk.Progressbar(frm4, mode="indeterminate")
        self.m_prog.pack(side="left", fill="x", expand=True, padx=10)

        frm5 = ttk.LabelFrame(parent, text="Log")
        frm5.pack(fill="both", expand=True, **pad)

        self.m_log = tk.Text(frm5, height=10, wrap="word")
        self.m_log.pack(fill="both", expand=True, padx=10, pady=8)

        self._m_log("Matching mode: whole folder segments (e.g. \\VECTOR\\ or /VECTOR/).\n")

    def _m_log(self, msg: str):
        self.m_log.insert("end", msg + "\n")
        self.m_log.see("end")
        self.update_idletasks()

    def _browse_merge_input(self):
        folder = filedialog.askdirectory(title="Select folder containing Excel files")
        if folder:
            self.m_in_var.set(folder)
            if not self.m_out_var.get():
                self.m_out_var.set(folder)

    def _browse_merge_output(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.m_out_var.set(folder)

    def _run_merge_split(self):
        in_dir = self.m_in_var.get().strip()
        out_dir = self.m_out_var.get().strip()
        prefix = self.m_prefix_var.get().strip()
        sheet_mode = self.m_sheet_var.get().strip()

        keywords = [l.strip() for l in self.m_keywords.get("1.0", "end").splitlines() if l.strip()]

        if not in_dir or not os.path.isdir(in_dir):
            messagebox.showerror("Missing input", "Please choose a valid input folder.")
            return
        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showerror("Missing output", "Please choose a valid output folder.")
            return
        if not keywords:
            messagebox.showerror("Missing keywords", "Please enter at least one keyword/category (one per line).")
            return

        self.m_run_btn.config(state="disabled")
        self.m_prog.start(10)
        self._m_log("Starting…")

        def worker():
            try:
                merge_and_split(
                    in_dir, out_dir, prefix,
                    keywords,
                    allow_spaces_for_underscores=self.m_allow_us_var.get(),
                    sheet_mode=sheet_mode,
                    log_fn=self._m_log
                )
                messagebox.showinfo("Done", f"Saved outputs in:\n{out_dir}")
            except Exception as e:
                self._m_log(f"ERROR: {e}")
                messagebox.showerror("Error", str(e))
            finally:
                self.m_prog.stop()
                self.m_run_btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()

    # -------- Report tab --------
    def _build_report_tab(self, parent):
        pad = {"padx": 10, "pady": 6}

        self.r_folder_var = tk.StringVar()
        self.r_output_var = tk.StringVar()
        self.r_sheet_var = tk.StringVar(value="First sheet")
        self.r_pathcol_var = tk.StringVar()
        self.r_include_details_var = tk.BooleanVar(value=True)
        self.r_allow_us_var = tk.BooleanVar(value=True)

        frm1 = ttk.LabelFrame(parent, text="1) Input folder (batch spreadsheets: xlsx/xlsm/xls/csv)")
        frm1.pack(fill="x", **pad)

        row = ttk.Frame(frm1)
        row.pack(fill="x", padx=10, pady=8)
        ttk.Label(row, text="Folder:").pack(side="left")
        ttk.Entry(row, textvariable=self.r_folder_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Browse…", command=self._browse_report_folder).pack(side="left")

        frm2 = ttk.LabelFrame(parent, text="2) Options")
        frm2.pack(fill="x", **pad)

        opts = ttk.Frame(frm2)
        opts.pack(fill="x", padx=10, pady=8)

        ttk.Label(opts, text="Sheet:").grid(row=0, column=0, sticky="w")
        self.r_sheet_combo = ttk.Combobox(
            opts, textvariable=self.r_sheet_var,
            values=["First sheet", "All sheets"], state="readonly", width=18
        )
        self.r_sheet_combo.grid(row=0, column=1, sticky="w", padx=8)

        ttk.Label(opts, text="Path column:").grid(row=0, column=2, sticky="w", padx=(20, 0))
        self.r_col_combo = ttk.Combobox(opts, textvariable=self.r_pathcol_var, values=[], state="readonly", width=42)
        self.r_col_combo.grid(row=0, column=3, sticky="w", padx=8)

        ttk.Button(opts, text="Detect columns", command=self._detect_report_columns).grid(row=0, column=4, sticky="w", padx=8)

        ttk.Checkbutton(opts, text="Include detailed Matches sheet", variable=self.r_include_details_var).grid(
            row=1, column=1, columnspan=3, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(
            opts,
            text="Treat underscores like spaces when matching",
            variable=self.r_allow_us_var
        ).grid(row=2, column=1, columnspan=3, sticky="w", pady=(6, 0))

        frm3 = ttk.LabelFrame(parent, text="3) Folders / keywords to look out for (one per line)")
        frm3.pack(fill="both", expand=True, **pad)

        self.r_keywords = tk.Text(frm3, height=10)
        self.r_keywords.pack(fill="both", expand=True, padx=10, pady=8)
        self.r_keywords.insert("1.0", "AAA_GRM_VAULT\nVECTOR\nOFFSHORE\n")

        frm4 = ttk.LabelFrame(parent, text="4) Output")
        frm4.pack(fill="x", **pad)

        out = ttk.Frame(frm4)
        out.pack(fill="x", padx=10, pady=8)
        ttk.Label(out, text="Save as:").pack(side="left")
        ttk.Entry(out, textvariable=self.r_output_var).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(out, text="Browse…", command=self._browse_report_output).pack(side="left")

        frm5 = ttk.Frame(parent)
        frm5.pack(fill="x", padx=10, pady=(0, 10))

        self.r_run_btn = ttk.Button(frm5, text="Run report", command=self._run_report)
        self.r_run_btn.pack(side="left")

        self.r_stop_btn = ttk.Button(frm5, text="Stop", command=self._stop_report, state="disabled")
        self.r_stop_btn.pack(side="left", padx=8)

        self.r_progress = ttk.Progressbar(frm5, mode="determinate")
        self.r_progress.pack(side="left", fill="x", expand=True, padx=10)

        self.r_status_var = tk.StringVar(value="Ready.")
        ttk.Label(frm5, textvariable=self.r_status_var).pack(side="left")

        frm6 = ttk.LabelFrame(parent, text="Log")
        frm6.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.r_log = tk.Text(frm6, height=10)
        self.r_log.pack(fill="both", expand=True, padx=10, pady=8)

    def _r_log(self, msg: str):
        self.r_log.insert("end", msg + "\n")
        self.r_log.see("end")

    def _browse_report_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.r_folder_var.set(folder)
            self._r_log(f"Selected folder: {folder}")
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            self.r_output_var.set(os.path.join(folder, f"folder_finder_report_{ts}.xlsx"))

    def _browse_report_output(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if path:
            self.r_output_var.set(path)

    def _detect_report_columns(self):
        folder = self.r_folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Missing folder", "Pick a valid folder first.")
            return

        files = list_files(folder, excel_only=False)
        if not files:
            messagebox.showerror("No files", "No .xlsx/.xlsm/.xls/.csv files found in that folder.")
            return

        col_list = get_column_candidates_from_samples(folder)
        if not col_list:
            messagebox.showwarning("No columns found", "Could not read columns from sample files.")
            return

        self.r_col_combo["values"] = col_list

        best = None
        for c in col_list:
            if looks_like_path_col(c):
                best = c
                break
        if not best:
            best = col_list[0]
        self.r_pathcol_var.set(best)

        self._r_log("Detected columns (top picks loaded).")
        self._r_log(f"Auto-selected path column: {best}")

        first_xlsx = next((f for f in files if not f.lower().endswith(".csv")), None)
        if first_xlsx:
            try:
                xl = pd.ExcelFile(first_xlsx)
                self.r_sheet_combo["values"] = ["First sheet", "All sheets"] + xl.sheet_names
            except Exception:
                pass

    def _run_report(self):
        folder = self.r_folder_var.get().strip()
        out = self.r_output_var.get().strip()
        path_col = self.r_pathcol_var.get().strip()
        sheet_choice = self.r_sheet_var.get().strip()
        keywords = [l.strip() for l in self.r_keywords.get("1.0", "end").splitlines() if l.strip()]

        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Missing folder", "Pick a valid input folder.")
            return
        if not out:
            messagebox.showerror("Missing output", "Pick an output .xlsx path.")
            return
        if not keywords:
            messagebox.showerror("Missing keywords", "Enter at least one keyword/folder name.")
            return
        if not path_col:
            messagebox.showerror("Missing column", "Detect/select the column that contains file paths.")
            return

        files = list_files(folder, excel_only=False)
        if not files:
            messagebox.showerror("No spreadsheets found", "No spreadsheets found in that folder.")
            return

        self.stop_flag.clear()
        self.r_run_btn.configure(state="disabled")
        self.r_stop_btn.configure(state="normal")
        self.r_progress.configure(maximum=len(files), value=0)
        self.r_status_var.set("Running…")
        self._r_log(f"Starting run on {len(files)} files…")

        args = (
            files, keywords, sheet_choice, path_col, out,
            self.r_include_details_var.get(),
            self.r_allow_us_var.get(),
            self.stop_flag, self.msg_queue
        )

        self.worker_thread = threading.Thread(target=build_report, args=args, daemon=True)
        self.worker_thread.start()

    def _stop_report(self):
        self.stop_flag.set()
        self._r_log("Stop requested…")

    def _poll_queue(self):
        try:
            while True:
                msg_type, payload = self.msg_queue.get_nowait()
                if msg_type == "log":
                    self._r_log(payload)
                elif msg_type == "progress":
                    self.r_progress.configure(value=payload)
                    self.r_status_var.set(f"Running… {payload}/{int(self.r_progress['maximum'])}")
                elif msg_type == "status":
                    self.r_status_var.set(payload)
                elif msg_type == "done":
                    self.r_run_btn.configure(state="normal")
                    self.r_stop_btn.configure(state="disabled")
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)


if __name__ == "__main__":
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass

    CombinedApp().mainloop()