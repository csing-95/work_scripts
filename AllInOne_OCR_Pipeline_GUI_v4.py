
import os
import shutil
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

# =========================================================
# Shared helpers
# =========================================================

def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)

def normalize_path(p: str) -> str:
    # Windows-friendly, tolerant of /
    return str(p).strip().replace("/", "\\")

def strip_drive_and_leading_slashes(p: str) -> str:
    """
    Turn full paths into relative-ish paths by stripping:
      - drive letters like C:\ or E:\
      - leading slashes
    """
    p = normalize_path(p)
    if len(p) >= 2 and p[1] == ":":
        p = p[2:]
    return p.lstrip("\\")

def parse_extensions(ext_text: str) -> tuple[str, ...]:
    """
    "docx,xlsx,dwg" or ".docx, .xlsx" -> (".docx",".xlsx",".dwg")
    """
    parts = [p.strip().lower() for p in str(ext_text).split(",") if p.strip()]
    exts = []
    for p in parts:
        if not p.startswith("."):
            p = "." + p
        exts.append(p)
    # unique, preserve order
    return tuple(dict.fromkeys(exts))

def build_flat_index(flat_folder: str, allowed_exts: tuple[str, ...] | None = None):
    """
    Index files in a (typically flat) folder.

    Returns:
        unique_lookup: filename_lower -> full path (only for unique filenames)
        duplicates: dict filename_lower -> list of full paths (for duplicates)
        all_files: dict filename_lower -> list of full paths (for reference)
    """
    all_files: dict[str, list[str]] = {}
    for name in os.listdir(flat_folder):
        full = os.path.join(flat_folder, name)
        if os.path.isfile(full):
            if allowed_exts is not None:
                if not name.lower().endswith(allowed_exts):
                    continue
            key = name.lower()
            all_files.setdefault(key, []).append(full)

    unique_lookup: dict[str, str] = {}
    duplicates: dict[str, list[str]] = {}

    for key, paths in all_files.items():
        if len(paths) == 1:
            unique_lookup[key] = paths[0]
        else:
            duplicates[key] = paths

    return unique_lookup, duplicates, all_files

def write_report_one_workbook(report_path: str, summary_row, missing_rows, duplicate_rows, orphan_rows):
    """
    ONE workbook with multiple sheets (Summary/Missing/Duplicates/Orphans).
    """
    if not report_path:
        return

    out_dir = os.path.dirname(report_path)
    if out_dir:
        ensure_dir(out_dir)

    df_summary = pd.DataFrame([summary_row])

    df_missing = pd.DataFrame(missing_rows) if missing_rows else pd.DataFrame(
        columns=["row_index", "source", "target_folder_rel", "expected_output_name", "reason"]
    )

    df_dupes = pd.DataFrame(duplicate_rows) if duplicate_rows else pd.DataFrame(
        columns=["flat_filename", "duplicate_count", "flat_full_path"]
    )

    df_orphans = pd.DataFrame(orphan_rows) if orphan_rows else pd.DataFrame(
        columns=["flat_filename", "flat_full_path"]
    )

    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        df_summary.to_excel(writer, index=False, sheet_name="Summary")
        df_missing.to_excel(writer, index=False, sheet_name="Missing")
        df_dupes.to_excel(writer, index=False, sheet_name="Duplicates")
        df_orphans.to_excel(writer, index=False, sheet_name="Orphans")

def write_rename_report(report_path: str, summary_row: dict, rows: list[dict]):
    """
    ONE workbook for rename/copy step: Summary + Renames
    """
    if not report_path:
        return
    out_dir = os.path.dirname(report_path)
    if out_dir:
        ensure_dir(out_dir)

    df_summary = pd.DataFrame([summary_row])
    df_rows = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["source_full_path", "new_full_path", "original_name", "new_name", "token", "action", "status", "reason"]
    )

    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        df_summary.to_excel(writer, index=False, sheet_name="Summary")
        df_rows.to_excel(writer, index=False, sheet_name="Renames")

# =========================================================
# Tab 0: Pre-OCR rename/copy (append __size OR __counter)
# =========================================================

def build_new_name(filename: str, token: str, separator="__"):
    base, ext = os.path.splitext(filename)
    return f"{base}{separator}{token}{ext}"

def pre_ocr_append_token(
    source_root: str,
    mode: str,  # "copy" or "rename"
    staging_root: str,
    filter_enabled: bool,
    include_exts: tuple[str, ...],
    separator: str,
    token_mode: str,  # "counter" or "size"
    counter_padding: int,
    report_path: str,
    log_fn=print,
):
    """
    Creates a staging set of files whose names include a unique token:
      - Counter mode: base__000001.ext (global counter in walk order)
      - Size mode:    base__922032.ext (source file size in bytes)

    Copy mode preserves original files and writes renamed copies to staging root.
    Rename mode renames in-place (riskier).
    """
    if not os.path.isdir(source_root):
        raise ValueError("Source root folder not found.")

    if mode == "copy":
        if not staging_root:
            raise ValueError("Staging root is required for Copy mode.")
        ensure_dir(staging_root)

    if token_mode not in ("counter", "size"):
        raise ValueError("Token mode must be 'counter' or 'size'.")

    scanned = 0
    processed = 0
    skipped = 0
    collisions = 0
    errors = 0
    rows: list[dict] = []

    counter = 1

    for root, _, files in os.walk(source_root):
        files.sort(key=lambda s: s.lower())
        for name in files:
            scanned += 1
            src_full = os.path.join(root, name)

            if filter_enabled and include_exts:
                if not name.lower().endswith(include_exts):
                    skipped += 1
                    continue

            try:
                if token_mode == "size":
                    token = str(os.path.getsize(src_full))
                else:
                    token = str(counter).zfill(max(1, int(counter_padding)))
                    counter += 1
            except Exception as e:
                errors += 1
                rows.append({
                    "source_full_path": src_full,
                    "new_full_path": "",
                    "original_name": name,
                    "new_name": "",
                    "token": "",
                    "action": mode.upper(),
                    "status": "ERROR",
                    "reason": f"Could not create token: {e}",
                })
                log_fn(f"❌ ERROR token: {src_full} -> {e}")
                continue

            new_name = build_new_name(name, token, separator=separator)

            if mode == "rename":
                dest_full = os.path.join(root, new_name)
            else:
                rel_dir = os.path.relpath(root, source_root)
                rel_dir = "" if rel_dir == "." else rel_dir
                dest_dir = os.path.join(staging_root, rel_dir)
                ensure_dir(dest_dir)
                dest_full = os.path.join(dest_dir, new_name)

            # Collision handling: append _1, _2...
            final_dest = dest_full
            if os.path.exists(final_dest):
                collisions += 1
                base2, ext2 = os.path.splitext(final_dest)
                n = 1
                while os.path.exists(f"{base2}_{n}{ext2}"):
                    n += 1
                final_dest = f"{base2}_{n}{ext2}"

            try:
                if mode == "rename":
                    os.rename(src_full, final_dest)
                    action = "RENAMED"
                else:
                    shutil.copy2(src_full, final_dest)
                    action = "COPIED"

                processed += 1
                rows.append({
                    "source_full_path": src_full,
                    "new_full_path": final_dest,
                    "original_name": name,
                    "new_name": os.path.basename(final_dest),
                    "token": token,
                    "action": action,
                    "status": "OK",
                    "reason": ""
                })
                log_fn(f"✅ {action}: {name} -> {os.path.basename(final_dest)}")
            except Exception as e:
                errors += 1
                rows.append({
                    "source_full_path": src_full,
                    "new_full_path": final_dest,
                    "original_name": name,
                    "new_name": os.path.basename(final_dest),
                    "token": token,
                    "action": mode.upper(),
                    "status": "ERROR",
                    "reason": str(e),
                })
                log_fn(f"❌ ERROR copy/rename: {src_full} -> {e}")

    summary = {
        "mode": mode,
        "token_mode": token_mode,
        "counter_padding": counter_padding if token_mode == "counter" else "",
        "source_root": source_root,
        "staging_root": staging_root if mode == "copy" else "",
        "separator": separator,
        "filter_enabled": filter_enabled,
        "include_exts": ",".join(include_exts) if include_exts else "",
        "files_scanned": scanned,
        "files_processed": processed,
        "skipped": skipped,
        "collisions": collisions,
        "errors": errors,
    }

    if report_path:
        write_rename_report(report_path, summary, rows)
        log_fn(f"📄 Report written: {report_path}")

    log_fn("")
    log_fn("===== Summary =====")
    log_fn(f"Scanned:     {scanned}")
    log_fn(f"Processed:   {processed}")
    log_fn(f"Skipped:     {skipped}")
    log_fn(f"Collisions:  {collisions}")
    log_fn(f"Errors:      {errors}")
    log_fn("===================")

# =========================================================
# Mode A: Rebuild using ORIGINAL folder structure (template)
# =========================================================

def recreate_from_original_tree(
    original_root: str,
    flat_output_folder: str,
    destination_root: str,
    move_files: bool = True,
    scan_all: bool = True,
    include_exts: tuple[str, ...] = (),
    output_ext: str = ".pdf",
    output_suffix: str = "",
    report_path: str = "",
    log_fn=print,
):
    """
    Walk original_root, recreate folder structure under destination_root,
    and place matching output files (typically PDFs) from flat_output_folder.

    Matching rule: base filename from original -> base + output_suffix + output_ext in flat folder.
    """
    if not os.path.isdir(original_root):
        raise ValueError("Original root folder doesn't exist.")
    if not os.path.isdir(flat_output_folder):
        raise ValueError("Flat output folder doesn't exist.")
    ensure_dir(destination_root)

    if not output_ext.startswith("."):
        output_ext = "." + output_ext
    output_ext = output_ext.lower()
    output_suffix = (output_suffix or "").strip()

    # Index flat outputs (only output_ext)
    unique_lookup, duplicates, all_files = build_flat_index(flat_output_folder, allowed_exts=(output_ext,))

    duplicate_rows = []
    for key, paths in duplicates.items():
        for p in paths:
            duplicate_rows.append({
                "flat_filename": os.path.basename(p),
                "duplicate_count": len(paths),
                "flat_full_path": p
            })

    used_unique = set()
    missing_rows = []

    def should_include_file(pathname: str) -> bool:
        n = pathname.lower()
        if n.endswith(output_ext):
            return False
        if scan_all:
            return True
        return n.endswith(include_exts)

    total_inputs = 0
    moved_or_copied = 0
    missing_not_found = 0
    skipped_duplicates = 0
    errors = 0

    for root, _, files in os.walk(original_root):
        files.sort(key=lambda s: s.lower())
        for file in files:
            if not should_include_file(file):
                continue

            total_inputs += 1
            base = os.path.splitext(file)[0]
            expected_name = f"{base}{output_suffix}{output_ext}"
            expected_key = expected_name.lower()

            source_full = os.path.join(root, file)
            rel_dir = os.path.relpath(root, original_root)
            target_folder_rel = rel_dir if rel_dir != "." else ""

            # Duplicate handling: skip if duplicates exist in flat folder
            if expected_key in duplicates:
                skipped_duplicates += 1
                missing_rows.append({
                    "row_index": None,
                    "source": source_full,
                    "target_folder_rel": target_folder_rel,
                    "expected_output_name": expected_name,
                    "reason": "Duplicate filename in flat folder (skipped)"
                })
                log_fn(f"⚠️ Duplicate in flat (skipped): {expected_name} -> from {source_full}")
                continue

            src = unique_lookup.get(expected_key)
            if not src:
                missing_not_found += 1
                missing_rows.append({
                    "row_index": None,
                    "source": source_full,
                    "target_folder_rel": target_folder_rel,
                    "expected_output_name": expected_name,
                    "reason": "Not found in flat folder"
                })
                log_fn(f"⚠️ Missing in flat: {expected_name} -> from {source_full}")
                continue

            dest_dir = os.path.join(destination_root, target_folder_rel)
            ensure_dir(dest_dir)
            dest_path = os.path.join(dest_dir, expected_name)

            try:
                if move_files:
                    shutil.move(src, dest_path)
                    log_fn(f"✅ Moved: {expected_name} -> {dest_path}")
                else:
                    shutil.copy2(src, dest_path)
                    log_fn(f"✅ Copied: {expected_name} -> {dest_path}")
                moved_or_copied += 1
                used_unique.add(expected_key)
            except Exception as e:
                errors += 1
                log_fn(f"❌ Error for {expected_name}: {e}")

    # Orphans: any flat output not used + all duplicates (never touched)
    orphan_rows = []
    for key, paths in all_files.items():
        for p in paths:
            if key in duplicates or key not in used_unique:
                orphan_rows.append({"flat_filename": os.path.basename(p), "flat_full_path": p})

    summary = {
        "mode": "Original structure template",
        "original_root": original_root,
        "flat_output_folder": flat_output_folder,
        "destination_root": destination_root,
        "action": "move" if move_files else "copy",
        "output_ext": output_ext,
        "output_suffix": output_suffix,
        "scan_all": scan_all,
        "include_exts": ",".join(include_exts) if include_exts else "",
        "inputs_scanned": total_inputs,
        "moved_or_copied": moved_or_copied,
        "missing_not_found": missing_not_found,
        "skipped_duplicates": skipped_duplicates,
        "duplicate_filenames_in_flat": len(duplicates),
        "duplicate_files_in_flat": sum(len(v) for v in duplicates.values()),
        "orphans_total": len(orphan_rows),
        "errors": errors,
    }

    if report_path:
        write_report_one_workbook(report_path, summary, missing_rows, duplicate_rows, orphan_rows)
        log_fn(f"📄 Report written: {report_path}")

    log_fn("")
    log_fn("===== Summary =====")
    log_fn(f"Inputs scanned:             {total_inputs}")
    log_fn(f"Moved/Copied:               {moved_or_copied}")
    log_fn(f"Missing (not found):        {missing_not_found}")
    log_fn(f"Skipped (duplicates):       {skipped_duplicates}")
    log_fn(f"Duplicate filenames (flat): {len(duplicates)}")
    log_fn(f"Duplicate files (flat):     {sum(len(v) for v in duplicates.values())}")
    log_fn(f"Orphans total:              {len(orphan_rows)}")
    log_fn(f"Errors:                     {errors}")
    log_fn("===================")

# =========================================================
# Mode B: Rebuild using XLSX mapping (Directory + Name etc.)
# =========================================================

def process_from_excel_mapping(
    spreadsheet_path: str,
    sheet_name: str,
    flat_folder: str,
    destination_root: str,
    directory_col: str,
    name_col: str,
    output_ext: str = ".pdf",
    output_suffix: str = "",
    directory_is_full_path: bool = False,
    move_files: bool = True,
    report_path: str = "",
    log_fn=print,
):
    """
    Uses spreadsheet mapping to place outputs from flat folder into recreated structure.

    IMPORTANT: ext mismatch is expected (e.g. DGN/TIF -> PDF). We match on base name only:
      Name: "thing" OR "thing.dgn" -> expected output "thing" + suffix + ".pdf"
    """
    if not os.path.isfile(spreadsheet_path):
        raise ValueError("Spreadsheet file not found.")
    if not os.path.isdir(flat_folder):
        raise ValueError("Flat folder not found.")
    ensure_dir(destination_root)

    if not output_ext.startswith("."):
        output_ext = "." + output_ext
    output_ext = output_ext.lower()
    output_suffix = (output_suffix or "").strip()

    df = pd.read_excel(spreadsheet_path, sheet_name=sheet_name)

    if directory_col not in df.columns:
        raise ValueError(f"Directory column '{directory_col}' not found in sheet '{sheet_name}'.")
    if name_col not in df.columns:
        raise ValueError(f"Name column '{name_col}' not found in sheet '{sheet_name}'.")

    # Index flat outputs (only output_ext)
    unique_lookup, duplicates, all_files = build_flat_index(flat_folder, allowed_exts=(output_ext,))

    duplicate_rows = []
    for key, paths in duplicates.items():
        for p in paths:
            duplicate_rows.append({
                "flat_filename": os.path.basename(p),
                "duplicate_count": len(paths),
                "flat_full_path": p
            })

    used_unique = set()
    missing_rows = []

    rows_considered = 0
    moved_or_copied = 0
    missing_not_found = 0
    skipped_duplicates = 0
    errors = 0

    for idx, row in df.iterrows():
        dir_val = row.get(directory_col, "")
        name_val = row.get(name_col, "")

        if pd.isna(dir_val) or str(dir_val).strip() == "" or pd.isna(name_val) or str(name_val).strip() == "":
            continue

        directory_value = normalize_path(dir_val)

        base_name_raw = str(name_val).strip()
        base_name = os.path.splitext(base_name_raw)[0].strip()

        expected_name = f"{base_name}{output_suffix}{output_ext}"
        expected_key = expected_name.lower()

        rows_considered += 1

        target_folder = os.path.dirname(directory_value)

        if directory_is_full_path:
            target_folder_rel = strip_drive_and_leading_slashes(target_folder)
        else:
            target_folder_rel = target_folder.lstrip("\\").rstrip("\\")

        if expected_key in duplicates:
            skipped_duplicates += 1
            missing_rows.append({
                "row_index": idx,
                "source": directory_value,
                "target_folder_rel": target_folder_rel,
                "expected_output_name": expected_name,
                "reason": "Duplicate filename in flat folder (skipped)"
            })
            log_fn(f"⚠️ Duplicate in flat (skipped): {expected_name} (row {idx})")
            continue

        src = unique_lookup.get(expected_key)
        if not src:
            missing_not_found += 1
            missing_rows.append({
                "row_index": idx,
                "source": directory_value,
                "target_folder_rel": target_folder_rel,
                "expected_output_name": expected_name,
                "reason": "Not found in flat folder"
            })
            log_fn(f"⚠️ Missing in flat: {expected_name} (row {idx})")
            continue

        dest_dir = os.path.join(destination_root, target_folder_rel)
        ensure_dir(dest_dir)
        dest_path = os.path.join(dest_dir, expected_name)

        try:
            if move_files:
                shutil.move(src, dest_path)
                log_fn(f"✅ Moved: {expected_name} -> {dest_path}")
            else:
                shutil.copy2(src, dest_path)
                log_fn(f"✅ Copied: {expected_name} -> {dest_path}")
            moved_or_copied += 1
            used_unique.add(expected_key)
        except Exception as e:
            errors += 1
            log_fn(f"❌ Error for {expected_name}: {e}")

    orphan_rows = []
    for key, paths in all_files.items():
        for p in paths:
            if key in duplicates or key not in used_unique:
                orphan_rows.append({"flat_filename": os.path.basename(p), "flat_full_path": p})

    summary = {
        "mode": "XLSX mapping",
        "spreadsheet": spreadsheet_path,
        "worksheet": sheet_name,
        "directory_col": directory_col,
        "name_col": name_col,
        "flat_folder": flat_folder,
        "destination_root": destination_root,
        "action": "move" if move_files else "copy",
        "output_ext": output_ext,
        "output_suffix": output_suffix,
        "rows_considered": rows_considered,
        "moved_or_copied": moved_or_copied,
        "missing_not_found": missing_not_found,
        "skipped_duplicates": skipped_duplicates,
        "duplicate_filenames_in_flat": len(duplicates),
        "duplicate_files_in_flat": sum(len(v) for v in duplicates.values()),
        "orphans_total": len(orphan_rows),
        "errors": errors,
    }

    if report_path:
        write_report_one_workbook(report_path, summary, missing_rows, duplicate_rows, orphan_rows)
        log_fn(f"📄 Report written: {report_path}")

    log_fn("")
    log_fn("===== Summary =====")
    log_fn(f"Rows considered:            {rows_considered}")
    log_fn(f"Moved/Copied:               {moved_or_copied}")
    log_fn(f"Missing (not found):        {missing_not_found}")
    log_fn(f"Skipped (duplicates):       {skipped_duplicates}")
    log_fn(f"Duplicate filenames (flat): {len(duplicates)}")
    log_fn(f"Duplicate files (flat):     {sum(len(v) for v in duplicates.values())}")
    log_fn(f"Orphans total:              {len(orphan_rows)}")
    log_fn(f"Errors:                     {errors}")
    log_fn("===================")

# =========================================================
# GUI (Notebook with 3 tabs, styled like earlier scripts)
# =========================================================

class AllInOneApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("All-in-One: Pre-OCR Token Rename + Rebuild Structure (Template + XLSX)")
        self.geometry("1150x900")

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True)

        self.tab_token = ttk.Frame(nb)
        self.tab_template = ttk.Frame(nb)
        self.tab_xlsx = ttk.Frame(nb)

        nb.add(self.tab_token, text="Pre-OCR: Append Token")
        nb.add(self.tab_template, text="Rebuild: Original Folder Structure")
        nb.add(self.tab_xlsx, text="Rebuild: XLSX Mapping")

        self._build_tab_token()
        self._build_tab_template()
        self._build_tab_xlsx()

    # ---------------- Tab 0 ----------------
    def _build_tab_token(self):
        pad = {"padx": 10, "pady": 6}

        self.p_source = tk.StringVar()
        self.p_staging = tk.StringVar()
        self.p_report = tk.StringVar()

        self.p_mode = tk.StringVar(value="copy")  # safer default
        self.p_separator = tk.StringVar(value="__")
        self.p_token_mode = tk.StringVar(value="counter")  # counter or size
        self.p_counter_padding = tk.IntVar(value=6)

        self.p_filter = tk.BooleanVar(value=False)
        self.p_exts = tk.StringVar(value="tif,tiff,dgn,dwg,doc,docx,xls,xlsx,msg,jpg,jpeg,png,pdf")

        frm = ttk.Frame(self.tab_token)
        frm.pack(fill="x", **pad)
        frm.columnconfigure(0, weight=1)

        ttk.Label(frm, text="Source root (original folder structure):").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.p_source, width=120).grid(row=1, column=0, sticky="we")
        ttk.Button(frm, text="Browse…", command=lambda: self._pick_folder(self.p_source)).grid(row=1, column=1, padx=8)

        ttk.Label(frm, text="Staging root (only used in Copy mode):").grid(row=2, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.p_staging, width=120).grid(row=3, column=0, sticky="we")
        ttk.Button(frm, text="Browse…", command=lambda: self._pick_folder(self.p_staging)).grid(row=3, column=1, padx=8)

        ttk.Label(frm, text="Report workbook (.xlsx):").grid(row=4, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.p_report, width=120).grid(row=5, column=0, sticky="we")
        ttk.Button(frm, text="Save as…", command=lambda: self._pick_report(self.p_report)).grid(row=5, column=1, padx=8)

        opts = ttk.LabelFrame(self.tab_token, text="Token options")
        opts.pack(fill="x", padx=10, pady=6)

        inner = ttk.Frame(opts)
        inner.pack(fill="x", padx=10, pady=10)
        inner.columnconfigure(3, weight=1)

        ttk.Label(inner, text="Mode:").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(inner, text="Copy to staging (recommended)", variable=self.p_mode, value="copy").grid(row=0, column=1, sticky="w", padx=8)
        ttk.Radiobutton(inner, text="Rename in place (risky)", variable=self.p_mode, value="rename").grid(row=0, column=2, sticky="w", padx=8)

        ttk.Label(inner, text="Token:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Radiobutton(inner, text="Global counter (walk order)", variable=self.p_token_mode, value="counter").grid(row=1, column=1, sticky="w", padx=8, pady=(8, 0))
        ttk.Radiobutton(inner, text="File size (bytes)", variable=self.p_token_mode, value="size").grid(row=1, column=2, sticky="w", padx=8, pady=(8, 0))

        ttk.Label(inner, text="Counter padding:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(inner, textvariable=self.p_counter_padding, width=8).grid(row=2, column=1, sticky="w", padx=8, pady=(8, 0))
        ttk.Label(inner, text="(e.g. 6 -> 000001)").grid(row=2, column=2, sticky="w", pady=(8, 0))

        ttk.Label(inner, text="Separator:").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(inner, textvariable=self.p_separator, width=8).grid(row=3, column=1, sticky="w", padx=8, pady=(8, 0))

        filt = ttk.LabelFrame(self.tab_token, text="Optional: only process certain file types")
        filt.pack(fill="x", padx=10, pady=6)

        f2 = ttk.Frame(filt)
        f2.pack(fill="x", padx=10, pady=10)
        ttk.Checkbutton(f2, text="Enable extension filter", variable=self.p_filter, command=self._toggle_token_exts).pack(side="left")
        self.p_ext_entry = ttk.Entry(f2, textvariable=self.p_exts, width=90)
        self.p_ext_entry.pack(side="left", padx=10, fill="x", expand=True)
        self._toggle_token_exts()

        btns = ttk.Frame(self.tab_token)
        btns.pack(fill="x", **pad)
        self.p_run_btn = ttk.Button(btns, text="Run", command=self._run_token)
        self.p_run_btn.pack(side="left")
        ttk.Button(btns, text="Clear Log", command=lambda: self._clear_log(self.p_log)).pack(side="left", padx=10)

        self.p_log = tk.Text(self.tab_token, height=22, wrap="word")
        self.p_log.pack(fill="both", expand=True, padx=10, pady=10)
        self._log(self.p_log, "Tab 0: Create a staging copy where filenames become unique (recommended: counter).")

    def _toggle_token_exts(self):
        self.p_ext_entry.configure(state=("normal" if self.p_filter.get() else "disabled"))

    def _run_token(self):
        source = self.p_source.get().strip()
        staging = self.p_staging.get().strip()
        report = self.p_report.get().strip()
        mode = self.p_mode.get()
        sep = self.p_separator.get().strip() or "__"
        token_mode = self.p_token_mode.get()
        padding = int(self.p_counter_padding.get() or 6)

        if not source:
            messagebox.showerror("Missing info", "Pick a Source root folder.")
            return
        if mode == "copy" and not staging:
            messagebox.showerror("Missing info", "Copy mode needs a Staging root folder.")
            return

        filter_enabled = bool(self.p_filter.get())
        include_exts = parse_extensions(self.p_exts.get()) if filter_enabled else ()

        self.p_run_btn.config(state="disabled")
        self._log(self.p_log, "")
        self._log(self.p_log, "Running…")

        def worker():
            try:
                pre_ocr_append_token(
                    source_root=source,
                    mode=mode,
                    staging_root=staging,
                    filter_enabled=filter_enabled,
                    include_exts=include_exts,
                    separator=sep,
                    token_mode=token_mode,
                    counter_padding=padding,
                    report_path=report,
                    log_fn=lambda m: self.after(0, self._log, self.p_log, m),
                )
                self.after(0, lambda: messagebox.showinfo("Done", "Finished pre-OCR token rename/copy."))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally:
                self.after(0, lambda: self.p_run_btn.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    # ---------------- Tab 1 ----------------
    def _build_tab_template(self):
        pad = {"padx": 10, "pady": 6}

        self.t_original = tk.StringVar()
        self.t_flat = tk.StringVar()
        self.t_dest = tk.StringVar()
        self.t_report = tk.StringVar()

        self.t_action = tk.StringVar(value="move")
        self.t_scan_mode = tk.StringVar(value="all")
        self.t_exts = tk.StringVar(value="doc,docx,xls,xlsx,dwg,dgn,msg,jpg,jpeg,png,tif,tiff")
        self.t_output_ext = tk.StringVar(value=".pdf")
        self.t_output_suffix = tk.StringVar(value="_OCR")

        frm = ttk.Frame(self.tab_template)
        frm.pack(fill="x", **pad)
        frm.columnconfigure(0, weight=1)

        ttk.Label(frm, text="Original root (nested structure):").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.t_original, width=120).grid(row=1, column=0, sticky="we")
        ttk.Button(frm, text="Browse…", command=lambda: self._pick_folder(self.t_original)).grid(row=1, column=1, padx=8)

        ttk.Label(frm, text="Flat output folder (PDF Tools output):").grid(row=2, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.t_flat, width=120).grid(row=3, column=0, sticky="we")
        ttk.Button(frm, text="Browse…", command=lambda: self._pick_folder(self.t_flat)).grid(row=3, column=1, padx=8)

        ttk.Label(frm, text="Destination root (recreated structure):").grid(row=4, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.t_dest, width=120).grid(row=5, column=0, sticky="we")
        ttk.Button(frm, text="Browse…", command=lambda: self._pick_folder(self.t_dest)).grid(row=5, column=1, padx=8)

        ttk.Label(frm, text="Report workbook (.xlsx) (one file with sheets):").grid(row=6, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.t_report, width=120).grid(row=7, column=0, sticky="we")
        ttk.Button(frm, text="Save as…", command=lambda: self._pick_report(self.t_report)).grid(row=7, column=1, padx=8)

        opts = ttk.Frame(self.tab_template)
        opts.pack(fill="x", **pad)

        ttk.Label(opts, text="Action:").pack(side="left")
        ttk.Radiobutton(opts, text="Move", variable=self.t_action, value="move").pack(side="left", padx=10)
        ttk.Radiobutton(opts, text="Copy", variable=self.t_action, value="copy").pack(side="left", padx=10)

        ttk.Label(opts, text="Output ext:").pack(side="left", padx=(30, 0))
        ttk.Entry(opts, textvariable=self.t_output_ext, width=10).pack(side="left", padx=8)

        ttk.Label(opts, text="Output suffix:").pack(side="left", padx=(18, 0))
        ttk.Entry(opts, textvariable=self.t_output_suffix, width=10).pack(side="left", padx=8)

        scan = ttk.LabelFrame(self.tab_template, text="Scan which source files")
        scan.pack(fill="x", padx=10, pady=6)

        top = ttk.Frame(scan)
        top.pack(fill="x", padx=10, pady=10)

        ttk.Radiobutton(top, text="Scan ALL file types (except output ext)", variable=self.t_scan_mode, value="all",
                        command=self._toggle_template_exts).pack(side="left")
        ttk.Radiobutton(top, text="Only scan these extensions:", variable=self.t_scan_mode, value="filter",
                        command=self._toggle_template_exts).pack(side="left", padx=18)

        self.t_ext_entry = ttk.Entry(scan, textvariable=self.t_exts)
        self.t_ext_entry.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Label(scan, text="Example: docx,xlsx,dwg,dgn,tif").pack(anchor="w", padx=10, pady=(0, 10))
        self._toggle_template_exts()

        btns = ttk.Frame(self.tab_template)
        btns.pack(fill="x", **pad)

        self.t_run_btn = ttk.Button(btns, text="Run", command=self._run_template)
        self.t_run_btn.pack(side="left")
        ttk.Button(btns, text="Clear Log", command=lambda: self._clear_log(self.t_log)).pack(side="left", padx=10)

        self.t_log = tk.Text(self.tab_template, height=22, wrap="word")
        self.t_log.pack(fill="both", expand=True, padx=10, pady=10)
        self._log(self.t_log, "Tab 1: Uses original folder structure as template. Matches base + suffix + ext (e.g. __000001_OCR.pdf).")

    def _toggle_template_exts(self):
        is_filter = (self.t_scan_mode.get() == "filter")
        self.t_ext_entry.configure(state=("normal" if is_filter else "disabled"))

    def _run_template(self):
        original = self.t_original.get().strip()
        flat = self.t_flat.get().strip()
        dest = self.t_dest.get().strip()
        report = self.t_report.get().strip()

        if not original or not flat or not dest:
            messagebox.showerror("Missing info", "Please select original root, flat output folder, and destination root.")
            return

        move_files = (self.t_action.get() == "move")
        if os.path.abspath(dest) == os.path.abspath(flat) and move_files:
            messagebox.showerror("Bad destination", "Destination cannot be the same as the flat folder when moving.")
            return

        output_ext = self.t_output_ext.get().strip() or ".pdf"
        output_suffix = self.t_output_suffix.get().strip()

        scan_all = (self.t_scan_mode.get() == "all")
        include_exts = ()
        if not scan_all:
            include_exts = parse_extensions(self.t_exts.get())
            if not include_exts:
                messagebox.showerror("Extensions needed", "Please enter at least one extension for filtered scan.")
                return

        self.t_run_btn.config(state="disabled")
        self._log(self.t_log, "")
        self._log(self.t_log, "Running…")

        def worker():
            try:
                recreate_from_original_tree(
                    original_root=original,
                    flat_output_folder=flat,
                    destination_root=dest,
                    move_files=move_files,
                    scan_all=scan_all,
                    include_exts=include_exts,
                    output_ext=output_ext,
                    output_suffix=output_suffix,
                    report_path=report,
                    log_fn=lambda m: self.after(0, self._log, self.t_log, m),
                )
                self.after(0, lambda: messagebox.showinfo("Done", "Finished rebuilding folder structure (template mode)."))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally:
                self.after(0, lambda: self.t_run_btn.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    # ---------------- Tab 2 ----------------
    def _build_tab_xlsx(self):
        pad = {"padx": 10, "pady": 6}

        self.x_file = tk.StringVar()
        self.x_sheet = tk.StringVar()
        self.x_flat = tk.StringVar()
        self.x_dest = tk.StringVar()
        self.x_report = tk.StringVar()

        self.x_directory_col = tk.StringVar()
        self.x_name_col = tk.StringVar()
        self.x_output_ext = tk.StringVar(value=".pdf")
        self.x_output_suffix = tk.StringVar(value="_OCR")

        self.x_action = tk.StringVar(value="move")
        self.x_dir_fullpath = tk.BooleanVar(value=False)

        self.x_sheet_names = []
        self.x_columns = []

        frm = ttk.Frame(self.tab_xlsx)
        frm.pack(fill="x", **pad)
        frm.columnconfigure(0, weight=1)

        ttk.Label(frm, text="Mapping spreadsheet (.xlsx):").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.x_file, width=120).grid(row=1, column=0, sticky="we")
        ttk.Button(frm, text="Browse…", command=self._pick_excel).grid(row=1, column=1, padx=8)

        ttk.Label(frm, text="Worksheet tab:").grid(row=2, column=0, sticky="w")
        self.x_sheet_combo = ttk.Combobox(frm, textvariable=self.x_sheet, state="readonly", width=50)
        self.x_sheet_combo.grid(row=3, column=0, sticky="w")
        self.x_sheet_combo.bind("<<ComboboxSelected>>", self._load_columns_for_sheet)

        ttk.Label(frm, text="Flat output folder (PDF Tools output):").grid(row=4, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.x_flat, width=120).grid(row=5, column=0, sticky="we")
        ttk.Button(frm, text="Browse…", command=lambda: self._pick_folder(self.x_flat)).grid(row=5, column=1, padx=8)

        ttk.Label(frm, text="Destination root (recreated structure):").grid(row=6, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.x_dest, width=120).grid(row=7, column=0, sticky="we")
        ttk.Button(frm, text="Browse…", command=lambda: self._pick_folder(self.x_dest)).grid(row=7, column=1, padx=8)

        ttk.Label(frm, text="Report workbook (.xlsx) (one file with sheets):").grid(row=8, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.x_report, width=120).grid(row=9, column=0, sticky="we")
        ttk.Button(frm, text="Save as…", command=lambda: self._pick_report(self.x_report)).grid(row=9, column=1, padx=8)

        cols = ttk.LabelFrame(self.tab_xlsx, text="Select headers (from chosen worksheet)")
        cols.pack(fill="x", padx=10, pady=6)

        inner = ttk.Frame(cols)
        inner.pack(fill="x", padx=10, pady=10)
        inner.columnconfigure(1, weight=1)

        ttk.Label(inner, text="Directory column (full file path incl filename):").grid(row=0, column=0, sticky="w")
        self.x_dir_combo = ttk.Combobox(inner, textvariable=self.x_directory_col, values=self.x_columns, state="readonly", width=60)
        self.x_dir_combo.grid(row=0, column=1, sticky="w", padx=10)

        ttk.Label(inner, text="Name column (base name; may include source ext):").grid(row=1, column=0, sticky="w")
        self.x_name_combo = ttk.Combobox(inner, textvariable=self.x_name_col, values=self.x_columns, state="readonly", width=60)
        self.x_name_combo.grid(row=1, column=1, sticky="w", padx=10)

        ttk.Label(inner, text="Output ext:").grid(row=2, column=0, sticky="w")
        ttk.Entry(inner, textvariable=self.x_output_ext, width=10).grid(row=2, column=1, sticky="w", padx=10)

        ttk.Label(inner, text="Output suffix (e.g. _OCR):").grid(row=3, column=0, sticky="w")
        ttk.Entry(inner, textvariable=self.x_output_suffix, width=10).grid(row=3, column=1, sticky="w", padx=10)

        ttk.Checkbutton(
            inner,
            text="Directory column is a FULL path (has drive letter like E:\\...)",
            variable=self.x_dir_fullpath
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(8, 0))

        opts = ttk.Frame(self.tab_xlsx)
        opts.pack(fill="x", padx=10, pady=6)

        ttk.Label(opts, text="Action:").pack(side="left")
        ttk.Radiobutton(opts, text="Move", variable=self.x_action, value="move").pack(side="left", padx=10)
        ttk.Radiobutton(opts, text="Copy", variable=self.x_action, value="copy").pack(side="left", padx=10)

        ttk.Label(
            opts,
            text="Matching rule: base filename (+ suffix) (so DGN/TIF/etc -> PDF is fine). Duplicates in flat are skipped."
        ).pack(side="left", padx=(25, 0))

        btns = ttk.Frame(self.tab_xlsx)
        btns.pack(fill="x", **pad)

        self.x_run_btn = ttk.Button(btns, text="Run", command=self._run_xlsx)
        self.x_run_btn.pack(side="left")
        ttk.Button(btns, text="Clear Log", command=lambda: self._clear_log(self.x_log)).pack(side="left", padx=10)

        self.x_log = tk.Text(self.tab_xlsx, height=22, wrap="word")
        self.x_log.pack(fill="both", expand=True, padx=10, pady=10)
        self._log(self.x_log, "Tab 2: Uses XLSX mapping. Pick worksheet + headers, then Run.")

    # ---------------- Shared GUI actions ----------------
    def _log(self, widget: tk.Text, msg: str):
        widget.insert("end", msg + "\n")
        widget.see("end")

    def _clear_log(self, widget: tk.Text):
        widget.delete("1.0", "end")

    def _pick_folder(self, var: tk.StringVar):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def _pick_report(self, var: tk.StringVar):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            title="Save report workbook as",
        )
        if path:
            if not path.lower().endswith(".xlsx"):
                path += ".xlsx"
            var.set(path)

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel Workbook", "*.xlsx *.xls")],
            title="Select mapping spreadsheet",
        )
        if not path:
            return

        self.x_file.set(path)

        try:
            xls = pd.ExcelFile(path)
            self.x_sheet_names = xls.sheet_names
            self.x_sheet_combo["values"] = self.x_sheet_names
            if self.x_sheet_names:
                self.x_sheet.set(self.x_sheet_names[0])
                self._load_columns_for_sheet()
        except Exception as e:
            messagebox.showerror("Error", f"Could not read Excel file:\n{e}")

    def _load_columns_for_sheet(self, event=None):
        path = self.x_file.get().strip()
        sheet = self.x_sheet.get().strip()
        if not path or not sheet:
            return

        try:
            df = pd.read_excel(path, sheet_name=sheet, nrows=1)
            self.x_columns = list(df.columns)
            self.x_dir_combo["values"] = self.x_columns
            self.x_name_combo["values"] = self.x_columns

            # sensible defaults for your sheet
            if "Directory" in self.x_columns:
                self.x_directory_col.set("Directory")
            elif self.x_columns:
                self.x_directory_col.set(self.x_columns[0])

            if "Name" in self.x_columns:
                self.x_name_col.set("Name")
            elif len(self.x_columns) > 1:
                self.x_name_col.set(self.x_columns[1])

        except Exception as e:
            messagebox.showerror("Error", f"Could not read sheet '{sheet}':\n{e}")

    def _run_xlsx(self):
        excel_path = self.x_file.get().strip()
        sheet = self.x_sheet.get().strip()
        flat = self.x_flat.get().strip()
        dest = self.x_dest.get().strip()
        report = self.x_report.get().strip()

        directory_col = self.x_directory_col.get().strip()
        name_col = self.x_name_col.get().strip()
        output_ext = self.x_output_ext.get().strip() or ".pdf"
        output_suffix = self.x_output_suffix.get().strip()

        if not excel_path or not sheet:
            messagebox.showerror("Missing info", "Please select an Excel file and worksheet tab.")
            return
        if not flat or not dest:
            messagebox.showerror("Missing info", "Please select flat folder and destination root.")
            return
        if not directory_col or not name_col:
            messagebox.showerror("Missing info", "Please select Directory and Name headers.")
            return

        move_files = (self.x_action.get() == "move")
        if os.path.abspath(dest) == os.path.abspath(flat) and move_files:
            messagebox.showerror("Bad destination", "Destination cannot be the same as the flat folder when moving.")
            return

        self.x_run_btn.config(state="disabled")
        self._log(self.x_log, "")
        self._log(self.x_log, "Running…")

        def worker():
            try:
                process_from_excel_mapping(
                    spreadsheet_path=excel_path,
                    sheet_name=sheet,
                    flat_folder=flat,
                    destination_root=dest,
                    directory_col=directory_col,
                    name_col=name_col,
                    output_ext=output_ext,
                    output_suffix=output_suffix,
                    directory_is_full_path=bool(self.x_dir_fullpath.get()),
                    move_files=move_files,
                    report_path=report,
                    log_fn=lambda m: self.after(0, self._log, self.x_log, m),
                )
                self.after(0, lambda: messagebox.showinfo("Done", "Finished placing files from XLSX mapping."))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", str(e)))
            finally:
                self.after(0, lambda: self.x_run_btn.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = AllInOneApp()
    app.mainloop()
