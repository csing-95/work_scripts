import os
import threading
import shutil
from datetime import datetime
from collections import defaultdict

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# -----------------------------
# Helpers
# -----------------------------
def bytes_to_mb(b: int) -> float:
    return round(b / (1024 ** 2), 4)


def bytes_to_gb(b: int) -> float:
    return round(b / (1024 ** 3), 6)


def ts_to_str(ts: float) -> str:
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")


def safe_relpath(full_path: str, root: str) -> str:
    rel = os.path.relpath(full_path, root)
    return rel.replace("\\", "/")


def should_skip_file(full_path: str, skip_hidden: bool, skip_extensions: set[str]) -> bool:
    base = os.path.basename(full_path)
    if skip_hidden and base.startswith("."):
        return True
    if base.startswith("~$"):
        return True
    ext = os.path.splitext(base)[1].lower()
    return ext in skip_extensions


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def scan_all_files(root: str, skip_hidden: bool, skip_extensions: set[str], progress_cb=None):
    """
    Returns list of dicts: one per file, with metadata + relative path pieces.
    """
    root = os.path.normpath(root)

    total = 0
    for _, _, files in os.walk(root):
        total += len(files)

    seen = 0
    rows = []

    for dirpath, _, filenames in os.walk(root):
        for fn in filenames:
            full = os.path.join(dirpath, fn)

            seen += 1
            if progress_cb:
                progress_cb(seen, total, full)

            if should_skip_file(full, skip_hidden, skip_extensions):
                continue

            try:
                st = os.stat(full)
            except OSError:
                st = None

            rel = safe_relpath(full, root)
            rel_dir = os.path.dirname(rel).replace("\\", "/")
            stem, ext = os.path.splitext(os.path.basename(rel))
            ext = ext.lower()

            rows.append({
                "File Name": os.path.basename(rel),
                "Stem": stem,
                "Extension": ext,
                "Relative Path": rel,
                "Relative Dir": rel_dir,  # folder within root
                "Full Path": os.path.normpath(full),
                "Size (Bytes)": st.st_size if st else None,
                "Size (MB)": bytes_to_mb(st.st_size) if st else None,
                "Size (GB)": bytes_to_gb(st.st_size) if st else None,
                "Created": ts_to_str(st.st_ctime) if st else None,
                "Modified": ts_to_str(st.st_mtime) if st else None,
                "Accessed": ts_to_str(st.st_atime) if st else None,
            })

    return rows


def build_expected_pdfs(original_rows):
    """
    Build expected searchable PDFs from Original:
    Expected key = (Relative Dir, Stem)
    Expect revised to have: Relative Dir / Stem.pdf
    Also keep list of source originals for extraction (all extensions for that key).
    """
    expected = {}
    sources_by_key = defaultdict(list)

    for r in original_rows:
        key = (r["Relative Dir"], r["Stem"])
        sources_by_key[key].append(r)

    for (rel_dir, stem), src_rows in sources_by_key.items():
        expected_rel = f"{rel_dir}/{stem}.pdf" if rel_dir else f"{stem}.pdf"
        expected_key = (rel_dir, stem)

        # representative metadata: pick the largest file (often the “real” one) just for reporting
        rep = max(
            src_rows,
            key=lambda x: (x["Size (Bytes)"] or -1, x["Extension"] == ".pdf")
        )

        expected[expected_key] = {
            "Relative Dir": rel_dir,
            "Stem": stem,
            "Expected Revised Relative Path": expected_rel,
            "Expected Revised File Name": f"{stem}.pdf",
            "Source Variants Count": len(src_rows),
            "Source Extensions": ", ".join(sorted({x["Extension"] for x in src_rows})),
            "Source Example Full Path": rep["Full Path"],
        }

    return expected, sources_by_key


def index_revised_pdfs(revised_rows):
    """
    Index Revised PDFs by (Relative Dir, Stem) if file is .pdf
    """
    idx = {}
    for r in revised_rows:
        if r["Extension"] != ".pdf":
            continue
        key = (r["Relative Dir"], r["Stem"])
        idx[key] = r
    return idx


def export_excel(out_path, original_root, revised_root, missing_rows, extra_rows, summary_dict):
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        pd.DataFrame([summary_dict]).to_excel(writer, sheet_name="Summary", index=False)

        df_missing = pd.DataFrame(missing_rows)
        df_missing.to_excel(writer, sheet_name="Missing_Searchable_PDFs", index=False)

        df_extra = pd.DataFrame(extra_rows)
        df_extra.to_excel(writer, sheet_name="Extra_PDFs_in_Revised", index=False)


def copy_sources_for_missing(missing_rows, sources_by_key, original_root, dest_root, progress_cb=None, stop_flag_fn=None):
    """
    For each missing expected PDF key (Relative Dir, Stem), copy ALL source variants from original
    into dest_root, preserving the original relative folder structure.
    """
    # Build list of files to copy
    to_copy = []
    for m in missing_rows:
        key = (m["Relative Dir"], m["Stem"])
        for src in sources_by_key.get(key, []):
            rel = src["Relative Path"]
            src_full = os.path.join(original_root, rel.replace("/", os.sep))
            dst_full = os.path.join(dest_root, rel.replace("/", os.sep))
            to_copy.append((src_full, dst_full))

    total = len(to_copy)
    copied = 0
    failed = 0

    for i, (src_full, dst_full) in enumerate(to_copy, start=1):
        if stop_flag_fn and stop_flag_fn():
            return {"status": "cancelled", "copied": copied, "failed": failed, "total": total}

        ensure_dir(os.path.dirname(dst_full))
        try:
            shutil.copy2(src_full, dst_full)
            copied += 1
        except Exception:
            failed += 1

        if progress_cb:
            progress_cb(i, total, dst_full)

    return {"status": "done", "copied": copied, "failed": failed, "total": total}


# -----------------------------
# GUI
# -----------------------------
class SearchablePdfCompareApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Original → Searchable PDFs: Compare + Extract Missing Sources")
        self.geometry("940x520")
        self.resizable(False, False)

        self.original_var = tk.StringVar()
        self.revised_var = tk.StringVar()
        self.outfile_var = tk.StringVar()

        self.extract_dest_var = tk.StringVar()

        self.skip_hidden_var = tk.BooleanVar(value=True)
        self.skip_ext_var = tk.StringVar(value=".tmp,.log,.bak")

        self.status_var = tk.StringVar(value="Pick folders to compare.")
        self.progress_var = tk.DoubleVar(value=0)

        self.stop_flag = False

        # In-memory results
        self.last_missing = []
        self.last_extra = []
        self.last_sources_by_key = None

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True, padx=12, pady=12)

        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True)

        self.tab_compare = ttk.Frame(self.nb)
        self.tab_extract = ttk.Frame(self.nb)
        self.nb.add(self.tab_compare, text="1) Compare")
        self.nb.add(self.tab_extract, text="2) Extract Missing Sources")

        # ---- Compare tab
        pad = {"padx": 10, "pady": 6}
        frm = self.tab_compare

        ttk.Label(frm, text="Original folder (mixed types)").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.original_var, width=84).grid(row=0, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_original).grid(row=0, column=2, **pad)

        ttk.Label(frm, text="Revised folder (PDFs only)").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.revised_var, width=84).grid(row=1, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_revised).grid(row=1, column=2, **pad)

        ttk.Label(frm, text="Output Excel (.xlsx)").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.outfile_var, width=84).grid(row=2, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Save as…", command=self.pick_outfile).grid(row=2, column=2, **pad)

        opt = ttk.LabelFrame(frm, text="Options")
        opt.grid(row=3, column=0, columnspan=3, sticky="we", padx=10, pady=10)

        ttk.Checkbutton(opt, text="Skip hidden dotfiles", variable=self.skip_hidden_var)\
            .grid(row=0, column=0, sticky="w", padx=10, pady=6)

        ttk.Label(opt, text="Skip extensions (comma-separated):").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        ttk.Entry(opt, textvariable=self.skip_ext_var, width=45).grid(row=1, column=1, sticky="w", padx=10, pady=6)

        expl = (
            "Matching logic:\n"
            "- We treat a document as: (relative folder path) + (base filename without extension)\n"
            "- For every stem found in Original, we expect: stem.pdf in Revised (same relative folder)\n"
            "- If Original has stem.dwg and stem.pdf, Revised still only needs one stem.pdf"
        )
        ttk.Label(frm, text=expl, justify="left").grid(row=4, column=0, columnspan=3, sticky="w", padx=10, pady=6)

        self.compare_btn = ttk.Button(frm, text="Run comparison", command=self.run_compare)
        self.compare_btn.grid(row=5, column=2, sticky="e", padx=10, pady=10)

        # ---- Extract tab
        frm2 = self.tab_extract

        info = (
            "This will copy the ORIGINAL source files for any missing searchable PDFs.\n"
            "It preserves the original folder structure and copies ALL source variants (e.g., .dwg + .pdf).\n"
            "Run Compare first so the missing list is loaded."
        )
        ttk.Label(frm2, text=info, justify="left").grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=10)

        ttk.Label(frm2, text="Destination folder (Missing Docs)").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm2, textvariable=self.extract_dest_var, width=84).grid(row=1, column=1, sticky="w", **pad)
        ttk.Button(frm2, text="Browse…", command=self.pick_extract_dest).grid(row=1, column=2, **pad)

        self.missing_count_label = ttk.Label(frm2, text="Missing stems loaded: 0")
        self.missing_count_label.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=10)

        self.extract_btn = ttk.Button(frm2, text="Extract missing sources", command=self.run_extract, state="disabled")
        self.extract_btn.grid(row=2, column=2, sticky="e", padx=10, pady=10)

        # ---- Shared progress + status
        self.progress = ttk.Progressbar(root, variable=self.progress_var, maximum=100)
        self.progress.pack(fill="x", padx=10, pady=6)

        ttk.Label(root, textvariable=self.status_var).pack(fill="x", padx=10, pady=6)

        self.cancel_btn = ttk.Button(root, text="Cancel current job", command=self.cancel, state="disabled")
        self.cancel_btn.pack(anchor="e", padx=10, pady=6)

    # -----------------------------
    # UI helpers
    # -----------------------------
    def set_busy(self, busy: bool):
        self.compare_btn.configure(state="disabled" if busy else "normal")
        # extract button should stay disabled if no missing loaded
        if busy:
            self.extract_btn.configure(state="disabled")
        else:
            self.extract_btn.configure(state="normal" if self.last_missing else "disabled")
        self.cancel_btn.configure(state="normal" if busy else "disabled")

    def cancel(self):
        self.stop_flag = True
        self.status_var.set("Cancelling… (finishes current file)")

    def _progress_cb(self, done, total, current_path):
        if total > 0:
            self.progress_var.set(min(100, (done / total) * 100))
        self.status_var.set(f"{done}/{total} … {os.path.basename(current_path)}")
        self.update_idletasks()

    def _make_skip_exts(self):
        skip_exts = {e.strip().lower() for e in self.skip_ext_var.get().split(",") if e.strip()}
        return {e if e.startswith(".") else f".{e}" for e in skip_exts}

    def _maybe_autofill_outfile(self):
        if self.outfile_var.get().strip():
            return
        o = self.original_var.get().strip()
        r = self.revised_var.get().strip()
        if o and r:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            default = os.path.join(os.path.expanduser("~"), "Desktop", f"searchable_pdf_compare_{ts}.xlsx")
            self.outfile_var.set(default)

    # -----------------------------
    # Browse
    # -----------------------------
    def pick_original(self):
        p = filedialog.askdirectory(title="Select ORIGINAL folder")
        if p:
            self.original_var.set(p)
            self._maybe_autofill_outfile()

    def pick_revised(self):
        p = filedialog.askdirectory(title="Select REVISED folder")
        if p:
            self.revised_var.set(p)
            self._maybe_autofill_outfile()

    def pick_outfile(self):
        p = filedialog.asksaveasfilename(
            title="Save Excel output",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if p:
            self.outfile_var.set(p)

    def pick_extract_dest(self):
        p = filedialog.askdirectory(title="Select destination folder for missing docs")
        if p:
            self.extract_dest_var.set(p)

    # -----------------------------
    # Compare logic
    # -----------------------------
    def run_compare(self):
        original = self.original_var.get().strip()
        revised = self.revised_var.get().strip()
        outpath = self.outfile_var.get().strip()

        if not original or not os.path.isdir(original):
            messagebox.showerror("Missing info", "Please select a valid ORIGINAL folder.")
            return
        if not revised or not os.path.isdir(revised):
            messagebox.showerror("Missing info", "Please select a valid REVISED folder.")
            return
        if not outpath:
            messagebox.showerror("Missing info", "Please choose an output Excel file path.")
            return

        self.stop_flag = False
        self.set_busy(True)
        self.progress_var.set(0)
        self.status_var.set("Scanning ORIGINAL…")

        skip_hidden = self.skip_hidden_var.get()
        skip_exts = self._make_skip_exts()

        def worker():
            try:
                orig_rows = scan_all_files(original, skip_hidden, skip_exts, progress_cb=self._progress_cb)
                if self.stop_flag:
                    return

                self.progress_var.set(0)
                self.status_var.set("Scanning REVISED…")
                rev_rows = scan_all_files(revised, skip_hidden, skip_exts, progress_cb=self._progress_cb)
                if self.stop_flag:
                    return

                self.status_var.set("Building expected PDFs…")
                expected, sources_by_key = build_expected_pdfs(orig_rows)
                revised_idx = index_revised_pdfs(rev_rows)

                # Missing expected PDFs
                missing = []
                for key, exp in expected.items():
                    if key not in revised_idx:
                        rel_dir, stem = key
                        # include source list and expected output path
                        row = {
                            "Relative Dir": rel_dir,
                            "Stem": stem,
                            "Expected Revised Relative Path": exp["Expected Revised Relative Path"],
                            "Expected Revised File Name": exp["Expected Revised File Name"],
                            "Source Extensions": exp["Source Extensions"],
                            "Source Variants Count": exp["Source Variants Count"],
                            "Source Example Full Path": exp["Source Example Full Path"],
                        }
                        missing.append(row)

                # Extra PDFs in revised (pdfs that don't correspond to any original stem)
                extra = []
                expected_keys = set(expected.keys())
                for key, r in revised_idx.items():
                    if key not in expected_keys:
                        rel_dir, stem = key
                        extra.append({
                            "Relative Dir": rel_dir,
                            "Stem": stem,
                            "Revised Relative Path": r["Relative Path"],
                            "Revised Full Path": r["Full Path"],
                            "Size (MB)": r["Size (MB)"],
                            "Modified": r["Modified"],
                        })

                summary = {
                    "Original Folder": original,
                    "Revised Folder": revised,
                    "Unique Document Stems in Original": len(expected),
                    "Missing Searchable PDFs": len(missing),
                    "Extra PDFs in Revised": len(extra),
                    "Exported At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }

                self.status_var.set("Exporting Excel…")
                self.progress_var.set(0)
                export_excel(outpath, original, revised, missing, extra, summary)

                # store for extract tab
                self.last_missing = missing
                self.last_extra = extra
                self.last_sources_by_key = sources_by_key

                self.missing_count_label.configure(text=f"Missing stems loaded: {len(missing)}")
                self.extract_btn.configure(state="normal" if missing else "disabled")

                self.progress_var.set(100)
                self.status_var.set(f"Done. Missing: {len(missing)} | Extra PDFs: {len(extra)}")
                messagebox.showinfo(
                    "Success",
                    f"Exported:\n{outpath}\n\nMissing searchable PDFs: {len(missing)}\nExtra PDFs in revised: {len(extra)}"
                )

            except Exception as e:
                messagebox.showerror("Error", str(e))
                self.status_var.set("Error occurred.")
            finally:
                self.set_busy(False)

        threading.Thread(target=worker, daemon=True).start()

    # -----------------------------
    # Extract sources
    # -----------------------------
    def run_extract(self):
        if not self.last_missing or not self.last_sources_by_key:
            messagebox.showwarning("Nothing to extract", "No missing list loaded. Run Compare first.")
            return

        original = self.original_var.get().strip()
        if not original or not os.path.isdir(original):
            messagebox.showerror("Missing info", "Original folder is not set/valid.")
            return

        dest = self.extract_dest_var.get().strip()
        if not dest:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = os.path.join(os.path.expanduser("~"), "Desktop", f"Missing_Sources_{ts}")
            self.extract_dest_var.set(dest)
        ensure_dir(dest)

        self.stop_flag = False
        self.set_busy(True)
        self.progress_var.set(0)
        self.status_var.set("Copying missing sources…")

        def worker():
            try:
                res = copy_sources_for_missing(
                    missing_rows=self.last_missing,
                    sources_by_key=self.last_sources_by_key,
                    original_root=original,
                    dest_root=dest,
                    progress_cb=self._progress_cb,
                    stop_flag_fn=lambda: self.stop_flag
                )

                if res["status"] == "cancelled":
                    self.status_var.set("Extraction cancelled.")
                    return

                self.progress_var.set(100)
                self.status_var.set(f"Extraction complete → {dest}")
                messagebox.showinfo(
                    "Done",
                    f"Copied {res['copied']}/{res['total']} files.\nFailed: {res['failed']}\n\nDestination:\n{dest}"
                )
            except Exception as e:
                messagebox.showerror("Error", str(e))
                self.status_var.set("Error occurred.")
            finally:
                self.set_busy(False)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = SearchablePdfCompareApp()
    app.mainloop()
