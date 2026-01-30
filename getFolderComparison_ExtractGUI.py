import os
import threading
import shutil
from datetime import datetime

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


def should_skip_file(name: str, skip_hidden: bool, skip_extensions: set[str]) -> bool:
    base = os.path.basename(name)
    if skip_hidden and base.startswith("."):
        return True
    if base.startswith("~$"):  # Office temp file
        return True
    ext = os.path.splitext(base)[1].lower()
    return ext in skip_extensions


def scan_folder(root: str, skip_hidden: bool, skip_extensions: set[str], progress_cb=None):
    """
    Builds a dict keyed by RELATIVE PATH (includes filename), so matching is by filepath+filename.
    """
    root = os.path.normpath(root)
    results = {}

    total_files = 0
    for _, _, files in os.walk(root):
        total_files += len(files)

    seen = 0
    for dirpath, _, filenames in os.walk(root):
        for fn in filenames:
            full = os.path.join(dirpath, fn)

            seen += 1
            if progress_cb:
                progress_cb(seen, total_files, full)

            if should_skip_file(full, skip_hidden, skip_extensions):
                continue

            try:
                st = os.stat(full)
            except OSError:
                st = None

            rel = safe_relpath(full, root)
            key = rel  # EXACT relative path match (filepath + filename)

            ext = os.path.splitext(fn)[1].lower()

            meta = {
                "File Name": fn,
                "Extension": ext,
                "Relative Path": rel,
                "Full Path": os.path.normpath(full),
                "Parent Folder": os.path.basename(os.path.dirname(full)),
                "Size (Bytes)": st.st_size if st else None,
                "Size (MB)": bytes_to_mb(st.st_size) if st else None,
                "Size (GB)": bytes_to_gb(st.st_size) if st else None,
                "Created": ts_to_str(st.st_ctime) if st else None,
                "Modified": ts_to_str(st.st_mtime) if st else None,
                "Accessed": ts_to_str(st.st_atime) if st else None,
            }

            results[key] = meta

    return results


def export_excel(missing_rows, extra_rows, out_path: str, original_root: str, revised_root: str):
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        summary = pd.DataFrame([{
            "Original Folder": original_root,
            "Revised Folder": revised_root,
            "Missing Count (in Revised)": len(missing_rows),
            "Extra Count (only in Revised)": len(extra_rows),
            "Exported At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        summary.to_excel(writer, sheet_name="Summary", index=False)

        df_missing = pd.DataFrame(missing_rows)
        df_missing.to_excel(writer, sheet_name="Missing_in_Revised", index=False)

        df_extra = pd.DataFrame(extra_rows)
        df_extra.to_excel(writer, sheet_name="Extra_in_Revised", index=False)


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def copy_missing_files(missing_rows: list[dict], original_root: str, dest_root: str, progress_cb=None, stop_flag_fn=None):
    """
    Copies missing files FROM original_root into dest_root, keeping Relative Path structure.
    """
    total = len(missing_rows)
    for i, row in enumerate(missing_rows, start=1):
        if stop_flag_fn and stop_flag_fn():
            return "cancelled"

        rel = row.get("Relative Path")
        if not rel:
            continue

        src = os.path.join(original_root, rel.replace("/", os.sep))
        dst = os.path.join(dest_root, rel.replace("/", os.sep))
        dst_parent = os.path.dirname(dst)
        ensure_dir(dst_parent)

        try:
            shutil.copy2(src, dst)  # keeps modified time etc
        except Exception:
            # If a file fails, keep going — you’ll see it in the log/status
            pass

        if progress_cb:
            progress_cb(i, total, dst)

    return "done"


# -----------------------------
# GUI App (2 tabs)
# -----------------------------
class FolderCompareApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Folder Compare + Extract Missing")
        self.geometry("900x460")
        self.resizable(False, False)

        # Inputs
        self.original_var = tk.StringVar()
        self.revised_var = tk.StringVar()
        self.outfile_var = tk.StringVar()

        # Extract destination
        self.extract_dest_var = tk.StringVar()

        # Options
        self.skip_hidden_var = tk.BooleanVar(value=True)
        self.skip_ext_var = tk.StringVar(value=".tmp,.log,.bak")

        # Status/progress
        self.status_var = tk.StringVar(value="Pick folders to compare.")
        self.progress_var = tk.DoubleVar(value=0)

        # In-memory results from last run
        self.last_missing_rows = []
        self.last_extra_rows = []

        # threading control
        self.stop_flag = False

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        root = ttk.Frame(self)
        root.pack(fill="both", expand=True, padx=12, pady=12)

        # Notebook tabs
        self.nb = ttk.Notebook(root)
        self.nb.pack(fill="both", expand=True)

        self.tab_compare = ttk.Frame(self.nb)
        self.tab_extract = ttk.Frame(self.nb)

        self.nb.add(self.tab_compare, text="1) Compare")
        self.nb.add(self.tab_extract, text="2) Extract Missing")

        # --- Compare tab UI ---
        frm = self.tab_compare

        ttk.Label(frm, text="Original folder").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.original_var, width=82).grid(row=0, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_original).grid(row=0, column=2, **pad)

        ttk.Label(frm, text="Revised folder").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.revised_var, width=82).grid(row=1, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_revised).grid(row=1, column=2, **pad)

        ttk.Label(frm, text="Output Excel (.xlsx)").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.outfile_var, width=82).grid(row=2, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Save as…", command=self.pick_outfile).grid(row=2, column=2, **pad)

        opt = ttk.LabelFrame(frm, text="Options")
        opt.grid(row=3, column=0, columnspan=3, sticky="we", padx=10, pady=10)

        ttk.Checkbutton(opt, text="Skip hidden dotfiles", variable=self.skip_hidden_var)\
            .grid(row=0, column=0, sticky="w", padx=10, pady=6)

        ttk.Label(opt, text="Skip extensions (comma-separated):").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        ttk.Entry(opt, textvariable=self.skip_ext_var, width=45).grid(row=1, column=1, sticky="w", padx=10, pady=6)

        self.compare_btn = ttk.Button(frm, text="Run comparison", command=self.run_compare)
        self.compare_btn.grid(row=4, column=2, sticky="e", padx=10, pady=10)

        # --- Extract tab UI ---
        frm2 = self.tab_extract

        info = (
            "This tab uses the most recent comparison results.\n"
            "It will copy the missing files FROM the Original folder into a new folder,\n"
            "keeping the same folder structure (Relative Path)."
        )
        ttk.Label(frm2, text=info, justify="left").grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=10)

        ttk.Label(frm2, text="Destination folder (Missing Docs)").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm2, textvariable=self.extract_dest_var, width=82).grid(row=1, column=1, sticky="w", **pad)
        ttk.Button(frm2, text="Browse…", command=self.pick_extract_dest).grid(row=1, column=2, **pad)

        self.extract_btn = ttk.Button(frm2, text="Extract missing files", command=self.run_extract, state="disabled")
        self.extract_btn.grid(row=2, column=2, sticky="e", padx=10, pady=10)

        self.missing_count_label = ttk.Label(frm2, text="Missing files loaded: 0")
        self.missing_count_label.grid(row=2, column=0, columnspan=2, sticky="w", padx=10, pady=10)

        # --- Shared progress + status (bottom) ---
        self.progress = ttk.Progressbar(root, variable=self.progress_var, maximum=100)
        self.progress.pack(fill="x", padx=10, pady=6)

        ttk.Label(root, textvariable=self.status_var).pack(fill="x", padx=10, pady=6)

        # Cancel button
        self.cancel_btn = ttk.Button(root, text="Cancel current job", command=self.cancel, state="disabled")
        self.cancel_btn.pack(anchor="e", padx=10, pady=6)

    # -----------------------------
    # Browse actions
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

    def _maybe_autofill_outfile(self):
        if self.outfile_var.get().strip():
            return
        o = self.original_var.get().strip()
        r = self.revised_var.get().strip()
        if o and r:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            default = os.path.join(os.path.expanduser("~"), "Desktop", f"folder_compare_{ts}.xlsx")
            self.outfile_var.set(default)

    # -----------------------------
    # Busy/cancel + progress
    # -----------------------------
    def set_busy(self, busy: bool):
        self.compare_btn.configure(state="disabled" if busy else "normal")
        self.extract_btn.configure(state="disabled" if busy else self.extract_btn["state"])
        self.cancel_btn.configure(state="normal" if busy else "disabled")

    def cancel(self):
        self.stop_flag = True
        self.status_var.set("Cancelling… (finishes current file)")

    def _make_skip_exts(self):
        skip_exts = {e.strip().lower() for e in self.skip_ext_var.get().split(",") if e.strip()}
        return {e if e.startswith(".") else f".{e}" for e in skip_exts}

    def _progress_cb(self, done, total, current_path):
        if total > 0:
            pct = (done / total) * 100
            self.progress_var.set(min(100, pct))
        self.status_var.set(f"{done}/{total} … {os.path.basename(current_path)}")
        self.update_idletasks()

    # -----------------------------
    # Compare
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
                original_map = scan_folder(original, skip_hidden, skip_exts, progress_cb=self._progress_cb)
                if self.stop_flag:
                    return

                self.progress_var.set(0)
                self.status_var.set("Scanning REVISED…")
                revised_map = scan_folder(revised, skip_hidden, skip_exts, progress_cb=self._progress_cb)
                if self.stop_flag:
                    return

                self.status_var.set("Comparing…")
                missing = [meta for k, meta in original_map.items() if k not in revised_map]
                extra = [meta for k, meta in revised_map.items() if k not in original_map]

                self.status_var.set("Exporting Excel…")
                self.progress_var.set(0)
                export_excel(missing, extra, outpath, original, revised)

                self.last_missing_rows = missing
                self.last_extra_rows = extra

                # Update extract tab UI
                self.missing_count_label.configure(text=f"Missing files loaded: {len(missing)}")
                if len(missing) > 0:
                    self.extract_btn.configure(state="normal")
                else:
                    self.extract_btn.configure(state="disabled")

                self.progress_var.set(100)
                self.status_var.set(f"Done. Missing: {len(missing)} | Extra: {len(extra)}")
                messagebox.showinfo(
                    "Success",
                    f"Exported:\n{outpath}\n\nMissing: {len(missing)}\nExtra: {len(extra)}"
                )

            except Exception as e:
                messagebox.showerror("Error", str(e))
                self.status_var.set("Error occurred.")
            finally:
                self.set_busy(False)

        threading.Thread(target=worker, daemon=True).start()

    # -----------------------------
    # Extract missing
    # -----------------------------
    def run_extract(self):
        if not self.last_missing_rows:
            messagebox.showwarning("Nothing to extract", "No missing files are loaded. Run Compare first.")
            return

        original = self.original_var.get().strip()
        if not original or not os.path.isdir(original):
            messagebox.showerror("Missing info", "Original folder is not set/valid. Go to Compare tab and set it.")
            return

        dest = self.extract_dest_var.get().strip()
        if not dest:
            # default: Desktop/Missing_Docs_<timestamp>
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            dest = os.path.join(os.path.expanduser("~"), "Desktop", f"Missing_Docs_{ts}")
            self.extract_dest_var.set(dest)

        ensure_dir(dest)

        self.stop_flag = False
        self.set_busy(True)
        self.progress_var.set(0)
        self.status_var.set("Copying missing files…")

        def worker():
            try:
                result = copy_missing_files(
                    self.last_missing_rows,
                    original_root=original,
                    dest_root=dest,
                    progress_cb=self._progress_cb,
                    stop_flag_fn=lambda: self.stop_flag
                )

                if result == "cancelled":
                    self.status_var.set("Extraction cancelled.")
                    return

                self.progress_var.set(100)
                self.status_var.set(f"Extraction complete → {dest}")
                messagebox.showinfo("Done", f"Copied {len(self.last_missing_rows)} missing files into:\n{dest}")

            except Exception as e:
                messagebox.showerror("Error", str(e))
                self.status_var.set("Error occurred.")
            finally:
                self.set_busy(False)

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    app = FolderCompareApp()
    app.mainloop()
