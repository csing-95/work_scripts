import os
import threading
import time
from datetime import datetime
from pathlib import Path

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
    # local time
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M:%S")


def safe_relpath(full_path: str, root: str) -> str:
    # Always use forward slashes in keys to be consistent across OS weirdness
    rel = os.path.relpath(full_path, root)
    rel = rel.replace("\\", "/")
    return rel


def should_skip_file(name: str, skip_hidden: bool, skip_extensions: set[str]) -> bool:
    base = os.path.basename(name)
    if skip_hidden and base.startswith("."):
        return True
    ext = os.path.splitext(base)[1].lower()
    if ext in skip_extensions:
        return True
    # common temp/lock patterns
    if base.startswith("~$"):  # office temp file
        return True
    return False


def scan_folder(root: str, skip_hidden: bool, skip_extensions: set[str], progress_cb=None):
    """
    Returns dict: key -> metadata
    key is normalized relative path (casefolded on Windows-ish behavior).
    """
    root = os.path.normpath(root)
    results = {}

    # Count files first (for nicer progress). If too slow, we still proceed.
    total_files = 0
    for _, _, files in os.walk(root):
        total_files += len(files)

    seen = 0
    for dirpath, _, filenames in os.walk(root):
        for fn in filenames:
            full = os.path.join(dirpath, fn)
            if should_skip_file(full, skip_hidden, skip_extensions):
                seen += 1
                if progress_cb:
                    progress_cb(seen, total_files, full)
                continue

            try:
                st = os.stat(full)
            except OSError:
                # Can't read stats — still record basic info
                st = None

            rel = safe_relpath(full, root)
            key = rel.casefold()  # makes matching case-insensitive (good for Windows shares)

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

            seen += 1
            if progress_cb:
                progress_cb(seen, total_files, full)

    return results


def export_excel(missing_rows, extra_rows, out_path: str, original_root: str, revised_root: str):
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Summary sheet
        summary = pd.DataFrame([{
            "Original Folder": original_root,
            "Revised Folder": revised_root,
            "Missing Count (in Revised)": len(missing_rows),
            "Extra Count (only in Revised)": len(extra_rows),
            "Exported At": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        summary.to_excel(writer, sheet_name="Summary", index=False)

        # Missing
        df_missing = pd.DataFrame(missing_rows)
        if df_missing.empty:
            df_missing = pd.DataFrame(columns=[
                "File Name", "Extension", "Relative Path", "Full Path", "Parent Folder",
                "Size (Bytes)", "Size (MB)", "Size (GB)", "Created", "Modified", "Accessed"
            ])
        df_missing.to_excel(writer, sheet_name="Missing_in_Revised", index=False)

        # Extra (optional but useful)
        df_extra = pd.DataFrame(extra_rows)
        if df_extra.empty:
            df_extra = pd.DataFrame(columns=[
                "File Name", "Extension", "Relative Path", "Full Path", "Parent Folder",
                "Size (Bytes)", "Size (MB)", "Size (GB)", "Created", "Modified", "Accessed"
            ])
        df_extra.to_excel(writer, sheet_name="Extra_in_Revised", index=False)


# -----------------------------
# GUI App
# -----------------------------
class FolderCompareApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Folder Compare → Missing Files Excel Export")
        self.geometry("820x360")
        self.resizable(False, False)

        self.original_var = tk.StringVar()
        self.revised_var = tk.StringVar()
        self.outfile_var = tk.StringVar()

        self.skip_hidden_var = tk.BooleanVar(value=True)
        self.skip_ext_var = tk.StringVar(value=".tmp,.log,.bak")  # comma separated

        self.status_var = tk.StringVar(value="Pick folders to compare.")
        self.progress_var = tk.DoubleVar(value=0)

        self._build_ui()

        self.worker_thread = None
        self.stop_flag = False

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        # Original folder
        ttk.Label(frm, text="Original folder").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.original_var, width=80).grid(row=0, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_original).grid(row=0, column=2, **pad)

        # Revised folder
        ttk.Label(frm, text="Revised folder").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.revised_var, width=80).grid(row=1, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_revised).grid(row=1, column=2, **pad)

        # Output file
        ttk.Label(frm, text="Output Excel (.xlsx)").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(frm, textvariable=self.outfile_var, width=80).grid(row=2, column=1, sticky="w", **pad)
        ttk.Button(frm, text="Save as…", command=self.pick_outfile).grid(row=2, column=2, **pad)

        # Options
        opt = ttk.LabelFrame(frm, text="Options")
        opt.grid(row=3, column=0, columnspan=3, sticky="we", padx=10, pady=10)

        ttk.Checkbutton(opt, text="Skip hidden dotfiles (recommended)", variable=self.skip_hidden_var)\
            .grid(row=0, column=0, sticky="w", padx=10, pady=6)

        ttk.Label(opt, text="Skip extensions (comma-separated):").grid(row=1, column=0, sticky="w", padx=10, pady=6)
        ttk.Entry(opt, textvariable=self.skip_ext_var, width=45).grid(row=1, column=1, sticky="w", padx=10, pady=6)

        # Progress + status
        self.progress = ttk.Progressbar(frm, variable=self.progress_var, maximum=100)
        self.progress.grid(row=4, column=0, columnspan=3, sticky="we", padx=10, pady=8)

        ttk.Label(frm, textvariable=self.status_var).grid(row=5, column=0, columnspan=3, sticky="w", padx=10, pady=6)

        # Buttons
        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=3, sticky="e", padx=10, pady=10)

        self.run_btn = ttk.Button(btns, text="Run comparison", command=self.run)
        self.run_btn.pack(side="right", padx=8)

        self.cancel_btn = ttk.Button(btns, text="Cancel", command=self.cancel, state="disabled")
        self.cancel_btn.pack(side="right")

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

    def _maybe_autofill_outfile(self):
        if self.outfile_var.get().strip():
            return
        o = self.original_var.get().strip()
        r = self.revised_var.get().strip()
        if o and r:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            default = os.path.join(os.path.expanduser("~"), "Desktop", f"folder_compare_{ts}.xlsx")
            self.outfile_var.set(default)

    def set_busy(self, busy: bool):
        self.run_btn.configure(state="disabled" if busy else "normal")
        self.cancel_btn.configure(state="normal" if busy else "disabled")

    def cancel(self):
        self.stop_flag = True
        self.status_var.set("Cancelling… (finishes current file)")

    def run(self):
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
        self.status_var.set("Scanning folders…")

        skip_hidden = self.skip_hidden_var.get()
        skip_exts = {e.strip().lower() for e in self.skip_ext_var.get().split(",") if e.strip()}
        # ensure they start with dot
        skip_exts = {e if e.startswith(".") else f".{e}" for e in skip_exts}

        def progress_cb(done, total, current_path):
            if self.stop_flag:
                return
            if total > 0:
                pct = (done / total) * 100
                self.progress_var.set(min(100, pct))
            # keep status short-ish
            self.status_var.set(f"{done}/{total} … {os.path.basename(current_path)}")
            self.update_idletasks()

        def worker():
            try:
                # Scan original
                if self.stop_flag:
                    return
                self.status_var.set("Scanning ORIGINAL folder…")
                original_map = scan_folder(original, skip_hidden, skip_exts, progress_cb=progress_cb)

                # Scan revised
                if self.stop_flag:
                    return
                self.progress_var.set(0)
                self.status_var.set("Scanning REVISED folder…")
                revised_map = scan_folder(revised, skip_hidden, skip_exts, progress_cb=progress_cb)

                if self.stop_flag:
                    return

                # Compare
                missing = []
                for k, meta in original_map.items():
                    if k not in revised_map:
                        missing.append(meta)

                extra = []
                for k, meta in revised_map.items():
                    if k not in original_map:
                        extra.append(meta)

                # Export
                self.status_var.set("Exporting Excel…")
                self.progress_var.set(0)
                export_excel(missing, extra, outpath, original, revised)

                self.progress_var.set(100)
                self.status_var.set(f"Done. Missing: {len(missing)} | Extra: {len(extra)}")
                messagebox.showinfo("Success", f"Exported:\n{outpath}\n\nMissing: {len(missing)}\nExtra: {len(extra)}")

            except Exception as e:
                messagebox.showerror("Error", str(e))
                self.status_var.set("Error occurred.")
            finally:
                self.set_busy(False)

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()


if __name__ == "__main__":
    app = FolderCompareApp()
    app.mainloop()
