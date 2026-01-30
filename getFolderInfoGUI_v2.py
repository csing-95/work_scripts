import os
import re
import threading
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ============================================================
# Helpers
# ============================================================

def safe_filename(name: str, max_len: int = 120) -> str:
    """
    Make a Windows-safe-ish filename (no \ / : * ? " < > | etc).
    """
    name = (name or "").strip().rstrip(".")
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    if not name:
        name = "folder"
    return name[:max_len]


def list_immediate_subfolders(parent_folder: str) -> list[str]:
    """
    Return full paths of immediate subfolders (one level deep only).
    """
    subfolders = []
    with os.scandir(parent_folder) as it:
        for entry in it:
            if entry.is_dir():
                subfolders.append(entry.path)
    return sorted(subfolders, key=lambda p: os.path.basename(p).lower())


# ============================================================
# Core logic
# ============================================================

def get_file_metadata(folder_path: str) -> pd.DataFrame:
    data = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            filepath = os.path.normpath(os.path.join(root, file))
            try:
                stats = os.stat(filepath)
                _, file_ext = os.path.splitext(file)

                data.append({
                    "File Path": filepath,
                    "File Name": file,
                    "File Extension": file_ext.lower() if file_ext else "No Extension",
                    "File Size (Bytes)": stats.st_size,
                    "File Size (KB)": round(stats.st_size / 1024, 2),
                    "File Size (MB)": round(stats.st_size / (1024 * 1024), 4),
                    "Created Date": datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                    "Modified Date": datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                    "Accessed Date": datetime.fromtimestamp(stats.st_atime).strftime('%Y-%m-%d %H:%M:%S'),
                    "Error": ""
                })
            except Exception as e:
                data.append({
                    "File Path": filepath,
                    "File Name": file,
                    "File Extension": "ERROR",
                    "File Size (Bytes)": 0,
                    "File Size (KB)": 0,
                    "File Size (MB)": 0,
                    "Created Date": "",
                    "Modified Date": "",
                    "Accessed Date": "",
                    "Error": str(e)
                })

    df = pd.DataFrame(data)

    # Keep columns stable even if empty
    expected_cols = [
        "File Path", "File Name", "File Extension",
        "File Size (Bytes)", "File Size (KB)", "File Size (MB)",
        "Created Date", "Modified Date", "Accessed Date", "Error"
    ]
    if df.empty:
        df = pd.DataFrame(columns=expected_cols)
    else:
        for c in expected_cols:
            if c not in df.columns:
                df[c] = ""

    return df


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    total_files = len(df)

    if total_files == 0:
        return pd.DataFrame([{
            "File Extension": "TOTAL",
            "File_Count": 0,
            "Total_Size_Bytes": 0,
            "Total_Size_KB": 0,
            "Total_Size_MB": 0,
            "% of Files": 100.00
        }])

    summary_counts = (
        df.groupby("File Extension", dropna=False)
          .agg(
              File_Count=("File Extension", "size"),
              Total_Size_Bytes=("File Size (Bytes)", "sum"),
              Total_Size_KB=("File Size (KB)", "sum"),
          )
          .reset_index()
    )

    summary_counts["% of Files"] = (summary_counts["File_Count"] / total_files * 100).round(2)
    summary_counts["Total_Size_KB"] = summary_counts["Total_Size_KB"].round(2)
    summary_counts["Total_Size_MB"] = (summary_counts["Total_Size_Bytes"] / (1024 * 1024)).round(4)
    summary_counts = summary_counts.sort_values(by="File_Count", ascending=False)

    total_row = pd.DataFrame([{
        "File Extension": "TOTAL",
        "File_Count": total_files,
        "Total_Size_Bytes": int(summary_counts["Total_Size_Bytes"].sum()),
        "Total_Size_KB": round(float(summary_counts["Total_Size_KB"].sum()), 2),
        "Total_Size_MB": round(float(summary_counts["Total_Size_Bytes"].sum()) / (1024 * 1024), 4),
        "% of Files": 100.00
    }])

    return pd.concat([summary_counts, total_row], ignore_index=True)


def save_single_output(df: pd.DataFrame, summary: pd.DataFrame, out_dir: str, out_name: str, out_format: str) -> list[str]:
    """
    Single-folder mode output (one Excel OR two CSVs).
    Returns list of saved file paths.
    """
    os.makedirs(out_dir, exist_ok=True)
    saved = []
    base = os.path.join(out_dir, out_name)

    if out_format == "csv":
        meta_csv = base + "_metadata.csv"
        summary_csv = base + "_summary.csv"
        df.to_csv(meta_csv, index=False)
        summary.to_csv(summary_csv, index=False)
        saved.extend([meta_csv, summary_csv])
    else:
        xlsx = base + ".xlsx"
        with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="File Metadata", index=False)
            summary.to_excel(writer, sheet_name="Summary", index=False)
        saved.append(xlsx)

    return saved


def save_batch_excel(df: pd.DataFrame, summary: pd.DataFrame, out_dir: str, batch_folder_name: str) -> str:
    """
    Batch-mode output: always one Excel per immediate subfolder.
    Returns the saved xlsx path.
    """
    os.makedirs(out_dir, exist_ok=True)
    filename = safe_filename(batch_folder_name) + ".xlsx"
    xlsx = os.path.join(out_dir, filename)

    with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="File Metadata", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)

    return xlsx


# ============================================================
# GUI
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("File Metadata Scanner")
        self.geometry("820x560")
        self.minsize(760, 520)

        self.folder_var = tk.StringVar()
        self.out_dir_var = tk.StringVar()
        self.out_name_var = tk.StringVar(value="sample_file_metadata")
        self.format_var = tk.StringVar(value="excel")

        # NEW: scan mode
        # "single" = scan selected folder into one output
        # "batch"  = scan immediate subfolders into one Excel per subfolder
        self.scan_mode_var = tk.StringVar(value="single")

        self.status_var = tk.StringVar(value="Ready.")
        self.total_files_var = tk.StringVar(value="Total files: -")

        self._build_ui()
        self._apply_mode_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        # Scan mode
        ttk.Label(frm, text="Scan mode").grid(row=0, column=0, sticky="w", **pad)

        mode_frame = ttk.Frame(frm)
        mode_frame.grid(row=0, column=1, columnspan=2, sticky="w", **pad)

        ttk.Radiobutton(
            mode_frame,
            text="Single folder → one output",
            variable=self.scan_mode_var,
            value="single",
            command=self._apply_mode_ui
        ).pack(side="left", padx=(0, 16))

        ttk.Radiobutton(
            mode_frame,
            text="Immediate subfolders → one Excel per subfolder",
            variable=self.scan_mode_var,
            value="batch",
            command=self._apply_mode_ui
        ).pack(side="left")

        # Folder picker (label changes with mode)
        self.folder_label = ttk.Label(frm, text="Folder to scan")
        self.folder_label.grid(row=1, column=0, sticky="w", **pad)

        folder_entry = ttk.Entry(frm, textvariable=self.folder_var)
        folder_entry.grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_folder).grid(row=1, column=2, sticky="ew", **pad)

        # Output directory picker
        ttk.Label(frm, text="Save output in").grid(row=2, column=0, sticky="w", **pad)

        outdir_entry = ttk.Entry(frm, textvariable=self.out_dir_var)
        outdir_entry.grid(row=2, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_out_dir).grid(row=2, column=2, sticky="ew", **pad)

        # Output name (single mode only)
        self.out_name_label = ttk.Label(frm, text="Output file name")
        self.out_name_label.grid(row=3, column=0, sticky="w", **pad)

        self.name_entry = ttk.Entry(frm, textvariable=self.out_name_var)
        self.name_entry.grid(row=3, column=1, sticky="ew", **pad)

        self.name_hint = ttk.Label(frm, text="(no extension)")
        self.name_hint.grid(row=3, column=2, sticky="w", **pad)

        # Format (single mode supports Excel or CSV; batch mode forced to Excel)
        ttk.Label(frm, text="Format").grid(row=4, column=0, sticky="w", **pad)

        fmt_frame = ttk.Frame(frm)
        fmt_frame.grid(row=4, column=1, columnspan=2, sticky="w", **pad)

        self.rb_excel = ttk.Radiobutton(fmt_frame, text="Excel (.xlsx)", variable=self.format_var, value="excel")
        self.rb_excel.pack(side="left", padx=(0, 12))

        self.rb_csv = ttk.Radiobutton(fmt_frame, text="CSV (two files)", variable=self.format_var, value="csv")
        self.rb_csv.pack(side="left")

        # Run button + progress
        self.run_btn = ttk.Button(frm, text="Run scan", command=self.on_run)
        self.run_btn.grid(row=5, column=0, sticky="ew", **pad)

        self.progress = ttk.Progressbar(frm, mode="indeterminate")
        self.progress.grid(row=5, column=1, columnspan=2, sticky="ew", **pad)

        # Status + total
        ttk.Label(frm, textvariable=self.status_var).grid(row=6, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(frm, textvariable=self.total_files_var).grid(row=7, column=0, columnspan=3, sticky="w", **pad)

        # Log output
        ttk.Label(frm, text="Log").grid(row=8, column=0, sticky="w", **pad)
        self.log = tk.Text(frm, height=16, wrap="word")
        self.log.grid(row=9, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))

        # Layout expand
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(9, weight=1)

        # defaults
        if not self.folder_var.get():
            self.folder_var.set(os.path.expanduser("~"))
        if not self.out_dir_var.get():
            self.out_dir_var.set(os.path.expanduser("~"))

    def _apply_mode_ui(self):
        mode = self.scan_mode_var.get()

        if mode == "single":
            self.folder_label.config(text="Folder to scan")
            self.out_name_label.config(state="normal")
            self.name_entry.config(state="normal")
            self.name_hint.config(state="normal")

            # allow CSV in single mode
            self.rb_csv.config(state="normal")

        else:
            # batch mode
            self.folder_label.config(text="Parent folder (contains Batch folders)")
            # output name not used in batch mode (files are named after each batch folder)
            self.out_name_label.config(state="disabled")
            self.name_entry.config(state="disabled")
            self.name_hint.config(state="disabled")

            # force Excel for batch mode
            self.format_var.set("excel")
            self.rb_csv.config(state="disabled")

    def pick_folder(self):
        title = "Select folder to scan"
        if self.scan_mode_var.get() == "batch":
            title = "Select PARENT folder (contains Batch folders)"
        path = filedialog.askdirectory(title=title)
        if path:
            self.folder_var.set(path)

    def pick_out_dir(self):
        path = filedialog.askdirectory(title="Select output folder")
        if path:
            self.out_dir_var.set(path)

    def log_line(self, msg: str):
        self.log.insert("end", msg + "\n")
        self.log.see("end")

    def set_running(self, running: bool):
        if running:
            self.run_btn.config(state="disabled")
            self.progress.start(12)
        else:
            self.progress.stop()
            self.run_btn.config(state="normal")

    def on_run(self):
        folder = self.folder_var.get().strip()
        out_dir = self.out_dir_var.get().strip()
        out_name = self.out_name_var.get().strip()
        out_format = self.format_var.get().strip()
        mode = self.scan_mode_var.get().strip()

        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Missing/Invalid folder", "Please choose a valid folder.")
            return
        if not out_dir:
            messagebox.showerror("Missing output folder", "Please choose where to save the output.")
            return
        if mode == "single" and not out_name:
            messagebox.showerror("Missing output name", "Please enter an output file name.")
            return

        # Clear log + start worker thread
        self.log.delete("1.0", "end")
        self.status_var.set("Running scan…")
        self.total_files_var.set("Total files: -")
        self.set_running(True)

        worker = threading.Thread(
            target=self._run_worker,
            args=(mode, folder, out_dir, out_name, out_format),
            daemon=True
        )
        worker.start()

    def _run_worker(self, mode, folder, out_dir, out_name, out_format):
        try:
            if mode == "single":
                self._ui(lambda: self.log_line(f"Mode: Single folder"))
                self._ui(lambda: self.log_line(f"Scanning: {folder}"))

                df = get_file_metadata(folder)
                summary = build_summary(df)

                self._ui(lambda: self.log_line("Saving output…"))
                saved = save_single_output(df, summary, out_dir, out_name, out_format)

                self._ui(lambda: self.total_files_var.set(f"Total files: {len(df)}"))
                self._ui(lambda: self.status_var.set("Done ✅"))
                self._ui(lambda: self.log_line(""))
                self._ui(lambda: self.log_line("Saved files:"))
                for p in saved:
                    self._ui(lambda p=p: self.log_line(f" - {p}"))

                errors = (df["File Extension"] == "ERROR").sum() if "File Extension" in df.columns else 0
                if errors:
                    self._ui(lambda: self.log_line(f"\nNote: {errors} file(s) had errors. Check the 'Error' column."))

            else:
                # batch mode: immediate subfolders only, one Excel per subfolder
                self._ui(lambda: self.log_line("Mode: Immediate subfolders (Batch mode)"))
                self._ui(lambda: self.log_line(f"Parent folder: {folder}"))
                self._ui(lambda: self.log_line("Finding immediate subfolders…"))

                subfolders = list_immediate_subfolders(folder)

                if not subfolders:
                    self._ui(lambda: self.status_var.set("No subfolders found."))
                    self._ui(lambda: self.log_line("No immediate subfolders found under the selected parent folder."))
                    self._ui(lambda: self.total_files_var.set("Total files: -"))
                    return

                self._ui(lambda: self.log_line(f"Found {len(subfolders)} subfolder(s)."))
                self._ui(lambda: self.log_line(""))

                total_files_all = 0
                saved_paths = []

                for idx, sub in enumerate(subfolders, start=1):
                    batch_name = os.path.basename(sub)
                    self._ui(lambda idx=idx, n=len(subfolders), bn=batch_name:
                             self.status_var.set(f"Scanning {idx}/{n}: {bn} …"))
                    self._ui(lambda bn=batch_name: self.log_line(f"[{bn}] Scanning…"))

                    df = get_file_metadata(sub)
                    summary = build_summary(df)

                    xlsx = save_batch_excel(df, summary, out_dir, batch_name)
                    saved_paths.append(xlsx)

                    total_files_all += len(df)

                    errors = (df["File Extension"] == "ERROR").sum() if "File Extension" in df.columns else 0
                    self._ui(lambda x=xlsx: self.log_line(f"[{batch_name}] Saved: {x}"))
                    if errors:
                        self._ui(lambda e=errors: self.log_line(f"[{batch_name}] Note: {e} file(s) had errors."))

                    self._ui(lambda: self.log_line(""))

                self._ui(lambda: self.total_files_var.set(f"Total files (all batches): {total_files_all}"))
                self._ui(lambda: self.status_var.set("Done ✅"))
                self._ui(lambda: self.log_line("Finished. Saved Excel files:"))
                for p in saved_paths:
                    self._ui(lambda p=p: self.log_line(f" - {p}"))

        except Exception as e:
            self._ui(lambda: self.status_var.set("Failed ❌"))
            self._ui(lambda: self.log_line(f"ERROR: {e}"))
            self._ui(lambda: messagebox.showerror("Error", str(e)))
        finally:
            self._ui(lambda: self.set_running(False))

    def _ui(self, fn):
        self.after(0, fn)


if __name__ == "__main__":
    # Windows: makes it look a bit nicer if available
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = App()
    app.mainloop()
