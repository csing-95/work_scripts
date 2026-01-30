import os
import threading
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# -----------------------------
# Core logic (your existing code)
# -----------------------------
def get_file_metadata(folder_path: str) -> pd.DataFrame:
    data = []

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            filepath = os.path.join(root, file)
            filepath = os.path.normpath(filepath)
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
                    "Accessed Date": datetime.fromtimestamp(stats.st_atime).strftime('%Y-%m-%d %H:%M:%S')
                })
            except Exception as e:
                # In GUI mode we’ll surface errors in the log
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
    # ensure optional Error column exists consistently
    if "Error" not in df.columns:
        df["Error"] = ""
    return df


def build_summary(df: pd.DataFrame) -> pd.DataFrame:
    total_files = len(df)

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
        "Total_Size_Bytes": summary_counts["Total_Size_Bytes"].sum(),
        "Total_Size_KB": round(summary_counts["Total_Size_KB"].sum(), 2),
        "Total_Size_MB": round(summary_counts["Total_Size_Bytes"].sum() / (1024 * 1024), 4),
        "% of Files": 100.00
    }])

    return pd.concat([summary_counts, total_row], ignore_index=True)


def save_outputs(df: pd.DataFrame, summary: pd.DataFrame, out_dir: str, out_name: str, out_format: str) -> list[str]:
    """
    Returns list of saved file paths.
    out_format: 'excel' or 'csv'
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


# -----------------------------
# GUI
# -----------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("File Metadata Scanner")
        self.geometry("780x520")
        self.minsize(740, 480)

        self.folder_var = tk.StringVar()
        self.out_dir_var = tk.StringVar()
        self.out_name_var = tk.StringVar(value="sample_file_metadata")
        self.format_var = tk.StringVar(value="excel")

        self.status_var = tk.StringVar(value="Ready.")
        self.total_files_var = tk.StringVar(value="Total files: -")

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        frm = ttk.Frame(self)
        frm.pack(fill="both", expand=True, padx=12, pady=12)

        # Folder picker
        ttk.Label(frm, text="Folder to scan").grid(row=0, column=0, sticky="w", **pad)
        folder_entry = ttk.Entry(frm, textvariable=self.folder_var)
        folder_entry.grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_folder).grid(row=0, column=2, sticky="ew", **pad)

        # Output directory picker
        ttk.Label(frm, text="Save output in").grid(row=1, column=0, sticky="w", **pad)
        outdir_entry = ttk.Entry(frm, textvariable=self.out_dir_var)
        outdir_entry.grid(row=1, column=1, sticky="ew", **pad)
        ttk.Button(frm, text="Browse…", command=self.pick_out_dir).grid(row=1, column=2, sticky="ew", **pad)

        # Output name
        ttk.Label(frm, text="Output file name").grid(row=2, column=0, sticky="w", **pad)
        name_entry = ttk.Entry(frm, textvariable=self.out_name_var)
        name_entry.grid(row=2, column=1, sticky="ew", **pad)
        ttk.Label(frm, text="(no extension)").grid(row=2, column=2, sticky="w", **pad)

        # Format
        ttk.Label(frm, text="Format").grid(row=3, column=0, sticky="w", **pad)
        fmt_frame = ttk.Frame(frm)
        fmt_frame.grid(row=3, column=1, columnspan=2, sticky="w", **pad)

        ttk.Radiobutton(fmt_frame, text="Excel (.xlsx)", variable=self.format_var, value="excel").pack(side="left", padx=(0, 12))
        ttk.Radiobutton(fmt_frame, text="CSV (two files)", variable=self.format_var, value="csv").pack(side="left")

        # Run button + progress
        self.run_btn = ttk.Button(frm, text="Run scan", command=self.on_run)
        self.run_btn.grid(row=4, column=0, sticky="ew", **pad)

        self.progress = ttk.Progressbar(frm, mode="indeterminate")
        self.progress.grid(row=4, column=1, columnspan=2, sticky="ew", **pad)

        # Status + total
        ttk.Label(frm, textvariable=self.status_var).grid(row=5, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(frm, textvariable=self.total_files_var).grid(row=6, column=0, columnspan=3, sticky="w", **pad)

        # Log output
        ttk.Label(frm, text="Log").grid(row=7, column=0, sticky="w", **pad)
        self.log = tk.Text(frm, height=14, wrap="word")
        self.log.grid(row=8, column=0, columnspan=3, sticky="nsew", padx=10, pady=(0, 10))

        # Make columns/rows expand nicely
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(8, weight=1)

        # Nice defaults (optional)
        if not self.folder_var.get():
            self.folder_var.set(os.path.expanduser("~"))
        if not self.out_dir_var.get():
            self.out_dir_var.set(os.path.expanduser("~"))

    def pick_folder(self):
        path = filedialog.askdirectory(title="Select folder to scan")
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

        if not folder or not os.path.isdir(folder):
            messagebox.showerror("Missing/Invalid folder", "Please choose a valid folder to scan.")
            return
        if not out_dir:
            messagebox.showerror("Missing output folder", "Please choose where to save the output.")
            return
        if not out_name:
            messagebox.showerror("Missing output name", "Please enter an output file name.")
            return

        # Clear log + start worker thread
        self.log.delete("1.0", "end")
        self.status_var.set("Running scan…")
        self.total_files_var.set("Total files: -")
        self.set_running(True)

        worker = threading.Thread(
            target=self._run_worker,
            args=(folder, out_dir, out_name, out_format),
            daemon=True
        )
        worker.start()

    def _run_worker(self, folder, out_dir, out_name, out_format):
        try:
            self._ui(lambda: self.log_line(f"Scanning: {folder}"))
            df = get_file_metadata(folder)

            if df.empty:
                self._ui(lambda: self.status_var.set("No files found."))
                self._ui(lambda: self.log_line("No files found in the selected folder."))
                return

            summary = build_summary(df)

            self._ui(lambda: self.log_line("Building summary…"))
            saved = save_outputs(df, summary, out_dir, out_name, out_format)

            self._ui(lambda: self.total_files_var.set(f"Total files: {len(df)}"))
            self._ui(lambda: self.status_var.set("Done ✅"))
            self._ui(lambda: self.log_line(""))
            self._ui(lambda: self.log_line("Saved files:"))
            for p in saved:
                self._ui(lambda p=p: self.log_line(f" - {p}"))

            # quick extra info
            errors = (df["File Extension"] == "ERROR").sum() if "File Extension" in df.columns else 0
            if errors:
                self._ui(lambda: self.log_line(f"\nNote: {errors} file(s) had errors. Check the 'Error' column in output."))

        except Exception as e:
            self._ui(lambda: self.status_var.set("Failed ❌"))
            self._ui(lambda: self.log_line(f"ERROR: {e}"))
            self._ui(lambda: messagebox.showerror("Error", str(e)))
        finally:
            self._ui(lambda: self.set_running(False))

    def _ui(self, fn):
        # schedule fn on main UI thread
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
