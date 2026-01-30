import os
import threading
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


def count_files(folder: str, recursive: bool = True) -> int:
    if recursive:
        total = 0
        for _, _, files in os.walk(folder):
            total += len(files)
        return total
    else:
        # only files directly inside folder (not inside nested subfolders)
        try:
            return sum(
                1 for entry in os.scandir(folder)
                if entry.is_file(follow_symlinks=False)
            )
        except PermissionError:
            return 0


def get_immediate_subfolders(root_folder: str):
    subfolders = []
    try:
        with os.scandir(root_folder) as it:
            for entry in it:
                if entry.is_dir(follow_symlinks=False):
                    subfolders.append(entry.path)
    except PermissionError:
        return []
    return sorted(subfolders, key=lambda p: os.path.basename(p).lower())


class FolderCountApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Root Folder File Counter")
        self.geometry("820x520")
        self.minsize(820, 520)

        self.root_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.recursive_var = tk.BooleanVar(value=True)

        self._build_ui()

        self._worker_thread = None
        self._stop_flag = False

    def _build_ui(self):
        pad = 10

        header = ttk.Frame(self)
        header.pack(fill="x", padx=pad, pady=(pad, 0))

        title = ttk.Label(header, text="Folder file counts (root → subfolders)", font=("Segoe UI", 14, "bold"))
        title.pack(anchor="w")

        subtitle = ttk.Label(
            header,
            text="Select a root folder. This will count files for each immediate subfolder and export to Excel.",
            font=("Segoe UI", 10),
        )
        subtitle.pack(anchor="w", pady=(4, 0))

        # Root folder selection
        root_frame = ttk.LabelFrame(self, text="1) Choose root folder")
        root_frame.pack(fill="x", padx=pad, pady=(pad, 0))

        row = ttk.Frame(root_frame)
        row.pack(fill="x", padx=pad, pady=pad)

        root_entry = ttk.Entry(row, textvariable=self.root_path_var)
        root_entry.pack(side="left", fill="x", expand=True)

        browse_btn = ttk.Button(row, text="Browse…", command=self.choose_root_folder)
        browse_btn.pack(side="left", padx=(8, 0))

        # Options
        opt_frame = ttk.LabelFrame(self, text="2) Options")
        opt_frame.pack(fill="x", padx=pad, pady=(pad, 0))

        opt_row = ttk.Frame(opt_frame)
        opt_row.pack(fill="x", padx=pad, pady=pad)

        recursive_cb = ttk.Checkbutton(
            opt_row,
            text="Count files recursively inside each subfolder (recommended)",
            variable=self.recursive_var
        )
        recursive_cb.pack(side="left")

        # Output selection
        out_frame = ttk.LabelFrame(self, text="3) Output")
        out_frame.pack(fill="x", padx=pad, pady=(pad, 0))

        out_row = ttk.Frame(out_frame)
        out_row.pack(fill="x", padx=pad, pady=pad)

        out_entry = ttk.Entry(out_row, textvariable=self.output_path_var)
        out_entry.pack(side="left", fill="x", expand=True)

        out_btn = ttk.Button(out_row, text="Choose output…", command=self.choose_output_file)
        out_btn.pack(side="left", padx=(8, 0))

        # Actions
        action_frame = ttk.Frame(self)
        action_frame.pack(fill="x", padx=pad, pady=(pad, 0))

        self.run_btn = ttk.Button(action_frame, text="Run scan + export", command=self.run_scan)
        self.run_btn.pack(side="left")

        self.stop_btn = ttk.Button(action_frame, text="Stop", command=self.stop_scan, state="disabled")
        self.stop_btn.pack(side="left", padx=(8, 0))

        # Progress
        prog_frame = ttk.Frame(self)
        prog_frame.pack(fill="x", padx=pad, pady=(10, 0))

        self.progress = ttk.Progressbar(prog_frame, mode="determinate")
        self.progress.pack(fill="x")

        self.status_var = tk.StringVar(value="Ready.")
        status_lbl = ttk.Label(self, textvariable=self.status_var)
        status_lbl.pack(fill="x", padx=pad, pady=(6, 0))

        # Preview table
        table_frame = ttk.LabelFrame(self, text="Preview")
        table_frame.pack(fill="both", expand=True, padx=pad, pady=pad)

        cols = ("Folder", "File Count")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        self.tree.heading("Folder", text="Folder")
        self.tree.heading("File Count", text="File Count")
        self.tree.column("Folder", width=620, anchor="w")
        self.tree.column("File Count", width=120, anchor="e")

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)

        self.tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

    def choose_root_folder(self):
        path = filedialog.askdirectory(title="Select root folder")
        if path:
            self.root_path_var.set(path)
            # Auto-suggest output path
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            suggested = os.path.join(path, f"folder_file_counts_{stamp}.xlsx")
            self.output_path_var.set(suggested)

    def choose_output_file(self):
        path = filedialog.asksaveasfilename(
            title="Save Excel output",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if path:
            self.output_path_var.set(path)

    def run_scan(self):
        root_folder = self.root_path_var.get().strip()
        out_file = self.output_path_var.get().strip()

        if not root_folder or not os.path.isdir(root_folder):
            messagebox.showerror("Missing root folder", "Please choose a valid root folder.")
            return

        if not out_file:
            messagebox.showerror("Missing output file", "Please choose an output .xlsx file path.")
            return

        # Reset UI
        self._stop_flag = False
        self._clear_tree()
        self.progress["value"] = 0
        self.status_var.set("Scanning…")

        self.run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")

        self._worker_thread = threading.Thread(
            target=self._scan_worker,
            args=(root_folder, out_file, self.recursive_var.get()),
            daemon=True
        )
        self._worker_thread.start()

    def stop_scan(self):
        self._stop_flag = True
        self.status_var.set("Stopping… (will stop after current folder)")

    def _scan_worker(self, root_folder: str, out_file: str, recursive: bool):
        try:
            subfolders = get_immediate_subfolders(root_folder)
            total = len(subfolders)

            if total == 0:
                self._ui_done_no_folders()
                return

            rows = []
            self._ui_set_progress_max(total)

            for i, folder in enumerate(subfolders, start=1):
                if self._stop_flag:
                    self._ui_stopped(i - 1, total)
                    return

                folder_name = os.path.basename(folder)
                self._ui_status(f"Counting: {folder_name} ({i}/{total})")

                try:
                    fc = count_files(folder, recursive=recursive)
                except Exception:
                    fc = None

                rows.append({
                    "Folder Name": folder_name,
                    "Folder Path": folder,
                    "File Count": fc if fc is not None else "ERROR",
                })

                self._ui_add_row(folder, fc if fc is not None else "ERROR")
                self._ui_progress(i)

            df = pd.DataFrame(rows)
            df.to_excel(out_file, index=False, sheet_name="FolderCounts")

            self._ui_done(out_file, total)

        except Exception as e:
            self._ui_error(str(e))

    # ---------- UI helpers (thread-safe via after) ----------

    def _ui_set_progress_max(self, max_value: int):
        self.after(0, lambda: self.progress.config(maximum=max_value))

    def _ui_progress(self, value: int):
        self.after(0, lambda: self.progress.config(value=value))

    def _ui_status(self, text: str):
        self.after(0, lambda: self.status_var.set(text))

    def _ui_add_row(self, folder_path: str, file_count):
        def add():
            self.tree.insert("", "end", values=(os.path.basename(folder_path), file_count))
        self.after(0, add)

    def _ui_done(self, out_file: str, total: int):
        def done():
            self.status_var.set(f"Done. Scanned {total} folders. Exported: {out_file}")
            self.run_btn.config(state="normal")
            self.stop_btn.config(state="disabled")
            messagebox.showinfo("Finished", f"Export complete!\n\n{out_file}")
        self.after(0, done)

    def _ui_done_no_folders(self):
        def done():
            self.status_var.set("No subfolders found in the chosen root folder.")
            self.run_btn.config(state="normal")
            self.stop_btn.config(state="disabled")
            messagebox.showwarning("No folders", "No immediate subfolders were found in that root folder.")
        self.after(0, done)

    def _ui_stopped(self, scanned: int, total: int):
        def stopped():
            self.status_var.set(f"Stopped. Scanned {scanned}/{total} folders.")
            self.run_btn.config(state="normal")
            self.stop_btn.config(state="disabled")
        self.after(0, stopped)

    def _ui_error(self, msg: str):
        def err():
            self.status_var.set("Error occurred.")
            self.run_btn.config(state="normal")
            self.stop_btn.config(state="disabled")
            messagebox.showerror("Error", msg)
        self.after(0, err)

    def _clear_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)


if __name__ == "__main__":
    # slightly nicer default theme where available
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = FolderCountApp()
    app.mainloop()
