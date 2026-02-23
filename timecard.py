import json
import os
import time
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from openpyxl import Workbook
    from openpyxl.styles import numbers
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


APP_NAME = "Task Time Tracker"
DATA_FILE = "time_tracker_data.json"


def now_iso():
    return datetime.now().replace(microsecond=0).isoformat(sep=" ")


def seconds_to_hhmmss(seconds: int) -> str:
    seconds = max(0, int(seconds))
    td = timedelta(seconds=seconds)
    total = int(td.total_seconds())
    h = total // 3600
    m = (total % 3600) // 60
    s = total % 60
    return f"{h:02d}:{m:02d}:{s:02d}"


def seconds_to_hhmm(seconds: int) -> str:
    seconds = max(0, int(seconds))
    h = seconds // 3600
    m = (seconds % 3600) // 60
    return f"{h:02d}:{m:02d}"


def seconds_to_decimal_hours(seconds: int, places: int = 2) -> Decimal:
    # Use Decimal to avoid float weirdness (keeps 0.00 as 0.00 when formatted)
    hrs = (Decimal(seconds) / Decimal(3600))
    q = Decimal("1." + ("0" * places))
    return hrs.quantize(q, rounding=ROUND_HALF_UP)


@dataclass
class Session:
    task: str
    start: str
    end: str
    seconds: int
    note: str = ""


class TimeTrackerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("980x620")

        self.data = {
            "tasks": ["Admin"],
            "sessions": []
        }

        self.current_task = tk.StringVar(value="Admin")
        self.current_note = tk.StringVar(value="")
        self.timer_running = False
        self.timer_start_epoch = None
        self.timer_start_iso = None
        self.timer_elapsed_seconds = 0

        self._build_ui()
        self._load_data()
        self._refresh_task_dropdown()
        self._refresh_tables()
        self._tick()

    # ---------- UI ----------
    def _build_ui(self):
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        top = ttk.LabelFrame(main, text="Current Timer", padding=10)
        top.pack(fill="x")

        row1 = ttk.Frame(top)
        row1.pack(fill="x", pady=(0, 6))

        ttk.Label(row1, text="Task:").pack(side="left")
        self.task_combo = ttk.Combobox(row1, textvariable=self.current_task, state="readonly", width=35)
        self.task_combo.pack(side="left", padx=(8, 12))

        ttk.Button(row1, text="Add Task", command=self.add_task).pack(side="left")
        ttk.Button(row1, text="Rename Task", command=self.rename_task).pack(side="left", padx=6)
        ttk.Button(row1, text="Delete Task", command=self.delete_task).pack(side="left")

        row2 = ttk.Frame(top)
        row2.pack(fill="x", pady=(0, 6))

        ttk.Label(row2, text="Note (optional):").pack(side="left")
        self.note_entry = ttk.Entry(row2, textvariable=self.current_note, width=80)
        self.note_entry.pack(side="left", padx=(8, 0), fill="x", expand=True)

        row3 = ttk.Frame(top)
        row3.pack(fill="x")

        self.time_label = ttk.Label(row3, text="00:00:00", font=("Segoe UI", 22, "bold"))
        self.time_label.pack(side="left")

        self.dec_label = ttk.Label(row3, text="(0.00 hrs)", font=("Segoe UI", 12))
        self.dec_label.pack(side="left", padx=(12, 0))

        btns = ttk.Frame(row3)
        btns.pack(side="right")

        self.start_btn = ttk.Button(btns, text="Start", command=self.start_timer)
        self.start_btn.pack(side="left")

        self.stop_btn = ttk.Button(btns, text="Stop", command=self.stop_timer, state="disabled")
        self.stop_btn.pack(side="left", padx=8)

        ttk.Button(btns, text="Cancel (don’t save)", command=self.cancel_timer).pack(side="left")

        mid = ttk.Frame(main)
        mid.pack(fill="both", expand=True, pady=12)

        left = ttk.LabelFrame(mid, text="Totals by Task", padding=10)
        left.pack(side="left", fill="both", expand=True, padx=(0, 6))

        right = ttk.LabelFrame(mid, text="Session Log", padding=10)
        right.pack(side="left", fill="both", expand=True, padx=(6, 0))

        # Totals table
        self.totals_tree = ttk.Treeview(left, columns=("task", "hhmm", "hhmmss", "decimal"), show="headings", height=12)
        for col, w in [("task", 200), ("hhmm", 90), ("hhmmss", 110), ("decimal", 100)]:
            self.totals_tree.heading(col, text=col.upper())
            self.totals_tree.column(col, width=w, anchor="w")
        self.totals_tree.pack(fill="both", expand=True)

        # Sessions table
        self.sessions_tree = ttk.Treeview(
            right,
            columns=("task", "start", "end", "hhmmss", "decimal", "note"),
            show="headings",
            height=12
        )
        headings = [
            ("task", 120),
            ("start", 160),
            ("end", 160),
            ("hhmmss", 90),
            ("decimal", 90),
            ("note", 260),
        ]
        for col, w in headings:
            self.sessions_tree.heading(col, text=col.upper())
            self.sessions_tree.column(col, width=w, anchor="w")
        self.sessions_tree.pack(fill="both", expand=True)

        # Bottom actions
        bottom = ttk.Frame(main)
        bottom.pack(fill="x")

        ttk.Button(bottom, text="Delete Selected Session", command=self.delete_selected_session).pack(side="left")
        ttk.Button(bottom, text="Clear All Sessions", command=self.clear_all_sessions).pack(side="left", padx=8)

        ttk.Button(bottom, text="Export CSV", command=self.export_csv).pack(side="right")
        ttk.Button(bottom, text="Export Excel (.xlsx)", command=self.export_xlsx).pack(side="right", padx=8)

        self.status = ttk.Label(main, text="", foreground="#444")
        self.status.pack(fill="x", pady=(8, 0))

    # ---------- Timer ----------
    def start_timer(self):
        if self.timer_running:
            return
        task = (self.current_task.get() or "").strip()
        if not task:
            messagebox.showwarning(APP_NAME, "Pick or create a task first.")
            return

        self.timer_running = True
        self.timer_start_epoch = time.time()
        self.timer_start_iso = now_iso()
        self.timer_elapsed_seconds = 0

        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self._set_status(f"Running: {task}")

    def stop_timer(self):
        if not self.timer_running:
            return

        end_iso = now_iso()
        elapsed = int(round(time.time() - self.timer_start_epoch))
        task = self.current_task.get()
        note = self.current_note.get().strip()

        sess = Session(
            task=task,
            start=self.timer_start_iso,
            end=end_iso,
            seconds=elapsed,
            note=note
        )
        self.data["sessions"].append(asdict(sess))
        self._save_data()

        # Reset timer
        self.timer_running = False
        self.timer_start_epoch = None
        self.timer_start_iso = None
        self.timer_elapsed_seconds = 0
        self.current_note.set("")

        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")

        self._refresh_tables()
        self._set_status(f"Saved: {task} — {seconds_to_hhmmss(elapsed)} ({seconds_to_decimal_hours(elapsed)} hrs)")

    def cancel_timer(self):
        if not self.timer_running:
            return
        if not messagebox.askyesno(APP_NAME, "Cancel this timer without saving?"):
            return

        self.timer_running = False
        self.timer_start_epoch = None
        self.timer_start_iso = None
        self.timer_elapsed_seconds = 0

        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self._set_status("Timer cancelled (not saved).")

    def _tick(self):
        if self.timer_running and self.timer_start_epoch is not None:
            self.timer_elapsed_seconds = int(round(time.time() - self.timer_start_epoch))
        self.time_label.config(text=seconds_to_hhmmss(self.timer_elapsed_seconds))
        self.dec_label.config(text=f"({seconds_to_decimal_hours(self.timer_elapsed_seconds)} hrs)")
        self.root.after(250, self._tick)

    # ---------- Tasks ----------
    def _refresh_task_dropdown(self):
        tasks = self.data.get("tasks", [])
        if not tasks:
            tasks = ["Admin"]
            self.data["tasks"] = tasks

        self.task_combo["values"] = tasks
        if self.current_task.get() not in tasks:
            self.current_task.set(tasks[0])

    def add_task(self):
        name = self._prompt("Add Task", "Task name:")
        if name is None:
            return
        name = name.strip()
        if not name:
            return
        if name in self.data["tasks"]:
            messagebox.showinfo(APP_NAME, "That task already exists.")
            return
        self.data["tasks"].append(name)
        self._save_data()
        self._refresh_task_dropdown()
        self.current_task.set(name)
        self._refresh_tables()

    def rename_task(self):
        old = self.current_task.get()
        new = self._prompt("Rename Task", f"Rename '{old}' to:")
        if new is None:
            return
        new = new.strip()
        if not new or new == old:
            return
        if new in self.data["tasks"]:
            messagebox.showwarning(APP_NAME, "That name already exists.")
            return

        # rename in tasks
        self.data["tasks"] = [new if t == old else t for t in self.data["tasks"]]
        # rename in sessions
        for s in self.data["sessions"]:
            if s.get("task") == old:
                s["task"] = new

        self._save_data()
        self._refresh_task_dropdown()
        self.current_task.set(new)
        self._refresh_tables()

    def delete_task(self):
        task = self.current_task.get()
        if task == "Admin":
            messagebox.showinfo(APP_NAME, "You can’t delete the default 'Admin' task.")
            return
        if self.timer_running:
            messagebox.showwarning(APP_NAME, "Stop the timer before deleting a task.")
            return
        if not messagebox.askyesno(APP_NAME, f"Delete task '{task}'?\nSessions will be kept but still labelled '{task}'."):
            return

        self.data["tasks"] = [t for t in self.data["tasks"] if t != task]
        self._save_data()
        self._refresh_task_dropdown()
        self._refresh_tables()

    # ---------- Sessions / tables ----------
    def _refresh_tables(self):
        # clear
        for tv in (self.totals_tree, self.sessions_tree):
            for item in tv.get_children():
                tv.delete(item)

        # totals
        totals = {}
        for s in self.data["sessions"]:
            totals.setdefault(s["task"], 0)
            totals[s["task"]] += int(s.get("seconds", 0))

        # include tasks with zero time too
        for t in self.data.get("tasks", []):
            totals.setdefault(t, 0)

        for task in sorted(totals.keys(), key=str.lower):
            sec = totals[task]
            dec = seconds_to_decimal_hours(sec)
            self.totals_tree.insert("", "end", values=(task, seconds_to_hhmm(sec), seconds_to_hhmmss(sec), f"{dec}"))

        # sessions log (newest first)
        for s in reversed(self.data["sessions"]):
            sec = int(s.get("seconds", 0))
            dec = seconds_to_decimal_hours(sec)
            self.sessions_tree.insert(
                "",
                "end",
                values=(s.get("task", ""), s.get("start", ""), s.get("end", ""), seconds_to_hhmmss(sec), f"{dec}", s.get("note", ""))
            )

    def delete_selected_session(self):
        sel = self.sessions_tree.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "Select a session row first.")
            return
        if not messagebox.askyesno(APP_NAME, "Delete selected session(s)?"):
            return

        # Since we show newest-first, we need to map selected rows back to session dicts reliably.
        # We'll match on a tuple key.
        selected_keys = set()
        for item in sel:
            vals = self.sessions_tree.item(item, "values")
            # (task, start, end, seconds, note)
            selected_keys.add((vals[0], vals[1], vals[2], vals[3], vals[5]))

        kept = []
        for s in self.data["sessions"]:
            sec = int(s.get("seconds", 0))
            key = (s.get("task", ""), s.get("start", ""), s.get("end", ""), seconds_to_hhmmss(sec), s.get("note", ""))
            if key not in selected_keys:
                kept.append(s)

        self.data["sessions"] = kept
        self._save_data()
        self._refresh_tables()
        self._set_status("Deleted selected session(s).")

    def clear_all_sessions(self):
        if self.timer_running:
            messagebox.showwarning(APP_NAME, "Stop the timer before clearing sessions.")
            return
        if not messagebox.askyesno(APP_NAME, "Clear ALL sessions? This can’t be undone."):
            return
        self.data["sessions"] = []
        self._save_data()
        self._refresh_tables()
        self._set_status("All sessions cleared.")

    # ---------- Export ----------
    def export_csv(self):
        path = filedialog.asksaveasfilename(
            title="Export CSV",
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")]
        )
        if not path:
            return

        # Simple CSV writer (no extra deps). Use tab-safe quoting lightly.
        # CSV fields: task, start, end, seconds, hh:mm:ss, decimal_hours, note
        import csv
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["task", "start", "end", "seconds", "hh:mm:ss", "decimal_hours", "note"])
            for s in self.data["sessions"]:
                sec = int(s.get("seconds", 0))
                dec = seconds_to_decimal_hours(sec)
                w.writerow([s.get("task",""), s.get("start",""), s.get("end",""), sec, seconds_to_hhmmss(sec), f"{dec}", s.get("note","")])

        self._set_status(f"Exported CSV: {path}")

    def export_xlsx(self):
        if not HAS_OPENPYXL:
            messagebox.showerror(APP_NAME, "openpyxl isn’t installed.\nRun: pip install openpyxl")
            return

        path = filedialog.asksaveasfilename(
            title="Export Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")]
        )
        if not path:
            return

        wb = Workbook()
        ws1 = wb.active
        ws1.title = "Sessions"

        ws1.append(["task", "start", "end", "seconds", "hh:mm:ss", "decimal_hours", "note"])

        for s in self.data["sessions"]:
            sec = int(s.get("seconds", 0))
            dec = seconds_to_decimal_hours(sec)  # Decimal
            ws1.append([
                s.get("task",""),
                s.get("start",""),
                s.get("end",""),
                sec,
                seconds_to_hhmmss(sec),
                float(dec),  # store as numeric; formatting below ensures 0.00 stays displayed as 0.00
                s.get("note","")
            ])

        # Format decimal_hours column to 2dp
        for cell in ws1["F"][1:]:
            cell.number_format = "0.00"

        # Totals sheet
        ws2 = wb.create_sheet("Totals")
        ws2.append(["task", "total_seconds", "hh:mm", "hh:mm:ss", "decimal_hours"])

        totals = {}
        for s in self.data["sessions"]:
            totals.setdefault(s["task"], 0)
            totals[s["task"]] += int(s.get("seconds", 0))

        for t in self.data.get("tasks", []):
            totals.setdefault(t, 0)

        for task in sorted(totals.keys(), key=str.lower):
            sec = totals[task]
            dec = seconds_to_decimal_hours(sec)
            ws2.append([task, sec, seconds_to_hhmm(sec), seconds_to_hhmmss(sec), float(dec)])

        for cell in ws2["E"][1:]:
            cell.number_format = "0.00"

        wb.save(path)
        self._set_status(f"Exported Excel: {path}")

    # ---------- Persistence ----------
    def _load_data(self):
        if not os.path.exists(DATA_FILE):
            self._save_data()
            return
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                self.data = json.load(f)
            # Basic sanity
            self.data.setdefault("tasks", ["Admin"])
            self.data.setdefault("sessions", [])
        except Exception as e:
            messagebox.showwarning(APP_NAME, f"Couldn’t load data file. Starting fresh.\n\n{e}")
            self.data = {"tasks": ["Admin"], "sessions": []}
            self._save_data()

    def _save_data(self):
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror(APP_NAME, f"Couldn’t save data:\n{e}")

    # ---------- Helpers ----------
    def _prompt(self, title: str, label: str):
        win = tk.Toplevel(self.root)
        win.title(title)
        win.transient(self.root)
        win.grab_set()
        win.resizable(False, False)

        frm = ttk.Frame(win, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text=label).pack(anchor="w")
        var = tk.StringVar()

        ent = ttk.Entry(frm, textvariable=var, width=40)
        ent.pack(pady=8)
        ent.focus_set()

        out = {"value": None}

        def ok():
            out["value"] = var.get()
            win.destroy()

        def cancel():
            out["value"] = None
            win.destroy()

        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(6, 0))

        ttk.Button(btns, text="OK", command=ok).pack(side="left")
        ttk.Button(btns, text="Cancel", command=cancel).pack(side="left", padx=8)

        win.bind("<Return>", lambda e: ok())
        win.bind("<Escape>", lambda e: cancel())

        self.root.wait_window(win)
        return out["value"]

    def _set_status(self, msg: str):
        self.status.config(text=msg)


def main():
    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    app = TimeTrackerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()