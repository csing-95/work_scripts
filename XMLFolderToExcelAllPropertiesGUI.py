# Run: python '.\XMLFolderToExcel_AllProperties_GUI.py'
# A GUI tool to extract all properties from all XML files in a folder into a single Excel file.

import os
import glob
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from collections import Counter
import xml.etree.ElementTree as ET
import pandas as pd


def make_unique_headers(headers):
    counts = Counter()
    unique = []
    for h in headers:
        counts[h] += 1
        unique.append(h if counts[h] == 1 else f"{h} ({counts[h]})")
    return unique


def extract_xml_folder_to_excel(xml_folder: str, output_filename: str, status_cb=None) -> str:
    if not os.path.isdir(xml_folder):
        raise ValueError("The selected XML folder does not exist.")

    xml_files = glob.glob(os.path.join(xml_folder, "*.xml"))
    if status_cb:
        status_cb(f"Found {len(xml_files)} XML file(s).")

    if not xml_files:
        raise ValueError("No .xml files found in the selected folder.")

    # Ensure .xlsx extension
    output_filename = output_filename.strip()
    if not output_filename:
        raise ValueError("Please enter an output file name.")
    if not output_filename.lower().endswith(".xlsx"):
        output_filename += ".xlsx"

    all_rows = []

    for i, xml_file in enumerate(xml_files, start=1):
        if status_cb:
            status_cb(f"Reading {i}/{len(xml_files)}: {os.path.basename(xml_file)}")

        try:
            tree = ET.parse(xml_file)
            root = tree.getroot()

            # ---- Build ordered property definitions ----
            with_ix = []
            no_ix = []

            for pd_node in root.findall(".//PropertyDefs/PropertyDef"):
                display = pd_node.get("_DISPLAYNAME") or pd_node.get("_INTERNALNAME") or "Unnamed"
                internal = pd_node.get("_INTERNALNAME") or ""
                ix = pd_node.get("ix")

                if ix is not None:
                    with_ix.append((int(ix), display, internal))
                else:
                    no_ix.append((display, internal))

            with_ix.sort(key=lambda x: x[0])

            max_ix = with_ix[-1][0] if with_ix else -1
            next_ix = max_ix + 1

            all_defs = [(ix, display, internal) for ix, display, internal in with_ix]
            for display, internal in no_ix:
                all_defs.append((next_ix, display, internal))
                next_ix += 1

            headers = [d[1] for d in all_defs]
            unique_headers = make_unique_headers(headers)

            # ---- Extract each Record ----
            for record in root.findall(".//Record"):
                props = record.findall("./Property")
                row = {"SourceFile": os.path.basename(xml_file)}

                for (ix, _display, _internal), header in zip(all_defs, unique_headers):
                    row[header] = props[ix].text.strip() if ix < len(props) and props[ix].text else ""

                if len(props) > len(all_defs):
                    for extra_i in range(len(all_defs), len(props)):
                        extra_val = props[extra_i].text.strip() if props[extra_i].text else ""
                        row[f"ExtraProperty_{extra_i}"] = extra_val

                all_rows.append(row)

        except Exception as e:
            # Keep going; log error as a row so you can see what failed
            all_rows.append({
                "SourceFile": os.path.basename(xml_file),
                "ERROR": str(e)
            })

    df = pd.DataFrame(all_rows)

    output_path = os.path.join(xml_folder, output_filename)
    df.to_excel(output_path, index=False)

    if status_cb:
        status_cb(f"Done! Wrote {len(df)} row(s) to {output_filename}")

    return output_path


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("XML → Excel Extractor")
        self.geometry("640x260")
        self.resizable(False, False)

        self.xml_folder_var = tk.StringVar()
        self.output_name_var = tk.StringVar(value="combined_output.xlsx")
        self.status_var = tk.StringVar(value="Pick a folder of XML files, choose an output name, then Run.")

        pad = {"padx": 10, "pady": 6}

        # Folder row
        tk.Label(self, text="XML folder:").grid(row=0, column=0, sticky="w", **pad)
        tk.Entry(self, textvariable=self.xml_folder_var, width=60).grid(row=0, column=1, sticky="w", **pad)
        tk.Button(self, text="Browse…", command=self.browse_folder).grid(row=0, column=2, sticky="w", **pad)

        # Output name row
        tk.Label(self, text="Output Excel name:").grid(row=1, column=0, sticky="w", **pad)
        tk.Entry(self, textvariable=self.output_name_var, width=60).grid(row=1, column=1, sticky="w", **pad)
        tk.Label(self, text="(.xlsx)").grid(row=1, column=2, sticky="w", **pad)

        # Run button
        self.run_btn = tk.Button(self, text="Run", command=self.run_clicked, width=12)
        self.run_btn.grid(row=2, column=1, sticky="w", **pad)

        # Status box
        tk.Label(self, text="Status:").grid(row=3, column=0, sticky="nw", **pad)
        self.status_label = tk.Label(self, textvariable=self.status_var, justify="left", wraplength=520, anchor="w")
        self.status_label.grid(row=3, column=1, columnspan=2, sticky="w", **pad)

        # Little footer tip
        tk.Label(self, text="Tip: The output file will be saved inside the selected XML folder.").grid(
            row=4, column=1, columnspan=2, sticky="w", padx=10, pady=2
        )

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select folder containing XML files")
        if folder:
            self.xml_folder_var.set(folder)

    def set_status(self, msg: str):
        self.status_var.set(msg)
        self.update_idletasks()

    def run_clicked(self):
        xml_folder = self.xml_folder_var.get().strip()
        output_name = self.output_name_var.get().strip()

        if not xml_folder:
            messagebox.showerror("Missing folder", "Please select an XML folder.")
            return
        if not output_name:
            messagebox.showerror("Missing output name", "Please enter an output Excel file name.")
            return

        self.run_btn.config(state="disabled")
        self.set_status("Starting...")

        def worker():
            try:
                out_path = extract_xml_folder_to_excel(xml_folder, output_name, status_cb=self.set_status)
                messagebox.showinfo("Success", f"Export complete!\n\nSaved to:\n{out_path}")
            except Exception as e:
                messagebox.showerror("Error", str(e))
                self.set_status(f"Error: {e}")
            finally:
                self.run_btn.config(state="normal")

        threading.Thread(target=worker, daemon=True).start()


if __name__ == "__main__":
    App().mainloop()
