import os
import pandas as pd
from datetime import datetime

def get_file_metadata(folder_path):
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
                    "Created Date": datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                    "Modified Date": datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                    "Accessed Date": datetime.fromtimestamp(stats.st_atime).strftime('%Y-%m-%d %H:%M:%S')
                })
            except Exception as e:
                print(f"Error processing file {filepath}: {e}")

    return pd.DataFrame(data)

def build_summary(df):
    total_files = len(df)

    summary_counts = (
        df.groupby("File Extension")
          .agg(
              File_Count=("File Extension", "size"),
              Total_Size_Bytes=("File Size (Bytes)", "sum"),
              Total_Size_KB=("File Size (KB)", "sum"),
          )
          .reset_index()
    )

    summary_counts["% of Files"] = (summary_counts["File_Count"] / total_files * 100).round(2)
    summary_counts["Total_Size_KB"] = summary_counts["Total_Size_KB"].round(2)
    summary_counts = summary_counts.sort_values(by="File_Count", ascending=False)

    total_row = pd.DataFrame([{
        "File Extension": "TOTAL",
        "File_Count": total_files,
        "Total_Size_Bytes": summary_counts["Total_Size_Bytes"].sum(),
        "Total_Size_KB": round(summary_counts["Total_Size_KB"].sum(), 2),
        "% of Files": 100.00
    }])

    return pd.concat([summary_counts, total_row], ignore_index=True)

def ask(prompt, default=None):
    suffix = f" [{default}]" if default else ""
    value = input(f"{prompt}{suffix}: ").strip()
    value = value.strip('"').strip("'")  # 👈 remove quotes
    return value if value else default


def ask_format():
    while True:
        choice = input("Save as (1) Excel or (2) CSV? [1/2]: ").strip()
        if choice in ("1", "2"):
            return "excel" if choice == "1" else "csv"
        print("Please type 1 for Excel or 2 for CSV.")

if __name__ == "__main__":
    folder_path = ask("Enter folder path to scan", r"C:\Users\corri\Documents\sample_files")
    out_base = ask("Enter output file path (no extension)", r"C:\Users\corri\Documents\sample_file_metadata")
    out_format = ask_format()

    if not os.path.isdir(folder_path):
        print(f"Folder not found: {folder_path}")
        raise SystemExit(1)

    df = get_file_metadata(folder_path)

    if df.empty:
        print("No files found in the specified folder.")
        raise SystemExit(0)

    summary_counts = build_summary(df)

    if out_format == "csv":
        meta_csv = out_base + "_metadata.csv"
        summary_csv = out_base + "_summary.csv"
        df.to_csv(meta_csv, index=False)
        summary_counts.to_csv(summary_csv, index=False)
        print(f"\nSaved:\n- {meta_csv}\n- {summary_csv}")
    else:
        xlsx = out_base + ".xlsx"
        with pd.ExcelWriter(xlsx, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="File Metadata", index=False)
            summary_counts.to_excel(writer, sheet_name="Summary", index=False)
        print(f"\nSaved Excel: {xlsx}")

    print(f"Total files: {len(df)}")
