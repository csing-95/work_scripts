import os
import pandas as pd
from datetime import datetime

def get_file_metadata(folder_path):
    data = []
    
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            filepath = os.path.join(root, file)
            try:
                stats = os.stat(filepath)
                _, file_ext = os.path.splitext(file)
                
                data.append({
                    "File Path": filepath,
                    "File Name": file,
                    "File Extension": file_ext.lower() if file_ext else "No Extension",
                    "File Size (Bytes)": stats.st_size,
                    "Created Date": datetime.fromtimestamp(stats.st_ctime).strftime('%Y-%m-%d %H:%M:%S'),
                    "Modified Date": datetime.fromtimestamp(stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                    "Accessed Date": datetime.fromtimestamp(stats.st_atime).strftime('%Y-%m-%d %H:%M:%S')
                })
            except Exception as e:
                print(f"Error processing file {filepath}: {e}")
    
    return pd.DataFrame(data)

if __name__ == "__main__":
    folder_path = r"\\?\C:\Users\corri\Desktop\to_sort"
    output_file = "file_metadata.xlsx"
    
    # Get detailed metadata
    df = get_file_metadata(folder_path)
    
    if df.empty:
        print("No files found in the specified folder.")
    else:
        # Summary stats
        total_files = len(df)
        summary_counts = (
            df.groupby("File Extension")
              .agg(File_Count=("File Extension", "size"), Total_Size=("File Size (Bytes)", "sum"))
              .reset_index()
        )
        
        # Add percentage
        summary_counts["% of Files"] = (summary_counts["File_Count"] / total_files * 100).round(2)
        
        # Sort by File_Count descending
        summary_counts = summary_counts.sort_values(by="File_Count", ascending=False)
        
        # Add total row
        total_row = pd.DataFrame([{
            "File Extension": "TOTAL",
            "File_Count": total_files,
            "Total_Size": summary_counts["Total_Size"].sum(),
            "% of Files": 100.00
        }])
        
        summary_counts = pd.concat([summary_counts, total_row], ignore_index=True)
        
        # Export to Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name="File Metadata", index=False)
            summary_counts.to_excel(writer, sheet_name="Summary", index=False)
        
        print(f"\nMetadata and summary exported to {output_file}")
        print(f"Total files: {total_files}")
        print(summary_counts)
