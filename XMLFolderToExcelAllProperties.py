# tool for extracting all properties from XML files in a folder and exporting to Excel

import xml.etree.ElementTree as ET
import pandas as pd
import glob
import os
from collections import Counter

# Folder containing XML files (must be a folder)
xml_folder = r"C:\Users\corri\Kinsmen Group\P25025 Plains Adept to Meridian Migration Project - Documents\6. Data Migration\PROD Location Loadsheets\"  # <-- change if needed

xml_files = glob.glob(os.path.join(xml_folder, "*.xml"))
print(f"Found {len(xml_files)} XML files.")

all_rows = []

def make_unique_headers(headers):
    """
    If display names repeat (e.g. 'Project ID'), make them unique:
    'Project ID', 'Project ID (2)', 'Project ID (3)'...
    """
    counts = Counter()
    unique = []
    for h in headers:
        counts[h] += 1
        unique.append(h if counts[h] == 1 else f"{h} ({counts[h]})")
    return unique

for xml_file in xml_files:
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
                no_ix.append((display, internal))  # keep in document order

        with_ix.sort(key=lambda x: x[0])

        max_ix = with_ix[-1][0] if with_ix else -1
        next_ix = max_ix + 1

        all_defs = []
        for ix, display, internal in with_ix:
            all_defs.append((ix, display, internal))

        for display, internal in no_ix:
            all_defs.append((next_ix, display, internal))
            next_ix += 1

        # Make headers unique (because display names can repeat)
        headers = [d[1] for d in all_defs]
        unique_headers = make_unique_headers(headers)

        # ---- Extract each Record ----
        for record in root.findall(".//Record"):
            props = record.findall("./Property")

            row = {"SourceFile": os.path.basename(xml_file)}

            # Fill values based on defs
            for (ix, _display, _internal), header in zip(all_defs, unique_headers):
                row[header] = props[ix].text.strip() if ix < len(props) and props[ix].text else ""

            # If there are MORE Property values than defs, keep them too
            if len(props) > len(all_defs):
                for extra_i in range(len(all_defs), len(props)):
                    extra_val = props[extra_i].text.strip() if props[extra_i].text else ""
                    row[f"ExtraProperty_{extra_i}"] = extra_val

            all_rows.append(row)

    except Exception as e:
        print(f"⚠ Error reading {xml_file}: {e}")

df = pd.DataFrame(all_rows)

output_path = os.path.join(xml_folder, "combined_output_missingtitles.xlsx")
df.to_excel(output_path, index=False)

print(f"✅ Done! Extracted {len(df)} records from {len(xml_files)} XML files.")
print(f"📄 Output saved to: {output_path}")
