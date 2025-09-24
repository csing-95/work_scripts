import pandas as pd
import re

# ===== CONFIG =====
# file_path    = r"filepath/filename.xlsx"
file_path    = r"C:\Users\corri\Kinsmen Group\P25025 Plains Adept to Meridian Migration Project - Documents\6. Data Migration\Location Loadsheets\Projects\Working\000 - Plains Midstream Canada Projects Loadsheet.xlsx"
sheet_name   = "Documents"
output_sheet = "Dupe Decisions"

COL_DOC    = "Document Name"
COL_SIZE   = "File size"
COL_REV    = "Revision Number"          # or "Rev"
COL_LATEST = "isLatest"                 # boolean or truthy strings
COL_MAJMIN = "Legacy Version Number"    # e.g., "2.10", "1.3", can be blank
COL_EXT    = "Ext"                      # optional; only used if GROUP_BY needs it

# Choose how to group duplicates:
#   "doc_size_rev"  -> Document + File size + Revision (original behavior, but normalized)
#   "doc_rev"       -> Document + Revision (ignores size & format)
#   "doc_rev_ext"   -> Document + Revision + Ext (ignores size, keeps format separate)
GROUP_BY = "doc_size_rev"

# ===== Helpers =====
def truthy(x):
    if pd.isna(x): return False
    if isinstance(x, bool): return x
    return str(x).strip().lower() in {"true","yes","y","1","t"}

def parse_majmin(s):
    """'2.10' -> (2,10), '1.3.5' -> (1,3,5), blanks -> (0,)"""
    if pd.isna(s) or str(s).strip()=="":
        return (0,)
    parts = re.split(r"[^\d]+", str(s).strip())
    nums = tuple(int(p) for p in parts if p.isdigit())
    return nums or (0,)

def pad_tuple(t, n=4):
    return t + (0,)*(n-len(t)) if len(t) < n else t

def norm_num_str(x, null_value="-1"):
    """Return a canonical string for numeric-like values:
       77428 -> '77428', 77428.0 -> '77428', 1.1 -> '1.1', NaN -> null_value."""
    if pd.isna(x): return null_value
    try:
        v = float(x)
        return str(int(v)) if v.is_integer() else str(v)
    except Exception:
        s = str(x).strip()
        return null_value if s=="" else s

def norm_text(s):
    return "" if pd.isna(s) else str(s).strip()

# ===== Load =====
df = pd.read_excel(file_path, sheet_name=sheet_name)
print(f"Loaded {len(df)} rows from '{sheet_name}' - {file_path}")
print("Now running duplicate-check process...")

# Preserve original order
df["_orig_order"] = range(len(df))

# Checks
for c in [COL_DOC, COL_SIZE, COL_REV]:
    if c not in df.columns:
        raise ValueError(f"Missing required column: {c}")

# Normalise core fields
df["_doc_norm"]  = df[COL_DOC].apply(norm_text)
df["_rev_norm"]  = df[COL_REV].apply(norm_text)

# size as numeric first, then canonical string
df[COL_SIZE] = pd.to_numeric(df[COL_SIZE], errors="coerce")
df["_size_norm"] = df[COL_SIZE]  # keep numeric for reference
df["_size_norm_str"] = df["_size_norm"].apply(lambda x: norm_num_str(x, null_value="-1"))

# Optional fields
if COL_EXT in df.columns:
    df["_ext_norm"] = df[COL_EXT].astype(str).str.strip().str.lower()
else:
    df["_ext_norm"] = ""

df["_is_latest_bool"] = df[COL_LATEST].apply(truthy) if COL_LATEST in df.columns else False
df["_majmin_tuple"]   = (
    df[COL_MAJMIN].apply(parse_majmin).apply(lambda t: pad_tuple(t, 4))
    if COL_MAJMIN in df.columns else [(0,)*4]*len(df)
)

# ===== Dupe key (switchable) =====
if GROUP_BY == "doc_rev":
    df["_dupe_key"] = df["_doc_norm"] + "||" + df["_rev_norm"]
elif GROUP_BY == "doc_rev_ext":
    df["_dupe_key"] = df["_doc_norm"] + "||" + df["_rev_norm"] + "||" + df["_ext_norm"]
else:  # "doc_size_rev"
    df["_dupe_key"] = df["_doc_norm"] + "||" + df["_size_norm_str"] + "||" + df["_rev_norm"]

# Defaults
df["Dupe Action"]   = "keep"
df["Dupe Reason"]   = "unique (no duplicates)"
df["Dupe Conflict"] = ""

# ===== Decide per group =====
for key, g in df.groupby("_dupe_key", sort=False):  # don't reorder groups
    if len(g) == 1:
        continue

    g_latest = g[g["_is_latest_bool"]]
    if len(g_latest) >= 1:
        max_mm = g_latest["_majmin_tuple"].max()
        latest_max = g_latest[g_latest["_majmin_tuple"] == max_mm]
        if len(latest_max) > 1:
            idx_keep = latest_max.index.min()  # keep the first seen
            df.loc[latest_max.index, "Dupe Conflict"] = "both isLatest=True & same Maj&Min"
            reason_keep   = "kept (both IsLatest=True & same Maj&Min; kept first)"
            reason_remove = "removed (duplicate: IsLatest=True tie; not first)"
        else:
            idx_keep = latest_max.index[0]
            reason_keep   = "kept (IsLatest=True and highest Maj&Min)"
            reason_remove = "removed (duplicate: lower IsLatest/Maj&Min)"
    else:
        max_mm = g["_majmin_tuple"].max()
        ties = g[g["_majmin_tuple"] == max_mm]
        if len(ties) > 1:
            idx_keep = ties.index.min()
            reason_keep   = "kept (tie on Maj&Min; kept first)"
            reason_remove = "removed (duplicate: tie on Maj&Min; not first)"
        else:
            idx_keep = ties.index[0]
            reason_keep   = "kept (highest Maj&Min)"
            reason_remove = "removed (older Maj&Min)"

    df.loc[g.index, "Dupe Action"] = "remove"
    df.loc[g.index, "Dupe Reason"] = reason_remove
    df.loc[idx_keep, ["Dupe Action","Dupe Reason"]] = ["keep", reason_keep]

# ===== Output (preserve original order) =====
preferred = [
    COL_DOC, COL_SIZE, COL_REV,
    COL_MAJMIN if COL_MAJMIN in df.columns else None,
    COL_LATEST if COL_LATEST in df.columns else None,
    "Dupe Action", "Dupe Reason", "Dupe Conflict"
]
preferred = [c for c in preferred if c is not None]

aux_drop = {"_dupe_key", "_is_latest_bool", "_majmin_tuple"}
# keep debug columns to help diagnose grouping
debug_cols = ["_doc_norm","_rev_norm","_size_norm_str","_ext_norm"]
rest = [c for c in df.columns if c not in set(preferred) | aux_drop | set(debug_cols) | {"_orig_order"}]

out = df.sort_values("_orig_order")[preferred + debug_cols + rest]

with pd.ExcelWriter(file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
    out.to_excel(w, sheet_name=output_sheet, index=False)

print(f"Wrote '{output_sheet}' with keep/remove decisions, reasons, and conflict flags. Grouping: {GROUP_BY} - {file_path}")

