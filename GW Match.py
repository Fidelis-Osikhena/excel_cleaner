import pandas as pd
from pathlib import Path

# -----------------------------
# SETTINGS
# -----------------------------

input_file = Path("input.xlsx")
output_file = Path("aligned_output_fixed.xlsx")
sheet_name = "Sheet1"

# Your left table is columns A:C
left_usecols = ["A Feature Type", "A Weld Id", "A Odometer Main"]

# Your right table is columns G:I
right_usecols = ["B Feature Type", "B Weld Id", "B Odometer Main"]

left_feature_col = "A Feature Type"
left_weld_col = "A Weld Id"

right_feature_col = "B Feature Type"
right_weld_col = "B Weld Id"

girth_value = "Girth Weld"


# -----------------------------
# READ LEFT AND RIGHT TABLES
# -----------------------------

left = pd.read_excel(input_file, sheet_name=sheet_name, usecols=left_usecols)
right = pd.read_excel(input_file, sheet_name=sheet_name, usecols=right_usecols)

# Remove completely empty rows at the bottom only
left = left.dropna(how="all").reset_index(drop=True)
right = right.dropna(how="all").reset_index(drop=True)


# -----------------------------
# HELPER FUNCTIONS
# -----------------------------

def is_girth(row, feature_col):
    return str(row[feature_col]).strip().lower() == girth_value.lower()


def clean_weld_id(value):
    if pd.isna(value):
        return ""
    return str(value).strip().replace(".0", "")


def blank_row(columns):
    return pd.Series({col: "" for col in columns})


def get_girth_positions(df, feature_col, weld_col):
    """
    Returns a list of:
    [
        {"index": row_number, "weld_id": weld_id},
        ...
    ]
    """
    girths = []

    for idx, row in df.iterrows():
        if is_girth(row, feature_col):
            girths.append({
                "index": idx,
                "weld_id": clean_weld_id(row[weld_col])
            })

    return girths


# -----------------------------
# FIND GIRTH WELDS IN BOTH TABLES
# -----------------------------

left_girths = get_girth_positions(left, left_feature_col, left_weld_col)
right_girths = get_girth_positions(right, right_feature_col, right_weld_col)

left_ids = {g["weld_id"] for g in left_girths}
right_ids = {g["weld_id"] for g in right_girths}

# Only align girth welds that exist in BOTH tables
common_ids = left_ids.intersection(right_ids)

# Preserve left-side order
matched_ids = [g["weld_id"] for g in left_girths if g["weld_id"] in common_ids]


# -----------------------------
# BUILD ALIGNED OUTPUT
# -----------------------------

aligned_left = []
aligned_right = []

left_start = 0
right_start = 0

for weld_id in matched_ids:
    # Find the next matching girth weld in each table
    left_match_index = next(
        g["index"] for g in left_girths
        if g["weld_id"] == weld_id and g["index"] >= left_start
    )

    right_match_index = next(
        g["index"] for g in right_girths
        if g["weld_id"] == weld_id and g["index"] >= right_start
    )

    # Rows before this matching girth weld
    left_block = left.iloc[left_start:left_match_index + 1]
    right_block = right.iloc[right_start:right_match_index + 1]

    # Make both blocks the same height by padding the shorter side
    max_len = max(len(left_block), len(right_block))

    for i in range(max_len):
        if i < len(left_block):
            aligned_left.append(left_block.iloc[i])
        else:
            aligned_left.append(blank_row(left.columns))

        if i < len(right_block):
            aligned_right.append(right_block.iloc[i])
        else:
            aligned_right.append(blank_row(right.columns))

    # Move past the matched girth weld
    left_start = left_match_index + 1
    right_start = right_match_index + 1


# -----------------------------
# ADD REMAINING ROWS AFTER LAST MATCH
# -----------------------------

left_remaining = left.iloc[left_start:]
right_remaining = right.iloc[right_start:]

max_remaining = max(len(left_remaining), len(right_remaining))

for i in range(max_remaining):
    if i < len(left_remaining):
        aligned_left.append(left_remaining.iloc[i])
    else:
        aligned_left.append(blank_row(left.columns))

    if i < len(right_remaining):
        aligned_right.append(right_remaining.iloc[i])
    else:
        aligned_right.append(blank_row(right.columns))


# -----------------------------
# COMBINE SIDE BY SIDE
# -----------------------------

aligned_left_df = pd.DataFrame(aligned_left).reset_index(drop=True)
aligned_right_df = pd.DataFrame(aligned_right).reset_index(drop=True)

# Add blank spacer columns like your original file
spacer = pd.DataFrame({"": [""] * len(aligned_left_df), " ": [""] * len(aligned_left_df), "  ": [""] * len(aligned_left_df)})

final_df = pd.concat([aligned_left_df, spacer, aligned_right_df], axis=1)


# -----------------------------
# SAVE OUTPUT
# -----------------------------

final_df.to_excel(output_file, sheet_name=sheet_name, index=False)

print(f"Done. Saved as: {output_file}")