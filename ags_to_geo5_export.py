import pandas as pd
import os
import csv
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ---- PARAMETERS ----
AGS_FILE = "input_file.ags"  # Set your AGS file path
TEMPLATE_FILE = "FieldTestImportTemplate.xlsx"  # Set your template path
OUTPUT_FILE = "Geo5_Import.xlsx"


# ---- AGS PARSING LOGIC (adapted from data_loader.py) ----
def parse_group(content, group_name):
    lines = content.splitlines()
    parsed = list(csv.reader(lines, delimiter=",", quotechar='"'))
    headings = []
    data = []
    in_group = False
    for row in parsed:
        if row and row[0] == "GROUP" and len(row) > 1 and row[1] == group_name:
            in_group = True
            continue
        if in_group and row and row[0] == "HEADING":
            headings = row[1:]
            continue
        if in_group and row and row[0] == "DATA":
            data.append(row[1 : len(headings) + 1])
            continue
        if (
            in_group
            and row
            and row[0] == "GROUP"
            and (len(row) < 2 or row[1] != group_name)
        ):
            break
    df = pd.DataFrame(data, columns=headings)
    return df


# ---- LOAD AGS FILE ----
def load_ags_tables(ags_path):
    with open(ags_path, encoding="utf-8") as f:
        content = f.read()
    geol = parse_group(content, "GEOL")
    point = parse_group(content, "POINT")
    loca = parse_group(content, "LOCA")
    return {"GEOL": geol, "POINT": point, "LOCA": loca}


# ---- DATA LOADING ----
nags_tables = load_ags_tables(AGS_FILE)
df_geol = nags_tables["GEOL"]
df_point = nags_tables["POINT"]
df_loca = nags_tables["LOCA"]

# Convert numeric columns
for col in ["GEOL_DEPTH", "GEOL_BASE"]:
    if col in df_geol.columns:
        df_geol[col] = pd.to_numeric(df_geol[col], errors="coerce")

for col in ["LOCA_NATE", "LOCA_NATN", "LOCA_GL"]:
    if col in df_loca.columns:
        df_loca[col] = pd.to_numeric(df_loca[col], errors="coerce")

# Merge POINT with LOCA for coordinates
if "POINT_ID" in df_point.columns and "LOCA_ID" in df_loca.columns:
    df_point = df_point.merge(
        df_loca,
        left_on="POINT_ID",
        right_on="LOCA_ID",
        how="left",
        suffixes=("", "_loca"),
    )

# ---- LAYER NAME EXTRACTION ----
regex = re.compile(r"^([A-Z]+(?:\s+[A-Z]+)*)")


def extract_layer_name(desc):
    match = regex.match(str(desc))
    return match.group(1).strip() if match and len(match.group(1)) > 1 else ""


if "GEOL_GEOL" in df_geol.columns:
    df_geol["GEOL_GEOL"] = df_geol.apply(
        lambda row: (
            row["GEOL_GEOL"]
            if pd.notna(row.get("GEOL_GEOL", "")) and row["GEOL_GEOL"]
            else extract_layer_name(row.get("GEOL_DESC", ""))
        ),
        axis=1,
    )
else:
    df_geol["GEOL_GEOL"] = df_geol.apply(
        lambda row: extract_layer_name(row.get("GEOL_DESC", "")), axis=1
    )


# ---- SOIL CLASSIFICATION AUTO ASSIGNMENT ----
def auto_classify(geol_leg):
    soil_map = {
        "CLAY": "Clay, fine grained",
        "SAND": "Sand, coarse grained",
        "SILT": "Silt",
        "GRAVEL": "Gravel",
        # Add more mappings as needed
    }
    return soil_map.get(str(geol_leg).upper(), "")


df_geol["Soil Classification"] = df_geol.get("GEOL_LEG", pd.Series("")).apply(
    auto_classify
)

# ---- WRITE TO TEMPLATE ----
wb = load_workbook(TEMPLATE_FILE)
ws_fieldtests = wb["FieldTests"]
ws_layers = wb["Layers"]

# ---- CLEAR AND PREPARE FIELDTESTS SHEET ----
ws_fieldtests.delete_rows(2, ws_fieldtests.max_row)  # keep header

# Map POINT table columns to FieldTests
row_idx = 2
for i, row in df_point.iterrows():
    ws_fieldtests.cell(
        row=row_idx, column=1, value=row.get("POINT_ID", row.get("LOCA_ID", ""))
    )  # "Test name"
    ws_fieldtests.cell(
        row=row_idx, column=2, value="(local set) : Borehole"
    )  # "Template"
    ws_fieldtests.cell(
        row=row_idx, column=3, value=row.get("LOCA_NATN", "")
    )  # "Coordinate X"
    ws_fieldtests.cell(
        row=row_idx, column=4, value=row.get("LOCA_NATE", "")
    )  # "Coordinate Y"
    ws_fieldtests.cell(row=row_idx, column=5, value="input")  # "Elevation"
    ws_fieldtests.cell(
        row=row_idx, column=6, value=row.get("LOCA_GL", "")
    )  # "Coordinate Z"
    row_idx += 1

# ---- CLEAR AND PREPARE LAYERS SHEET ----
ws_layers.delete_rows(2, ws_layers.max_row)  # keep header

# Group GEOL table by borehole and layer, calculate thickness and combine descriptions
layer_row = 2
for borehole_id, bh_data in (
    df_geol.groupby("GEOL_BHID") if "GEOL_BHID" in df_geol.columns else []
):
    for layer, layer_data in bh_data.groupby("GEOL_GEOL"):
        start_depth = (
            layer_data["GEOL_DEPTH"].min()
            if "GEOL_DEPTH" in layer_data.columns
            else None
        )
        end_depth = (
            layer_data["GEOL_BASE"].max() if "GEOL_BASE" in layer_data.columns else None
        )
        thickness = (
            end_depth - start_depth
            if start_depth is not None and end_depth is not None
            else None
        )
        desc = "; ".join(
            layer_data.get("GEOL_DESC", pd.Series("")).astype(str).tolist()
        )
        geol_leg = layer_data.get("GEOL_LEG", pd.Series("")).iloc[0]
        soil_class = layer_data["Soil Classification"].iloc[0]

        ws_layers.cell(row=layer_row, column=1, value=borehole_id)  # Test name
        ws_layers.cell(row=layer_row, column=2, value=thickness)  # Thickness
        ws_layers.cell(row=layer_row, column=3, value=layer)  # Soil name
        ws_layers.cell(
            row=layer_row, column=4, value="GEO_CLAY"
        )  # Soil pattern|Pattern (adjust as needed)
        ws_layers.cell(
            row=layer_row, column=5, value=""
        )  # Soil pattern|Color (fill later)
        ws_layers.cell(
            row=layer_row, column=6, value="clDefault"
        )  # Soil pattern|Background
        ws_layers.cell(row=layer_row, column=7, value="50")  # Soil pattern|Saturation
        ws_layers.cell(row=layer_row, column=8, value=desc)  # Layer description
        ws_layers.cell(
            row=layer_row, column=9, value=soil_class
        )  # EN ISO 14688-1 Classification
        layer_row += 1

# ---- COLOUR FORMATTING (OPTIONAL) ----
yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
for row in ws_layers.iter_rows(
    min_row=2, max_row=ws_layers.max_row, min_col=3, max_col=3
):
    for cell in row:
        if cell.value:
            cell.fill = yellow

# ---- SAVE OUTPUT ----
wb.save(OUTPUT_FILE)
print(f"AGS parsed and exported to {OUTPUT_FILE}")
