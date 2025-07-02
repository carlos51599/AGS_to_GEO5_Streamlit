import pandas as pd
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def extract_layer_name(desc):
    regex = re.compile(r"^([A-Z]+(?:\s+[A-Z]+)*)")
    match = regex.match(str(desc))
    return match.group(1).strip() if match and len(match.group(1)) > 1 else ""


def auto_classify(geol_leg):
    soil_map = {
        "CLAY": "Clay, fine grained",
        "SAND": "Sand, coarse grained",
        "SILT": "Silt",
        "GRAVEL": "Gravel",
        # Add more mappings as needed
    }
    return soil_map.get(str(geol_leg).upper(), "")


def export_to_excel(df_geol, df_loca, df_abbr, template_path, output_path):
    # Prepare data
    for col in ["GEOL_DEPTH", "GEOL_BASE"]:
        if col in df_geol.columns:
            df_geol[col] = pd.to_numeric(df_geol[col], errors="coerce")
    for col in ["LOCA_NATE", "LOCA_NATN", "LOCA_GL"]:
        if col in df_loca.columns:
            df_loca[col] = pd.to_numeric(df_loca[col], errors="coerce")
    # Write to Excel
    wb = load_workbook(template_path)
    ws_fieldtests = wb["FieldTests"]
    ws_layers = wb["Layers"]
    ws_fieldtests.delete_rows(2, ws_fieldtests.max_row)
    row_idx = 2
    for i, row in df_loca.iterrows():
        ws_fieldtests.cell(row=row_idx, column=1, value=row.get("LOCA_ID", ""))
        ws_fieldtests.cell(row=row_idx, column=2, value="(local set) : Borehole")
        ws_fieldtests.cell(row=row_idx, column=3, value=row.get("LOCA_NATN", ""))
        ws_fieldtests.cell(row=row_idx, column=4, value=row.get("LOCA_NATE", ""))
        ws_fieldtests.cell(row=row_idx, column=5, value="input")
        ws_fieldtests.cell(row=row_idx, column=6, value=row.get("LOCA_GL", ""))
        row_idx += 1
    ws_layers.delete_rows(2, ws_layers.max_row)
    layer_row = 2
    # Group by LOCA_ID and GEOL_LEG (soil code)
    if "LOCA_ID" in df_geol.columns and "GEOL_LEG" in df_geol.columns:
        for borehole_id, bh_data in df_geol.groupby("LOCA_ID"):
            for leg, layer_data in bh_data.groupby("GEOL_LEG"):
                start_depth = (
                    layer_data["GEOL_DEPTH"].min()
                    if "GEOL_DEPTH" in layer_data.columns
                    else None
                )
                end_depth = (
                    layer_data["GEOL_BASE"].max()
                    if "GEOL_BASE" in layer_data.columns
                    else None
                )
                thickness = (
                    end_depth - start_depth
                    if start_depth is not None and end_depth is not None
                    else None
                )
                desc = "; ".join(
                    layer_data.get("GEOL_DESC", pd.Series("")).astype(str).tolist()
                )
                # Soil name from ABBR_DESC
                soil_name = leg
                if (
                    df_abbr is not None
                    and "ABBR_CODE" in df_abbr.columns
                    and "ABBR_DESC" in df_abbr.columns
                ):
                    abbr_match = df_abbr[df_abbr["ABBR_CODE"] == str(leg)]
                    if not abbr_match.empty:
                        soil_name = abbr_match["ABBR_DESC"].iloc[0]
                ws_layers.cell(row=layer_row, column=1, value=borehole_id)
                ws_layers.cell(row=layer_row, column=2, value=thickness)
                ws_layers.cell(row=layer_row, column=3, value=soil_name)
                ws_layers.cell(row=layer_row, column=4, value=leg)
                ws_layers.cell(row=layer_row, column=5, value="")
                ws_layers.cell(row=layer_row, column=6, value="clDefault")
                ws_layers.cell(row=layer_row, column=7, value="50")
                ws_layers.cell(row=layer_row, column=8, value=desc)
                ws_layers.cell(row=layer_row, column=9, value=soil_name)
                layer_row += 1
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for row in ws_layers.iter_rows(
        min_row=2, max_row=ws_layers.max_row, min_col=3, max_col=3
    ):
        for cell in row:
            if cell.value:
                cell.fill = yellow
    wb.save(output_path)
