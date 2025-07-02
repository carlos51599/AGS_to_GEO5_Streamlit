import streamlit as st
import pandas as pd
from ags_to_geo5.ags_parser import load_ags_tables
from ags_to_geo5.exporter import export_to_excel
import tempfile
import os
from streamlit_pdf_viewer import pdf_viewer

st.set_page_config(layout="wide")

st.title("AGS to GEO5 Excel Converter")
st.write("Upload your AGS file and download the GEO5 import Excel file.")

TEMPLATE_FILE = "FieldTestImportTemplate.xlsx"

uploaded_file = st.file_uploader("Choose an AGS file", type=["ags"])

if uploaded_file is not None:
    ags_content = uploaded_file.read().decode("utf-8")
    ags_tables = load_ags_tables(ags_content)
    df_geol = ags_tables["GEOL"]
    df_loca = ags_tables["LOCA"]
    df_abbr = ags_tables["ABBR"] if "ABBR" in ags_tables else None
    # Ensure numeric columns for subtraction
    for col in ["GEOL_TOP", "GEOL_BASE", "GEOL_DEPTH"]:
        if col in df_geol.columns:
            df_geol[col] = pd.to_numeric(df_geol[col], errors="coerce")
    for col in ["LOCA_NATE", "LOCA_NATN", "LOCA_GL"]:
        if col in df_loca.columns:
            df_loca[col] = pd.to_numeric(df_loca[col], errors="coerce")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        export_to_excel(df_geol, df_loca, df_abbr, TEMPLATE_FILE, tmp.name)
        tmp.seek(0)
        st.success("Conversion complete! Download your file below.")
        with open(tmp.name, "rb") as f:
            st.download_button(
                label="Download GEO5 Excel File",
                data=f,
                file_name="Geo5_Import.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    os.unlink(tmp.name)

# ---- User Guide Section ----
st.markdown("---")
user_guide_pdf = "UserGuide.pdf"
if os.path.exists(user_guide_pdf):
    pdf_viewer(user_guide_pdf, width=0)  # 0 means full width in streamlit-pdf-viewer
else:
    st.info("User guide not found. Please add 'UserGuide.pdf' to the app folder.")
