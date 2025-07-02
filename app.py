import streamlit as st
import pandas as pd
from ags_to_geo5.ags_parser import load_ags_tables
from ags_to_geo5.exporter import export_to_excel
from user_guide_utils import word_to_image
import tempfile
import os

st.title("AGS to GEO5 Excel Converter")
st.write("Upload your AGS file and download the GEO5 import Excel file.")

TEMPLATE_FILE = "FieldTestImportTemplate.xlsx"

uploaded_file = st.file_uploader("Choose an AGS file", type=["ags"])

if uploaded_file is not None:
    ags_content = uploaded_file.read().decode("utf-8")
    ags_tables = load_ags_tables(ags_content)
    df_geol = ags_tables["GEOL"]
    df_loca = ags_tables["LOCA"]
    df_abbr = ags_tables["ABBR"]
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
st.header("User Guide")
user_guide_docx = "UserGuide.docx"
output_dir = "user_guide_images"
if os.path.exists(user_guide_docx):
    image_paths = word_to_image(user_guide_docx, output_dir)
    for img_path in image_paths:
        st.image(img_path, use_column_width=True)
else:
    st.info("User guide not found. Please add 'UserGuide.docx' to the app folder.")
