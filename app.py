import streamlit as st
import pandas as pd
from ags_to_geo5.ags_parser import load_ags_tables
from ags_to_geo5.exporter import export_to_excel
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
    df_point = ags_tables["POINT"]
    df_loca = ags_tables["LOCA"]
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        export_to_excel(df_geol, df_point, df_loca, TEMPLATE_FILE, tmp.name)
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
