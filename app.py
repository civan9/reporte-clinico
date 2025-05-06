import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import zipfile
import os
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Generador de Reportes Clínicos", layout="centered")

st.title("📄 Generador Automático de Reportes Clínicos")
st.write("Sube tu archivo Excel con datos de pacientes y genera reportes personalizados en segundos.")

# Cargar archivo Excel
uploaded_excel = st.file_uploader("📤 Sube tu archivo Excel (.xlsx)", type=["xlsx"])

# Plantilla Word
uploaded_template = st.file_uploader("📄 Sube tu plantilla Word (.docx)", type=["docx"])

if uploaded_excel and uploaded_template:
    df = pd.read_excel(uploaded_excel)
    template = DocxTemplate(uploaded_template)

    # Botón para generar reportes
    if st.button("⚙️ Generar Reportes"):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zipf:
            for _, row in df.iterrows():
                context = {
                    "nombre": row["Nombre"],
                    "edad": row["Edad"],
                    "fecha": pd.to_datetime(row["Fecha"]).strftime('%d de %B de %Y'),
                    "diagnostico": row["Diagnóstico"],
                    "tratamiento": row["Tratamiento"]
                }
                template.render(context)
                buffer = BytesIO()
                template.save(buffer)
                buffer.seek(0)
                filename = f"Reporte_{row['Nombre'].replace(' ', '_')}.docx"
                zipf.writestr(filename, buffer.read())
        
        st.success("✅ Reportes generados con éxito.")
        st.download_button(
            label="📥 Descargar ZIP con reportes",
            data=zip_buffer.getvalue(),
            file_name="reportes_clinicos.zip",
            mime="application/zip"
        )
else:
    st.info("Por favor, sube el Excel de pacientes y la plantilla Word para comenzar.")
