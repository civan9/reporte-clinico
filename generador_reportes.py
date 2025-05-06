import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os

# Cargar Excel
df = pd.read_excel("pacientes_ejemplo.xlsx")

# Cargar plantilla Word
template = DocxTemplate("plantilla_reporte.docx")

# Crear carpeta de salida
os.makedirs("reportes_generados", exist_ok=True)

# Generar reportes
for _, row in df.iterrows():
    context = {
        "nombre": row["Nombre"],
        "edad": row["Edad"],
        "fecha": pd.to_datetime(row["Fecha"]).strftime('%d de %B de %Y'),
        "diagnostico": row["Diagnóstico"],
        "tratamiento": row["Tratamiento"]
    }
    template.render(context)
    filename = f"Reporte_{row['Nombre'].replace(' ', '_')}.docx"
    template.save(os.path.join("reportes_generados", filename))

print("✅ Reportes generados correctamente.")
