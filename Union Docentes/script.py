import pandas as pd
import os
import unicodedata

# Cargar los archivos Excel y las hojas correspondientes

script_dir = os.path.dirname(os.path.abspath(__file__))
df1 = pd.read_excel(os.path.join(script_dir, "docentes.xlsx"), sheet_name="list")
df2 = pd.read_excel(os.path.join(script_dir, "LISTA OFICIAL DOCENTES.xlsx"), sheet_name="LISTA DOCENTES 25")

# Normalizar los nombres de los docentes para quitar tildes y hacer el merge correctamente
def normalize_str(s):
	if pd.isnull(s):
		return s
	return ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn').lower().strip()

df1["Docente_normalizado"] = df1["Docente"].apply(normalize_str)
df2["Docente_normalizado"] = df2["Docente"].apply(normalize_str)

# Seleccionar las columnas relevantes de df2
cols_to_merge = ["Docente", "Sede", "Nro. Documento", "Celular", "Direccion", "Correo Institucional", "Correo personal", "Docente_normalizado", "Categoria (Monto)", "Categoria (Letra)"]
df2_subset = df2[cols_to_merge]

# Unir los datos según el nombre del docente normalizado
df_merged = df1.drop(columns=["Nro. Documento", "Sede", "Celular", "Direccion", "Correo Institucional", "Correo personal", "Categoria (Monto)", "Categoria (Letra)"], errors='ignore')
df_merged = df_merged.merge(df2_subset, on="Docente_normalizado", how="left", suffixes=('', '_nuevo'))

# Opcional: reemplazar los datos antiguos por los nuevos si existen
for col in ["Nro. Documento", "Sede", "Celular", "Direccion", "Correo Institucional", "Correo personal", "Docente_normalizado", "Categoria (Monto)", "Categoria (Letra)"]:
	df_merged[col] = df_merged[col].combine_first(df_merged.get(col))

# Eliminar columnas auxiliares
df_merged = df_merged.drop(columns=[c for c in df_merged.columns if c.endswith("_nuevo") or c == "Docente_normalizado"])

# Guardar el resultado en un nuevo archivo Excel
df_merged.to_excel("docentes_actualizado.xlsx", index=False)