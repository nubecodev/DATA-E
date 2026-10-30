
import pandas as pd
from pathlib import Path

# LECCION 1: LEER ARCHIVOS EXCEL CON PANDAS
# Isidora Cortes Cataldo
# Ingeniería en Informática

# 1. Ubicar el archivo Excel
carpeta = Path(__file__).parent
archivo = carpeta / "estudiantes_practica_30.xlsx"

# 2. Leer el archivo Excel
df = pd.read_excel(
    archivo,
    sheet_name="Estudiantes"
)

# 3. Mostrar los primeros 5 estudiantes
print("\n--- PRIMEROS 5 ESTUDIANTES ---")
print(df.head())

# 4. Mostrar los ultimos 5 estudiantes
print("\n--- ULTIMOS 5 ESTUDIANTES ---")
print(df.tail())

# 5. Mostrar la cantidad de filas y columnas
print("\n--- DIMENSIONES DEL EXCEL ---")
print(df.shape)

# 6. Contar los registros
print("\n--- TOTAL DE ESTUDIANTES ---")
print(len(df))

# 7. Mostrar los nombres de las columnas
print("\n--- COLUMNAS DEL EXCEL ---")
print(df.columns.tolist())

# 8. Mostrar informacion general
print("\n--- INFORMACION GENERAL ---")
df.info()

# 9. Mostrar solamente los nombres
print("\n--- NOMBRES DE ESTUDIANTES ---")
print(df["Nombre"])

# 10. Mostrar nombre y carrera
print("\n--- NOMBRE Y CARRERA ---")
print(df[["Nombre", "Carrera"]])

print("\n¡Leccion 1 completada!")
