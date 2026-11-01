
import pandas as pd
from pathlib import Path

# LECCION 1: LEER ARCHIVOS EXCEL CON PANDAS
# Isidora Cortes Cataldo
# Ingeniería en Informática

# 1. OBTENER LA RUTA DEL ARCHIVO
# __file__ representa el archivo Python actual.
# parent.parent permite llegar a la carpeta DATA-E.

carpeta = Path(__file__).resolve().parent.parent
archivo = carpeta / "data" / "estudiantes_practica_30.xlsx"

# 2. Leer el archivo Excel
# read_excel() carga los datos en un DataFrame.
# sheet_name indica la hoja que queremos leer.

df = pd.read_excel(
    archivo,
    sheet_name="Estudiantes"
)

# 3. Mostrar los primeros 5 estudiantes
print("\n--- PRIMEROS 5 ESTUDIANTES ---")
print(df.head().to_string(index=False))

# 4. Mostrar los ultimos 5 estudiantes
print("\n--- ULTIMOS 5 ESTUDIANTES ---")
print(df.tail().to_string(index=False))

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
