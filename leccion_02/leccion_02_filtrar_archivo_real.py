"""Data-E - Leccion 2: buscar y filtrar estudiantes con Pandas.

INSTRUCCIONES:
1. Guarda este .py y tu Excel estudiantes_practica_30(1).xlsx
   en la misma carpeta.
2. Ejecuta en VS Code: python leccion_02_filtrar_archivo_real.py
3. Se creara el archivo reporte_filtrado_data_e.xlsx.

El programa NO modifica el Excel original.

"""
from pathlib import Path
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.workbook.properties import CalcProperties
from openpyxl.utils import get_column_letter

# PARTE 1. LEER EL ARCHIVO EXCEL

carpeta = Path(__file__).resolve().parent.parent
archivo = carpeta / "data" / "estudiantes_practica_30(1).xlsx"

# Tambien admite el nombre del archivo sin "(1)".
if not archivo.exists():
    archivo = carpeta / "data" / "estudiantes_practica_30.xlsx"
if not archivo.exists():
    raise FileNotFoundError(
        "Pon estudiantes_practica_30(1).xlsx en la misma carpeta que este codigo."
    )

df = pd.read_excel(archivo, sheet_name="Estudiantes")
columnas_necesarias = [
    "ID_Estudiante", "Nombre", "Carrera", "Anio_Ingreso",
    "Estado", "Promedio", "Asistencia_Pct", "Beca"
]
faltan = [nombre for nombre in columnas_necesarias if nombre not in df.columns]
if faltan:
    raise ValueError(f"Faltan columnas necesarias: {faltan}")

# PARTE 2. BUSCAR Y FILTRAR (la parte central de esta leccion)

# .loc[] selecciona filas que cumplen una condicion.
# & significa Y; .isin() permite buscar varios valores.

# 1. Activos cuyo promedio es 6,0 o superior.
activos_prom_6 = df.loc[
    (df["Estado"] == "Activo") & (df["Promedio"] >= 6)
].copy()

# 2. Todos los registros de Informatica.
informatica = df.loc[df["Carrera"] == "Informatica"].copy()

# 3. Informatica Y estado Activo.
activos_informatica = df.loc[
    (df["Carrera"] == "Informatica") & (df["Estado"] == "Activo")
].copy()

# 4. Psicologia O Derecho: isin() permite listar varias carreras.
psicologia_derecho = df.loc[
    df["Carrera"].isin(["Psicologia", "Derecho"])
].copy()

# 5. Estudiantes activos con beca.
activos_con_beca = df.loc[
    (df["Estado"] == "Activo") & (df["Beca"] == "Si")
].copy()

# 6. Personas que ingresaron en 2025.
ingreso_2025 = df.loc[df["Anio_Ingreso"] == 2025].copy()

# 7. Buscar "ana" dentro del nombre, sin distinguir mayusculas.
buscar_ana = df.loc[
    df["Nombre"].str.contains("ana", case=False, na=False, regex=False)
].copy()

# 8. Valores vacios en promedio o asistencia.
datos_faltantes = df.loc[
    df[["Promedio", "Asistencia_Pct"]].isna().any(axis=1)
].copy()

# 9. Mostrar TODAS las filas cuyo ID se repite.
# No eliminamos nada: la limpieza se vera en la siguiente leccion.
ids_repetidos = df.loc[
    df.duplicated(subset="ID_Estudiante", keep=False)
].copy()

# Asociamos cada filtro con una hoja del reporte.
filtros = [
    ("Activos prom 6", "Estado = Activo y Promedio >= 6,0", activos_prom_6),
    ("Informatica", "Carrera = Informatica", informatica),
    ("Activos Informatica", "Carrera = Informatica y Estado = Activo", activos_informatica),
    ("Psicologia Derecho", "Carrera en Psicologia o Derecho", psicologia_derecho),
    ("Activos con beca", "Estado = Activo y Beca = Si", activos_con_beca),
    ("Ingreso 2025", "Anio_Ingreso = 2025", ingreso_2025),
    ("Buscar ana", 'Nombre contiene "ana"', buscar_ana),
    ("Datos faltantes", "Promedio o Asistencia_Pct sin dato", datos_faltantes),
    ("IDs repetidos", "ID_Estudiante repetido", ids_repetidos),
]

# PARTE 3. GENERAR EL REPORTE SIN MODIFICAR EL ORIGINAL

salida = carpeta / "leccion_02" / "reporte_filtrado_data_e.xlsx"
with pd.ExcelWriter(salida, engine="openpyxl", date_format="DD/MM/YYYY", datetime_format="DD/MM/YYYY") as writer:
    # Guardamos una copia de la tabla original como referencia.
    df.to_excel(writer, sheet_name="Base original", index=False)

    # Guardamos cada resultado en una hoja independiente.
    for nombre_hoja, criterio, datos in filtros:
        datos.to_excel(writer, sheet_name=nombre_hoja, index=False)
        print(f"{nombre_hoja}: {len(datos)} registros")

    # Resumen con formulas de Excel (recalculadas al abrir el archivo).
    libro = writer.book
    libro.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)
    resumen = libro.create_sheet("Resumen", 0)
    resumen.merge_cells("A1:C1")
    resumen["A1"] = "DATA-E | REPORTE DE BUSQUEDA Y FILTRADO"
    resumen.merge_cells("A2:C2")
    resumen["A2"] = f"Fuente: {archivo.name}  |  Los filtros mantienen los registros originales"
    resumen.append([])  # Fila 3
    resumen.append(["Conjunto de datos", "Condicion del filtro", "N. de registros"])
    resumen.append(["Base original", "Todos los registros cargados", f"=COUNTA('Base original'!A2:A{max(2, len(df)+1)})"])
    for nombre_hoja, criterio, datos in filtros:
        resumen.append([
            nombre_hoja, criterio, f"=COUNTA('{nombre_hoja}'!A2:A{max(2, len(datos)+1)})"
        ])
    resumen["A17"] = "Observaciones"
    resumen.merge_cells("A18:C18")
    resumen["A18"] = "Los conteos incluyen las filas duplicadas. No se elimino ningun registro."
    resumen.merge_cells("A19:C19")
    resumen["A19"] = "Los valores vacios no se reemplazaron: se muestran en Datos faltantes."
    resumen.merge_cells("A20:C20")
    resumen["A20"] = "La hoja Base original es una copia; el Excel de entrada no se modifica."

    # PARTE 4. FORMATO PARA FACILITAR LA LECTURA EN EXCEL
    # Usamos openpyxl para dar formato a las hojas del reporte.

    morado = "514267"
    morado_claro = "EFEAF5"
    verde = "28654D"
    gris = "5B6470"
    naranja_claro = "FFF1DE"
    fuente = "Aptos"

    for hoja in libro:
        hoja.sheet_view.showGridLines = False
        if hoja.title == "Resumen":
            hoja.freeze_panes = "A5"
            hoja.sheet_properties.tabColor = morado
            hoja.row_dimensions[1].height = 32
            hoja.row_dimensions[2].height = 28
            hoja["A1"].font = Font(name=fuente, size=16, bold=True, color="FFFFFF")
            hoja["A1"].fill = PatternFill("solid", fgColor=morado)
            hoja["A1"].alignment = Alignment(vertical="center", indent=1)
            hoja["A2"].font = Font(name=fuente, size=10, italic=True, color=gris)
            for celda in hoja[4]:
                celda.fill = PatternFill("solid", fgColor=morado)
                celda.font = Font(name=fuente, size=11, bold=True, color="FFFFFF")
                celda.alignment = Alignment(vertical="center")
            hoja.row_dimensions[4].height = 24
            for fila in range(5, 15):
                if fila % 2 == 0:
                    for celda in hoja[fila][:3]:
                        celda.fill = PatternFill("solid", fgColor=morado_claro)
                hoja.cell(fila, 1).font = Font(name=fuente, bold=True, color=morado, size=10)
                hoja.cell(fila, 2).font = Font(name=fuente, color=gris, size=10)
                hoja.cell(fila, 3).font = Font(name=fuente, bold=True, color=verde, size=12)
                hoja.cell(fila, 3).number_format = "0"
                hoja.cell(fila, 3).alignment = Alignment(horizontal="center")
                hoja.row_dimensions[fila].height = 23
            hoja["A17"].font = Font(name=fuente, size=12, color=morado, bold=True)
            for fila in [18, 19, 20]:
                hoja.cell(fila, 1).font = Font(name=fuente, size=10, color=gris)
            hoja.column_dimensions["A"].width = 25
            hoja.column_dimensions["B"].width = 52
            hoja.column_dimensions["C"].width = 20
        else:
            hoja.freeze_panes = "A2"
            hoja.auto_filter.ref = hoja.dimensions
            hoja.sheet_properties.tabColor = "63A18C" if hoja.title == "Base original" else morado
            hoja.row_dimensions[1].height = 27
            for celda in hoja[1]:
                celda.fill = PatternFill("solid", fgColor=morado)
                celda.font = Font(name=fuente, size=10, bold=True, color="FFFFFF")
                celda.alignment = Alignment(vertical="center", wrap_text=True)
            for i, fila in enumerate(hoja.iter_rows(min_row=2), start=2):
                hoja.row_dimensions[i].height = 21
                for celda in fila:
                    if i % 2 == 0:
                        celda.fill = PatternFill("solid", fgColor="F7F4FA")
                    celda.font = Font(name=fuente, size=10, color=verde)
                    celda.alignment = Alignment(vertical="center")
                    if celda.value is None and hoja.title == "Datos faltantes":
                        celda.fill = PatternFill("solid", fgColor=naranja_claro)
                if hoja.title == "IDs repetidos":
                    fila[0].fill = PatternFill("solid", fgColor=naranja_claro)
            for col, ancho in {
                "A": 17, "B": 23, "C": 25, "D": 16, "E": 18,
                "F": 12, "G": 16, "H": 14, "I": 19, "J": 12
            }.items():
                hoja.column_dimensions[col].width = ancho
            for fila in hoja.iter_rows(min_row=2):
                hoja.cell(fila[0].row, 5).number_format = "dd/mm/yyyy"
                hoja.cell(fila[0].row, 8).number_format = "0.0"
                hoja.cell(fila[0].row, 9).number_format = '0"%"'

print(f"\nREPORTE CREADO: {salida}")
print(f"Registros en el Excel original: {len(df)}")
print("Nota: los conteos incluyen los duplicados; el archivo original no cambia.")
