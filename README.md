# DATA-E | Python, Pandas y automatización de datos

Bienvenido/a al proyecto del ramo de Responsabilidad Social Empresarial y Ética Laboral en donde nuestro objetivo es ayudar a **DATA-E**. Este es un proyecto educativo gratuito con lecciones y ejercicios prácticos para aprender a trabajar con archivos Excel mediante Python: Pandas, OpenPyXL, NumPy.

Está pensado para personas que están comenzando a programar: puedes descargar los ejemplos, ejecutarlos y modificarlos en tu propio computador mientras aprendes.

## Objetivo

Facilitar el aprendizaje, paso a paso, de herramientas de análisis y automatización de datos. A través de las lecciones podrás practicar tareas como leer planillas Excel, explorar registros y, a medida que avance el proyecto, filtrar información, detectar duplicados y generar reportes.

## Contenido del repositorio

```text
DATA-E/
├── data/
│   ├── estudiantes_practica_30.xlsx
├── leccion_01/
│   ├── leccion_01_pdf.pdf
│   └── leccion_01.py
├── leccion_02/
│   ├── leccion_02_filtrar_archivo_real.py
│   ├── leccion_02_pdf.pdf
│   └── reporte_filtrado_data_e.xlsx
├── Instalacion_Entorno_Trabajo_DATA_E.pdf
├── LICENSE-CODE.md
├── LICENSE-CONTENT.md
└── README.md
```

**Lección 1 — Lectura de archivos Excel con Pandas:** utilizarás el archivo `estudiantes_practica_30.xlsx` para cargar datos con `pd.read_excel()` y visualizar registros, por ejemplo con `df.head()` y `df.tail()`.

### Lección 2 — Búsqueda y filtrado de datos con Pandas
En esta lección aprenderás a extraer información precisa de tus DataFrames aplicando filtros lógicos y métodos avanzados de Pandas:
* **Filtrado por condiciones múltiples (`&`):** Para combinar criterios como estado y promedio mínimo.
* **Búsqueda por listas (`.isin()`):** Ideal para filtrar múltiples categorías (ej. Psicología o Derecho).
* **Búsqueda de texto (`.str.contains()`):** Para localizar patrones o nombres específicos (ej. "ana") sin distinguir mayúsculas.
* **Control de calidad de datos:** Detección de valores faltantes (`.isna()`) y registros duplicados (`duplicated()`).
* **Exportación profesional:** Creación de un reporte multipestaña con `pd.ExcelWriter` y formato personalizado usando `openpyxl`.

> La estructura se ampliará a medida que se publiquen nuevas lecciones. Los archivos de licencia indicados arriba deben incorporarse a la raíz del repositorio.

---

## Guía de instalación del entorno de trabajo (Windows)

Esta guía está adaptada de **«Guía de Instalación del Entorno de Trabajo»**, elaborada por **Isidora Cortes** (29 de octubre de 2026). Sigue el mismo recorrido: instalar Python, comprobar su funcionamiento, instalar Pandas/NumPy/OpenPyXL, instalar Visual Studio Code y configurar su extensión de Python.

### 1. Instalar Python

1. Ingresa a la [página oficial de descargas de Python para Windows](https://www.python.org/downloads/windows/).
2. Descarga el **Python Install Manager** para Windows (formato **MSIX**), tal como se muestra en la guía de instalación.
3. Abre el archivo descargado y selecciona **Install / Instalar**.
4. Sigue las instrucciones que aparezcan en la ventana de configuración. Es posible que se soliciten respuestas como `Y` (sí) o `N` (no) para configurar rutas de acceso o instalar una versión de Python: **lee cada pregunta antes de responder**, pues puede variar según tu equipo.
5. Instala la versión estable de Python disponible. La guía original utiliza **Python 3.15.2**.

**Comprobar la instalación:** presiona `Windows + R`, escribe `cmd`, pulsa `Enter` y ejecuta:

```powershell
python --version
```

Si instalaste la versión de la guía, verás un resultado similar a:

```text
Python 3.15.2
```

Si aparece otra versión estable de Python 3, el número mostrado puede ser distinto.

### 2. Instalar Pandas, NumPy y OpenPyXL

En la misma terminal (`cmd`), ejecuta:

```powershell
python -m pip install pandas numpy openpyxl
```

Estas bibliotecas se utilizan para:

| Biblioteca | ¿Para qué la utilizaremos? |
| --- | --- |
| **Pandas** | Leer, organizar, filtrar y analizar tablas de datos. |
| **NumPy** | Realizar operaciones numéricas y trabajar con arreglos. |
| **OpenPyXL** | Leer y escribir archivos de Excel con formato `.xlsx`. |

Una vez terminada la instalación, verifica que las tres puedan importarse:

```powershell
python -c "import pandas, numpy, openpyxl; print('Data-E: bibliotecas funcionando correctamente')"
```

Si todo está instalado correctamente, la terminal mostrará:

```text
Data-E: bibliotecas funcionando correctamente
```

### 3. Descargar e instalar Visual Studio Code

**Visual Studio Code (VS Code)** es el editor en el que escribiremos y ejecutaremos los ejercicios. **VS Code no incluye Python**: ambos programas se instalan por separado.

1. Abre el [sitio oficial de Visual Studio Code](https://code.visualstudio.com/).
2. Selecciona **Download for Windows** y descarga la versión estable. Se recomienda **User Installer**, que normalmente no necesita permisos de administrador.
3. Abre el instalador descargado (`VSCodeUserSetup-x64-...exe`, o un nombre parecido).
4. Lee el acuerdo de licencia y, si estás de acuerdo, acéptalo. Haz clic en **Next / Siguiente**.
5. Puedes mantener la ubicación de instalación predeterminada.
6. En las tareas adicionales, se recomienda activar **Add to PATH**, si aparece. Opcionalmente, marca **Create a desktop icon** y **Open with Code**, si están disponibles.
7. Selecciona **Install / Instalar** y, al terminar, **Finish / Finalizar** para abrir VS Code.

### 4. Instalar la extensión de Python en VS Code

1. Abre **Visual Studio Code**.
2. En el menú lateral, selecciona **Extensions / Extensiones** (o presiona `Ctrl + Shift + X`).
3. Busca **Python** y selecciona la extensión publicada por **Microsoft**.
4. Haz clic en **Install / Instalar**.
5. Si aparecen las opciones **Disable / Deshabilitar** y **Uninstall / Desinstalar**, significa que la extensión está instalada.

### 5. Seleccionar el intérprete de Python

1. Presiona `Ctrl + Shift + P` en VS Code.
2. Escribe **`Python: Select Interpreter`**.
3. Selecciona la instalación de Python que configuraste anteriormente (por ejemplo, **Python 3.15.2**).

Con este paso, VS Code sabrá qué instalación de Python utilizar para ejecutar tus programas.

### 6. Descargar y ejecutar la primera lección

1. En la página principal del repositorio de GitHub, selecciona **Code → Download ZIP**. Descomprime el archivo en una carpeta de tu computador.
2. En VS Code, selecciona **File / Archivo → Open Folder / Abrir carpeta** y abre la carpeta `DATA-E`.
3. Comprueba que `Leccion_uno` contenga **ambos archivos**: `leccion_1.py` y `estudiantes_practica_30.xlsx`.
4. Abre la terminal integrada desde **Terminal → New Terminal / Nueva terminal**.
5. Desde la carpeta principal `DATA-E`, ejecuta:

```powershell
python Leccion_uno/leccion_1.py
```

El programa leerá el Excel y mostrará información de los estudiantes en la terminal. Si en el código aparece `sheet_name="Estudiantes"`, esa hoja debe existir en el archivo Excel utilizado.

### Problemas frecuentes

| Problema | Qué revisar |
| --- | --- |
| `python` no se reconoce como comando | Comprueba la instalación de Python y su disponibilidad en la terminal. Cierra y abre otra terminal después de instalarlo. |
| `ModuleNotFoundError: No module named 'pandas'` | Vuelve a ejecutar `python -m pip install pandas numpy openpyxl` con el Python seleccionado en VS Code. |
| No aparece Python en VS Code | Verifica que instalaste la extensión **Python** de Microsoft y vuelve a usar **Python: Select Interpreter**. |
| No se encuentra el Excel | Comprueba que el archivo `.xlsx` esté dentro de `Leccion_uno` y que su nombre coincida con el utilizado en `leccion_1.py`. |

---

## Licencias del proyecto

DATA-E es un proyecto educativo gratuito para aprender Python, Pandas y automatización de datos.

### Código fuente

Los archivos `.py` están disponibles bajo **PolyForm Noncommercial 1.0.0**.

Consulta [LICENSE-CODE.md](LICENSE-CODE.md).

### Material educativo

Las guías, explicaciones, presentaciones y archivos Excel de práctica originales están disponibles bajo **Creative Commons CC BY-NC 4.0**.

Consulta [LICENSE-CONTENT.md](LICENSE-CONTENT.md).

### Condiciones de uso

Se permite descargar, estudiar, ejecutar, adaptar y compartir los materiales conforme a sus respectivas licencias y únicamente para los usos autorizados.

Los usos comerciales requieren autorización adicional del titular de los derechos.

© 2026 Isidora Cortes Cataldo.
