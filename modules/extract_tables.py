import os
import pandas as pd
from io import BytesIO
import tempfile

try:
    import camelot  # Para PDFs digitales
except ImportError:
    camelot = None

try:
    from docx import Document  # Para Word
except ImportError:
    Document = None


def detectar_fila_encabezado(dataframe):
    """
    Detecta la fila que contiene los encabezados correctos.
    Usa heurísticas simples: más celdas no vacías, presencia de palabras clave comunes.
    """
    palabras_clave = ["código", "enunciado", "denominación", "objetivo", "indicador", "meta"]
    mejor_fila = 0
    mejor_puntaje = 0

    for i, fila in enumerate(dataframe.values):
        texto = [str(x).lower() for x in fila]
        puntaje = sum(1 for x in texto if any(p in x for p in palabras_clave))
        puntaje += sum(1 for x in texto if x.strip() != "")
        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_fila = i

    return mejor_fila


def extraer_tablas(archivo):
    """
    Extrae las tablas OEI y AEI de un archivo PDF o Word del PEI.
    Retorna un diccionario con DataFrames.
    """
    nombre_archivo = archivo.name
    extension = os.path.splitext(nombre_archivo)[1].lower()

    tablas_objetivo = {
        "OEI": ["OEI.0", "Objetivos Estratégicos Institucionales"],
        "AEI": ["AEI.0", "Acciones Estratégicas Institucionales"]
    }

    tablas_encontradas = {}

    # === PDF ===
    if extension == ".pdf":
        if camelot is None:
            raise ImportError("Falta instalar camelot: pip install camelot-py[cv]")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
            tmp.write(archivo.read())
            tmp_path = tmp.name

        try:
            tablas = camelot.read_pdf(tmp_path, pages="all")
        except Exception as e:
            raise RuntimeError(f"Error al leer el PDF con Camelot: {e}")

        for nombre, palabras_clave in tablas_objetivo.items():
            for i, tabla in enumerate(tablas):
                df = tabla.df
                texto_tabla = " ".join(df.astype(str).values.flatten())
                if any(p.lower() in texto_tabla.lower() for p in palabras_clave):
                    fila_header = detectar_fila_encabezado(df)
                    df.columns = df.iloc[fila_header]
                    df = df[fila_header + 1:].reset_index(drop=True)
                    df = df.loc[:, ~df.columns.duplicated()]
                    tablas_encontradas[nombre] = df
                    break

        os.remove(tmp_path)

    # === WORD ===
    elif extension == ".docx":
        if Document is None:
            raise ImportError("Falta instalar python-docx: pip install python-docx")

        doc = Document(archivo)
        for nombre, palabras_clave in tablas_objetivo.items():
            for i, tabla in enumerate(doc.tables):
                try:
                    data = [[celda.text.strip() for celda in fila.cells] for fila in tabla.rows]
                    texto_tabla = " ".join(" ".join(fila) for fila in data)
                    if any(p.lower() in texto_tabla.lower() for p in palabras_clave):
                        df = pd.DataFrame(data)
                        fila_header = detectar_fila_encabezado(df)
                        df.columns = df.iloc[fila_header]
                        df = df[fila_header + 1:].reset_index(drop=True)
                        df = df.loc[:, ~df.columns.duplicated()]
                        tablas_encontradas[nombre] = df
                        break
                except Exception as e:
                    print(f"⚠️ Error al procesar tabla {i}: {e}")

    else:
        raise ValueError(f"Formato de archivo no soportado: {extension}")

    if not tablas_encontradas:
        print("⚠️ No se encontraron tablas OEI o AEI en el documento.")

    return tablas_encontradas
