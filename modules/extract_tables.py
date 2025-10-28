import os
import pandas as pd
from io import BytesIO

try:
    import camelot  # Para PDFs digitales
except ImportError:
    camelot = None

try:
    from docx import Document  # Para Word
except ImportError:
    Document = None


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

        # Guardar temporalmente el archivo para Camelot
        with open("temp.pdf", "wb") as f:
            f.write(archivo.read())

        tablas = camelot.read_pdf("temp.pdf", pages="all")

        for nombre, palabras_clave in tablas_objetivo.items():
            for i, tabla in enumerate(tablas):
                df = tabla.df
                texto_tabla = " ".join(df.astype(str).values.flatten())
                if any(p.lower() in texto_tabla.lower() for p in palabras_clave):
                    df.columns = df.iloc[0]
                    df = df[1:]
                    tablas_encontradas[nombre] = df
                    break

    # === WORD ===
    elif extension == ".docx":
        if Document is None:
            raise ImportError("Falta instalar python-docx: pip install python-docx")

        doc = Document(archivo)
        for nombre, palabras_clave in tablas_objetivo.items():
            for i, tabla in enumerate(doc.tables):
                data = [[celda.text.strip() for celda in fila.cells] for fila in tabla.rows]
                texto_tabla = " ".join(" ".join(fila) for fila in data)
                if any(p.lower() in texto_tabla.lower() for p in palabras_clave):
                    if len(data) > 1:
                        df = pd.DataFrame(data[1:], columns=data[0])
                    else:
                        df = pd.DataFrame(data)
                    tablas_encontradas[nombre] = df
                    break

    return tablas_encontradas
