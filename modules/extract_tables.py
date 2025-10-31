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


def _clean_cell(s):
    """Normaliza una celda a string simple para comparar."""
    if pd.isna(s):
        return ""
    return str(s).strip().lower()


def _combinar_encabezado(df, start_row, height):
    """
    Combina 'height' filas empezando en start_row para formar un nombre por cada columna.
    Devuelve lista de strings (una por columna), donde cada string es la concatenación
    de las celdas no vacías (separadas por espacio).
    """
    ncols = df.shape[1]
    combined = []
    for col in range(ncols):
        parts = []
        for r in range(start_row, start_row + height):
            if r >= len(df):
                break
            cell = _clean_cell(df.iloc[r, col])
            if cell:
                parts.append(cell)
        combined_name = " ".join(parts) if parts else ""
        combined.append(combined_name)
    return combined


def _score_encabezado(combined_cells, palabras_clave):
    """
    Calcula un puntaje para una lista de nombres combinados de columnas:
    - +10 por aparición de palabra clave en una celda (para priorizar filas con términos relevantes)
    - +1 por cada celda no vacía
    """
    puntaje = 0
    for cell in combined_cells:
        if not cell:
            continue
        # contar palabras clave
        for pk in palabras_clave:
            if pk in cell:
                puntaje += 10
        # contar no vacías
        puntaje += 1
    return puntaje


def detectar_encabezado_multi(
    df,
    palabras_clave=None,
    max_start_row=4,
    max_header_height=3,
    require_min_non_empty=1,
):
    """
    Detecta el encabezado permitiendo:
    - buscar el inicio entre las primeras `max_start_row` filas (0-based)
    - combinar hasta `max_header_height` filas contiguas como encabezado

    Retorna (start_row, height, combined_header_list) del mejor candidato elegido.
    Si no encuentra un candidato válido devuelve (None, None, None).

    Parámetros:
    - df: DataFrame leído (sin interpretar encabezado)
    - palabras_clave: lista de strings para dar prioridad a filas que contengan términos relevantes
    - max_start_row: buscar inicio del header entre filas [0..max_start_row-1]
    - max_header_height: máximo número de filas a combinar
    - require_min_non_empty: mínimo de celdas no vacías en el encabezado combinado para aceptarlo
    """
    if palabras_clave is None:
        palabras_clave = ["denominación", "acción", "objetivo", "indicador", "meta"]

    n_rows = len(df)
    if n_rows == 0:
        return None, None, None

    max_start = min(n_rows, max_start_row)
    best = None  # (start, height, combined, score, non_empty_count)
    for start in range(0, max_start):
        max_height = min(max_header_height, n_rows - start)
        for height in range(1, max_height + 1):
            combined = _combinar_encabezado(df, start, height)
            non_empty = sum(1 for c in combined if c.strip() != "")
            if non_empty < require_min_non_empty:
                # no tiene suficientes celdas con texto, saltar
                continue
            score = _score_encabezado(combined, palabras_clave)
            # preferir mayor score, en empate preferir menor start (más arriba) y mayor non_empty
            if best is None or score > best[3] or (score == best[3] and non_empty > best[4]) or (score == best[3] and non_empty == best[4] and start < best[0]):
                best = (start, height, combined, score, non_empty)

    if best is None:
        return None, None, None

    return best[0], best[1], best[2]


def _apply_header_and_clean(df, start_row, height, combined_header):
    """
    Aplica el encabezado combinado al DataFrame:
    - Elimina las filas del encabezado (start_row .. start_row+height-1)
    - Asigna columnas con los nombres combinados (si hay nombre vacío genera col_{i})
    - Ajusta columnas si el número de columnas del resto no coincide
    - Elimina columnas completamente vacías
    """
    # Crear nombres seguros
    new_cols = []
    for i, name in enumerate(combined_header):
        if name and name.strip() != "":
            safe = name.strip().lower().replace(" ", "_")
            new_cols.append(safe)
        else:
            new_cols.append(f"col_{i}")

    # Tomar el resto del df debajo del header combinado
    df_tail = df.iloc[start_row + height :].copy().reset_index(drop=True)

    # Alinear número de columnas
    needed = len(new_cols)
    current = df_tail.shape[1]
    if current > needed:
        df_tail = df_tail.iloc[:, :needed]
    elif current < needed:
        for j in range(current, needed):
            df_tail[f"__pad_{j}"] = ""

    df_tail.columns = new_cols

    # Eliminar columnas vacías
    df_tail = df_tail.dropna(axis=1, how="all")

    return df_tail


def extraer_tablas(archivo):
    """
    Extrae las tablas OEI y AEI de un archivo PDF o Word del PEI.
    Retorna un diccionario con DataFrames.
    """
    nombre_archivo = archivo.name
    extension = os.path.splitext(nombre_archivo)[1].lower()

    tablas_objetivo = {
        "OEI": ["OEI.0", "Objetivos Estratégicos Institucionales", "objetivo estratégico"],
        "AEI": ["AEI.0", "Acciones Estratégicas Institucionales", "acción estratégica", "acción"]
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
            # limpiar el tmp y propagar
            try:
                os.remove(tmp_path)
            except Exception:
                pass
            raise RuntimeError(f"Error al leer el PDF con Camelot: {e}")

        for nombre, palabras_clave in tablas_objetivo.items():
            for i, tabla in enumerate(tablas):
                df = tabla.df
                texto_tabla = " ".join(df.astype(str).values.flatten())
                if any(p.lower() in texto_tabla.lower() for p in palabras_clave):
                    # detectar encabezado multi-fila
                    start, height, combined = detectar_encabezado_multi(
                        df,
                        palabras_clave=palabras_clave,
                        max_start_row=4,
                        max_header_height=3,
                        require_min_non_empty=1,
                    )
                    if start is None:
                        # fallback: usar heurística simple original (fila con más texto/no vacío)
                        fila_header = detectar_fila_encabezado_simple(df)
                        df.columns = df.iloc[fila_header]
                        df = df[fila_header + 1 :].reset_index(drop=True)
                        df = df.loc[:, ~df.columns.duplicated()]
                        tablas_encontradas[nombre] = df
                    else:
                        df_clean = _apply_header_and_clean(df, start, height, combined)
                        tablas_encontradas[nombre] = df_clean
                    break

        try:
            os.remove(tmp_path)
        except Exception:
            pass

    # === WORD ===
    elif extension == ".docx":
        if Document is None:
            raise ImportError("Falta instalar python-docx: pip install python-docx")

        # Si 'archivo' es un UploadedFile de Streamlit, Document acepta file-like. Si es path, usar Document(path)
        doc = Document(archivo)
        for nombre, palabras_clave in tablas_objetivo.items():
            for i, tabla in enumerate(doc.tables):
                try:
                    data = [[celda.text.strip() for celda in fila.cells] for fila in tabla.rows]
                    df = pd.DataFrame(data)
                    texto_tabla = " ".join(" ".join(fila) for fila in data)
                    if any(p.lower() in texto_tabla.lower() for p in palabras_clave):
                        start, height, combined = detectar_encabezado_multi(
                            df,
                            palabras_clave=palabras_clave,
                            max_start_row=4,
                            max_header_height=3,
                            require_min_non_empty=1,
                        )
                        if start is None:
                            fila_header = detectar_fila_encabezado_simple(df)
                            df.columns = df.iloc[fila_header]
                            df = df[fila_header + 1 :].reset_index(drop=True)
                            df = df.loc[:, ~df.columns.duplicated()]
                            tablas_encontradas[nombre] = df
                        else:
                            df_clean = _apply_header_and_clean(df, start, height, combined)
                            tablas_encontradas[nombre] = df_clean
                        break
                except Exception as e:
                    # no bloquear toda la extracción si una tabla falla
                    print(f"⚠️ Error al procesar tabla {i}: {e}")

    else:
        raise ValueError(f"Formato de archivo no soportado: {extension}")

    if not tablas_encontradas:
        print("⚠️ No se encontraron tablas OEI o AEI en el documento.")

    return tablas_encontradas


# --- Auxiliar: tu heurística original aislada como fallback ---
def detectar_fila_encabezado_simple(dataframe):
    """
    Tu función original para detectar una fila 'mejor' por heurística simple.
    Devuelve el índice de fila elegido.
    """
    palabras_clave = ["denominación", "acción", "objetivo", "indicador", "meta"]
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
