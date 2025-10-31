"""
Extracción y detección de encabezados (incluyendo multi-fila) para tablas OEI/AEI
Soporta mapeo/normalización de nombres de columnas a una lista de columnas esperadas
usando coincidencia exacta, alias y coincidencia difusa (rapidfuzz si está disponible,
sino difflib).

Uso:
- extraer_tablas(file_like) -> dict con DataFrames detectados (claves "OEI" y/o "AEI")
- Ajusta los parámetros EXPECTED_COLUMNS y alias en la sección de configuración
  según tus necesidades.
"""

import os
import pandas as pd
import tempfile
import unicodedata
import re

try:
    import camelot  # Para PDFs digitales
except Exception:
    camelot = None

try:
    from docx import Document  # Para Word
except Exception:
    Document = None

# Intentar rapidfuzz para fuzzy matching (más robusto y rápido)
try:
    from rapidfuzz import process, fuzz
    _HAS_RAPIDFUZZ = True
except Exception:
    _HAS_RAPIDFUZZ = False
    import difflib

# -----------------------------
# Configuración: columnas esperadas
# -----------------------------
EXPECTED_COLUMNS_RAW = [
    "Código", "Denominación de OEI", "Indicador", "Denominación",
    "OBJETIVOS ESTRATÉGICOS INSTITUCIONALES", "NOMBRE DEL INDICADOR",
    "ACCIONES ESTRATÉGICAS INSTITUCIONALES", "CODIGO", "OEI", "AEI",
    "OBJETIVOS ESTRATÉGICOS INSTITUCIONAL", "Denominación del OEI/AEI",
    "Denominación del OEI / AEI", "INDICADOR", "Descripción",
    "Nombre del Indicador", "OEI/AEI", "Nombre", "Enunciado y Indicador"
]

# Puedes añadir aliases adicionales si hay formas comunes que no están en EXPECTED_COLUMNS_RAW
ALIAS_MAP_RAW = {
    # ejemplos: variantes frecuentes -> canonical (se normalizarán internamente)
    "cod": "Código",
    "codigo": "Código",
    "denominación": "Denominación",
    "denominacion": "Denominación",
    "denominacion del oei/aei": "Denominación del OEI/AEI",
    "denominación del oei / aei": "Denominación del OEI / AEI",
    "nombre del indicador": "NOMBRE DEL INDICADOR",
    "nombre indicador": "NOMBRE DEL INDICADOR",
    "objetivos estrategicos institucionales": "OBJETIVOS ESTRATÉGICOS INSTITUCIONALES",
    "objetivo estrategico": "OBJETIVOS ESTRATÉGICOS INSTITUCIONALES",
    "acciones estrategicas institucionales": "ACCIONES ESTRATÉGICAS INSTITUCIONALES",
    "enunciado e indicador": "Enunciado y Indicador",
    # añade más según tus documentos
}

# -----------------------------
# Utilidades de normalización y mapeo
# -----------------------------
def _normalize_text(s: str) -> str:
    """Normaliza texto: lower, strip, quita tildes/diacríticos, colapsa espacios y sustituye símbolos."""
    if s is None:
        return ""
    s = str(s)
    s = s.strip().lower()
    # reemplazar ciertos separadores por espacio
    s = re.sub(r"[/\\\-_]+", " ", s)
    # quitar acentos
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    # colapsar espacios
    s = re.sub(r"\s+", " ", s)
    return s.strip()


# Preparar listados normalizados de columnas esperadas y aliases
_CANONICAL_EXPECTED = []
_CANONICAL_TO_RAW = {}
for raw in EXPECTED_COLUMNS_RAW:
    norm = _normalize_text(raw)
    # mantener primera aparición como canonical raw representative
    if norm not in _CANONICAL_TO_RAW:
        _CANONICAL_TO_RAW[norm] = raw
        _CANONICAL_EXPECTED.append(norm)

_ALIAS_MAP_PROC = {}
for k, v in ALIAS_MAP_RAW.items():
    _AL = _normalize_text(k)
    _VAL = _normalize_text(v)
    # si alias apunta a raw presente, respetarlo. Sino apuntar a normalized v.
    if _VAL in _CANONICAL_TO_RAW:
        _ALIAS_MAP_PROC[_AL] = _VAL
    else:
        _ALIAS_MAP_PROC[_AL] = _VAL

def _best_fuzzy_match(target: str, candidates: list) -> tuple:
    """Devuelve (best_candidate_norm, score_int) usando rapidfuzz si está,
    sino difflib SequenceMatcher convertido a [0..100]."""
    if not target:
        return None, 0
    if _HAS_RAPIDFUZZ:
        match = process.extractOne(target, candidates, scorer=fuzz.ratio)
        if match:
            return match[0], int(match[1])
        return None, 0
    else:
        best = None
        best_score = 0.0
        for c in candidates:
            score = difflib.SequenceMatcher(None, target, c).ratio()
            if score > best_score:
                best_score = score
                best = c
        return best, int(best_score * 100)


def map_combined_to_expected(
    combined_header: list,
    expected_norm_list: list = None,
    alias_map_proc: dict = None,
    match_threshold: int = 80,
):
    """
    Mapea cada nombre combinado a una columna esperada normalizada (o None si no hay match).
    - combined_header: lista de strings (ya lowercase o no)
    - expected_norm_list: lista de expected normalizados (si None usa _CANONICAL_EXPECTED)
    - alias_map_proc: diccionario normalized_alias -> normalized_expected (si None usa _ALIAS_MAP_PROC)
    Returns: mapped_list (same length), matches_count, score_sum
    """
    if expected_norm_list is None:
        expected_norm_list = _CANONICAL_EXPECTED
    if alias_map_proc is None:
        alias_map_proc = _ALIAS_MAP_PROC

    mapped = []
    matches = 0
    score_sum = 0

    for cell in combined_header:
        norm_cell = _normalize_text(cell)
        if not norm_cell:
            mapped.append((cell, None, 0))
            continue
        # 1) alias exacto
        if norm_cell in alias_map_proc:
            mapped_to_norm = alias_map_proc[norm_cell]
            mapped.append((cell, mapped_to_norm, 100))
            matches += 1
            score_sum += 100
            continue
        # 2) exacto con expected
        if norm_cell in expected_norm_list:
            mapped.append((cell, norm_cell, 100))
            matches += 1
            score_sum += 100
            continue
        # 3) fuzzy match
        best, score = _best_fuzzy_match(norm_cell, expected_norm_list)
        if best and score >= match_threshold:
            mapped.append((cell, best, score))
            matches += 1
            score_sum += score
        else:
            mapped.append((cell, None, score))

    return mapped, matches, score_sum


# -----------------------------
# Detección de encabezado multi-fila
# -----------------------------
def _clean_cell(s):
    if pd.isna(s):
        return ""
    return str(s).strip().lower()


def _combinar_encabezado(df, start_row, height):
    """Combina height filas empezando en start_row -> lista de strings por columna"""
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
    """Puntaje simple: +10 por palabra clave en celda, +1 por celda no vacía"""
    puntaje = 0
    for cell in combined_cells:
        if not cell:
            continue
        for pk in palabras_clave:
            if pk in cell:
                puntaje += 10
        puntaje += 1
    return puntaje


def detectar_encabezado_multi(
    df,
    palabras_clave=None,
    max_start_row=4,
    max_header_height=3,
    require_min_non_empty=1,
    # parámetros nuevos para usar expected columns en la elección
    use_expected_for_selection=False,
    match_threshold_for_selection=80,
):
    """
    Detecta encabezado multi-fila y devuelve start_row, height, combined_header, mapped_info

    Si use_expected_for_selection=True, además de la heurística de palabras_clave se comprobará
    cuántas columnas del header combinado se mapean a las columnas esperadas (fuzzy).
    Esto ayuda a elegir mejor cuando hay varias filas similares.
    """
    if palabras_clave is None:
        palabras_clave = ["denominación", "acción", "objetivo", "indicador", "meta"]

    n_rows = len(df)
    if n_rows == 0:
        return None, None, None, None

    max_start = min(n_rows, max_start_row)
    best = None  # (start, height, combined, score, non_empty, mapped_matches, mapped_score_sum)

    for start in range(0, max_start):
        max_height = min(max_header_height, n_rows - start)
        for height in range(1, max_height + 1):
            combined = _combinar_encabezado(df, start, height)
            non_empty = sum(1 for c in combined if c.strip() != "")
            if non_empty < require_min_non_empty:
                continue
            score = _score_encabezado(combined, palabras_clave)

            mapped_matches = 0
            mapped_score_sum = 0
            if use_expected_for_selection:
                mapped, mapped_matches, mapped_score_sum = map_combined_to_expected(
                    combined,
                    expected_norm_list=_CANONICAL_EXPECTED,
                    alias_map_proc=_ALIAS_MAP_PROC,
                    match_threshold=match_threshold_for_selection,
                )

            # Heurística de selección:
            # preferir mayor mapped_matches, luego mayor score, luego mayor non_empty, luego menor start.
            cmp_key = (mapped_matches if use_expected_for_selection else 0, score, non_empty, -start)
            if best is None:
                best = (start, height, combined, score, non_empty, mapped_matches, mapped_score_sum)
            else:
                best_key = (best[5] if use_expected_for_selection else 0, best[3], best[4], -best[0])
                if cmp_key > best_key:
                    best = (start, height, combined, score, non_empty, mapped_matches, mapped_score_sum)

    if best is None:
        return None, None, None, None

    # devolver también el mapping para el mejor candidate (con umbral por defecto)
    mapped_info, mmatches, mscore = map_combined_to_expected(
        best[2],
        expected_norm_list=_CANONICAL_EXPECTED,
        alias_map_proc=_ALIAS_MAP_PROC,
        match_threshold=match_threshold_for_selection,
    )

    return best[0], best[1], best[2], mapped_info


# -----------------------------
# Aplicar header y limpiar + renombrar usando mapeo
# -----------------------------
def _safe_col_name_from_norm(norm):
    """Convierte nombre normalizado a un identificador de columna seguro (snake_case)."""
    if not norm:
        return ""
    s = norm.strip().lower()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^\w_]", "", s)
    return s


def _apply_header_and_clean(df, start_row, height, combined_header, mapped_info=None):
    """
    Aplica el encabezado combinado al DataFrame y renombra columnas:
    - Si mapped_info (lista de tuples (orig, mapped_norm, score)) está presente,
      usa mapped_norm para nombrar columnas cuando exista.
    - Si mapped_norm es None, genera nombre seguro a partir del texto combinado.
    """
    # Preparar nombres resultantes
    new_cols = []
    if mapped_info is None:
        # fallback: usar combined_header directo
        for i, name in enumerate(combined_header):
            if name and name.strip():
                new_cols.append(_safe_col_name_from_norm(_normalize_text(name)))
            else:
                new_cols.append(f"col_{i}")
    else:
        # mapped_info es lista de tuplas (orig, mapped_norm, score)
        for i, (orig, mapped_norm, score) in enumerate(mapped_info):
            if mapped_norm:
                # obtener raw representative si existe
                raw = _CANONICAL_TO_RAW.get(mapped_norm, mapped_norm)
                new_cols.append(_safe_col_name_from_norm(mapped_norm))
            else:
                if orig and orig.strip():
                    new_cols.append(_safe_col_name_from_norm(_normalize_text(orig)))
                else:
                    new_cols.append(f"col_{i}")

    # Tomar resto del df
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


# -----------------------------
# Extracción principal (PDF / DOCX)
# -----------------------------
def extraer_tablas(archivo, match_threshold=80, use_expected_for_selection=True):
    """
    Extrae las tablas OEI y AEI de un archivo PDF o Word del PEI.
    Parámetros:
    - archivo: file-like (e.g., Streamlit UploadedFile) o ruta con atributo .name
    - match_threshold: umbral (0-100) para aceptar coincidencias difusas al mapear encabezados
    - use_expected_for_selection: si True, tener en cuenta cuántas columnas del header mapean
      a las columnas esperadas al elegir el mejor candidato de header.

    Retorna diccionario con DataFrames encontrados: claves 'OEI' y/o 'AEI'.
    """
    nombre_archivo = getattr(archivo, "name", str(archivo))
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
                    start, height, combined, mapped_info = detectar_encabezado_multi(
                        df,
                        palabras_clave=palabras_clave,
                        max_start_row=4,
                        max_header_height=3,
                        require_min_non_empty=1,
                        use_expected_for_selection=use_expected_for_selection,
                        match_threshold_for_selection=match_threshold,
                    )
                    if start is None:
                        # fallback simple
                        fila_header = detectar_fila_encabezado_simple(df)
                        df.columns = df.iloc[fila_header]
                        df = df[fila_header + 1 :].reset_index(drop=True)
                        df = df.loc[:, ~df.columns.duplicated()]
                        tablas_encontradas[nombre] = df
                    else:
                        # mapped_info puede ser una lista de tuples
                        # Si mapped_info no contiene match_threshold en su cálculo, recalcular con threshold proporcionado
                        # mapped_info devuelve (orig, mapped_norm, score) por columna
                        # pero la función ya usó match_threshold_for_selection
                        df_clean = _apply_header_and_clean(df, start, height, combined, mapped_info=mapped_info)
                        tablas_encontradas[nombre] = df_clean
                    break

        try:
            os.remove(tmp_path)
        except Exception:
            pass

    # === WORD (.docx) ===
    elif extension == ".docx":
        if Document is None:
            raise ImportError("Falta instalar python-docx: pip install python-docx")

        doc = Document(archivo)
        for nombre, palabras_clave in tablas_objetivo.items():
            for i, tabla in enumerate(doc.tables):
                try:
                    data = [[celda.text.strip() for celda in fila.cells] for fila in tabla.rows]
                    df = pd.DataFrame(data)
                    texto_tabla = " ".join(" ".join(fila) for fila in data)
                    if any(p.lower() in texto_tabla.lower() for p in palabras_clave):
                        start, height, combined, mapped_info = detectar_encabezado_multi(
                            df,
                            palabras_clave=palabras_clave,
                            max_start_row=4,
                            max_header_height=3,
                            require_min_non_empty=1,
                            use_expected_for_selection=use_expected_for_selection,
                            match_threshold_for_selection=match_threshold,
                        )
                        if start is None:
                            fila_header = detectar_fila_encabezado_simple(df)
                            df.columns = df.iloc[fila_header]
                            df = df[fila_header + 1 :].reset_index(drop=True)
                            df = df.loc[:, ~df.columns.duplicated()]
                            tablas_encontradas[nombre] = df
                        else:
                            df_clean = _apply_header_and_clean(df, start, height, combined, mapped_info=mapped_info)
                            tablas_encontradas[nombre] = df_clean
                        break
                except Exception as e:
                    print(f"⚠️ Error al procesar tabla {i}: {e}")

    else:
        raise ValueError(f"Formato de archivo no soportado: {extension}")

    if not tablas_encontradas:
        print("⚠️ No se encontraron tablas OEI o AEI en el documento.")

    return tablas_encontradas


# --- Auxiliar: tu heurística original aislada como fallback ---
def detectar_fila_encabezado_simple(dataframe):
    """
    Función original para detectar una fila 'mejor' por heurística simple.
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
