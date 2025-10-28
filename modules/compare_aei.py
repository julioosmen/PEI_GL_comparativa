import pandas as pd
from sentence_transformers import SentenceTransformer, util
import difflib

# === CONFIGURACIÓN ===
HOJA_ESTANDAR = "AEI"
COLUMNA_ESTANDAR_TEXTO = "denominación de oei / aei / ao"
COLUMNA_ESTANDAR_CODIGO = "código"
COLUMNA_COMPARAR_TEXTO = "denominación"  # nombre base que queremos encontrar
COLUMNA_COMPARAR_CODIGO = "código"
UMBRAL_SIMILITUD = 0.75


def leer_excel_con_encabezado_dinamico(ruta, sheet_name=None):
    """I ntenta leer el Excel considerando que el encabezado puede estar en la fila 1 o 2."""
    try:
        df = pd.read_excel(ruta, sheet_name=sheet_name)
        if len(df.columns) == 1:
            df = pd.read_excel(ruta, sheet_name=sheet_name, header=1)
    except Exception:
        df = pd.read_excel(ruta, sheet_name=sheet_name, header=1)
    return df


def buscar_columna_similar(nombre_columna_objetivo, columnas_disponibles):
    """Encuentra el nombre de columna más parecido al buscado (tolerante a tildes o diferencias leves)."""
    columnas_disponibles = [c.strip().lower() for c in columnas_disponibles]
    nombre_columna_objetivo = nombre_columna_objetivo.strip().lower()
    coincidencias = difflib.get_close_matches(nombre_columna_objetivo, columnas_disponibles, n=1, cutoff=0.6)
    return coincidencias[0] if coincidencias else None


def comparar_aei(ruta_estandar, archivo_comparar):
    print("🔹 Cargando modelo de embeddings...")
    modelo = SentenceTransformer('paraphrase-MiniLM-L6-v2')

    # === Leer ambos archivos ===
    df_estandar = leer_excel_con_encabezado_dinamico(ruta_estandar, sheet_name=HOJA_ESTANDAR)
    df_comparar = leer_excel_con_encabezado_dinamico(archivo_comparar)

    # Normalizar encabezados
    df_estandar.columns = df_estandar.columns.str.strip().str.lower()
    df_comparar.columns = df_comparar.columns.str.strip().str.lower()

    print("\n📋 Columnas detectadas en el archivo estándar:", df_estandar.columns.tolist())
    print("📋 Columnas detectadas en el archivo a comparar:", df_comparar.columns.tolist())

    # Buscar columnas más parecidas
    col_texto_estandar = buscar_columna_similar(COLUMNA_ESTANDAR_TEXTO, df_estandar.columns)
    col_codigo_estandar = buscar_columna_similar(COLUMNA_ESTANDAR_CODIGO, df_estandar.columns)
    col_texto_comparar = buscar_columna_similar(COLUMNA_COMPARAR_TEXTO, df_comparar.columns)
    col_codigo_comparar = buscar_columna_similar(COLUMNA_COMPARAR_CODIGO, df_comparar.columns)

    if not all([col_texto_estandar, col_codigo_estandar, col_texto_comparar, col_codigo_comparar]):
        raise KeyError("❌ No se encontraron las columnas necesarias. Verifica los encabezados del archivo Excel.")

    print(f"\n✅ Columna de texto (comparar): '{col_texto_comparar}'")
    print(f"✅ Columna de texto (estándar): '{col_texto_estandar}'")

    # === Limpieza de texto ===
    df_estandar[col_texto_estandar] = df_estandar[col_texto_estandar].astype(str).str.strip().str.lower()
    df_comparar[col_texto_comparar] = df_comparar[col_texto_comparar].astype(str).str.strip().str.lower()

    # === Generar embeddings ===
    print("\n🔹 Generando embeddings...")
    embeddings_estandar = modelo.encode(df_estandar[col_texto_estandar].tolist(), convert_to_tensor=True)
    embeddings_comparar = modelo.encode(df_comparar[col_texto_comparar].tolist(), convert_to_tensor=True)

    resultados = []

    for i, texto in enumerate(df_comparar[col_texto_comparar]):
        emb_texto = embeddings_comparar[i]
        similitudes = util.cos_sim(emb_texto, embeddings_estandar)[0]
        indice_max = similitudes.argmax().item()
        valor_max = similitudes[indice_max].item()

        texto_estandar = df_estandar.loc[indice_max, col_texto_estandar]
        codigo_estandar = df_estandar.loc[indice_max, col_codigo_estandar]
        codigo_comparar = df_comparar.loc[i, col_codigo_comparar]

        # Clasificación
        if texto == texto_estandar:
            categoria = "Coincidencia exacta"
        elif valor_max >= UMBRAL_SIMILITUD:
            categoria = "Coincidencia parcial"
        else:
            categoria = "No coincide"

        resultados.append({
            "Código comparar": codigo_comparar,
            "Elemento a comparar": texto,
            "Código estándar más similar": codigo_estandar,
            "Elemento estándar más similar": texto_estandar,
            "Similitud": round(valor_max, 3),
            "Resultado": categoria
        })

    df_resultados = pd.DataFrame(resultados)
    print("\n✅ Comparación AEI completada correctamente.")
    return df_resultados
