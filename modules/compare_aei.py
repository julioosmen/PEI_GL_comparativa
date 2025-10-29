import pandas as pd
from sentence_transformers import SentenceTransformer, util
import difflib

# === CONFIGURACIÓN ===
HOJA_ESTANDAR = "AEI"
COLUMNA_ESTANDAR_TEXTO = "denominación de oei / aei / ao"
COLUMNA_ESTANDAR_CODIGO = "código"
COLUMNA_COMPARAR_TEXTO = "denominación"
COLUMNA_COMPARAR_CODIGO = "código"
UMBRAL_SIMILITUD = 0.75


def leer_excel_con_encabezado_dinamico(ruta, sheet_name=None):
    """
    Lee un Excel detectando automáticamente la fila que contiene el encabezado real.
    Considera casos donde la segunda fila es el encabezado (muy común en PEI).
    """

    # ✅ Si ya es un DataFrame, simplemente devuélvelo
    if isinstance(ruta, pd.DataFrame):
        return ruta

    # Leemos sin encabezado para poder inspeccionar las primeras filas
    df_preview = pd.read_excel(ruta, sheet_name=sheet_name, header=None, nrows=5)

    # Posibles palabras clave que suelen aparecer en los encabezados
    palabras_clave = ["código", "denominación", "indicador", "denominación de OEI / AEI", "descripción"]

    # Buscar fila que tenga coincidencias con las palabras clave
    fila_encabezado = 0
    for i, fila in df_preview.iterrows():
        texto_fila = " ".join(str(x).lower() for x in fila.tolist())
        if any(p in texto_fila for p in palabras_clave):
            fila_encabezado = i
            break

    # Leer nuevamente usando esa fila como encabezado
    df = pd.read_excel(ruta, sheet_name=sheet_name, header=fila_encabezado)

    return df



def comparar_aei(ruta_estandar, archivo_comparar, usar_streamlit=False):
    """
    Compara las AEI del PEI con la tabla estándar.
    Si usar_streamlit=True, usa st.error() en lugar de raise KeyError.
    """
    if usar_streamlit:
        import streamlit as st

    print("🔹 Cargando modelo de embeddings...")
    modelo = SentenceTransformer('paraphrase-MiniLM-L6-v2')

    # === Leer ambos archivos ===
    df_estandar = leer_excel_con_encabezado_dinamico(ruta_estandar, sheet_name=HOJA_ESTANDAR)
    df_comparar = leer_excel_con_encabezado_dinamico(archivo_comparar)

    # === Normalizar encabezados ===
    def limpiar_encabezados(df):
        df.columns = (
            df.columns.astype(str)
            .str.strip()
            .str.lower()
            .str.normalize("NFKD")
            .str.encode("ascii", errors="ignore")
            .str.decode("utf-8")
            .str.replace(r"[^a-z0-9 ]", "", regex=True)
        )
        return df

    df_estandar = limpiar_encabezados(df_estandar)
    df_comparar = limpiar_encabezados(df_comparar)

    print("\n📋 Columnas estándar:", df_estandar.columns.tolist())
    print("📋 Columnas comparar:", df_comparar.columns.tolist())

    # === Buscar columnas más parecidas ===
    def buscar_columna_similar(nombre_columna_objetivo, columnas_disponibles):
        columnas_disponibles = [c.strip().lower() for c in columnas_disponibles]
        nombre_columna_objetivo = nombre_columna_objetivo.strip().lower()
        coincidencias = difflib.get_close_matches(nombre_columna_objetivo, columnas_disponibles, n=1, cutoff=0.4)
        return coincidencias[0] if coincidencias else None

    col_texto_estandar = buscar_columna_similar(COLUMNA_ESTANDAR_TEXTO, df_estandar.columns)
    col_codigo_estandar = buscar_columna_similar(COLUMNA_ESTANDAR_CODIGO, df_estandar.columns)
    col_texto_comparar = buscar_columna_similar(COLUMNA_COMPARAR_TEXTO, df_comparar.columns)
    col_codigo_comparar = buscar_columna_similar(COLUMNA_COMPARAR_CODIGO, df_comparar.columns)

    print(f"\n🔎 Texto estándar: {col_texto_estandar}")
    print(f"🔎 Código estándar: {col_codigo_estandar}")
    print(f"🔎 Texto comparar: {col_texto_comparar}")
    print(f"🔎 Código comparar: {col_codigo_comparar}")

    # === Validar columnas ===
    if not all([col_texto_estandar, col_codigo_estandar, col_texto_comparar, col_codigo_comparar]):
        mensaje_error = (
            "❌ No se encontraron las columnas necesarias.\n"
            f"🧠 Columnas estándar: {df_estandar.columns.tolist()}\n"
            f"🧠 Columnas comparar: {df_comparar.columns.tolist()}\n"
            "Verifica que los encabezados contengan palabras similares a las configuradas."
        )
        if usar_streamlit:
            st.error(mensaje_error)
            return None
        else:
            raise KeyError(mensaje_error)

    # === Limpieza ===
    df_estandar[col_texto_estandar] = df_estandar[col_texto_estandar].astype(str).str.strip().str.lower()
    df_comparar[col_texto_comparar] = df_comparar[col_texto_comparar].astype(str).str.strip().str.lower()

    # === Embeddings ===
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
