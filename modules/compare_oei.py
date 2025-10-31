import pandas as pd
from sentence_transformers import SentenceTransformer, util

def comparar_oei(ruta_estandar, df_oei):
    """
    Compara la tabla OEI extraída del PEI con la tabla estándar.
    Soporta múltiples nombres posibles de columnas.
    Devuelve un DataFrame con los resultados de similitud.
    """

    # === CONFIGURACIÓN ===
    modelo = SentenceTransformer('paraphrase-MiniLM-L6-v2')
    HOJA_ESTANDAR = "OEI"
    UMBRAL_SIMILITUD = 0.75

    # Posibles nombres de columnas (texto y código)
    OPCIONES_TEXTO_COMPARAR = [
        "Denominación de OEI",
        "OBJETIVOS ESTRATÉGICOS INSTITUCIONALES",
        "Denominación de OEI / AEI / AO",
        "Descripción"
    ]
    OPCIONES_CODIGO_COMPARAR = [
        "Código",
        "CODIGO",
        "Código OEI",
        "Cod OEI"
    ]

    COLUMNA_ESTANDAR_TEXTO = "Denominación de OEI / AEI / AO"
    COLUMNA_ESTANDAR_CODIGO = "Código"

    # === CARGA DE ARCHIVOS ===
    df_estandar = pd.read_excel(ruta_estandar, sheet_name=HOJA_ESTANDAR)
    df_comparar = df_oei.copy()

    # === DETECCIÓN DE COLUMNAS ===
    def detectar_columna(df, opciones, tipo):
        for col in df.columns:
            if col.strip() in opciones:
                return col
        raise ValueError(f"No se encontró columna de {tipo} en las opciones: {opciones}")

    col_texto_comparar = detectar_columna(df_comparar, OPCIONES_TEXTO_COMPARAR, "texto a comparar")
    col_codigo_comparar = detectar_columna(df_comparar, OPCIONES_CODIGO_COMPARAR, "código a comparar")

    # === LIMPIEZA DE TEXTO ===
    df_estandar[COLUMNA_ESTANDAR_TEXTO] = df_estandar[COLUMNA_ESTANDAR_TEXTO].astype(str).str.strip()
    df_comparar[col_texto_comparar] = df_comparar[col_texto_comparar].astype(str).str.strip()

    # === EMBEDDINGS ===
    embeddings_estandar = modelo.encode(df_estandar[COLUMNA_ESTANDAR_TEXTO].tolist(), convert_to_tensor=True)
    embeddings_comparar = modelo.encode(df_comparar[col_texto_comparar].tolist(), convert_to_tensor=True)

    # === CÁLCULO DE SIMILITUD ===
    similitudes = util.cos_sim(embeddings_comparar, embeddings_estandar)

    resultados = []
    for i, fila in df_comparar.iterrows():
        idx_max = similitudes[i].argmax().item()
        similitud_max = similitudes[i][idx_max].item()

        if similitud_max >= 0.95:
            estado = "Coincidencia Exacta"
            color = "🟩"
        elif similitud_max >= UMBRAL_SIMILITUD:
            estado = "Coincidencia Parcial"
            color = "🟨"
        else:
            estado = "No Coincide"
            color = "🟥"

        resultados.append({
            "Código (PEI)": fila[col_codigo_comparar],
            "Denominación (PEI)": fila[col_texto_comparar],
            "Código (Estandar)": df_estandar.loc[idx_max, COLUMNA_ESTANDAR_CODIGO],
            "Denominación (Estandar)": df_estandar.loc[idx_max, COLUMNA_ESTANDAR_TEXTO],
            "Similitud": round(similitud_max, 3),
            "Resultado": estado,
            "Color": color
        })

    df_resultado = pd.DataFrame(resultados)
    return df_resultado
