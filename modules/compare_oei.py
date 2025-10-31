import pandas as pd
from sentence_transformers import SentenceTransformer, util
#import streamlit as st if usar_streamlit else None

def comparar_oei(ruta_estandar, df_oei):
    """
    Compara la tabla OEI extraída del PEI con la tabla estándar.
    Devuelve un DataFrame con los resultados.
    """
    modelo = SentenceTransformer('paraphrase-MiniLM-L6-v2')
       
    HOJA_ESTANDAR = "OEI"
    COLUMNA_ESTANDAR_TEXTO = "Denominación de OEI / AEI / AO"
    COLUMNA_ESTANDAR_CODIGO = "Código"
    COLUMNA_COMPARAR_TEXTO = [
        "Denominación de OEI",
        "OBJETIVOS ESTRATÉGICOS INSTITUCIONALES",
        "OBJETIVOS ESTRATÉGICOS INSTITUCIONAL",
        "Denominación de OEI / AEI / AO",
        "Descripción"
    ]    
    COLUMNA_COMPARAR_CODIGO = [
        "Código",
        "CODIGO",
        "Código OEI",
        "Cod OEI"
    ]
    UMBRAL_SIMILITUD = 0.75

    # === CARGA DE ARCHIVOS ===
    df_estandar = pd.read_excel(ruta_estandar, sheet_name=HOJA_ESTANDAR)
    df_comparar = df_oei.copy()

    # === DETECCIÓN DE COLUMNAS ===
    def detectar_columna(df, opciones, tipo):
        for col in df.columns:
            if col.strip() in opciones:
                return col
        raise ValueError(f"No se encontró columna de {tipo} en las opciones: {opciones}")

    col_texto_comparar = detectar_columna(df_comparar, COLUMNA_COMPARAR_TEXTO, "texto a comparar")
    col_codigo_comparar = detectar_columna(df_comparar, COLUMNA_COMPARAR_CODIGO, "código a comparar")

    # === LIMPIEZA DE TEXTO ===
    df_estandar[COLUMNA_ESTANDAR_TEXTO] = df_estandar[COLUMNA_ESTANDAR_TEXTO].astype(str).str.strip()
    df_comparar[col_texto_comparar] = df_comparar[col_texto_comparar].astype(str).str.strip()

    # === EMBEDDINGS ===
    embeddings_estandar = modelo.encode(df_estandar[COLUMNA_ESTANDAR_TEXTO].tolist(), convert_to_tensor=True)
    embeddings_comparar = modelo.encode(df_comparar[col_texto_comparar].tolist(), convert_to_tensor=True)

    # === CÁLCULO DE SIMILITUD ===
    similitudes = util.cos_sim(embeddings_comparar, embeddings_estandar)
    
    resultados = []
    for i, texto in enumerate(df_comparar[COLUMNA_COMPARAR_TEXTO]):
        emb_texto = embeddings_comparar[i]
        similitudes = util.cos_sim(emb_texto, embeddings_estandar)[0]
        indice_max = similitudes.argmax().item()
        valor_max = similitudes[indice_max].item()

        texto_estandar = df_estandar.loc[indice_max, COLUMNA_ESTANDAR_TEXTO]
        codigo_estandar = df_estandar.loc[indice_max, COLUMNA_ESTANDAR_CODIGO]
        codigo_comparar = df_comparar.loc[i, COLUMNA_COMPARAR_CODIGO]

        if texto.lower() == texto_estandar.lower():
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

    #return pd.DataFrame(resultados)
    df_resultado = pd.DataFrame(resultados)

    # ---- 🔵 APLICAR COLORES ----
    def color_fila(row):
        if row["Resultado"] == "Coincidencia exacta":
            color = "background-color: lightgreen"
        elif row["Resultado"] == "Coincidencia parcial":
            color = "background-color: khaki"
        else:
            color = "background-color: lightcoral"
        return [color] * len(row)

    df_styled = df_resultado.style.apply(color_fila, axis=1)

    return df_styled
