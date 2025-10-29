import pandas as pd
from sentence_transformers import SentenceTransformer, util
import streamlit as st if usar_streamlit else None

def comparar_oei(ruta_estandar, df_oei):
    """
    Compara la tabla OEI extraída del PEI con la tabla estándar.
    Devuelve un DataFrame con los resultados.
    """
    #modelo = SentenceTransformer('paraphrase-MiniLM-L6-v2')
    try:
        modelo = SentenceTransformer('paraphrase-MiniLM-L6-v2')
    except Exception as e:
        if usar_streamlit:
            st.error(f"❌ Error al cargar el modelo de comparación: {e}")
        else:
            raise e
        return None
        
    HOJA_ESTANDAR = "OEI"
    COLUMNA_ESTANDAR_TEXTO = "Denominación de OEI / AEI / AO"
    COLUMNA_ESTANDAR_CODIGO = "Código"
    COLUMNA_COMPARAR_TEXTO = "Denominación de OEI"
    COLUMNA_COMPARAR_CODIGO = "Código"
    UMBRAL_SIMILITUD = 0.75

    df_estandar = pd.read_excel(ruta_estandar, sheet_name=HOJA_ESTANDAR)
    df_comparar = df_oei.copy()

    df_estandar[COLUMNA_ESTANDAR_TEXTO] = df_estandar[COLUMNA_ESTANDAR_TEXTO].astype(str).str.strip()
    df_comparar[COLUMNA_COMPARAR_TEXTO] = df_comparar[COLUMNA_COMPARAR_TEXTO].astype(str).str.strip()

    embeddings_estandar = modelo.encode(df_estandar[COLUMNA_ESTANDAR_TEXTO].tolist(), convert_to_tensor=True)
    embeddings_comparar = modelo.encode(df_comparar[COLUMNA_COMPARAR_TEXTO].tolist(), convert_to_tensor=True)

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
