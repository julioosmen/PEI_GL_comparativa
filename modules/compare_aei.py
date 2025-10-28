import pandas as pd
from sentence_transformers import SentenceTransformer, util

def comparar_aei(archivo_estandar, df_aei):
    """
    Compara la tabla AEI extraída del PEI con la tabla estándar.
    Devuelve un DataFrame con los resultados.
    """
    modelo = SentenceTransformer('all-mpnet-base-v2')

    HOJA_ESTANDAR = "AEI"
    COLUMNA_ESTANDAR = "Denominación de OEI / AEI / AO"
    COLUMNA_COMPARAR = "Denominación"
    CODIGO_ESTANDAR = "Código"
    CODIGO_COMPARAR = "Código"
    PREFIJO_FILTRAR = "AEI"
    UMBRAL_SIMILITUD = 0.75

    df_estandar = pd.read_excel(archivo_estandar, sheet_name=HOJA_ESTANDAR)
    df_comparar = df_aei.copy()

    if CODIGO_COMPARAR in df_comparar.columns:
        df_comparar = df_comparar[df_comparar[CODIGO_COMPARAR].astype(str).str.startswith(PREFIJO_FILTRAR)]

    df_estandar[COLUMNA_ESTANDAR] = df_estandar[COLUMNA_ESTANDAR].astype(str).str.strip().str.lower()
    df_comparar[COLUMNA_COMPARAR] = df_comparar[COLUMNA_COMPARAR].astype(str).str.strip().str.lower()

    embeddings_estandar = modelo.encode(df_estandar[COLUMNA_ESTANDAR].tolist(), convert_to_tensor=True)
    embeddings_comparar = modelo.encode(df_comparar[COLUMNA_COMPARAR].tolist(), convert_to_tensor=True)

    resultados = []
    for i, texto in enumerate(df_comparar[COLUMNA_COMPARAR]):
        emb_texto = embeddings_comparar[i]
        similitudes = util.cos_sim(emb_texto, embeddings_estandar)[0]
        indice_max = similitudes.argmax().item()
        valor_max = similitudes[indice_max].item()

        texto_estandar = df_estandar.iloc[indice_max][COLUMNA_ESTANDAR]
        codigo_estandar = df_estandar.iloc[indice_max][CODIGO_ESTANDAR]
        codigo_comparar = df_comparar.iloc[i][CODIGO_COMPARAR]

        if texto == texto_estandar:
            categoria = "Coincidencia exacta"
        elif valor_max >= UMBRAL_SIMILITUD:
            categoria = "Coincidencia parcial"
        else:
            categoria = "No coincide"

        resultados.append({
            "Código comparar": codigo_comparar,
            f"{COLUMNA_COMPARAR} (comparar)": texto,
            "Código estándar más similar": codigo_estandar,
            f"{COLUMNA_ESTANDAR} (estándar)": texto_estandar,
            "Similitud": round(valor_max, 3),
            "Resultado": categoria
        })

    return pd.DataFrame(resultados)
