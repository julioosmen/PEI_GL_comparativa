import pandas as pd
from sentence_transformers import SentenceTransformer, util

def comparar_oei(archivo_estandar, df_oei):
    """
    Compara la tabla OEI extraída del PEI con la tabla estándar.
    Devuelve un DataFrame con los resultados.
    """
    modelo = SentenceTransformer('paraphrase-MiniLM-L6-v2')

    HOJA_ESTANDAR = "OEI"
    COLUMNA_ESTANDAR_TEXTO = "Denominación de OEI / AEI / AO"
    COLUMNA_ESTANDAR_CODIGO = "Código"
    COLUMNA_COMPARAR_TEXTO = "Denominación de OEI"
    COLUMNA_COMPARAR_CODIGO = "Código"
    UMBRAL_SIMILITUD = 0.75

    df_estandar = pd.read_excel(archivo_estandar, sheet_name=HOJA_ESTANDAR)
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

    return pd.DataFrame(resultados)
