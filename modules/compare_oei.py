import pandas as pd
import re
import unicodedata
from sentence_transformers import SentenceTransformer, util

def comparar_oei(ruta_estandar, df_oei, umbral=0.75):
    """
    Compara la tabla OEI extraída del PEI con la tabla estándar,
    ignorando diferencias en espacios, comas y tildes.
    Devuelve (df_resultado, df_estilizado).
    """

    modelo = SentenceTransformer('paraphrase-MiniLM-L6-v2')
    HOJA_ESTANDAR = "OEI"
    COL_EST_TEXTO = "Denominación de OEI / AEI / AO"
    COL_EST_CODIGO = "Código"

    COL_OPC_TEXTO = [
        "Denominación de OEI",
        "OBJETIVOS ESTRATÉGICOS INSTITUCIONALES",
        "OBJETIVOS ESTRATÉGICOS INSTITUCIONAL",
        "OBJETIVOS ESTRATEGICOS INSTITUCIONAL",
        "Denominación de OEI / AEI / AO",
        "Descripción"
    ]
    COL_OPC_CODIGO = [
        "Código", "CODIGO", "CÓDIGO", "Código OEI", "Cod OEI"
    ]

    # === FUNCIONES AUXILIARES ===
    def normalizar_texto(texto):
        if pd.isna(texto):
            return ""
        texto = str(texto).lower().strip()
        # Eliminar tildes
        texto = unicodedata.normalize("NFD", texto)
        texto = texto.encode("ascii", "ignore").decode("utf-8")
        # Quitar signos de puntuación comunes
        texto = re.sub(r"[.,;:!?¿¡()\"'”“]", "", texto)
        # Quitar espacios múltiples
        texto = re.sub(r"\s+", " ", texto)
        # Quitar espacios al inicio/fin
        texto = texto.strip()
        return texto

    def detectar_columna(df, opciones, tipo):
        for col in df.columns:
            for opc in opciones:
                if col.strip().lower() == opc.strip().lower():
                    return col
        raise ValueError(f"No se encontró columna de {tipo} en las opciones: {opciones}")

    # === CARGA ===
    df_estandar = pd.read_excel(ruta_estandar, sheet_name=HOJA_ESTANDAR)
    df_comparar = df_oei.copy()

    col_txt_cmp = detectar_columna(df_comparar, COL_OPC_TEXTO, "texto a comparar")
    col_cod_cmp = detectar_columna(df_comparar, COL_OPC_CODIGO, "código a comparar")

    # === LIMPIEZA ===
    df_estandar[COL_EST_TEXTO] = df_estandar[COL_EST_TEXTO].astype(str).str.strip()
    df_comparar[col_txt_cmp] = df_comparar[col_txt_cmp].astype(str).str.strip()

    # === NORMALIZACIÓN ===
    textos_estandar_norm = df_estandar[COL_EST_TEXTO].apply(normalizar_texto)
    textos_comparar_norm = df_comparar[col_txt_cmp].apply(normalizar_texto)

    # === EMBEDDINGS ===
    emb_estandar = modelo.encode(textos_estandar_norm.tolist(), convert_to_tensor=True)
    emb_comparar = modelo.encode(textos_comparar_norm.tolist(), convert_to_tensor=True)

    # === SIMILITUD ===
    matriz_sim = util.cos_sim(emb_comparar, emb_estandar)

    resultados = []
    for i, texto in enumerate(df_comparar[col_txt_cmp]):
        simil_row = matriz_sim[i]
        idx_max = simil_row.argmax().item()
        val_max = simil_row[idx_max].item()

        texto_estandar = df_estandar.loc[idx_max, COL_EST_TEXTO]
        codigo_estandar = df_estandar.loc[idx_max, COL_EST_CODIGO]
        codigo_comparar = df_comparar.loc[i, col_cod_cmp]

        if normalizar_texto(texto) == normalizar_texto(texto_estandar):
            categoria = "Coincidencia exacta"
        elif val_max >= umbral:
            categoria = "Coincidencia parcial"
        else:
            categoria = "No coincide"

        resultados.append({
            "Código comparar": codigo_comparar,
            "Elemento a comparar": texto,
            "Código estándar más similar": codigo_estandar,
            "Elemento estándar más similar": texto_estandar,
            "Similitud": round(val_max, 3),
            "Resultado": categoria
        })

    df_result = pd.DataFrame(resultados)

    # === COLOR VISUAL ===
    def color_fila(row):
        if row["Resultado"] == "Coincidencia exacta":
            return ["background-color: lightgreen"] * len(row)
        elif row["Resultado"] == "Coincidencia parcial":
            return ["background-color: khaki"] * len(row)
        else:
            return ["background-color: lightcoral"] * len(row)

    df_styled = df_result.style.apply(color_fila, axis=1)

    return df_styled
