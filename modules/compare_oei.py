import pandas as pd
import re
import unicodedata
import difflib
from sentence_transformers import SentenceTransformer, util

def comparar_oei(ruta_estandar, df_oei, umbral=0.75):
    """
    Compara la tabla OEI extraída del PEI con la tabla estándar,
    ignorando diferencias en tildes, espacios y puntuación.
    Además, muestra las palabras que difieren entre ambas frases.
    Devuelve (df_resultado, df_estilizado).
    """

    modelo = SentenceTransformer('paraphrase-MiniLM-L6-v2')
    HOJA_ESTANDAR = "OEI"
    COL_EST_TEXTO = "Denominación de OEI / AEI / AO"
    COL_EST_CODIGO = "Código"

    COL_OPC_TEXTO = [
        "Denominación",
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
        texto = unicodedata.normalize("NFD", texto)
        texto = texto.encode("ascii", "ignore").decode("utf-8")
        texto = re.sub(r"[.,;:!?¿¡()\"'”“]", "", texto)
        texto = re.sub(r"\s+", " ", texto).strip()
        return texto

    def detectar_columna(df, opciones, tipo):
        for col in df.columns:
            for opc in opciones:
                if col.strip().lower() == opc.strip().lower():
                    return col
        raise ValueError(f"No se encontró columna de {tipo} en las opciones: {opciones}")

    def obtener_diferencias(texto1, texto2):
        """
        Devuelve las palabras que difieren entre texto1 y texto2.
        """
        palabras1 = normalizar_texto(texto1).split()
        palabras2 = normalizar_texto(texto2).split()
        diffs = []
        sm = difflib.SequenceMatcher(None, palabras1, palabras2)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag in ["replace", "delete", "insert"]:
                parte1 = " ".join(palabras1[i1:i2])
                parte2 = " ".join(palabras2[j1:j2])
                if parte1 and parte2:
                    diffs.append(f"{parte1} → {parte2}")
                elif parte1:
                    diffs.append(f"– {parte1}")
                elif parte2:
                    diffs.append(f"+ {parte2}")
        return "; ".join(diffs) if diffs else "—"

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

        # Categoría
        if normalizar_texto(texto) == normalizar_texto(texto_estandar):
            categoria = "Coincidencia exacta"
        elif val_max >= umbral:
            categoria = "Coincidencia parcial"
        else:
            categoria = "No coincide"

        # Diferencias literales
        diferencias = obtener_diferencias(texto, texto_estandar)

        resultados.append({
            "Código comparar": codigo_comparar,
            "Elemento a comparar": texto,
            "Código estándar más similar": codigo_estandar,
            "Elemento estándar más similar": texto_estandar,
            #"Similitud": round(val_max, 3),
            "Resultado": categoria,
            "Diferencias": diferencias
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
