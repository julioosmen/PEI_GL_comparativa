import pandas as pd
from sentence_transformers import SentenceTransformer, util
from difflib import get_close_matches, ndiff

def comparar_aei(ruta_estandar, df_aei):
    """
    Compara la tabla AEI extraída del PEI con la tabla estándar.
    Devuelve un DataFrame con los resultados, incluyendo diferencias de texto.
    Excluye filas cuyo código comience con 'OEI'.
    """
    modelo = SentenceTransformer('paraphrase-MiniLM-L6-v2')
       
    HOJA_ESTANDAR = "AEI"
    COLUMNA_ESTANDAR_TEXTO = "Denominación de OEI / AEI / AO"
    COLUMNA_ESTANDAR_CODIGO = "Código"
    COLUMNA_COMPARAR_TEXTO = [
        "AEI",
        "ACCIONES ESTRATÉGICAS INSTITUCIONALES",
        "Denominación de OEI / AEI / AO",
        "Denominación del OEI/AEI",
        "Denominación de OEI / AEI",
        "Denominación de OEI/AEI",
        "Enunciado",
        "Descripción"        
    ]    
    COLUMNA_COMPARAR_CODIGO = [
        "Código",
        "CODIGO",
        "CÓDIGO",        
        "Código AEI",
        "Cod AEI"
    ]
    UMBRAL_SIMILITUD = 0.75

    # === CARGA DE ARCHIVOS ===
    df_estandar = pd.read_excel(ruta_estandar, sheet_name=HOJA_ESTANDAR)
    df_comparar = df_aei.copy()

    # === DETECCIÓN DE COLUMNAS ===
    def detectar_columna(df, opciones, tipo):
        cols_norm = {
            col: col.strip().lower()
            .replace("ó", "o").replace("í", "i")
            .replace("á", "a").replace("é", "e")
            .replace("ú", "u")
            for col in df.columns
        }

        for col_real, col_norm in cols_norm.items():
            for opt in opciones:
                opt_norm = opt.strip().lower()
                opt_norm = opt_norm.replace("ó", "o").replace("í", "i").replace("á", "a").replace("é", "e").replace("ú", "u")
                if opt_norm in col_norm or col_norm in opt_norm:
                    return col_real
            coincidencia = get_close_matches(col_norm, [o.lower() for o in opciones], n=1, cutoff=0.6)
            if coincidencia:
                return col_real

        raise KeyError(f"❌ No se encontró la columna de {tipo}.\nColumnas: {list(df.columns)}\nBuscadas: {opciones}")

    col_texto_comparar = detectar_columna(df_comparar, COLUMNA_COMPARAR_TEXTO, "texto a comparar")
    col_codigo_comparar = detectar_columna(df_comparar, COLUMNA_COMPARAR_CODIGO, "código a comparar")

    # === LIMPIEZA DE TEXTO (ignorando tildes, comas y espacios dobles) ===
    def normalizar_texto(texto):
        if not isinstance(texto, str):
            return ""
        texto = texto.lower().strip()
        texto = texto.replace(",", "").replace(".", "")
        texto = texto.replace("  ", " ")
        for a, b in zip("áéíóú", "aeiou"):
            texto = texto.replace(a, b)
        return texto

    df_estandar[COLUMNA_ESTANDAR_TEXTO] = df_estandar[COLUMNA_ESTANDAR_TEXTO].astype(str)
    df_comparar[col_texto_comparar] = df_comparar[col_texto_comparar].astype(str)

    # === EMBEDDINGS ===
    embeddings_estandar = modelo.encode(df_estandar[COLUMNA_ESTANDAR_TEXTO].tolist(), convert_to_tensor=True)
    embeddings_comparar = modelo.encode(df_comparar[col_texto_comparar].tolist(), convert_to_tensor=True)

    # === CÁLCULO DE SIMILITUD ===
    resultados = []
    for i, texto in enumerate(df_comparar[col_texto_comparar]):
        emb_texto = embeddings_comparar[i]
        similitudes = util.cos_sim(emb_texto, embeddings_estandar)[0]
        indice_max = similitudes.argmax().item()
        valor_max = similitudes[indice_max].item()

        texto_estandar = df_estandar.loc[indice_max, COLUMNA_ESTANDAR_TEXTO]
        codigo_estandar = df_estandar.loc[indice_max, COLUMNA_ESTANDAR_CODIGO]
        codigo_comparar = df_comparar.loc[i, col_codigo_comparar]

        texto_norm = normalizar_texto(texto)
        texto_estandar_norm = normalizar_texto(texto_estandar)

        # Detectar diferencias palabra por palabra
        diff = list(ndiff(texto_norm.split(), texto_estandar_norm.split()))
        diferencias = [d[2:] for d in diff if d.startswith('- ') or d.startswith('+ ')]

        if texto_norm == texto_estandar_norm:
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
            "Resultado": categoria,
            "Diferencias": ", ".join(diferencias) if diferencias else ""
        })

    # === CREAR DATAFRAME FINAL ===
    df_resultado = pd.DataFrame(resultados)

    # 🔴 FILTRAR FILAS QUE EMPIECEN CON "OEI"
    df_resultado = df_resultado[~df_resultado["Código comparar"].astype(str).str.startswith("OEI")].reset_index(drop=True)

    # === APLICAR COLORES ===
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
