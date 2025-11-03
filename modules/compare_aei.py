import pandas as pd
from sentence_transformers import SentenceTransformer, util
from difflib import get_close_matches
import re

def comparar_aei(ruta_estandar, df_aei, umbral=0.75):
    """
    Compara la tabla AEI extraída del PEI con la tabla estándar.
    Devuelve un DataFrame estilizado con:
      - Similitud semántica
      - Clasificación (exacta / parcial / no coincide)
      - Diferencias visuales (+ añadidas / – eliminadas)
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

    # === CARGA DE ARCHIVOS ===
    df_estandar = pd.read_excel(ruta_estandar, sheet_name=HOJA_ESTANDAR)
    df_comparar = df_aei.copy()

    # === DETECCIÓN DE COLUMNAS ===
    def detectar_columna(df, opciones, tipo):
        cols_norm = {col: col.strip().lower()
                     .replace("ó", "o").replace("í", "i")
                     .replace("á", "a").replace("é", "e")
                     .replace("ú", "u") for col in df.columns}

        for col_real, col_norm in cols_norm.items():
            for opt in opciones:
                opt_norm = opt.strip().lower().replace("ó", "o").replace("í", "i")\
                    .replace("á", "a").replace("é", "e").replace("ú", "u")
                if opt_norm in col_norm or col_norm in opt_norm:
                    return col_real

            coincidencia = get_close_matches(col_norm, [o.lower() for o in opciones], n=1, cutoff=0.6)
            if coincidencia:
                return col_real

        raise KeyError(
            f"❌ No se encontró la columna de {tipo}.\n"
            f"🧠 Columnas del archivo: {list(df.columns)}\n"
            f"🧩 Opciones buscadas: {opciones}"
        )

    col_texto_comparar = detectar_columna(df_comparar, COLUMNA_COMPARAR_TEXTO, "texto a comparar")
    col_codigo_comparar = detectar_columna(df_comparar, COLUMNA_COMPARAR_CODIGO, "código a comparar")

    # === LIMPIEZA Y NORMALIZACIÓN ===
    def limpiar_texto(t):
        if not isinstance(t, str):
            return ""
        t = t.lower().strip()
        t = re.sub(r"\s+", " ", t)  # quita dobles espacios
        t = re.sub(r"[.,;:¡!¿?\-\–]", "", t)  # quita puntuación
        t = re.sub(r"[áà]", "a", t)
        t = re.sub(r"[éè]", "e", t)
        t = re.sub(r"[íì]", "i", t)
        t = re.sub(r"[óò]", "o", t)
        t = re.sub(r"[úù]", "u", t)
        return t

    df_estandar[COLUMNA_ESTANDAR_TEXTO] = df_estandar[COLUMNA_ESTANDAR_TEXTO].astype(str).apply(limpiar_texto)
    df_comparar[col_texto_comparar] = df_comparar[col_texto_comparar].astype(str).apply(limpiar_texto)

    # 🧹 Eliminar filas sin código o sin texto (vacías o nulas)
    df_comparar = df_comparar[
        df_comparar[col_texto_comparar].str.strip().ne("") &
        df_comparar[col_codigo_comparar].astype(str).str.strip().ne("")
    ].reset_index(drop=True)
       
    # === EMBEDDINGS ===
    embeddings_estandar = modelo.encode(df_estandar[COLUMNA_ESTANDAR_TEXTO].tolist(), convert_to_tensor=True)
    embeddings_comparar = modelo.encode(df_comparar[col_texto_comparar].tolist(), convert_to_tensor=True)

    # === FUNCIÓN PARA DETECTAR DIFERENCIAS VISUALES ===
    def detectar_diferencias(texto_estandar, texto):
        palabras_estandar = texto_estandar.split()
        palabras_texto = texto.split()

        eliminadas = set(palabras_estandar) - set(palabras_texto)
        añadidas = set(palabras_texto) - set(palabras_estandar)

        diferencias = []
        if eliminadas:
            diferencias.append("– " + ", ".join(sorted(eliminadas)))
        if añadidas:
            diferencias.append("+ " + ", ".join(sorted(añadidas)))

        return "; ".join(diferencias) if diferencias else "(sin diferencias)"

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

        # Clasificación
        if texto.lower() == texto_estandar.lower():
            categoria = "Coincidencia exacta"
        elif valor_max >= umbral:
            categoria = "Coincidencia parcial"
        else:
            categoria = "No coincide"

        diferencias = detectar_diferencias(texto_estandar, texto)

        resultados.append({
            "Código del GL": codigo_comparar,
            "Elemento del GL": texto,
            "Código estándar más similar": codigo_estandar,
            "Elemento estándar más similar": texto_estandar,
            #"Similitud": round(valor_max, 3),
            "Resultado": categoria,
            "Diferencias detectadas": diferencias
        })

    df_resultado = pd.DataFrame(resultados)

    # === 🔍 FILTRO PARA EXCLUIR FILAS CON "OEI", "OIE" O SIMILARES ===
    df_resultado = df_resultado[
        ~df_resultado["Código del GL"]
        .astype(str)
        .str.contains(r"O.?E.?I|O.?I.?E", case=False, na=False)
    ].reset_index(drop=True)

    # === 🎨 COLORES DE RESULTADO ===
    def color_fila(row):
        if row["Resultado"] == "Coincidencia exacta":
            color = "background-color: lightgreen"
        elif row["Resultado"] == "Coincidencia parcial":
            color = "background-color: khaki"
        else:
            color = "background-color: lightcoral"
        return [color] * len(row)

    return df_resultado.style.apply(color_fila, axis=1)
