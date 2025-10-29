import streamlit as st
import pandas as pd
from modules.extract_tables import extraer_tablas
from modules.compare_oei import comparar_oei
from modules.compare_aei import comparar_aei

# === CONFIGURACIÓN INICIAL ===
st.set_page_config(page_title="Comparador de elementos PEI de los Gobiernos Locales", layout="wide")
st.title("📊 Analizador PEI – Extracción y Comparación de OEI/AEI")

# Ruta fija del archivo estándar (ya presente en el proyecto)
RUTA_ESTANDAR = "Extraer_por_elemento_MEGL.xlsx"

# === SECCIÓN DE SUBIDA DE ARCHIVO ===
st.sidebar.header("📂 Subir archivo PEI")
archivo_pei = st.sidebar.file_uploader(
    "Selecciona el archivo del PEI (Word o PDF)", 
    type=["pdf", "docx"]
)

tipo_comparacion = st.sidebar.radio("Selecciona tipo de comparación", ["OEI", "AEI"])

if archivo_pei:
    st.write(f"📁 Archivo subido: **{archivo_pei.name}**")
    if st.button("🚀 Iniciar procesamiento"):
        with st.spinner("Extrayendo tablas relevantes..."):
            tablas = extraer_tablas(archivo_pei)

        if not tablas:
            st.error("⚠️ No se encontraron tablas relevantes (OEI o AEI).")
            st.stop()

        # Ejecutar comparación según tipo seleccionado
        with st.spinner("Realizando comparación con estándar..."):
            df_result = None  # valor por defecto
        
            if tipo_comparacion == "OEI" and "OEI" in tablas:
                df_result = comparar_oei(RUTA_ESTANDAR, tablas["OEI"], usar_streamlit=True)
            elif tipo_comparacion == "AEI" and "AEI" in tablas:
                df_result = comparar_aei(RUTA_ESTANDAR, tablas["AEI"], usar_streamlit=True)
            else:
                st.error(f"No se encontró la tabla {tipo_comparacion} en el documento subido.")
                st.stop()
        
        # Mostrar resultado si existe
        if df_result is not None and not df_result.empty:
            st.success(f"✅ Comparación {tipo_comparacion} completada correctamente")
            st.dataframe(df_result, use_container_width=True)
        else:
            st.warning(f"⚠️ No se pudo realizar la comparación {tipo_comparacion}. "
                       "Verifica que las columnas sean correctas o que los datos estén completos.")

        # Descargar archivo resultante
        nombre_salida = f"resultado_comparacion_{tipo_comparacion}.xlsx"
        df_result.to_excel(nombre_salida, index=False)

        with open(nombre_salida, "rb") as f:
            st.download_button(
                label="⬇️ Descargar resultado en Excel",
                data=f,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Sube un archivo Word o PDF para comenzar.")
