import streamlit as st
import pandas as pd
from modules.extract_tables import extraer_tablas
from modules.compare_oei import comparar_oei, comparar_oei_ind
from modules.compare_aei import comparar_aei, comparar_aei_ind
from io import BytesIO

RUTA_ESTANDAR = "Extraer_por_elemento_MEGL.xlsx"

st.set_page_config(page_title="Comparador de elementos PEI de los Gobiernos Locales", layout="wide")
st.title("📊 Comparador de elementos PEI de los Gobiernos Locales")

# ===============================
# 1️⃣ Cargar archivo del usuario
# ===============================
uploaded_file = st.file_uploader("Sube tu archivo PEI (Word o PDF)", type=["docx", "pdf"])

if uploaded_file:
    tablas = extraer_tablas(uploaded_file)
    st.success("✅ Tablas extraídas correctamente")

    # ===============================
    # 2️⃣ Ejecutar todas las comparaciones
    # ===============================
    with st.spinner("Comparando tablas..."):
        df_oei_den = comparar_oei(RUTA_ESTANDAR, tablas.get("OEI"))
        df_oei_ind = comparar_oei_ind(RUTA_ESTANDAR, tablas.get("OEI"))
        df_aei_den = comparar_aei(RUTA_ESTANDAR, tablas.get("AEI"))
        df_aei_ind = comparar_aei_ind(RUTA_ESTANDAR, tablas.get("AEI"))

    # Guardar en session_state
    st.session_state.update({
        "df_result_oei_den": df_oei_den,
        "df_result_oei_ind": df_oei_ind,
        "df_result_aei_den": df_aei_den,
        "df_result_aei_ind": df_aei_ind,
    })

    st.success("✅ Comparaciones completadas")

    # =====================================
    # 3️⃣ Mostrar resultados individuales
    # =====================================
    st.subheader("📊 Resultados de comparaciones")
    
    if resultado_oei is not None:
        st.markdown("### Comparación OEI")
        df_result = resultado_oei
    
        # 🔹 Si es diccionario, tomar el DataFrame principal
        if isinstance(df_result, dict):
            df_result = df_result.get("data", None)
    
        # 🔹 Verificar si es un Styler (con formato de colores)
        if hasattr(df_result, "render") and hasattr(df_result, "data"):
            df_result.data.index = range(1, len(df_result.data) + 1)
            st.dataframe(df_result, use_container_width=True)
        elif isinstance(df_result, pd.DataFrame):
            df_result.index = range(1, len(df_result) + 1)
            st.dataframe(df_result, use_container_width=True)
    
    if resultado_aei is not None:
        st.markdown("### Comparación AEI")
        df_result = resultado_aei
    
        # 🔹 Si es diccionario, tomar el DataFrame principal
        if isinstance(df_result, dict):
            df_result = df_result.get("data", None)
    
        # 🔹 Verificar si es un Styler
        if hasattr(df_result, "render") and hasattr(df_result, "data"):
            df_result.data.index = range(1, len(df_result.data) + 1)
            st.dataframe(df_result, use_container_width=True)
        elif isinstance(df_result, pd.DataFrame):
            df_result.index = range(1, len(df_result) + 1)
            st.dataframe(df_result, use_container_width=True)

    
    # ===============================
    # 4️⃣ Resumen estadístico (sin promedio general)
    # ===============================
    st.header("📈 Resumen de Resultados")

    def calcular_estadisticas(df):
        if isinstance(df, pd.io.formats.style.Styler):
            df = df.data
        if df is None or df.empty:
            return 0
        total = len(df)
        exactas = (df["Resultado"] == "Coincidencia exacta").sum()
        parciales = (df["Resultado"] == "Coincidencia parcial").sum()
        no_coincide = (df["Resultado"] == "No coincide").sum()
        return {
            "Total": total,
            "Exactas": exactas,
            "Parciales": parciales,
            "No coincide": no_coincide,
            "% Exactas": round(exactas / total * 100, 1) if total else 0,
            "% Parciales": round(parciales / total * 100, 1) if total else 0,
            "% No coincide": round(no_coincide / total * 100, 1) if total else 0,
        }

    resumen_data = []
    comparaciones = {
        "OEI (Denominación)": "df_result_oei_den",
        "OEI (Indicador)": "df_result_oei_ind",
        "AEI (Denominación)": "df_result_aei_den",
        "AEI (Indicador)": "df_result_aei_ind"
    }

    for nombre, key in comparaciones.items():
        df = st.session_state[key]
        stats = calcular_estadisticas(df)
        resumen_data.append({
            "Comparación": nombre,
            **stats
        })

    df_resumen = pd.DataFrame(resumen_data)
    st.dataframe(df_resumen, use_container_width=True)
    st.session_state["df_resumen"] = df_resumen

    # ===============================
    # 5️⃣ Exportar a Excel consolidado con colores
    # ===============================
    st.header("📤 Exportar Resultados")
    
    def exportar_excel():
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # --- Hoja de resumen (sin estilos) ---
            df_resumen.to_excel(writer, sheet_name="Resumen", index=False)
    
            # --- Hojas con formato de color ---
            for nombre, key in comparaciones.items():
                df = st.session_state[key]
                sheet_name = nombre.replace(" ", "_")
    
                if isinstance(df, pd.io.formats.style.Styler):
                    # Exportar manteniendo los colores definidos en el Styler
                    df.to_excel(writer, sheet_name=sheet_name, index=False, engine="openpyxl")
                else:
                    # Si no tiene estilos, exportar normalmente
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
    
        output.seek(0)
        return output
    
    excel_bytes = exportar_excel()
    
    st.download_button(
        label="⬇️ Descargar Excel Consolidado",
        data=excel_bytes,
        file_name="Comparativo_PEIGL_Completo.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("📁 Sube un archivo Word o PDF para iniciar la comparación.")
