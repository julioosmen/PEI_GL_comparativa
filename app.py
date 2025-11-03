import streamlit as st
import pandas as pd
from io import BytesIO
from modules.extract_tables import extraer_tablas
from modules.compare_oei import comparar_oei
from modules.compare_aei import comparar_aei

# === FUNCIÓN PARA GENERAR RESUMEN ===
def generar_resumen(df_oei=None, df_aei=None):
    def procesar(df, tipo):
        if df is None or df.empty:
            return pd.Series({
                "Tipo": tipo,
                "Total de elementos": 0,
                "Coincidencia exacta": 0,
                "Coincidencia parcial": 0,
                "No coincide": 0,
                "% Coincidencia exacta": 0,
                "% Coincidencia parcial": 0,
                "% No coincide": 0
            })

        total = len(df)
        exactas = (df["Resultado"] == "Coincidencia exacta").sum()
        parciales = (df["Resultado"] == "Coincidencia parcial").sum()
        no_coincide = (df["Resultado"] == "No coincide").sum()

        return pd.Series({
            "Tipo": tipo,
            "Total de elementos": total,
            "Coincidencia exacta": exactas,
            "Coincidencia parcial": parciales,
            "No coincide": no_coincide,
            "% Coincidencia exacta": round(exactas / total * 100, 1),
            "% Coincidencia parcial": round(parciales / total * 100, 1),
            "% No coincide": round(no_coincide / total * 100, 1)
        })

    resumen_oei = procesar(df_oei, "OEI") if df_oei is not None else None
    resumen_aei = procesar(df_aei, "AEI") if df_aei is not None else None

    df_resumen = pd.DataFrame([r for r in [resumen_oei, resumen_aei] if r is not None])
    return df_resumen


# === CONFIGURACIÓN INICIAL ===
st.set_page_config(page_title="Comparador de elementos PEI de los Gobiernos Locales", layout="wide")
st.title("📊 Analizador PEI – Extracción y Comparación de OEI/AEI")

RUTA_ESTANDAR = "Extraer_por_elemento_MEGL.xlsx"

# === SECCIÓN DE SUBIDA ===
st.sidebar.header("📂 Subir archivo PEI")
archivo_pei = st.sidebar.file_uploader("Selecciona el archivo del PEI (Word o PDF)", type=["pdf", "docx"])
tipo_comparacion = st.sidebar.radio("Selecciona tipo de comparación", ["OEI", "AEI"])

if archivo_pei:
    st.write(f"📁 Archivo subido: **{archivo_pei.name}**")

    if st.button("🚀 Iniciar procesamiento"):
        with st.spinner("Extrayendo tablas relevantes..."):
            tablas = extraer_tablas(archivo_pei)

        if not tablas:
            st.error("⚠️ No se encontraron tablas relevantes (OEI o AEI).")
            st.stop()

        with st.spinner("Realizando comparación con estándar..."):
            if tipo_comparacion == "OEI" and "OEI" in tablas:
                df_result = comparar_oei(RUTA_ESTANDAR, tablas["OEI"])
                st.session_state["df_result_oei"] = df_result
            elif tipo_comparacion == "AEI" and "AEI" in tablas:
                df_result = comparar_aei(RUTA_ESTANDAR, tablas["AEI"])
                st.session_state["df_result_aei"] = df_result
            else:
                st.error(f"No se encontró la tabla {tipo_comparacion} en el documento subido.")
                st.stop()

        st.success(f"✅ Comparación completada para {tipo_comparacion}")
        st.dataframe(df_result, use_container_width=True)

        # Descarga individual
        nombre_salida = f"resultado_comparacion_{tipo_comparacion}.xlsx"
        df_result.to_excel(nombre_salida, index=False)
        with open(nombre_salida, "rb") as f:
            st.download_button(
                label=f"⬇️ Descargar resultado {tipo_comparacion}",
                data=f,
                file_name=nombre_salida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# === DESCARGA CONSOLIDADA (Fuera del botón) ===
if "df_result_oei" in st.session_state or "df_result_aei" in st.session_state:
    st.markdown("---")
    st.subheader("📘 Descarga consolidada (Resumen + OEI + AEI)")

    df_result_oei = st.session_state.get("df_result_oei")
    df_result_aei = st.session_state.get("df_result_aei")

    if df_result_oei is not None or df_result_aei is not None:
        resumen = generar_resumen(df_result_oei, df_result_aei)

        # === Contador visual ===
        st.markdown("### 📊 Resumen de resultados")
        for _, fila in resumen.iterrows():
            st.metric(
                label=f"{fila['Tipo']} - Coincidencia exacta",
                value=f"{fila['% Coincidencia exacta']}%",
                delta=f"{fila['Coincidencia exacta']} de {fila['Total de elementos']}"
            )

        # === Excel consolidado ===
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            resumen.to_excel(writer, index=False, sheet_name="Resumen")
            if df_result_oei is not None:
                df_result_oei.to_excel(writer, index=False, sheet_name="OEI")
            if df_result_aei is not None:
                df_result_aei.to_excel(writer, index=False, sheet_name="AEI")
        output.seek(0)

        st.download_button(
            label="⬇️ Descargar consolidado (Resumen + OEI + AEI)",
            data=output,
            file_name="comparacion_consolidada.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
