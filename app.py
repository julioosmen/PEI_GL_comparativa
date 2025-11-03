import streamlit as st
import pandas as pd
from io import BytesIO
from modules.compare_oei import comparar_oei
from modules.compare_aei import comparar_aei
from modules.utils import leer_documento

# === CONFIGURACIÓN GENERAL ===
st.set_page_config(page_title="Comparador PEI GL", layout="wide")

st.title("📊 Comparador PEI con Estándar General")

RUTA_ESTANDAR = "data/PEI_Estandar.xlsx"

# === FUNCIÓN PARA GENERAR RESUMEN CONSOLIDADO ===
def generar_resumen(df_oei=None, df_aei=None):
    resumen_data = []

    if df_oei is not None:
        resumen_data.append({
            "Tipo de Comparación": "OEI",
            "Total Filas": len(df_oei),
            "Coincidencias Exactas": (df_oei["Resultado"] == "Coincidencia Exacta").sum(),
            "Coincidencias Parciales": (df_oei["Resultado"] == "Coincidencia Parcial").sum(),
            "No Coinciden": (df_oei["Resultado"] == "No Coincide").sum()
        })

    if df_aei is not None:
        resumen_data.append({
            "Tipo de Comparación": "AEI",
            "Total Filas": len(df_aei),
            "Coincidencias Exactas": (df_aei["Resultado"] == "Coincidencia Exacta").sum(),
            "Coincidencias Parciales": (df_aei["Resultado"] == "Coincidencia Parcial").sum(),
            "No Coinciden": (df_aei["Resultado"] == "No Coincide").sum()
        })

    return pd.DataFrame(resumen_data)

# === SECCIÓN DE CARGA DE DOCUMENTO ===
st.sidebar.header("📁 Cargar documento PEI")
archivo = st.sidebar.file_uploader("Sube el archivo PEI (.docx o .pdf)", type=["docx", "pdf"])

if archivo:
    with st.spinner("Leyendo documento..."):
        tablas = leer_documento(archivo)
        st.success(f"✅ Se encontraron {len(tablas)} tablas en el documento.")

    tipo_comparacion = st.selectbox(
        "Selecciona el tipo de comparación",
        ["OEI", "AEI"]
    )

    # === COMPARACIÓN ===
    with st.spinner("Realizando comparación con estándar..."):
        if tipo_comparacion == "OEI" and "OEI" in tablas:
            df_result_oei = comparar_oei(RUTA_ESTANDAR, tablas["OEI"])
            st.session_state["df_result_oei"] = df_result_oei  # ✅ Guardar en memoria
            df_result = df_result_oei

        elif tipo_comparacion == "AEI" and "AEI" in tablas:
            df_result_aei = comparar_aei(RUTA_ESTANDAR, tablas["AEI"])
            st.session_state["df_result_aei"] = df_result_aei  # ✅ Guardar en memoria
            df_result = df_result_aei

        else:
            st.error(f"No se encontró la tabla {tipo_comparacion} en el documento subido.")
            st.stop()

    # === MOSTRAR RESULTADOS ===
    st.subheader(f"📋 Resultados de comparación {tipo_comparacion}")
    st.dataframe(df_result, use_container_width=True)

    # === DESCARGA INDIVIDUAL ===
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_result.to_excel(writer, index=False, sheet_name=tipo_comparacion)
    output.seek(0)

    st.download_button(
        label=f"⬇️ Descargar resultado {tipo_comparacion}",
        data=output,
        file_name=f"comparacion_{tipo_comparacion}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# === DESCARGA CONSOLIDADA (Resumen + OEI + AEI) ===
if "df_result_oei" in st.session_state or "df_result_aei" in st.session_state:
    st.markdown("---")
    st.subheader("📘 Descarga consolidada (Resumen + OEI + AEI)")

    df_result_oei = st.session_state.get("df_result_oei")
    df_result_aei = st.session_state.get("df_result_aei")

    if df_result_oei is not None or df_result_aei is not None:
        resumen = generar_resumen(df_result_oei, df_result_aei)

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
    else:
        st.info("Debes procesar al menos una comparación para habilitar la descarga consolidada.")
