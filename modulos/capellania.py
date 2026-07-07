import pandas as pd
import streamlit as st

from core import bot_carga
from core import excel_store
from core.excel_store import (
    EXCEL_FILE,
    GROUP_COLUMNS,
    inyectar_whatsapp,
    leer_referencias,
)
from core.parser_whatsapp import parsear_whatsapp


def _preview_dataframe(registros):
    filas = []
    for registro in registros:
        integrantes = registro.get("Integrantes", [])
        if not integrantes:
            filas.append({**{k: v for k, v in registro.items() if k != "Integrantes"}, "Nombre*": ""})
        for integrante in integrantes:
            filas.append({**{k: v for k, v in registro.items() if k != "Integrantes"}, **integrante})
    return pd.DataFrame(filas)


def _empty_queue_dataframe():
    return pd.DataFrame(columns=GROUP_COLUMNS)


def _normalize_queue_dataframe(df):
    normalized = df.copy()
    for column in GROUP_COLUMNS:
        if column not in normalized.columns:
            normalized[column] = "PENDIENTE" if column == "Estado_Carga" else ""

    normalized = normalized[GROUP_COLUMNS].fillna("")
    normalized["Estado_Carga"] = (
        normalized["Estado_Carga"]
        .replace("", "PENDIENTE")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    return normalized


def _load_initial_queue_dataframe():
    if EXCEL_FILE.exists():
        try:
            return _normalize_queue_dataframe(excel_store.leer_grupos())
        except Exception:
            return _empty_queue_dataframe()
    return _empty_queue_dataframe()


def _queue_status_counts(df):
    if df.empty or "Estado_Carga" not in df.columns:
        return {"PENDIENTE": 0, "PROCESADO": 0, "ERROR": 0}

    estados = df["Estado_Carga"].fillna("PENDIENTE").astype(str).str.upper().str.strip()
    return {
        "PENDIENTE": int((estados == "PENDIENTE").sum()),
        "PROCESADO": int((estados == "PROCESADO").sum()),
        "ERROR": int((estados == "ERROR").sum()),
    }


def render_capellania():
    st.title("Capellanía")

    if "df_grupos" not in st.session_state:
        st.session_state.df_grupos = _load_initial_queue_dataframe()
    if st.session_state.get("cola_upload_message"):
        st.success(st.session_state.pop("cola_upload_message"))

    estados = _queue_status_counts(st.session_state.df_grupos)
    col1, col2, col3, col4 = st.columns([1, 1, 1, 1.4])
    col1.metric("Pendientes", estados["PENDIENTE"])
    col2.metric("Procesados", estados["PROCESADO"])
    col3.metric("Errores", estados["ERROR"])

    with col4:
        st.write("")
        st.write("")
        if st.button("Ejecutar Carga Masiva", type="primary", width="stretch"):
            with st.spinner("Ejecutando carga masiva en WordPress..."):
                try:
                    bot_carga.procesar_carga_real()
                    st.session_state.df_grupos = excel_store.leer_grupos()
                    st.success("¡Carga finalizada con éxito!")
                except Exception as e:
                    st.error("❌ El bot de Playwright se detuvo por un error:")
                    st.exception(e)

    st.divider()

    st.subheader("Asistente de WhatsApp")
    texto = st.text_area(
        "Pegá el mensaje bruto",
        height=180,
        placeholder="Grupo: CTM Rio Verde\nEmpresa: ctm\nSucursal: rio verde\nMaterial: Cambia Tu Mundo\n• Juan Perez, 123456, juan@email.com\n- Maria Gomez",
    )

    registros = parsear_whatsapp(texto, leer_referencias()) if texto.strip() else []
    preview = _preview_dataframe(registros) if registros else pd.DataFrame()

    if not preview.empty:
        st.data_editor(preview, width="stretch", hide_index=True, disabled=True)
        if st.button("Confirmar e Inyectar a Excel", width="content"):
            grupos, integrantes = inyectar_whatsapp(registros)
            st.success(f"Inyectado: {grupos} grupo(s), {integrantes} integrante(s).")
            st.session_state.df_grupos = excel_store.leer_grupos()
            st.rerun()
    elif texto.strip():
        st.warning("No pude detectar una cabecera de grupo válida todavía.")

    st.divider()
    st.subheader("Cola actual")
    uploaded_excel = st.file_uploader(
        "Cargar archivo de Excel para previsualización",
        type=["xlsx", "xls"],
        key="uploader_cola_excel",
    )
    if uploaded_excel is not None:
        uploaded_key = f"{uploaded_excel.name}:{uploaded_excel.size}"
        if st.session_state.get("uploaded_cola_key") != uploaded_key:
            try:
                excel_store.guardar_archivo_subido(uploaded_excel)
                st.session_state.df_grupos = _normalize_queue_dataframe(excel_store.leer_grupos())
                st.session_state.uploaded_cola_key = uploaded_key
                st.session_state.cola_upload_message = "Archivo cargado y guardado como plantilla_flujo_completo.xlsx."
                st.rerun()
            except Exception as exc:
                st.error(f"No se pudo leer el archivo Excel cargado: {exc}")
                st.session_state.df_grupos = _empty_queue_dataframe()

    edited_df = st.data_editor(
        st.session_state.df_grupos,
        num_rows="dynamic",
        key="editor_grupos",
        width="stretch",
        hide_index=True,
    )
    st.session_state.df_grupos = edited_df.copy()

    if st.button("💾 Guardar Cambios de la Grilla"):
        excel_store.guardar_tabla_editada(edited_df)
        st.session_state.df_grupos = edited_df.copy()
        st.success("¡Cambios guardados en el Excel de forma exitosa!")
        st.rerun()
