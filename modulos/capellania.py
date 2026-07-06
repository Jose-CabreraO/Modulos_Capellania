import subprocess
import sys

import pandas as pd
import streamlit as st

from core import excel_store
from core.excel_store import (
    EXCEL_FILE,
    inyectar_whatsapp,
    leer_referencias,
    resumen_estados,
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


def _ejecutar_bot():
    proceso = subprocess.run(
        [sys.executable, "-m", "core.bot_carga"],
        cwd=EXCEL_FILE.parent,
        capture_output=True,
        text=True,
        check=False,
    )
    return proceso


def render_capellania():
    st.title("Capellanía")

    if "df_grupos" not in st.session_state:
        st.session_state.df_grupos = excel_store.leer_grupos()

    estados = resumen_estados()
    col1, col2, col3, col4 = st.columns([1, 1, 1, 1.4])
    col1.metric("Pendientes", estados["PENDIENTE"])
    col2.metric("Procesados", estados["PROCESADO"])
    col3.metric("Errores", estados["ERROR"])

    with col4:
        st.write("")
        st.write("")
        if st.button("Ejecutar Carga Masiva", type="primary", use_container_width=True):
            with st.spinner("Ejecutando Playwright sobre filas PENDIENTE..."):
                resultado = _ejecutar_bot()
            if resultado.returncode == 0:
                st.success("Carga finalizada.")
            else:
                st.error("La carga terminó con errores.")
            with st.expander("Log de ejecución", expanded=resultado.returncode != 0):
                st.code((resultado.stdout or "") + "\n" + (resultado.stderr or ""), language="text")
            st.session_state.df_grupos = excel_store.leer_grupos()
            st.rerun()

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
        st.data_editor(preview, use_container_width=True, hide_index=True, disabled=True)
        if st.button("Confirmar e Inyectar a Excel", use_container_width=False):
            grupos, integrantes = inyectar_whatsapp(registros)
            st.success(f"Inyectado: {grupos} grupo(s), {integrantes} integrante(s).")
            st.session_state.df_grupos = excel_store.leer_grupos()
            st.rerun()
    elif texto.strip():
        st.warning("No pude detectar una cabecera de grupo válida todavía.")

    st.divider()
    st.subheader("Cola actual")
    edited_df = st.data_editor(
        st.session_state.df_grupos,
        num_rows="dynamic",
        key="editor_grupos",
        use_container_width=True,
        hide_index=True,
    )
    st.session_state.df_grupos = edited_df.copy()

    if st.button("💾 Guardar Cambios de la Grilla"):
        excel_store.guardar_tabla_editada(edited_df)
        st.session_state.df_grupos = edited_df.copy()
        st.success("¡Cambios guardados en el Excel de forma exitosa!")
        st.rerun()
