import os
import subprocess
import sys

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import streamlit as st


@st.cache_resource
def garantizar_navegadores_playwright():
    try:
        subprocess.run(["playwright", "install", "chromium"], check=True)
    except Exception as e:
        st.error(f"Error instalando navegadores de Playwright: {e}")


garantizar_navegadores_playwright()

from modulos.capellania import render_capellania


st.set_page_config(
    page_title="Grupos de estudio",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)


st.sidebar.title("Grupos de estudio")
pagina = st.sidebar.radio("Navegación", ["Capellanía"], label_visibility="collapsed")

if pagina == "Capellanía":
    render_capellania()
