import os
import sys

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

import streamlit as st

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
