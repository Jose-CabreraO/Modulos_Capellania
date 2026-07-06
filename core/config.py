import os


def get_secret(name, default=""):
    value = os.getenv(name)
    if value:
        return value

    try:
        import streamlit as st

        return st.secrets.get(name, default)
    except Exception:
        return default


def capellania_credentials():
    return {
        "user": get_secret("CAPELLANIA_USER"),
        "password": get_secret("CAPELLANIA_PASS"),
    }
