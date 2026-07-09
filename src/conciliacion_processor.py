import pandas as pd
import streamlit as st


CENTURY_FACTURA_INDEX = 7
DNIT_COMPROBANTE_COL = "Número de Comprobante"
DNIT_FECHA_COL = "Fecha Emisión"


def _leer_archivo_subido(uploaded_file, *, header):
    nombre = uploaded_file.name.lower()
    if nombre.endswith(".xlsx"):
        return pd.read_excel(uploaded_file, header=header, dtype=str)
    if nombre.endswith(".csv"):
        return pd.read_csv(uploaded_file, header=header, dtype=str)
    raise ValueError("Formato no soportado. Subí un archivo .xlsx o .csv.")


def _normalizar_clave(serie):
    return serie.astype(str).str.strip()


def _validar_estructura(df_century, df_dnit):
    if df_century.shape[1] <= CENTURY_FACTURA_INDEX:
        raise ValueError("El archivo Century no contiene la columna H requerida.")

    columnas_faltantes = [
        columna
        for columna in [DNIT_COMPROBANTE_COL, DNIT_FECHA_COL]
        if columna not in df_dnit.columns
    ]
    if columnas_faltantes:
        raise ValueError("El archivo DNIT no contiene: " + ", ".join(columnas_faltantes))


def _procesar_conciliacion(df_century, df_dnit):
    _validar_estructura(df_century, df_dnit)

    century_keys = _normalizar_clave(df_century.iloc[:, CENTURY_FACTURA_INDEX])
    dnit_keys = _normalizar_clave(df_dnit[DNIT_COMPROBANTE_COL])

    century_key_set = set(century_keys)
    dnit_key_set = set(dnit_keys)

    df_dnit_filtrado = df_dnit[dnit_keys.isin(century_key_set)].copy()
    df_century_filtrado = df_century[~century_keys.isin(dnit_key_set)].copy()

    df_dnit_filtrado["_Fecha_Orden"] = pd.to_datetime(
        df_dnit_filtrado[DNIT_FECHA_COL],
        errors="coerce",
    )
    df_dnit_filtrado[DNIT_COMPROBANTE_COL] = _normalizar_clave(df_dnit_filtrado[DNIT_COMPROBANTE_COL])
    df_dnit_filtrado = (
        df_dnit_filtrado
        .sort_values(by=["_Fecha_Orden", DNIT_COMPROBANTE_COL], ascending=[True, True])
        .drop(columns=["_Fecha_Orden"])
    )

    return df_dnit_filtrado, df_century_filtrado


def render_modulo_conciliacion():
    st.header("📊 Módulo de Conciliación Automática: Century vs DNIT")
    st.write(
        "Subí ambos archivos para cruzarlos automáticamente. "
        "El módulo conserva en DNIT solo los comprobantes coincidentes y deja en Century solo los registros únicos."
    )

    col1, col2 = st.columns(2)

    with col1:
        file_century = st.file_uploader(
            "Archivo de Century",
            type=["xlsx", "csv"],
            key="conciliacion_file_century",
        )

    with col2:
        file_dnit = st.file_uploader(
            "Archivo de la DNIT",
            type=["xlsx", "csv"],
            key="conciliacion_file_dnit",
        )

    if not file_century or not file_dnit:
        return

    try:
        with st.spinner("Procesando conciliación..."):
            df_century = _leer_archivo_subido(file_century, header=None)
            df_dnit = _leer_archivo_subido(file_dnit, header=0)
            df_dnit_filtrado, df_century_filtrado = _procesar_conciliacion(df_century, df_dnit)

        total_dnit = len(df_dnit)
        total_coincidentes = len(df_dnit_filtrado)
        total_unicos_century = len(df_century_filtrado)
        eficiencia = (total_coincidentes / total_dnit * 100) if total_dnit else 0

        st.success("Cruce finalizado con éxito.")

        met1, met2, met3 = st.columns(3)
        met1.metric("Total Coincidentes", total_coincidentes)
        met2.metric("Únicos de Century", total_unicos_century)
        met3.metric("Eficiencia del Cruce", f"{eficiencia:.1f}%")

        csv_dnit = df_dnit_filtrado.to_csv(index=False, encoding="utf-8")
        if df_century_filtrado.shape[1] > 5:
            df_century_filtrado[5] = df_century_filtrado[5].astype(str).str.slice(0, 10)
        txt_century = df_century_filtrado.to_csv(
            sep="\t",
            index=False,
            header=False,
            encoding="utf-8",
        )

        st.subheader("Descargar resultados")
        down1, down2 = st.columns(2)

        with down1:
            st.download_button(
                label="Descargar DNIT filtrado CSV",
                data=csv_dnit,
                file_name="dnit_coincidentes.csv",
                mime="text/csv",
                width="stretch",
            )

        with down2:
            st.download_button(
                label="Descargar Century sin duplicados TXT",
                data=txt_century,
                file_name="century_sin_duplicados.txt",
                mime="text/plain",
                width="stretch",
            )

    except Exception as exc:
        st.error(
            "No se pudo procesar la conciliación. "
            f"Verificá que los archivos correspondan a Century y DNIT. Detalle: {exc}"
        )
