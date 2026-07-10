from io import BytesIO

import pandas as pd
import streamlit as st


CENTURY_FACTURA_INDEX = 7
DNIT_COMPROBANTE_COL = "Número de Comprobante"
DNIT_FECHA_COL = "Fecha Emisión"
SUPPORTED_INPUT_FORMATS = ["csv", "xlsx", "xls"]


def _extension_archivo(uploaded_file):
    nombre = uploaded_file.name.lower()
    return nombre.rsplit(".", 1)[-1] if "." in nombre else ""


def _leer_archivo_subido(uploaded_file, *, header):
    extension = _extension_archivo(uploaded_file)

    try:
        if extension == "csv":
            return pd.read_csv(uploaded_file, header=header, dtype=str)
        if extension in {"xlsx", "xls"}:
            return pd.read_excel(uploaded_file, header=header, dtype=str)
    except ImportError as exc:
        raise ValueError(
            "Falta una dependencia para leer este archivo. "
            "Use openpyxl para .xlsx y xlrd para .xls."
        ) from exc
    except pd.errors.EmptyDataError as exc:
        raise ValueError("El archivo está vacío o no contiene datos tabulares.") from exc
    except Exception as exc:
        raise ValueError(
            "No se pudo leer el archivo. Verificá que no esté corrupto y que su formato sea válido."
        ) from exc

    raise ValueError("Formato no soportado. Subí un archivo .csv, .xlsx o .xls.")


def _normalizar_clave(serie):
    return serie.astype(str).str.strip()


def _validar_dataframe_no_vacio(df, nombre):
    if df.empty:
        raise ValueError(f"El archivo {nombre} no contiene filas para procesar.")


def _validar_estructura(df_century, df_dnit):
    _validar_dataframe_no_vacio(df_century, "Century")
    _validar_dataframe_no_vacio(df_dnit, "DNIT")

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

    if df_century_filtrado.shape[1] > 5:
        df_century_filtrado[5] = df_century_filtrado[5].astype(str).str.slice(0, 10)

    return df_dnit_filtrado, df_century_filtrado


def _exportar_csv(df):
    return df.to_csv(index=False, encoding="utf-8")


def _exportar_txt_tabular(df):
    return df.to_csv(sep="\t", index=False, header=False, encoding="utf-8")


def _forzar_texto(df):
    return df.astype(str)


def _exportar_xlsx(df, *, header=True):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        _forzar_texto(df).to_excel(writer, index=False, header=header, sheet_name="Datos")
        worksheet = writer.sheets["Datos"]
        for row in worksheet.iter_rows():
            for cell in row:
                cell.number_format = "@"
    output.seek(0)
    return output


def _render_descargas_conciliacion(df_dnit_filtrado, df_century_filtrado):
    csv_dnit = _exportar_csv(df_dnit_filtrado)
    xlsx_dnit = _exportar_xlsx(df_dnit_filtrado, header=True)
    txt_century = _exportar_txt_tabular(df_century_filtrado)
    xlsx_century = _exportar_xlsx(df_century_filtrado, header=False)

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
            label="Descargar DNIT filtrado XLSX",
            data=xlsx_dnit,
            file_name="dnit_coincidentes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    down3, down4 = st.columns(2)
    with down3:
        st.download_button(
            label="Descargar Century sin duplicados TXT",
            data=txt_century,
            file_name="century_sin_duplicados.txt",
            mime="text/plain",
            width="stretch",
        )
    with down4:
        st.download_button(
            label="Descargar Century sin duplicados XLSX",
            data=xlsx_century,
            file_name="century_sin_duplicados.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )


def _convertir_archivo(uploaded_file, *, tiene_encabezados, formato_salida):
    header = 0 if tiene_encabezados else None
    df = _leer_archivo_subido(uploaded_file, header=header)
    _validar_dataframe_no_vacio(df, "convertido")

    if formato_salida == "CSV":
        data = df.to_csv(index=False, header=tiene_encabezados, encoding="utf-8")
        return data, "archivo_convertido.csv", "text/csv"

    if formato_salida == "XLSX":
        data = _exportar_xlsx(df, header=tiene_encabezados)
        return (
            data,
            "archivo_convertido.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    raise ValueError("La exportación a .xls no está disponible en este entorno; use .xlsx.")


def _render_convertidor_archivos():
    st.divider()
    st.header("Convertidor de archivos")
    st.write(
        "Convertí archivos CSV, XLSX o XLS manteniendo los datos como texto. "
        "La exportación a XLS heredado no está disponible; usá XLSX para Excel."
    )

    archivo = st.file_uploader(
        "Archivo a convertir",
        type=SUPPORTED_INPUT_FORMATS,
        key="convertidor_archivo",
    )

    col1, col2 = st.columns(2)
    with col1:
        tiene_encabezados = st.checkbox(
            "El archivo tiene encabezados",
            value=True,
            key="convertidor_tiene_encabezados",
        )
    with col2:
        formato_salida = st.selectbox(
            "Formato de salida",
            ["CSV", "XLSX", "XLS"],
            key="convertidor_formato_salida",
        )

    if formato_salida == "XLS":
        st.warning("La exportación a .xls no está disponible en este entorno; use .xlsx.")

    if not archivo:
        return

    try:
        data, file_name, mime = _convertir_archivo(
            archivo,
            tiene_encabezados=tiene_encabezados,
            formato_salida=formato_salida,
        )
    except Exception as exc:
        st.error(f"No se pudo convertir el archivo. Detalle: {exc}")
        return

    st.download_button(
        label=f"Descargar {formato_salida}",
        data=data,
        file_name=file_name,
        mime=mime,
        width="stretch",
    )


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
            type=SUPPORTED_INPUT_FORMATS,
            key="conciliacion_file_century",
        )

    with col2:
        file_dnit = st.file_uploader(
            "Archivo de la DNIT",
            type=SUPPORTED_INPUT_FORMATS,
            key="conciliacion_file_dnit",
        )

    if file_century and file_dnit:
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

            _render_descargas_conciliacion(df_dnit_filtrado, df_century_filtrado)

        except Exception as exc:
            st.error(
                "No se pudo procesar la conciliación. "
                f"Verificá que los archivos correspondan a Century y DNIT. Detalle: {exc}"
            )

    _render_convertidor_archivos()
