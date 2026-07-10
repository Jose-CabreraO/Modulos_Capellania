from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st


CENTURY_FACTURA_INDEX = 7
DNIT_COMPROBANTE_COL = "Número de Comprobante"
DNIT_FECHA_COL = "Fecha Emisión"
SUPPORTED_INPUT_FORMATS = ["csv", "xlsx", "xls"]
XLS_MAX_ROWS = 65536
XLS_MAX_COLUMNS = 256


def _extension_archivo(uploaded_file):
    nombre = uploaded_file.name.lower()
    return nombre.rsplit(".", 1)[-1] if "." in nombre else ""


def _nombre_base(uploaded_file):
    return Path(uploaded_file.name).stem


def _nombre_original(uploaded_file):
    return uploaded_file.name


def _nombre_con_extension(uploaded_file, extension):
    return f"{_nombre_base(uploaded_file)}.{extension}"


def _resetear_archivo(uploaded_file):
    try:
        uploaded_file.seek(0)
    except Exception:
        pass


def _leer_archivo_subido(uploaded_file, *, header):
    extension = _extension_archivo(uploaded_file)
    _resetear_archivo(uploaded_file)

    try:
        if extension == "csv":
            return pd.read_csv(uploaded_file, header=header, dtype=str)
        if extension == "xlsx":
            return pd.read_excel(uploaded_file, header=header, dtype=str, engine="openpyxl")
        if extension == "xls":
            return pd.read_excel(uploaded_file, header=header, dtype=str, engine="xlrd")
    except ImportError as exc:
        raise ValueError(
            "Falta una dependencia para leer este archivo. "
            "Use openpyxl para .xlsx y xlrd para .xls."
        ) from exc
    except pd.errors.EmptyDataError as exc:
        raise ValueError("El archivo está vacío o no contiene datos tabulares.") from exc
    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(
            "No se pudo leer el archivo. Verificá que no esté corrupto, "
            f"que la extensión .{extension} coincida con el contenido y que el formato sea válido. "
            f"Detalle técnico: {exc}"
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


def _exportar_csv(df, *, header=True):
    return df.to_csv(index=False, header=header, encoding="utf-8")


def _exportar_txt_tabular(df):
    return df.to_csv(sep="\t", index=False, header=False, encoding="utf-8")


def _forzar_texto(df):
    return df.astype(str)


def _valor_texto(value):
    if pd.isna(value):
        return ""
    return str(value)


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


def _exportar_xls(df, *, header=True):
    if len(df) + (1 if header else 0) > XLS_MAX_ROWS or len(df.columns) > XLS_MAX_COLUMNS:
        raise ValueError(
            "El formato .xls solo permite 65.536 filas y 256 columnas. "
            "Descargá el resultado en .xlsx."
        )

    try:
        import xlwt
    except ImportError as exc:
        raise ValueError(
            "La exportación a .xls requiere la dependencia xlwt. "
            "Descargá el resultado en .xlsx o instalá xlwt."
        ) from exc

    output = BytesIO()
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet("Datos")
    text_style = xlwt.XFStyle()
    text_style.num_format_str = "@"

    row_offset = 0
    if header:
        for col_idx, column in enumerate(df.columns):
            worksheet.write(0, col_idx, _valor_texto(column), text_style)
        row_offset = 1

    for row_idx, row in enumerate(_forzar_texto(df).itertuples(index=False), start=row_offset):
        for col_idx, value in enumerate(row):
            worksheet.write(row_idx, col_idx, _valor_texto(value), text_style)

    workbook.save(output)
    output.seek(0)
    return output


def _exportar_en_formato(df, extension, *, header=True):
    if extension == "csv":
        return _exportar_csv(df, header=header), "text/csv"
    if extension == "xlsx":
        return (
            _exportar_xlsx(df, header=header),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    if extension == "xls":
        return _exportar_xls(df, header=header), "application/vnd.ms-excel"
    raise ValueError("Formato de salida no soportado.")


def _render_download(label, data, file_name, mime):
    st.download_button(
        label=label,
        data=data,
        file_name=file_name,
        mime=mime,
        width="stretch",
    )


def _render_descarga_mismo_formato(uploaded_file, df, *, header, label):
    extension = _extension_archivo(uploaded_file)
    data, mime = _exportar_en_formato(df, extension, header=header)
    _render_download(
        label=label,
        data=data,
        file_name=_nombre_original(uploaded_file),
        mime=mime,
    )


def _render_descargas_conciliacion(df_dnit_filtrado, df_century_filtrado, file_dnit, file_century):
    st.subheader("Descargar resultados")

    st.caption("Descarga principal: conserva exactamente el nombre y formato del archivo original.")
    century_col, dnit_col = st.columns(2)

    with century_col:
        st.markdown("**Century**")
        try:
            _render_descarga_mismo_formato(
                file_century,
                df_century_filtrado,
                header=False,
                label="Descargar Century original",
            )
        except Exception as exc:
            st.warning(f"No se pudo preparar Century en formato original: {exc}")

    with dnit_col:
        st.markdown("**DNIT**")
        try:
            _render_descarga_mismo_formato(
                file_dnit,
                df_dnit_filtrado,
                header=True,
                label="Descargar DNIT original",
            )
        except Exception as exc:
            st.warning(f"No se pudo preparar DNIT en formato original: {exc}")

    csv_dnit = _exportar_csv(df_dnit_filtrado, header=True)
    xlsx_dnit = _exportar_xlsx(df_dnit_filtrado, header=True)
    txt_century = _exportar_txt_tabular(df_century_filtrado)
    xlsx_century = _exportar_xlsx(df_century_filtrado, header=False)

    st.caption("Descargas alternativas compatibles.")
    century_alt, dnit_alt = st.columns(2)
    with century_alt:
        _render_download(
            "Descargar Century TXT",
            txt_century,
            _nombre_con_extension(file_century, "txt"),
            "text/plain",
        )
        _render_download(
            "Descargar Century XLSX",
            xlsx_century,
            _nombre_con_extension(file_century, "xlsx"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with dnit_alt:
        _render_download(
            "Descargar DNIT CSV",
            csv_dnit,
            _nombre_con_extension(file_dnit, "csv"),
            "text/csv",
        )
        _render_download(
            "Descargar DNIT XLSX",
            xlsx_dnit,
            _nombre_con_extension(file_dnit, "xlsx"),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def _convertir_archivo(uploaded_file, *, tiene_encabezados, formato_salida):
    header = 0 if tiene_encabezados else None
    df = _leer_archivo_subido(uploaded_file, header=header)
    _validar_dataframe_no_vacio(df, "convertido")

    extension = formato_salida.lower()
    data, mime = _exportar_en_formato(df, extension, header=tiene_encabezados)
    return data, _nombre_con_extension(uploaded_file, extension), mime


def _render_convertidor_archivos():
    st.divider()
    st.header("Convertidor de archivos")
    st.write(
        "Convertí archivos CSV, XLSX o XLS manteniendo los datos como texto. "
        "Cuando sea posible, el archivo convertido conserva el nombre base original."
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

    _render_download(f"Descargar {formato_salida}", data, file_name, mime)


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

            _render_descargas_conciliacion(df_dnit_filtrado, df_century_filtrado, file_dnit, file_century)

        except Exception as exc:
            st.error(
                "No se pudo procesar la conciliación. "
                f"Verificá que los archivos correspondan a Century y DNIT. Detalle: {exc}"
            )

    _render_convertidor_archivos()
