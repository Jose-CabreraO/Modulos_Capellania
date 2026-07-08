import re
import textwrap
import unicodedata
from io import BytesIO

import pandas as pd
from fpdf import FPDF


DAY_ORDER = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"]
DAY_ALIASES = {
    "lunes": "Lunes",
    "martes": "Martes",
    "miercoles": "Miércoles",
    "miércoles": "Miércoles",
    "jueves": "Jueves",
    "viernes": "Viernes",
    "sabado": "Sábado",
    "sábado": "Sábado",
}
EXPECTED_COLUMNS = [
    "ListadoId",
    "Capellan",
    "Empresa",
    "Sucursal",
    "Sección",
    "Dia",
    "Hora",
    "RangoTiempoCarga",
    "FechaAct",
]


def normalize_text(value):
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFD", text.lower())
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    return re.sub(r"\s+", " ", text).strip()


def normalize_day(value):
    key = normalize_text(value)
    return DAY_ALIASES.get(key, str(value).strip())


def _is_day_token(value):
    return normalize_text(value) in {normalize_text(day) for day in DAY_ORDER}


def _read_by_wide_separator(text):
    rows = [line.strip() for line in text.splitlines() if line.strip()]
    if not rows:
        raise ValueError("No se detectaron filas para procesar.")

    header = re.split(r"\s{2,}|\t+", rows[0].strip())
    if len(header) < 5:
        return None

    parsed_rows = [re.split(r"\s{2,}|\t+", row.strip()) for row in rows[1:]]
    valid_rows = [row for row in parsed_rows if len(row) == len(header)]
    if not valid_rows:
        return None

    return pd.DataFrame(valid_rows, columns=header)


def _parse_line_fallback(line):
    tokens = line.split()
    if len(tokens) < 9 or not tokens[0].isdigit():
        return None

    day_idx = next((idx for idx, token in enumerate(tokens) if _is_day_token(token)), None)
    if day_idx is None or day_idx + 4 >= len(tokens):
        return None

    listado_id = tokens[0]
    dia = tokens[day_idx]
    hora = tokens[day_idx + 1]
    rango = " ".join(tokens[day_idx + 2:day_idx + 4])
    fecha = tokens[day_idx + 4]

    middle = tokens[1:day_idx]
    if len(middle) < 4:
        return None

    seccion = middle[-1]
    sucursal = middle[-2]
    capellan_tokens = middle[:4] if len(middle) >= 6 else middle[: max(1, len(middle) - 3)]
    empresa_tokens = middle[len(capellan_tokens):-2]

    return {
        "ListadoId": listado_id,
        "Capellan": " ".join(capellan_tokens),
        "Empresa": " ".join(empresa_tokens) or "---",
        "Sucursal": sucursal,
        "Sección": seccion,
        "Dia": dia,
        "Hora": hora,
        "RangoTiempoCarga": rango,
        "FechaAct": fecha,
    }


def parse_schedule_text(text):
    if not text or not text.strip():
        raise ValueError("Pegá el listado de horarios antes de procesar.")

    df = _read_by_wide_separator(text)
    if df is None:
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        data_lines = lines[1:] if lines and "ListadoId" in lines[0] else lines
        records = [_parse_line_fallback(line) for line in data_lines]
        records = [record for record in records if record]
        if not records:
            raise ValueError("No se pudo interpretar el texto pegado. Verificá encabezados y columnas.")
        df = pd.DataFrame(records)

    rename_map = {}
    normalized_columns = {normalize_text(col): col for col in df.columns}
    for expected in EXPECTED_COLUMNS:
        found = normalized_columns.get(normalize_text(expected))
        if found:
            rename_map[found] = expected
    df = df.rename(columns=rename_map)

    missing = [col for col in ["Empresa", "Sucursal", "Dia", "Hora"] if col not in df.columns]
    if missing:
        raise ValueError("Faltan columnas obligatorias: " + ", ".join(missing))

    df = df.dropna(subset=["Dia", "Hora"]).copy()
    df["Dia"] = df["Dia"].map(normalize_day)
    df = df[df["Dia"].isin(DAY_ORDER)]
    if df.empty:
        raise ValueError("No se detectaron filas con días válidos de Lunes a Sábado.")

    parsed_hours = pd.to_datetime(df["Hora"].astype(str), format="%H:%M", errors="coerce")
    df = df[parsed_hours.notna()].copy()
    if df.empty:
        raise ValueError("No se detectaron horas válidas con formato HH:MM.")

    df["Hora"] = parsed_hours[parsed_hours.notna()].dt.strftime("%H:%M")
    df["Celda"] = df["Empresa"].astype(str).str.strip() + " - " + df["Sucursal"].astype(str).str.strip()
    return df


def build_schedule_grid(df):
    hours = sorted(df["Hora"].unique(), key=lambda value: pd.to_datetime(value, format="%H:%M"))
    grid = pd.DataFrame("---", index=hours, columns=DAY_ORDER)
    grid.index.name = "Hora"

    for _, row in df.iterrows():
        current = grid.loc[row["Hora"], row["Dia"]]
        value = row["Celda"]
        grid.loc[row["Hora"], row["Dia"]] = value if current == "---" else f"{current}\n{value}"

    return grid.reset_index()


def highlight_occupied_cells(value):
    if value == "---":
        return "background-color: #FAFAFA; color: #9CA3AF;"
    return "background-color: #E8F3EE; color: #173B2F; font-weight: 600;"


def _wrapped_cell(text, width=18):
    if text == "---":
        return text
    return "\n".join(textwrap.wrap(str(text), width=width, break_long_words=False))


def generate_schedule_pdf(grid, title="Organizador de horarios"):
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=False)
    pdf.add_page()
    pdf.set_title(title)

    margin = 8
    page_width = 210
    usable_width = page_width - margin * 2
    hour_width = 18
    day_width = (usable_width - hour_width) / len(DAY_ORDER)
    row_height = max(12, min(22, (297 - 34) / max(1, len(grid))))

    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(31, 41, 55)
    pdf.cell(0, 10, title, align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    pdf.set_font("Helvetica", "B", 7)
    pdf.set_fill_color(219, 234, 254)
    pdf.set_draw_color(203, 213, 225)
    pdf.cell(hour_width, 9, "Hora", border=1, align="C", fill=True)
    for day in DAY_ORDER:
        pdf.cell(day_width, 9, day, border=1, align="C", fill=True)
    pdf.ln()

    pdf.set_font("Helvetica", "", 6)
    for _, row in grid.iterrows():
        y_start = pdf.get_y()
        if y_start + row_height > 287:
            pdf.add_page()
            y_start = pdf.get_y()

        pdf.set_fill_color(241, 245, 249)
        pdf.cell(hour_width, row_height, str(row["Hora"]), border=1, align="C", fill=True)

        for day in DAY_ORDER:
            text = _wrapped_cell(row[day])
            x = pdf.get_x()
            y = pdf.get_y()
            fill = row[day] != "---"
            pdf.set_fill_color(232, 243, 238) if fill else pdf.set_fill_color(250, 250, 250)
            pdf.multi_cell(day_width, 3.6, text, border=1, align="C", fill=True, max_line_height=3.6)
            pdf.set_xy(x + day_width, y)
        pdf.set_y(y_start + row_height)

    output = BytesIO()
    pdf.output(output)
    output.seek(0)
    return output


def build_ui():
    import streamlit as st

    st.title("Organizador de Horarios")
    raw_text = st.text_area(
        "Pegá el listado copiado desde la web",
        height=220,
        placeholder=(
            "ListadoId Capellan Empresa Sucursal Sección Dia Hora RangoTiempoCarga FechaAct\n"
            "1878 Souberlich Vidal Rodrigo Sebastin Inverfín S.A.E.C.A. Areguá General Martes 09:30 1 Hora/s 2026-05-12"
        ),
    )

    if not raw_text.strip():
        return

    try:
        parsed = parse_schedule_text(raw_text)
        grid = build_schedule_grid(parsed)
    except Exception as exc:
        st.error(f"No se pudo procesar el texto pegado: {exc}")
        return

    st.caption(f"Registros procesados: {len(parsed)}")
    styled = grid.style.map(highlight_occupied_cells, subset=DAY_ORDER)
    st.dataframe(styled, width="stretch", hide_index=True)

    pdf = generate_schedule_pdf(grid)
    st.download_button(
        "Descargar PDF para imprimir",
        data=pdf,
        file_name="organizador_horarios.pdf",
        mime="application/pdf",
    )
