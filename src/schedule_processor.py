import os
import re
import tempfile
import textwrap
import unicodedata
from io import BytesIO

os.environ.setdefault("MPLCONFIGDIR", tempfile.gettempdir())
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd


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


def get_capellan_title(df):
    if "Capellan" not in df.columns:
        return "Horario"
    capellanes = df["Capellan"].dropna().astype(str).str.strip()
    capellanes = capellanes[capellanes != ""]
    if capellanes.empty:
        return "Horario"
    return f"Horario - {capellanes.iloc[0]}"


def highlight_occupied_cells(value):
    return "background-color: #FFFFFF; color: #111827; font-weight: 500;"


def _wrapped_cell(text, width=18):
    if text == "---":
        return text
    return "\n".join(textwrap.wrap(str(text), width=width, break_long_words=False))


def generate_schedule_pdf(grid, title="Organizador de horarios"):
    fig, ax = plt.subplots(figsize=(11.69, 8.27))
    ax.axis("off")
    ax.set_title(title, fontsize=18, fontweight="bold", pad=12, color="#1F2937")

    columns = list(grid.columns)
    cell_text = [
        [_wrapped_cell(row[column], width=18 if column != "Hora" else 8) for column in columns]
        for _, row in grid.iterrows()
    ]
    cell_colours = []
    for _, row in grid.iterrows():
        row_colours = []
        for column in columns:
            if column == "Hora":
                row_colours.append("#E0F2FE")
            else:
                row_colours.append("#FFFFFF")
        cell_colours.append(row_colours)

    tabla = ax.table(
        cellText=cell_text,
        colLabels=columns,
        cellColours=cell_colours,
        colColours=["#BFDBFE"] * len(columns),
        colWidths=[0.09] + [0.151] * len(DAY_ORDER),
        cellLoc="center",
        loc="center",
        bbox=[0.0, 0.01, 1.0, 0.93],
    )
    tabla.auto_set_font_size(False)
    tabla.set_fontsize(12)
    tabla.scale(1.2, 3.2)

    for (row, col), cell in tabla.get_celld().items():
        cell.set_edgecolor("#CBD5E1")
        cell.set_linewidth(0.8)
        if row == 0:
            cell.set_text_props(weight="bold", color="#1E3A8A", fontsize=12)
        elif col == 0:
            cell.set_facecolor("#E0F2FE")
            cell.set_text_props(weight="bold", color="#075985", fontsize=12)
        else:
            cell.set_facecolor("#FFFFFF")
            cell.set_text_props(color="#111827", fontsize=12)

    fig.subplots_adjust(left=0.01, right=0.99, top=0.91, bottom=0.01)
    output = BytesIO()
    fig.savefig(output, format="pdf", orientation="landscape", bbox_inches="tight")
    plt.close(fig)
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
        title = get_capellan_title(parsed)
    except Exception as exc:
        st.error(f"No se pudo procesar el texto pegado: {exc}")
        return

    st.caption(f"Registros procesados: {len(parsed)}")
    styled = (
        grid.style
        .map(highlight_occupied_cells, subset=DAY_ORDER)
        .set_properties(subset=["Hora"], **{"background-color": "#E0F2FE", "font-weight": "700", "color": "#075985"})
        .set_table_styles([
            {"selector": "th", "props": [("background-color", "#BFDBFE"), ("color", "#1E3A8A"), ("font-weight", "700")]},
        ])
    )
    st.dataframe(styled, width="stretch", hide_index=True)

    pdf = generate_schedule_pdf(grid, title=title)
    st.download_button(
        "Descargar PDF para imprimir",
        data=pdf,
        file_name="organizador_horarios.pdf",
        mime="application/pdf",
    )
