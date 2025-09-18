# app.py — Pareto Comunidad (DELITO / RIESGO SOCIAL / OTROS FACTORES)
# Cuenta celdas con dato en columnas AI:ET (hoja 'matriz') y arma el Pareto.

import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Pareto Comunidad – MSP", layout="wide", initial_sidebar_state="collapsed")
TOP_N_GRAFICO = 60

# ===================== CATEGORÍAS PERMITIDAS =====================
CATEGORIAS_VALIDAS = {"DELITO", "RIESGO SOCIAL", "OTROS FACTORES"}

def _force_cat(x: str) -> str:
    n = (x or "").strip().upper()
    return n if n in CATEGORIAS_VALIDAS else "OTROS FACTORES"

# ===================== Diccionario embebido (AMPLIABLE por ti) =====================
# Si un descriptor no aparece aquí, se categoriza como "OTROS FACTORES".
CAT_MAP = {
    # ---- DELITO ----
    "Venta de drogas": "DELITO",
    "Consumo de drogas": "DELITO",
    "Bunker (Puntos de venta y consumo de drogas)": "DELITO",
    "Hurto": "DELITO",
    "Robo a personas": "DELITO",
    "Robo a vivienda (Tacha)": "DELITO",
    "Robo a vivienda (Intimidación)": "DELITO",
    "Robo a vehículos (Tacha)": "DELITO",
    "Robo a comercio (Intimidación)": "DELITO",
    "Robo a comercio (Tacha)": "DELITO",
    "Robo de vehículos": "DELITO",
    "Receptación": "DELITO",
    "Estafas o defraudación": "DELITO",
    "Daños/Vandalismo": "DELITO",
    "Lesiones": "DELITO",
    "Contrabando": "DELITO",
    "Homicidios": "DELITO",
    "Extorsión": "DELITO",
    "Delitos sexuales": "DELITO",
    "Estafa informática": "DELITO",
    "Robo de cable": "DELITO",
    "Robo de bienes agrícola": "DELITO",
    "Robo a edificacion (Tacha)": "DELITO",
    "Robo a transporte público con Intimidación": "DELITO",

    # ---- RIESGO SOCIAL ----
    "Falta de inversion social": "RIESGO SOCIAL",
    "Falta de oportunidades laborales.": "RIESGO SOCIAL",
    "Personas con exceso de tiempo de ocio": "RIESGO SOCIAL",
    "Violencia intrafamiliar": "RIESGO SOCIAL",
    "Desvinculación escolar": "RIESGO SOCIAL",
    "Abandono de personas (Menor de edad, adulto mayor o capacidades diferentes)": "RIESGO SOCIAL",

    # ---- OTROS FACTORES ----
    "Consumo de alcohol en vía pública": "OTROS FACTORES",
    "Deficiencia en la infraestructura vial": "OTROS FACTORES",
    "Contaminacion Sonica": "OTROS FACTORES",
    "Lotes baldíos.": "OTROS FACTORES",
    "Falta de salubridad publica": "OTROS FACTORES",
    "Disturbios(Riñas)": "OTROS FACTORES",
    "Personas en situación de calle.": "OTROS FACTORES",
    "Usurpacion de terrenos (Precarios)": "OTROS FACTORES",
    "Pérdida de espacios públicos": "OTROS FACTORES",
    "Deficiencias en el alumbrado publico": "OTROS FACTORES",
    "Acoso sexual callejero": "OTROS FACTORES",
    "Hospedajes ilegales (Cuarterías)": "OTROS FACTORES",
    "Ventas informales (Ambulantes)": "OTROS FACTORES",
    "Maltrato animal": "OTROS FACTORES",
    "Tala ilegal": "OTROS FACTORES",
    "Trata de personas": "OTROS FACTORES",
    "Explotacion Laboral infantil": "OTROS FACTORES",
    "Caza ilegal": "OTROS FACTORES",
    "Abigeato (Robo y destace de ganado)": "OTROS FACTORES",
    "Zona de prostitución": "OTROS FACTORES",
    "Explotacion Sexual infantil": "OTROS FACTORES",
    "Tráfico ilegal de personas": "OTROS FACTORES",
    "Robo de combustible": "OTROS FACTORES",
    "Pesca ilegal": "OTROS FACTORES",
}

# ===================== Utilidades =====================
def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def norm_text(s: Optional[str]) -> str:
    if s is None:
        return ""
    s = strip_accents(str(s)).lower().strip()
    s = re.sub(r"\s+", " ", s)
    return s

def excel_col_to_index(col: str) -> int:
    col = col.strip().upper()
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

def read_matriz(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(BytesIO(file_bytes), sheet_name="matriz", engine="openpyxl")

# celdas que cuentan como “dato”
NEGATIVOS = {"", "0", "no", "false", "nan", "ninguno", "n/a", "na"}

def cell_has_data(v) -> bool:
    if pd.isna(v): 
        return False
    # numérico ≠ 0
    try:
        f = float(str(v).replace(",", "."))
        return f != 0.0
    except Exception:
        pass
    # texto no negativo
    return norm_text(v) not in NEGATIVOS

# ===================== Núcleo: contar AI:ET por columna =====================
def counts_ai_et(df: pd.DataFrame) -> pd.DataFrame:
    # límites AI:ET (0-based)
    L = excel_col_to_index("AI") - 1
    R = excel_col_to_index("ET") - 1
    L = max(0, min(L, df.shape[1]-1))
    R = max(0, min(R, df.shape[1]-1))
    if L > R:
        L, R = 0, df.shape[1]-1

    sub = df.iloc[:, L:R+1]

    # evitar filas totalmente vacías dentro del rango
    rowmask = sub.applymap(cell_has_data).any(axis=1)
    sub = sub.loc[rowmask]

    rows = []
    for c in sub.columns:
        freq = int(sub[c].apply(cell_has_data).sum())
        if freq > 0:
            rows.append((str(df.columns[df.columns.get_loc(c)]), freq))
    if not rows:
        return pd.DataFrame({"Descriptor": [], "Frecuencia": []})
    out = pd.DataFrame(rows, columns=["Descriptor", "Frecuencia"])
    return out

# ===================== Pareto =====================
def build_pareto(base: pd.DataFrame) -> pd.DataFrame:
    if base.empty:
        return pd.DataFrame(columns=["Categoría","Descriptor","Frecuencia","Porcentaje","% acumulado","Acumulado","80/20"])

    df = base.copy()
    # Categoría (solo 3). Si no está en el mapa, cae en OTROS FACTORES.
    df["Categoría"] = df["Descriptor"].map(CAT_MAP).apply(_force_cat)

    total = int(df["Frecuencia"].sum())
    df = df.sort_values(["Frecuencia","Descriptor"], ascending=[False, True], ignore_index=True)
    df["Porcentaje"]  = df["Frecuencia"] / total      # fracción 0–1
    df["% acumulado"] = df["Porcentaje"].cumsum()     # fracción 0–1
    df["Acumulado"]   = df["Frecuencia"].cumsum()
    df["80/20"]       = "80%"

    assert int(df["Acumulado"].iloc[-1]) == total
    return df[["Categoría","Descriptor","Frecuencia","Porcentaje","% acumulado","Acumulado","80/20"]]

# ===================== Excel con formato + gráfico =====================
def export_excel(pareto: pd.DataFrame, titulo: str = "PARETO COMUNIDAD") -> bytes:
    from pandas import ExcelWriter
    import xlsxwriter  # noqa: F401

    out = BytesIO()
    with ExcelWriter(out, engine="xlsxwriter") as writer:
        sheet = "Pareto Comunidad"
        pareto.to_excel(writer, index=False, sheet_name=sheet)
        wb = writer.book
        ws = writer.sheets[sheet]
        n = len(pareto)
        if not n:
            return out.getvalue()

        fmt_head = wb.add_format({"bold": True, "align": "center", "bg_color": "#D9E1F2", "border": 1})
        fmt_pct  = wb.add_format({"num_format": "0,00%", "align": "right", "border": 1})
        fmt_int  = wb.add_format({"num_format": "#,##0", "align": "center", "border": 1})
        fmt_txt  = wb.add_format({"align": "left", "border": 1})
        fmt_cent = wb.add_format({"align": "center"})
        fmt_yel  = wb.add_format({"bg_color": "#FFF2CC"})

        ws.set_row(0, None, fmt_head)
        ws.set_column("A:A", 22, fmt_txt)
        ws.set_column("B:B", 52, fmt_txt)
        ws.set_column("C:C", 12, fmt_int)
        ws.set_column("D:D", 12, fmt_pct)   # fracción → 0,00%
        ws.set_column("E:E", 12, fmt_pct)
        ws.set_column("F:F", 12, fmt_int)
        ws.set_column("G:G", 8,  fmt_cent)

        cutoff_idx = int((pareto["% acumulado"] <= 0.80).sum())
        if cutoff_idx > 0:
            ws.conditional_format(1, 0, cutoff_idx, 6, {"type": "no_blanks", "format": fmt_yel})

        # Auxiliares para líneas
        ws.write(0, 9, "80/20");  ws.set_column("J:J", 6,  None, {"hidden": True})
        ws.write(0,10, "CorteX"); ws.set_column("K:K", 20, None, {"hidden": True})
        ws.write(0,11, "%");      ws.set_column("L:L", 6,  None, {"hidden": True})
        for i in range(n):
            ws.write_number(i+1, 9, 0.80)

        corte_row = max(1, cutoff_idx)
        xcat = pareto.iloc[corte_row-1]["Descriptor"]
        ws.write(1,10, xcat); ws.write(2,10, xcat)
        ws.write_number(1,11, 0.0); ws.write_number(2,11, 1.0)

        # Barras
        chart = wb.add_chart({'type': 'column'})
        points = [{"fill": {"color": "#5B9BD5"}} for _ in range(n)]
        for i in range(cutoff_idx, n):
            points[i] = {"fill": {"color": "#A6A6A6"}}
        chart.add_series({
            'name': 'Frecuencia',
            'categories': [sheet, 1, 1, n, 1],
            'values':     [sheet, 1, 2, n, 2],
            'points': points,
        })

        # % acumulado (eje secundario 0–100%)
        line = wb.add_chart({'type': 'line'})
        line.add_series({
            'name': '% acumulado',
            'categories': [sheet, 1, 1, n, 1],
            'values':     [sheet, 1, 4, n, 4],
            'y2_axis': True,
            'line': {'color': '#ED7D31', 'width': 2.0}
        })
        chart.combine(line)

        # 80%
        h80 = wb.add_chart({'type': 'line'})
        h80.add_series({
            'name': '80/20',
            'categories': [sheet, 1, 1, n, 1],
            'values':     [sheet, 1, 9, n, 9],
            'y2_axis': True,
            'line': {'color': '#7F7F7F', 'width': 1.25}
        })
        chart.combine(h80)

        # Línea vertical
        vline = wb.add_chart({'type': 'line'})
        vline.add_series({
            'name': '',
            'categories': [sheet, 1, 10, 2, 10],
            'values':     [sheet, 1, 11, 2, 11],
            'y2_axis': True,
            'line': {'color': '#C00000', 'width': 2.25},
            'marker': {'type': 'none'},
        })
        chart.combine(vline)

        chart.set_title({'name': titulo})
        chart.set_plotarea({'border': {'none': True}})
        chart.set_chartarea({'border': {'none': True}})
        chart.set_x_axis({'num_font': {'rotation': -50}})
        chart.set_y_axis({'major_gridlines': {'visible': False}})
        chart.set_y2_axis({'min': 0, 'max': 1, 'major_unit': 0.1, 'num_format': '0%'})
        chart.set_legend({'position': 'bottom'})
        ws.insert_chart(1, 9, chart, {'x_scale': 1.9, 'y_scale': 1.6})

    return out.getvalue()

# ===================== UI =====================
st.title("Pareto Comunidad – MSP (DELITO / RIESGO SOCIAL / OTROS FACTORES)")

up = st.file_uploader("📄 Subí la Plantilla (XLSX) – hoja `matriz`", type=["xlsx"])
if not up:
    st.info("Subí la Plantilla para procesar.")
    st.stop()

try:
    df = read_matriz(up.getvalue())
except Exception as e:
    st.error(f"Error leyendo 'matriz': {e}")
    st.stop()

st.caption(f"Vista previa (primeras 20 de {len(df)} filas)")
st.dataframe(df.head(20), use_container_width=True)

base = counts_ai_et(df)
if base.empty or base["Frecuencia"].sum() == 0:
    st.warning("No se hallaron datos en el rango AI:ET. Revisá la plantilla.")
    st.stop()

pareto = build_pareto(base)

# Mostrar con coma decimal
def pct_str(frac: float) -> str:
    return f"{frac*100:.2f}%".replace(".", ",")

display = pareto.copy()
display["Porcentaje"] = display["Porcentaje"].apply(pct_str)
display["% acumulado"] = display["% acumulado"].apply(pct_str)

TOTAL = int(pareto["Acumulado"].iloc[-1])
st.subheader(f"Pareto Comunidad (TOTAL = {TOTAL:,})")
st.dataframe(display, use_container_width=True)

# Gráfico rápido (pantalla)
import altair as alt
top_df = pareto.head(TOP_N_GRAFICO).copy()
bars = alt.Chart(top_df).mark_bar().encode(
    x=alt.X('Descriptor:N', sort=None, axis=alt.Axis(labelAngle=-50)),
    y=alt.Y('Frecuencia:Q')
)
line = alt.Chart(top_df).mark_line(point=True).encode(
    x='Descriptor:N',
    y=alt.Y('% acumulado:Q', axis=alt.Axis(format='%'), scale=alt.Scale(domain=[0,1])),
    color=alt.value('#ED7D31')
)
h80 = alt.Chart(pd.DataFrame({'y':[0.8]})).mark_rule().encode(y=alt.Y('y:Q', axis=alt.Axis(format='%')))
st.altair_chart((bars + line + h80).resolve_scale(y='independent'), use_container_width=True)

# Descargar Excel
st.subheader("Descargar Excel final")
st.download_button(
    "⬇️ Pareto Comunidad (Excel con formato y gráfico)",
    data=export_excel(pareto, titulo="PARETO COMUNIDAD"),
    file_name="Pareto_Comunidad.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)





