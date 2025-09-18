# Pareto Comunidad – MSP
# - 1 archivo obligatorio: Plantilla (hoja 'matriz')
# - 1 archivo opcional: 'DESCRIPTORES ACTUALIZADOS 2024 v2.xlsx' (diccionario Descriptor -> Categoría)

import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Pareto Comunidad – MSP", layout="wide", initial_sidebar_state="collapsed")
TOP_N_GRAFICO = 60

# ===================== Utilidades =====================
def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def norm_text(s: Optional[str]) -> str:
    if s is None:
        return ""
    s = str(s)
    s = strip_accents(s).lower().strip()
    s = re.sub(r"\s+", " ", s)
    return s

def make_unique_columns(cols: List[str]) -> List[str]:
    seen, out = {}, []
    for c in cols:
        nc = norm_text(c)
        seen[nc] = seen.get(nc, 0) + 1
        out.append(nc if seen[nc] == 1 else f"{nc}__{seen[nc]}")
    return out

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = make_unique_columns([str(c) for c in out.columns])
    return out

@st.cache_data(show_spinner=False)
def read_matriz(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(BytesIO(file_bytes), sheet_name="matriz", engine="openpyxl")

@st.cache_data(show_spinner=False)
def read_diccionario(file_bytes: bytes) -> pd.DataFrame:
    # intenta encontrar una hoja que tenga columnas 'Descriptor' y 'Categoria'
    xls = pd.ExcelFile(BytesIO(file_bytes))
    for sh in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sh)
        cols = {norm_text(c): c for c in df.columns}
        if "descriptor" in cols and ("categoria" in cols or "categoría" in cols):
            cat_col = cols.get("categoria", cols.get("categoría"))
            out = df[[cols["descriptor"], cat_col]].copy()
            out.columns = ["Descriptor", "Categoría"]
            out["Descriptor"] = out["Descriptor"].astype(str).str.strip()
            out["Categoría"] = out["Categoría"].astype(str).str.strip()
            out = out.dropna(subset=["Descriptor"]).drop_duplicates(subset=["Descriptor"], keep="first")
            return out
    raise ValueError("No se encontró hoja con columnas 'Descriptor' y 'Categoría'.")

# ===================== Catálogo/Sinónimos (fallback) =====================
CATEGORIA_POR_DESCRIPTOR_DEFAULT: Dict[str, str] = {
    # ← lista base (se usa si NO subes diccionario). Se puede ampliar.
    "Consumo de drogas": "DROGAS",
    "Venta de drogas": "DROGAS",
    "Hurto": "DELITOS CONTRA LA PROPIEDAD",
    "Robo a personas": "DELITOS CONTRA LA PROPIEDAD",
    "Falta de inversion social": "RIESGO SOCIAL",
    "Consumo de alcohol en vía pública": "ALCOHOL",
    "Deficiencia en la infraestructura vial": "ORDEN PÚBLICO",
    "Robo a vivienda (Tacha)": "DELITOS CONTRA LA PROPIEDAD",
    "Contaminacion Sonica": "ORDEN PÚBLICO",
    "Bunker (Puntos de venta y consumo de drogas)": "DROGAS",
    "Robo a vivienda (Intimidación)": "DELITOS CONTRA LA PROPIEDAD",
    "Disturbios(Riñas)": "ORDEN PÚBLICO",
    "Robo a vehículos (Tacha)": "DELITOS CONTRA LA PROPIEDAD",
    "Robo a vehiculos": "DELITOS CONTRA LA PROPIEDAD",
    "Robo a comercio (Intimidación)": "DELITOS CONTRA LA PROPIEDAD",
    "Daños/Vandalismo": "DELITOS CONTRA LA PROPIEDAD",
    "Robo de vehículos": "DELITOS CONTRA LA PROPIEDAD",
    "Personas en situación de calle.": "ORDEN PÚBLICO",
    "Personas con exceso de tiempo de ocio": "RIESGO SOCIAL",
    "Lesiones": "DELITOS CONTRA LA VIDA",
    "Estafas o defraudación": "DELITOS CONTRA LA PROPIEDAD",
    "Lotes baldíos.": "ORDEN PÚBLICO",
    "Falta de salubridad publica": "ORDEN PÚBLICO",
    "Falta de oportunidades laborales.": "RIESGO SOCIAL",
    "Contrabando": "DELITOS CONTRA LA PROPIEDAD",
    "Problemas Vecinales.": "RIESGO SOCIAL",
    "Robo a comercio (Tacha)": "DELITOS CONTRA LA PROPIEDAD",
    "Receptación": "DELITOS CONTRA LA PROPIEDAD",
}

SINONIMOS: Dict[str, List[str]] = {
    "Consumo de drogas": ["consumo de drogas", "consumen drogas", "consumo marihuana", "fumando piedra"],
    "Venta de drogas": ["venta de drogas", "punto de venta", "narcomenudeo"],
    "Hurto": ["hurto", "sustraccion"],
    "Robo a personas": ["robo a personas", "asalto a persona", "atraco a persona"],
    "Falta de inversion social": ["falta de inversion social"],
    "Consumo de alcohol en vía pública": ["consumo de alcohol en via publica", "licores en via publica"],
    "Deficiencia en la infraestructura vial": ["deficiencia en la infraestructura vial", "huecos", "baches"],
    "Robo a vivienda (Tacha)": ["robo a vivienda tacha"],
    "Contaminacion Sonica": ["contaminacion sonora", "ruido", "musica alta", "bulla"],
    "Bunker (Puntos de venta y consumo de drogas)": ["bunker", "bunquer", "búnker"],
    "Robo a vivienda (Intimidación)": ["robo a vivienda intimidacion", "asalto a vivienda"],
    "Disturbios(Riñas)": ["disturbios", "riñas", "riña", "peleas"],
    "Robo a vehículos (Tacha)": ["robo a vehiculos tacha"],
    "Robo a vehiculos": ["robo de vehiculos", "robo carro", "robo moto"],
    "Robo a comercio (Intimidación)": ["robo a comercio intimidacion", "asalto a comercio"],
    "Daños/Vandalismo": ["danos", "vandalismo", "grafiti", "daño a la propiedad"],
    "Robo de vehículos": ["robo de vehiculos"],
    "Personas en situación de calle.": ["personas en situacion de calle", "indigencia", "habitantes de calle"],
    "Personas con exceso de tiempo de ocio": ["exceso de tiempo de ocio", "ocio juvenil"],
    "Lesiones": ["lesiones", "lesionados", "golpiza"],
    "Estafas o defraudación": ["estafas", "defraudacion", "estafa"],
    "Lotes baldíos.": ["lotes baldios", "lote baldio"],
    "Falta de salubridad publica": ["falta de salubridad publica", "insalubridad"],
    "Falta de oportunidades laborales.": ["falta de oportunidades laborales", "desempleo"],
    "Contrabando": ["contrabando"],
    "Problemas Vecinales.": ["problemas vecinales", "conflictos vecinales"],
    "Robo a comercio (Tacha)": ["robo a comercio tacha"],
    "Receptación": ["receptacion", "compra de robado", "reduccion"],
}

# ===================== Detección =====================
def header_marked_series(s: pd.Series) -> pd.Series:
    num = pd.to_numeric(s, errors="coerce").fillna(0) != 0
    txt = s.astype(str).apply(norm_text)
    mask = ~txt.isin(["", "no", "0", "nan", "none", "false"])
    return num | mask

def build_regex_by_desc() -> Dict[str, re.Pattern]:
    compiled = {}
    for desc, keys in SINONIMOS.items():
        toks = [re.escape(norm_text(k)) for k in keys if norm_text(k)]
        if not toks:
            toks = [re.escape(norm_text(desc))]
        pat = r"(?:(?<=\s)|^)(" + "|".join(toks) + r")(?:(?=\s)|$)"
        compiled[desc] = re.compile(pat)
    return compiled

def detect_by_headers(df_norm: pd.DataFrame, regex_by_desc: Dict[str, re.Pattern]) -> Dict[str, int]:
    counts = {d: 0 for d in CATEGORIA_POR_DESCRIPTOR_DEFAULT.keys()}
    for desc, pat in regex_by_desc.items():
        hit_cols = [c for c in df_norm.columns if re.search(pat, " " + c + " ") is not None]
        if not hit_cols:
            continue
        mask_any = None
        for c in hit_cols:
            m = header_marked_series(df_norm[c])
            mask_any = m if mask_any is None else (mask_any | m)
        if mask_any is not None:
            counts[desc] += int(mask_any.sum())
    return counts

def guess_text_cols(df_norm: pd.DataFrame) -> List[str]:
    hints = ["observ", "descr", "coment", "suger", "porque", "por que", "por qué", "detalle", "problema", "actividad", "insegur"]
    out = []
    for c in df_norm.columns:
        s = df_norm[c]
        if getattr(s, "dtype", None) == object or any(h in c for h in hints):
            sample = s.astype(str).head(200).apply(norm_text)
            if (sample != "").mean() > 0.05 or any(h in c for h in hints):
                out.append(c)
    return out

def detect_in_text(df_norm: pd.DataFrame, regex_by_desc: Dict[str, re.Pattern]) -> Dict[str, int]:
    counts = {d: 0 for d in CATEGORIA_POR_DESCRIPTOR_DEFAULT.keys()}
    tcols = guess_text_cols(df_norm)
    if not tcols:
        return counts
    for desc, pat in regex_by_desc.items():
        mask_any = None
        for c in tcols:
            m = df_norm[c].astype(str).apply(norm_text).str.contains(pat, na=False)
            mask_any = m if mask_any is None else (mask_any | m)
        if mask_any is not None:
            counts[desc] += int(mask_any.sum())
    return counts

def build_copilado(counts_headers: Dict[str, int], counts_text: Dict[str, int]) -> pd.DataFrame:
    total = {}
    # unión de claves conocidas
    keys = set(counts_headers) | set(counts_text)
    for d in keys:
        total[d] = counts_headers.get(d, 0) + counts_text.get(d, 0)
    rows = [(d, f) for d, f in total.items() if f > 0]
    if not rows:
        return pd.DataFrame({"Descriptor": [], "Frecuencia": []})
    df = pd.DataFrame(rows, columns=["Descriptor", "Frecuencia"])
    return df.sort_values(["Frecuencia", "Descriptor"], ascending=[False, True], ignore_index=True)

def build_pareto(copilado: pd.DataFrame, cat_map: Dict[str, str]) -> pd.DataFrame:
    if copilado.empty:
        return pd.DataFrame(columns=["Categoría","Descriptor","Frecuencia","Porcentaje","% acumulado","Acumulado","80/20"])
    df = copilado.copy()
    # map de categoría
    df["Categoría"] = df["Descriptor"].map(cat_map).fillna("")
    # TOTAL sobre el que se sacan % (=> último acumulado)
    total = int(df["Frecuencia"].sum())
    df["Porcentaje"]  = (df["Frecuencia"] / total) * 100.0
    df = df.sort_values(["Frecuencia","Descriptor"], ascending=[False, True], ignore_index=True)
    df["% acumulado"] = df["Porcentaje"].cumsum()
    df["Acumulado"]   = df["Frecuencia"].cumsum()
    df["80/20"]       = "80%"
    # Garantía: último acumulado = TOTAL
    assert int(df["Acumulado"].iloc[-1]) == total
    return df[["Categoría","Descriptor","Frecuencia","Porcentaje","% acumulado","Acumulado","80/20"]]

# ===================== Excel (formato idéntico) =====================
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

        # formatos (coma decimal, miles)
        fmt_head = wb.add_format({"bold": True, "align": "center", "bg_color": "#D9E1F2", "border": 1})
        fmt_pct  = wb.add_format({"num_format": "0,00%", "align": "right", "border": 1})
        fmt_int  = wb.add_format({"num_format": "#,##0", "align": "center", "border": 1})
        fmt_txt  = wb.add_format({"align": "left", "border": 1})
        fmt_cent = wb.add_format({"align": "center"})
        fmt_yel  = wb.add_format({"bg_color": "#FFF2CC"})

        # columnas/encabezado
        ws.set_row(0, None, fmt_head)
        ws.set_column("A:A", 22, fmt_txt)
        ws.set_column("B:B", 52, fmt_txt)
        ws.set_column("C:C", 12, fmt_int)
        ws.set_column("D:D", 12, fmt_pct)   # Porcentaje
        ws.set_column("E:E", 12, fmt_pct)   # % acumulado
        ws.set_column("F:F", 12, fmt_int)   # Acumulado
        ws.set_column("G:G", 8,  fmt_cent)  # 80/20

        # pintar ≤80% en amarillo
        cutoff_idx = int((pareto["% acumulado"] <= 80).sum())
        if cutoff_idx > 0:
            ws.conditional_format(1, 0, cutoff_idx, 6, {"type": "no_blanks", "format": fmt_yel})

        # columnas auxiliares (ocultas)
        ws.write(0, 9, "80/20");  ws.set_column("J:J", 6,  None, {"hidden": True})
        ws.write(0,10, "CorteX"); ws.set_column("K:K", 20, None, {"hidden": True})
        ws.write(0,11, "%");      ws.set_column("L:L", 6,  None, {"hidden": True})
        for i in range(n):
            ws.write_number(i+1, 9, 0.80)  # 80%

        corte_row = max(1, cutoff_idx)
        xcat = pareto.iloc[corte_row-1]["Descriptor"]
        ws.write(1,10, xcat); ws.write(2,10, xcat)
        ws.write_number(1,11, 0.0); ws.write_number(2,11, 1.0)

        # gráfico grande
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

        line = wb.add_chart({'type': 'line'})
        line.add_series({
            'name': '% acumulado',
            'categories': [sheet, 1, 1, n, 1],
            'values':     [sheet, 1, 4, n, 4],
            'y2_axis': True,
            'line': {'color': '#ED7D31', 'width': 2.0}
        })
        chart.combine(line)

        h80 = wb.add_chart({'type': 'line'})
        h80.add_series({
            'name': '80/20',
            'categories': [sheet, 1, 1, n, 1],
            'values':     [sheet, 1, 9, n, 9],
            'y2_axis': True,
            'line': {'color': '#7F7F7F', 'width': 1.25}
        })
        chart.combine(h80)

        vline = wb.add_chart({'type': 'line'})
        vline.add_series({
            'name': '',
            'categories': [sheet, 1, 10, 2, 10],  # misma categoría → vertical
            'values':     [sheet, 1, 11, 2, 11],  # 0→1
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
st.title("Pareto Comunidad – MSP (total correcto y categorías del diccionario)")

colA, colB = st.columns([1.3, 1])
with colA:
    plantilla = st.file_uploader("📄 Subí la Plantilla (XLSX) – hoja `matriz`", type=["xlsx"])
with colB:
    dicc = st.file_uploader("🔎 (Opcional) Subí el diccionario de descriptores (XLSX)", type=["xlsx"])

if not plantilla:
    st.info("Subí la Plantilla para procesar.")
    st.stop()

# leer datos
try:
    df_raw = read_matriz(plantilla.getvalue())
except Exception as e:
    st.error(f"Error leyendo 'matriz': {e}")
    st.stop()

df = normalize_columns(df_raw)
st.caption(f"Vista previa (primeras 20 de {len(df)} filas)")
st.dataframe(df.head(20), use_container_width=True)

# construir catálogo desde diccionario (si lo suben)
cat_map = CATEGORIA_POR_DESCRIPTOR_DEFAULT.copy()
if dicc is not None:
    try:
        dic_df = read_diccionario(dicc.getvalue())
        # crear map exacto Descriptor->Categoría desde tu archivo
        cat_map = {row["Descriptor"]: row["Categoría"] for _, row in dic_df.iterrows()}
        st.success("Diccionario de descriptores cargado. Se usará Categoría del archivo.")
    except Exception as e:
        st.warning(f"No se pudo leer el diccionario: {e}. Se usa el catálogo por defecto.")

# detección
def build_regex_all():
    comp = {}
    base_keys = set(cat_map.keys()) | set(SINONIMOS.keys())
    for d in base_keys:
        keys = SINONIMOS.get(d, []) + [d]
        toks = [re.escape(norm_text(k)) for k in keys if norm_text(k)]
        if not toks:
            continue
        comp[d] = re.compile(r"(?:(?<=\s)|^)(" + "|".join(toks) + r")(?:(?=\s)|$)")
    return comp

with st.spinner("Procesando (encabezados + texto)…"):
    regex = build_regex_all()
    # para conteo por encabezados uso claves del catálogo default (para no explotar)
    counts_h = detect_by_headers(df, regex)
    counts_t = detect_in_text(df, regex)
    copilado = build_copilado(counts_h, counts_t)
    pareto   = build_pareto(copilado, cat_map)

if copilado.empty:
    st.warning("No se detectaron descriptores. Revisa sinónimos o comparte un ejemplo con texto.")
    st.stop()

# TOTAL visible (último acumulado)
TOTAL = int(pareto["Acumulado"].iloc[-1])

st.subheader(f"Pareto Comunidad (TOTAL = {TOTAL:,})")
st.dataframe(pareto, use_container_width=True)

# vista rápida
plot_df = pareto.head(TOP_N_GRAFICO).copy()
st.subheader("Gráfico (vista rápida)")
st.bar_chart(plot_df.set_index("Descriptor")["Frecuencia"])
st.line_chart(plot_df.set_index("Descriptor")["% acumulado"])

# descarga excel definitivo
st.subheader("Descargar Excel final")
st.download_button(
    "⬇️ Pareto Comunidad (Excel con formato y gráfico)",
    data=export_excel(pareto, titulo="PARETO COMUNIDAD"),
    file_name="Pareto_Comunidad.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


