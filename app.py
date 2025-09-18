# app.py — Pareto Comunidad (DELITO / RIESGO SOCIAL / OTROS FACTORES)
# - Detecta solo en columnas AI:ET
# - % en pantalla como "4,66%" y en Excel eje derecho 0–100%

import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from difflib import get_close_matches

st.set_page_config(page_title="Pareto Comunidad – MSP", layout="wide", initial_sidebar_state="collapsed")
TOP_N_GRAFICO = 60

# ===================== CATEGORÍAS PERMITIDAS =====================
CATEGORIAS_VALIDAS = {"DELITO", "RIESGO SOCIAL", "OTROS FACTORES"}

def _force_cat(x: str) -> str:
    n = (x or "").strip().upper()
    if n in CATEGORIAS_VALIDAS:
        return n
    return "OTROS FACTORES"

# ===================== Diccionario embebido (ejemplo; amplialo libremente) ======================
DICCIONARIO_EMBEBIDO = pd.DataFrame([
    # -------- DELITO --------
    ["Consumo de drogas", "DELITO"],
    ["Venta de drogas", "DELITO"],
    ["Bunker (Puntos de venta y consumo de drogas)", "DELITO"],
    ["Hurto", "DELITO"],
    ["Robo a personas", "DELITO"],
    ["Robo a vivienda (Tacha)", "DELITO"],
    ["Robo a vivienda (Intimidación)", "DELITO"],
    ["Robo a vehículos (Tacha)", "DELITO"],
    ["Robo a vehiculos", "DELITO"],
    ["Robo a comercio (Intimidación)", "DELITO"],
    ["Robo a comercio (Tacha)", "DELITO"],
    ["Robo de vehículos", "DELITO"],
    ["Receptación", "DELITO"],
    ["Estafas o defraudación", "DELITO"],
    ["Daños/Vandalismo", "DELITO"],
    ["Lesiones", "DELITO"],
    # ---- RIESGO SOCIAL ----
    ["Falta de oportunidades laborales.", "RIESGO SOCIAL"],
    ["Falta de inversion social", "RIESGO SOCIAL"],
    ["Personas con exceso de tiempo de ocio", "RIESGO SOCIAL"],
    ["Problemas Vecinales.", "RIESGO SOCIAL"],
    # ---- OTROS FACTORES ----
    ["Consumo de alcohol en vía pública", "OTROS FACTORES"],
    ["Contaminacion Sonica", "OTROS FACTORES"],
    ["Deficiencia en la infraestructura vial", "OTROS FACTORES"],
    ["Lotes baldíos.", "OTROS FACTORES"],
    ["Falta de salubridad publica", "OTROS FACTORES"],
    ["Disturbios(Riñas)", "OTROS FACTORES"],
    ["Personas en situación de calle.", "OTROS FACTORES"],
], columns=["Descriptor", "Categoría"]).assign(Categoría=lambda d: d["Categoría"].map(_force_cat))

# ===================== Sinónimos / variaciones =========================
SINONIMOS: Dict[str, List[str]] = {
    # DELITO
    "Consumo de drogas": ["consumo de drogas", "consumen drogas", "consumo marihuana", "fumando piedra"],
    "Venta de drogas": ["venta de drogas", "punto de venta", "narcomenudeo"],
    "Hurto": ["hurto", "sustraccion"],
    "Robo a personas": ["robo a personas", "asalto a persona", "atraco a persona"],
    "Robo a vivienda (Tacha)": ["tacha vivienda", "robo vivienda tacha"],
    "Robo a vivienda (Intimidación)": ["asalto a vivienda", "intimidacion vivienda"],
    "Robo a vehículos (Tacha)": ["tacha vehiculos", "robo tacha vehiculo"],
    "Robo a vehiculos": ["robo de vehiculos", "robo carro", "robo moto"],
    "Robo a comercio (Intimidación)": ["asalto a comercio"],
    "Robo a comercio (Tacha)": ["robo comercio tacha"],
    "Daños/Vandalismo": ["daños", "vandalismo", "grafiti", "daño a la propiedad"],
    "Receptación": ["receptacion", "compra de robado", "reduccion"],
    "Estafas o defraudación": ["estafas", "defraudacion", "estafa"],
    "Robo de vehículos": ["robo de vehiculos"],
    "Lesiones": ["lesiones", "golpiza"],
    # RIESGO SOCIAL
    "Falta de oportunidades laborales.": ["desempleo", "falta de empleo"],
    "Falta de inversion social": ["falta de inversión social"],
    "Personas con exceso de tiempo de ocio": ["ocio juvenil", "exceso de ocio"],
    "Problemas Vecinales.": ["conflictos vecinales", "problemas vecinales"],
    # OTROS FACTORES
    "Consumo de alcohol en vía pública": ["consumo de alcohol en via publica", "licores en via publica"],
    "Contaminacion Sonica": ["contaminacion sonora", "ruido", "musica alta", "bulla"],
    "Deficiencia en la infraestructura vial": ["infraestructura vial", "huecos", "baches"],
    "Lotes baldíos.": ["lote baldio", "lotes baldios"],
    "Falta de salubridad publica": ["insalubridad"],
    "Disturbios(Riñas)": ["disturbios", "riñas", "riña", "peleas"],
    "Personas en situación de calle.": ["situacion de calle", "indigentes", "habitantes de calle"],
}

# ===================== Utilidades =====================
import re as _re
def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def norm_text(s: Optional[str]) -> str:
    if s is None:
        return ""
    s = strip_accents(str(s)).lower().strip()
    s = _re.sub(r"\s+", " ", s)
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

# --- helpers para rango de columnas Excel (AI:ET) ---
def excel_col_to_index(col: str) -> int:
    """Convierte 'A'->1, 'Z'->26, 'AA'->27, ..., 'AI'->35, 'ET'->150 (1-based)."""
    col = col.strip().upper()
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

AI_IDX_1B = excel_col_to_index("AI")   # 35 (1-based)
ET_IDX_1B = excel_col_to_index("ET")   # 150 (1-based)
AI_IDX = AI_IDX_1B - 1                 # 0-based para iloc
ET_IDX = ET_IDX_1B - 1

def build_regex_all(dic_df: pd.DataFrame) -> Dict[str, re.Pattern]:
    compiled = {}
    base_keys = list(dic_df["Descriptor"].astype(str)) + list(SINONIMOS.keys())
    base_keys = sorted(set(base_keys))
    for d in base_keys:
        keys = SINONIMOS.get(d, []) + [d]
        toks = [re.escape(norm_text(k)) for k in keys if norm_text(k)]
        if toks:
            compiled[d] = re.compile(r"(?:(?<=\s)|^)(" + "|".join(toks) + r")(?:(?=\s)|$)")
    return compiled

def header_marked_series(s: pd.Series) -> pd.Series:
    num = pd.to_numeric(s, errors="coerce").fillna(0) != 0
    txt = s.astype(str).apply(norm_text)
    mask = ~txt.isin(["", "no", "0", "nan", "none", "false"])
    return num | mask

def detect_by_headers(df_norm: pd.DataFrame, regex_by_desc: Dict[str, re.Pattern]) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for desc, pat in regex_by_desc.items():
        hit_cols = [c for c in df_norm.columns if re.search(pat, " " + c + " ") is not None]
        if not hit_cols:
            continue
        mask_any = None
        for c in hit_cols:
            m = header_marked_series(df_norm[c])
            mask_any = m if mask_any is None else (mask_any | m)
        if mask_any is not None:
            counts[desc] = counts.get(desc, 0) + int(mask_any.sum())
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
    counts: Dict[str, int] = {}
    tcols = guess_text_cols(df_norm)
    if not tcols:
        return counts
    for desc, pat in regex_by_desc.items():
        mask_any = None
        for c in tcols:
            m = df_norm[c].astype(str).apply(norm_text).str.contains(pat, na=False)
            mask_any = m if mask_any is None else (mask_any | m)
        if mask_any is not None:
            counts[desc] = counts.get(desc, 0) + int(mask_any.sum())
    return counts

def build_copilado_from_counts(counts_headers: Dict[str, int], counts_text: Dict[str, int]) -> pd.DataFrame:
    total: Dict[str, int] = {}
    keys = set(counts_headers) | set(counts_text)
    for d in keys:
        total[d] = counts_headers.get(d, 0) + counts_text.get(d, 0)
    rows = [(d, f) for d, f in total.items() if f > 0]
    if not rows:
        return pd.DataFrame({"Descriptor": [], "Frecuencia": []})
    df = pd.DataFrame(rows, columns=["Descriptor", "Frecuencia"])
    return df.sort_values(["Frecuencia", "Descriptor"], ascending=[False, True], ignore_index=True)

# --- Canonización (fuzzy) a descriptor del catálogo ---
def build_canon_maps(dic_df: pd.DataFrame) -> Tuple[Dict[str,str], Dict[str,str]]:
    desc_by_norm = {}
    cat_by_norm  = {}
    for _, r in dic_df.iterrows():
        d = str(r["Descriptor"]).strip()
        c = _force_cat(str(r["Categoría"]))
        desc_by_norm[norm_text(d)] = d
        cat_by_norm[norm_text(d)]  = c
    return desc_by_norm, cat_by_norm

def canoniza(raw_desc: str, desc_by_norm: Dict[str,str]) -> str:
    n = norm_text(raw_desc)
    if n in desc_by_norm:
        return desc_by_norm[n]
    candidates = list(desc_by_norm.keys())
    hit = get_close_matches(n, candidates, n=1, cutoff=0.82)
    if hit:
        return desc_by_norm[hit[0]]
    return raw_desc.strip()

def build_pareto(copilado: pd.DataFrame, dic_df: pd.DataFrame) -> pd.DataFrame:
    if copilado.empty:
        return pd.DataFrame(columns=["Categoría","Descriptor","Frecuencia","Porcentaje","% acumulado","Acumulado","80/20"])

    desc_by_norm, cat_by_norm = build_canon_maps(dic_df)

    df = copilado.copy()
    df["Descriptor"] = df["Descriptor"].astype(str)
    df["Descriptor Canon"] = df["Descriptor"].apply(lambda x: canoniza(x, desc_by_norm))
    df["Categoría"] = df["Descriptor Canon"].apply(lambda d: _force_cat(cat_by_norm.get(norm_text(d), "OTROS FACTORES")))

    grp = df.groupby(["Categoría", "Descriptor Canon"], as_index=False)["Frecuencia"].sum()
    grp = grp.rename(columns={"Descriptor Canon":"Descriptor"})

    # TOTAL y % (fracciones 0–1)
    total = int(grp["Frecuencia"].sum())
    grp = grp.sort_values(["Frecuencia","Descriptor"], ascending=[False, True], ignore_index=True)
    grp["Porcentaje"]  = (grp["Frecuencia"] / total)     # fracción 0–1
    grp["% acumulado"] = grp["Porcentaje"].cumsum()      # fracción 0–1
    grp["Acumulado"]   = grp["Frecuencia"].cumsum()
    grp["80/20"]       = "80%"

    # Garantías
    assert int(grp["Acumulado"].iloc[-1]) == total
    assert abs(float(grp["% acumulado"].iloc[-1]) - 1.0) < 1e-9

    return grp[["Categoría","Descriptor","Frecuencia","Porcentaje","% acumulado","Acumulado","80/20"]]

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

        # columnas auxiliares
        ws.write(0, 9, "80/20");  ws.set_column("J:J", 6,  None, {"hidden": True})
        ws.write(0,10, "CorteX"); ws.set_column("K:K", 20, None, {"hidden": True})
        ws.write(0,11, "%");      ws.set_column("L:L", 6,  None, {"hidden": True})
        for i in range(n):
            ws.write_number(i+1, 9, 0.80)

        corte_row = max(1, cutoff_idx)
        xcat = pareto.iloc[corte_row-1]["Descriptor"]
        ws.write(1,10, xcat); ws.write(2,10, xcat)
        ws.write_number(1,11, 0.0); ws.write_number(2,11, 1.0)

        # Barras: Frecuencia
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

        # Línea: % acumulado (eje secundario, 0–100%)
        line = wb.add_chart({'type': 'line'})
        line.add_series({
            'name': '% acumulado',
            'categories': [sheet, 1, 1, n, 1],
            'values':     [sheet, 1, 4, n, 4],
            'y2_axis': True,
            'line': {'color': '#ED7D31', 'width': 2.0}
        })
        chart.combine(line)

        # Línea horizontal 80% (eje secundario)
        h80 = wb.add_chart({'type': 'line'})
        h80.add_series({
            'name': '80/20',
            'categories': [sheet, 1, 1, n, 1],
            'values':     [sheet, 1, 9, n, 9],
            'y2_axis': True,
            'line': {'color': '#7F7F7F', 'width': 1.25}
        })
        chart.combine(h80)

        # Línea vertical en el corte (eje secundario)
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
        chart.set_y2_axis({'min': 0, 'max': 1, 'major_unit': 0.1, 'num_format': '0%'})  # 0–100%
        chart.set_legend({'position': 'bottom'})
        ws.insert_chart(1, 9, chart, {'x_scale': 1.9, 'y_scale': 1.6})

    return out.getvalue()

# ===================== UI =====================
st.title("Pareto Comunidad – MSP (DELITO / RIESGO SOCIAL / OTROS FACTORES)")

plantilla = st.file_uploader("📄 Subí la Plantilla (XLSX) – hoja `matriz`", type=["xlsx"])
if not plantilla:
    st.info("Subí la Plantilla para procesar.")
    st.stop()

try:
    df_raw = read_matriz(plantilla.getvalue())
except Exception as e:
    st.error(f"Error leyendo 'matriz': {e}")
    st.stop()

df_all = normalize_columns(df_raw)
st.caption(f"Vista previa (primeras 20 de {len(df_all)} filas)")
st.dataframe(df_all.head(20), use_container_width=True)

# ---------- Tomar SOLO columnas AI:ET para la detección ----------
# (índices 0-based: AI=34, ET=149)
def excel_col_to_index(col: str) -> int:
    col = col.strip().upper()
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n

AI_IDX = excel_col_to_index("AI") - 1   # 35 -> 34 (0-based)
ET_IDX = excel_col_to_index("ET") - 1   # 150 -> 149 (0-based)
ncols = df_all.shape[1]
L = max(0, min(AI_IDX, ncols-1))
R = max(0, min(ET_IDX, ncols-1))
if L > R:
    L, R = 0, ncols-1  # fallback si la hoja tiene menos columnas

df_range = df_all.iloc[:, L:R+1].copy()
df_range = normalize_columns(df_range)   # normalizo nombres dentro del rango

# --- Si la hoja trae DESCRIPTOR + FRECUENCIA (en cualquier parte), usamos tal cual ---
cols_norm_all = {c: norm_text(c) for c in df_all.columns}
inv_all = {v: k for k, v in cols_norm_all.items()}

desc_candidates = [col for col, n in cols_norm_all.items()
                   if any(t in n for t in ["descriptor", "problema", "descriptor actualizado", "descripcion"])]

freq_col = next((inv_all[x] for x in ["frecuencia"] if x in inv_all), None)

if freq_col and desc_candidates:
    base = df_all[[desc_candidates[0], freq_col]].copy()
    base.columns = ["Descriptor", "Frecuencia"]
    base["Frecuencia"] = pd.to_numeric(base["Frecuencia"], errors="coerce").fillna(0).astype(int)
    base = base.groupby("Descriptor", as_index=False)["Frecuencia"].sum()
else:
    # Fallback: detectar SOLO usando columnas AI:ET
    regex = build_regex_all(DICCIONARIO_EMBEBIDO)
    counts_h = detect_by_headers(df_range, regex)
    counts_t = detect_in_text(df_range, regex)
    base = build_copilado_from_counts(counts_h, counts_t)

if base.empty or base["Frecuencia"].sum() == 0:
    st.warning("No se detectaron descriptores o las frecuencias son 0. Revisá el rango AI:ET y la plantilla.")
    st.stop()

pareto = build_pareto(base, DICCIONARIO_EMBEBIDO)

# ---------- FORMATO DE PORCENTAJES EN PANTALLA (4,66%) ----------
def pct_str(frac: float) -> str:
    s = f"{frac*100:.2f}%"
    return s.replace(".", ",")

display = pareto.copy()
display["Porcentaje"] = display["Porcentaje"].apply(pct_str)
display["% acumulado"] = display["% acumulado"].apply(pct_str)

TOTAL = int(pareto["Acumulado"].iloc[-1])
st.subheader(f"Pareto Comunidad (TOTAL = {TOTAL:,})")
st.dataframe(display, use_container_width=True)

# ======== Gráfico rápido (barras + % acumulado + 80%) ========
import altair as alt
top_df = pareto.head(TOP_N_GRAFICO).copy()
bars = alt.Chart(top_df).mark_bar().encode(
    x=alt.X('Descriptor:N', sort=None, axis=alt.Axis(labelAngle=-50)),
    y=alt.Y('Frecuencia:Q')
)
line = alt.Chart(top_df).mark_line(point=True).encode(
    x='Descriptor:N',
    y=alt.Y('% acumulado:Q', axis=alt.Axis(format='%'), scale=alt.Scale(domain=[0,1])),  # 0–100%
    color=alt.value('#ED7D31')
)
h80 = alt.Chart(pd.DataFrame({'y':[0.8]})).mark_rule().encode(
    y=alt.Y('y:Q', axis=alt.Axis(format='%'))
)
st.altair_chart((bars + line + h80).resolve_scale(y='independent'), use_container_width=True)

# ========= Descargar Excel definitivo =========
st.subheader("Descargar Excel final")
st.download_button(
    "⬇️ Pareto Comunidad (Excel con formato y gráfico)",
    data=export_excel(pareto, titulo="PARETO COMUNIDAD"),
    file_name="Pareto_Comunidad.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)





