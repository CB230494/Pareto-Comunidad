# =========================
# Pareto Comunidad – MSP (1 archivo, sin vueltas)
# =========================
# Flujo automático:
# 1) Subí la Plantilla (XLSX) con hoja 'matriz'.
# 2) La app lee TODAS las filas, mapea y categoriza (encabezados + códigos + texto abierto).
# 3) Muestra Copilado, Pareto, gráfico y botón de descarga (Excel con gráfico).

import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Pareto Comunidad – MSP", layout="wide", initial_sidebar_state="collapsed")
TOP_N_GRAFICO = 40
BLOCK_SIZE = 4000  # para columnas de texto grandes

# -------------------------------------------------------------------
# 1) Catálogo base (Descriptor → Categoría) y sinónimos (texto libre)
# -------------------------------------------------------------------
CATALOGO_BASE = [
    ("HURTO", "DELITOS CONTRA LA PROPIEDAD"),
    ("ROBO", "DELITOS CONTRA LA PROPIEDAD"),
    ("DAÑOS A LA PROPIEDAD", "DELITOS CONTRA LA PROPIEDAD"),
    ("ASALTO", "DELITOS CONTRA LA PROPIEDAD"),
    ("TENTATIVA DE ROBO", "DELITOS CONTRA LA PROPIEDAD"),

    ("VENTA DE DROGAS", "DROGAS"),
    ("TRÁFICO DE DROGAS", "DROGAS"),
    ("MICROTRÁFICO", "DROGAS"),
    ("CONSUMO DE DROGAS", "DROGAS"),
    ("BÚNKER", "DROGAS"),
    ("PUNTO DE VENTA", "DROGAS"),

    ("CONSUMO DE ALCOHOL EN VÍA PÚBLICA", "ALCOHOL"),

    ("HOMICIDIOS", "DELITOS CONTRA LA VIDA"),
    ("HERIDOS", "DELITOS CONTRA LA VIDA"),
    ("TENTATIVA DE HOMICIDIO", "DELITOS CONTRA LA VIDA"),

    ("VIOLENCIA DOMÉSTICA", "VIOLENCIA"),
    ("AGRESIÓN", "VIOLENCIA"),
    ("ABUSO SEXUAL", "VIOLENCIA"),
    ("VIOLACIÓN", "VIOLENCIA"),

    ("ACOSO SEXUAL CALLEJERO", "RIESGO SOCIAL"),
    ("ACOSO ESCOLAR (BULLYING)", "RIESGO SOCIAL"),
    ("ACTOS OBSCENOS EN VIA PUBLICA", "RIESGO SOCIAL"),

    ("PANDILLAS", "ORDEN PÚBLICO"),
    ("INDIGENCIA", "ORDEN PÚBLICO"),
    ("VAGANCIA", "ORDEN PÚBLICO"),
    ("CONTAMINACIÓN SONORA", "ORDEN PÚBLICO"),
    ("CARRERAS ILEGALES", "ORDEN PÚBLICO"),
    ("PORTACIÓN DE ARMA BLANCA", "ORDEN PÚBLICO"),
]

# Sinónimos para texto libre (normalizados: minúsculas, sin tildes)
SINONIMOS: Dict[str, List[str]] = {
    "HURTO": ["hurto", "sustraccion sin violencia"],
    "ROBO": ["robo", "robos", "asalto con violencia", "me robaron con violencia"],
    "DAÑOS A LA PROPIEDAD": ["danos a la propiedad", "vandalismo", "grafiti", "destruccion de propiedad"],
    "ASALTO": ["asalto", "asaltos", "atraco"],
    "TENTATIVA DE ROBO": ["tentativa de robo"],

    "VENTA DE DROGAS": ["venta de droga", "punto de venta", "narcomenudeo", "microtrafico"],
    "TRÁFICO DE DROGAS": ["trafico de drogas", "narco", "trasiego"],
    "MICROTRÁFICO": ["microtrafico", "micro trafico"],
    "CONSUMO DE DROGAS": ["consumo de droga", "fumando crack", "consumo marihuana", "consumiendo drogas"],
    "BÚNKER": ["bunker", "bunquer", "búnker"],
    "PUNTO DE VENTA": ["punto de venta", "puntos de venta"],

    "CONSUMO DE ALCOHOL EN VÍA PÚBLICA": ["consumo de alcohol", "licores en via publica", "tomando licor"],

    "HOMICIDIOS": ["homicidio", "homicidios"],
    "HERIDOS": ["herido", "heridos", "lesionados"],
    "TENTATIVA DE HOMICIDIO": ["tentativa de homicidio"],

    "VIOLENCIA DOMÉSTICA": ["violencia domestica", "violencia intrafamiliar", "maltrato en el hogar"],
    "AGRESIÓN": ["agresion", "agresiones", "pelea", "golpiza"],
    "ABUSO SEXUAL": ["abuso sexual", "tocamientos", "abuso a menor"],
    "VIOLACIÓN": ["violacion", "violada", "violador"],

    "ACOSO SEXUAL CALLEJERO": ["acoso sexual callejero", "acoso en la calle"],
    "ACOSO ESCOLAR (BULLYING)": ["acoso escolar", "bullying"],
    "ACTOS OBSCENOS EN VIA PUBLICA": ["actos obscenos", "exhibicionismo"],

    "PANDILLAS": ["pandillas", "bandas", "mareros"],
    "INDIGENCIA": ["indigencia", "habitantes de calle", "personas en situacion de calle"],
    "VAGANCIA": ["vagancia", "vagos"],
    "CONTAMINACIÓN SONORA": ["ruido", "contaminacion sonora", "musica alta", "bulla"],
    "CARRERAS ILEGALES": ["carreras ilegales", "piques", "piqueras"],
    "PORTACIÓN DE ARMA BLANCA": ["arma blanca", "portacion de cuchillo", "machete"],
}

# -------------------------------------------------------------------
# 2) Mapeo por CÓDIGOS (columna → código → descriptor/categoría)
#     -> Si tu plantilla usa códigos numéricos en columnas, se suman aquí.
#     -> Puedes ampliar/ajustar fácilmente esta lista.
# -------------------------------------------------------------------
# Formato: ("nombre_de_columna_normalizado", codigo, "DESCRIPTOR", "CATEGORÍA")
MAPEO_CODIGOS: List[Tuple[str, object, str, str]] = [
    # ======= ejemplos comunes =======
    # Columnas que codifican problemas/tipos en números (1=..., 2=..., etc.)
    # Usa nombres NORMALIZADOS de columnas (minúsculas, sin tildes, espacios simples).
    # Ajusta/añade si detectas columnas específicas de tu formulario.
    ("hurto", 1, "HURTO", "DELITOS CONTRA LA PROPIEDAD"),
    ("robo", 1, "ROBO", "DELITOS CONTRA LA PROPIEDAD"),
    ("danos a la propiedad", 1, "DAÑOS A LA PROPIEDAD", "DELITOS CONTRA LA PROPIEDAD"),
    ("asalto", 1, "ASALTO", "DELITOS CONTRA LA PROPIEDAD"),

    ("venta de drogas", 1, "VENTA DE DROGAS", "DROGAS"),
    ("trafico de drogas", 1, "TRÁFICO DE DROGAS", "DROGAS"),
    ("microtrafico", 1, "MICROTRÁFICO", "DROGAS"),
    ("consumo de drogas", 1, "CONSUMO DE DROGAS", "DROGAS"),
    ("bunker", 1, "BÚNKER", "DROGAS"),
    ("punto de venta", 1, "PUNTO DE VENTA", "DROGAS"),

    ("consumo de alcohol en via publica", 1, "CONSUMO DE ALCOHOL EN VÍA PÚBLICA", "ALCOHOL"),

    ("homicidios", 1, "HOMICIDIOS", "DELITOS CONTRA LA VIDA"),
    ("heridos", 1, "HERIDOS", "DELITOS CONTRA LA VIDA"),
    ("tentativa de homicidio", 1, "TENTATIVA DE HOMICIDIO", "DELITOS CONTRA LA VIDA"),

    ("violencia domestica", 1, "VIOLENCIA DOMÉSTICA", "VIOLENCIA"),
    ("agresion", 1, "AGRESIÓN", "VIOLENCIA"),
    ("abuso sexual", 1, "ABUSO SEXUAL", "VIOLENCIA"),
    ("violacion", 1, "VIOLACIÓN", "VIOLENCIA"),

    ("acoso sexual callejero", 1, "ACOSO SEXUAL CALLEJERO", "RIESGO SOCIAL"),
    ("acoso escolar", 1, "ACOSO ESCOLAR (BULLYING)", "RIESGO SOCIAL"),
    ("actos obscenos", 1, "ACTOS OBSCENOS EN VIA PUBLICA", "RIESGO SOCIAL"),

    ("pandillas", 1, "PANDILLAS", "ORDEN PÚBLICO"),
    ("indigencia", 1, "INDIGENCIA", "ORDEN PÚBLICO"),
    ("vagancia", 1, "VAGANCIA", "ORDEN PÚBLICO"),
    ("ruido", 1, "CONTAMINACIÓN SONORA", "ORDEN PÚBLICO"),
    ("carreras ilegales", 1, "CARRERAS ILEGALES", "ORDEN PÚBLICO"),
    ("armas blancas", 1, "PORTACIÓN DE ARMA BLANCA", "ORDEN PÚBLICO"),
]

# =========================
# Utilidades de normalización
# =========================
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

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [norm_text(c) for c in out.columns]
    return out

def split_multi(value: str) -> List[str]:
    if value is None:
        return []
    s = str(value)
    if s.strip() == "":
        return []
    parts = re.split(r"[;,/|]+", s)
    return [p.strip() for p in parts if p.strip() != ""]

# =========================
# Lectura (usa TODAS las filas)
# =========================
@st.cache_data(show_spinner=False)
def read_matriz(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(BytesIO(file_bytes), sheet_name="matriz", engine="openpyxl")

# =========================
# Motor de detección
# =========================
def build_cat_map() -> Dict[str, str]:
    return {d: c for d, c in CATALOGO_BASE}

def build_regex_by_desc() -> Dict[str, re.Pattern]:
    compiled: Dict[str, re.Pattern] = {}
    for desc, keys in SINONIMOS.items():
        keys = [re.escape(norm_text(k)) for k in keys if norm_text(k)]
        if not keys:
            continue
        # bordes “suaves” para español
        pat = r"(?:(?<=\s)|^)(" + "|".join(keys) + r")(?:(?=\s)|$)"
        compiled[desc] = re.compile(pat)
    return compiled

def header_marked_series(s: pd.Series) -> pd.Series:
    num = pd.to_numeric(s, errors="coerce").fillna(0) != 0
    txt = s.astype(str).apply(norm_text)
    mask = ~txt.isin(["", "no", "0", "nan", "none", "false"])
    return num | mask

def detect_by_headers(df_raw: pd.DataFrame) -> List[str]:
    df = normalize_columns(df_raw)
    hits: List[str] = []

    # mapeo de encabezados → descriptor (por nombre)
    desc_names = [norm_text(d) for d, _ in CATALOGO_BASE]
    for col in df.columns:
        ncol = norm_text(col)
        # si el encabezado contiene el nombre de un descriptor, cuenta filas marcadas
        for nd, (desc, _) in zip(desc_names, CATALOGO_BASE):
            if (nd == ncol) or (nd in ncol) or (ncol in nd):
                m = header_marked_series(df[col])
                c = int(m.sum())
                if c > 0:
                    hits.extend([desc] * c)
    return hits

def detect_by_codes(df_raw: pd.DataFrame) -> List[str]:
    """Cuenta a partir del MAPEO_CODIGOS (columna → código)"""
    if not MAPEO_CODIGOS:
        return []
    df = normalize_columns(df_raw)
    hits: List[str] = []
    colset = set(df.columns)
    for col_norm, code, desc, _cat in MAPEO_CODIGOS:
        if col_norm not in colset:
            continue
        s = df[col_norm]
        # cada celda puede tener multi-valores
        def match_value(cell) -> bool:
            if pd.isna(cell):
                return False
            # num puro
            try:
                ncell = pd.to_numeric(cell)
                return ncell == code
            except Exception:
                pass
            parts = split_multi(str(cell))
            if parts:
                for p in parts:
                    try:
                        pn = pd.to_numeric(p)
                        if pn == code:
                            return True
                    except Exception:
                        if norm_text(p) == norm_text(str(code)):
                            return True
                return False
            return norm_text(str(cell)) == norm_text(str(code))
        c = int(s.apply(match_value).sum())
        if c > 0:
            hits.extend([desc] * c)
    return hits

def guess_text_columns(df: pd.DataFrame) -> List[str]:
    hints = ["por que", "por qué", "observ", "descr", "coment", "suger", "detalle", "porque", "actividad", "insegur"]
    cols = []
    for col in df.columns:
        if df[col].dtype == object or any(h in norm_text(col) for h in hints):
            sample = df[col].astype(str).head(200).apply(norm_text)
            if (sample != "").mean() > 0.10 or any(h in norm_text(col) for h in hints):
                cols.append(col)
    return cols

def detect_in_text(df_raw: pd.DataFrame) -> List[str]:
    df = normalize_columns(df_raw)
    text_cols = guess_text_columns(df)
    if not text_cols:
        return []
    regex_by_desc = build_regex_by_desc()
    hits: List[str] = []
    for col in text_cols:
        col_norm = df[col].astype(str).apply(norm_text)
        for desc, pat in regex_by_desc.items():
            for i in range(0, len(col_norm), BLOCK_SIZE):
                part = col_norm.iloc[i:i+BLOCK_SIZE]
                c = int(part.str.contains(pat, na=False).sum())
                if c > 0:
                    hits.extend([desc] * c)
    return hits

# =========================
# Agregación y Pareto
# =========================
def make_copilado(hits: List[str]) -> pd.DataFrame:
    if not hits:
        return pd.DataFrame({"Descriptor": [], "Frecuencia": []})
    s = pd.Series(hits, name="Descriptor")
    df = s.value_counts(dropna=False).rename_axis("Descriptor").reset_index(name="Frecuencia")
    return df.sort_values(["Frecuencia","Descriptor"], ascending=[False, True], ignore_index=True)

def make_pareto(copilado_df: pd.DataFrame, cat_map: Dict[str, str]) -> pd.DataFrame:
    if copilado_df.empty:
        return pd.DataFrame(columns=["Categoría","Descriptor","Frecuencia","Porcentaje","% Acumulado","Acumulado","80/20"])
    df = copilado_df.copy()
    df["Categoría"] = df["Descriptor"].map(cat_map).fillna("")
    total = df["Frecuencia"].sum()
    df["Porcentaje"] = (df["Frecuencia"] / total) * 100.0
    df = df.sort_values(["Frecuencia","Descriptor"], ascending=[False, True], ignore_index=True)
    df["% Acumulado"] = df["Porcentaje"].cumsum()
    df["Acumulado"] = df["Frecuencia"].cumsum()
    df["80/20"] = np.where(df["% Acumulado"] <= 80.0, "≤80%", ">80%")
    return df[["Categoría","Descriptor","Frecuencia","Porcentaje","% Acumulado","Acumulado","80/20"]]

# =========================
# Exportación Excel con gráfico
# =========================
def export_excel(copilado: pd.DataFrame, pareto: pd.DataFrame) -> bytes:
    from pandas import ExcelWriter
    import xlsxwriter  # noqa
    out = BytesIO()
    with ExcelWriter(out, engine="xlsxwriter") as writer:
        copilado.to_excel(writer, index=False, sheet_name="Copilado Comunidad")
        pareto.to_excel(writer, index=False, sheet_name="Pareto Comunidad")
        wb  = writer.book
        wsP = writer.sheets["Pareto Comunidad"]
        n = len(pareto)
        if n:
            chart = wb.add_chart({'type': 'column'})
            chart.add_series({
                'name': 'Frecuencia',
                'categories': ['Pareto Comunidad', 1, 1, n, 1],  # B
                'values':     ['Pareto Comunidad', 1, 2, n, 2],  # C
            })
            line = wb.add_chart({'type': 'line'})
            line.add_series({
                'name': '% Acumulado',
                'categories': ['Pareto Comunidad', 1, 1, n, 1],  # B
                'values':     ['Pareto Comunidad', 1, 4, n, 4],  # E
                'y2_axis': True,
            })
            chart.combine(line)
            chart.set_title({'name': 'Pareto Comunidad'})
            chart.set_x_axis({'name': 'Descriptor'})
            chart.set_y_axis({'name': 'Frecuencia'})
            chart.set_y2_axis({'name': '% Acumulado', 'min': 0, 'max': 100})
            wsP.insert_chart(1, 9, chart, {'x_scale': 1.2, 'y_scale': 1.2})
    return out.getvalue()

# =========================
# UI mínima (1 solo archivo)
# =========================
st.title("Pareto Comunidad – MSP (automático)")
archivo = st.file_uploader("📄 Subí la Plantilla (XLSX) – debe tener hoja `matriz`", type=["xlsx"])

if not archivo:
    st.info("Subí la Plantilla para procesar.")
    st.stop()

# Leer TODAS las filas de 'matriz'
try:
    df_matriz = read_matriz(archivo.getvalue())
except Exception as e:
    st.error(f"Error al leer la hoja `matriz`: {e}")
    st.stop()

st.caption(f"Vista previa (primeras 20 de {len(df_matriz)} filas)")
st.dataframe(df_matriz.head(20), use_container_width=True)

# Detección combinada (sin pedir nada)
with st.spinner("Procesando y categorizando (encabezados + códigos + texto abierto)…"):
    cat_map = build_cat_map()
    hits = []
    hits += detect_by_headers(df_matriz)   # encabezados
    hits += detect_by_codes(df_matriz)     # códigos embebidos
    hits += detect_in_text(df_matriz)      # texto abierto
    copilado = make_copilado(hits)
    pareto = make_pareto(copilado, cat_map)

if copilado.empty:
    st.warning("No se detectaron descriptores con el catálogo/mapeo base. Si tu formulario usa otros nombres/códigos, dímelos y los dejo embebidos.")
    st.stop()

st.subheader("Copilado Comunidad")
st.dataframe(copilado, use_container_width=True)

st.subheader("Pareto Comunidad")
st.dataframe(pareto, use_container_width=True)

st.subheader("Gráfico Pareto")
plot_df = pareto.head(TOP_N_GRAFICO).copy()
st.bar_chart(plot_df.set_index("Descriptor")["Frecuencia"])
st.line_chart(plot_df.set_index("Descriptor")["% Acumulado"])

st.subheader("Descargar Excel")
st.download_button(
    "⬇️ Copilado + Pareto + gráfico",
    data=export_excel(copilado, pareto),
    file_name="Pareto_Comunidad.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)




