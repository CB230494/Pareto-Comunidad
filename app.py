# =========================
# Pareto Comunidad – MSP (1 archivo, sin vueltas)
# =========================
# Flujo:
# 1) Subí la Plantilla (XLSX) con hoja 'matriz'.
# 2) La app lee TODAS las filas, normaliza y desduplica encabezados.
# 3) Detecta descriptores por: (a) ENCABEZADOS que contengan el nombre/sinónimos,
#    y (b) TEXTO ABIERTO en columnas tipo texto.
# 4) Muestra Copilado + Pareto + Gráfico y permite descargar Excel con el gráfico.

import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Pareto Comunidad – MSP", layout="wide", initial_sidebar_state="collapsed")
TOP_N_GRAFICO = 50

# ---------------------------
# Utilidades de normalización
# ---------------------------
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
    """
    Hace los nombres de columnas normalizados y únicos:
    si hay duplicados, agrega sufijos __2, __3, ...
    """
    seen: Dict[str, int] = {}
    out: List[str] = []
    for c in cols:
        nc = norm_text(c)
        if nc in seen:
            seen[nc] += 1
            nc2 = f"{nc}__{seen[nc]}"
        else:
            seen[nc] = 1
            nc2 = nc
        out.append(nc2)
    return out

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy()
    df2.columns = make_unique_columns([str(c) for c in df2.columns])
    return df2

# ---------------------------
# Descriptores y categorías
# ---------------------------
# Catálogo principal (Descriptor -> Categoría)
CATEGORIA_POR_DESCRIPTOR: Dict[str, str] = {
    "Disturbios(Riñas)": "ORDEN PÚBLICO",
    "Daños/Vandalismo": "DELITOS CONTRA LA PROPIEDAD",
    "Extorsión": "DELITOS CONTRA LA PROPIEDAD",
    "Hurto": "DELITOS CONTRA LA PROPIEDAD",
    "Receptación": "DELITOS CONTRA LA PROPIEDAD",
    "Contrabando": "DELITOS CONTRA LA PROPIEDAD",
    "Maltrato animal": "ORDEN PÚBLICO",
    "Tráfico ilegal de personas": "DELITOS CONTRA LA VIDA",
    "Venta de drogas": "DROGAS",
    "Homicidios": "DELITOS CONTRA LA VIDA",
    "Lesiones": "DELITOS CONTRA LA VIDA",
    "Delitos sexuales": "VIOLENCIA",
    "Acoso sexual callejero": "RIESGO SOCIAL",
    "Robo a personas": "DELITOS CONTRA LA PROPIEDAD",
    "Robo a comercio (Intimidación)": "DELITOS CONTRA LA PROPIEDAD",
    "Robo a vivienda (Intimidación)": "DELITOS CONTRA LA PROPIEDAD",
    "Robo a transporte público con Intimidación": "DELITOS CONTRA LA PROPIEDAD",
    "Estafas o defraudación": "DELITOS CONTRA LA PROPIEDAD",
    "Estafa informática": "DELITOS CONTRA LA PROPIEDAD",
    "Robo a comercio (Tacha)": "DELITOS CONTRA LA PROPIEDAD",
    "Robo a edificacion (Tacha)": "DELITOS CONTRA LA PROPIEDAD",
    "Robo a vivienda (Tacha)": "DELITOS CONTRA LA PROPIEDAD",
    "Robo a vehículos (Tacha)": "DELITOS CONTRA LA PROPIEDAD",
    "Abigeato (Robo y destace de ganado)": "DELITOS CONTRA LA PROPIEDAD",
    "Robo de bienes agrícola": "DELITOS CONTRA LA PROPIEDAD",
    "Robo de vehículos": "DELITOS CONTRA LA PROPIEDAD",
    "Robo de cable": "DELITOS CONTRA LA PROPIEDAD",
    "Robo de combustible": "DELITOS CONTRA LA PROPIEDAD",
    "Abandono de personas (Menor de edad, adulto mayor o capacidades diferentes)": "RIESGO SOCIAL",
    "Explotacion Sexual infantil": "VIOLENCIA",
    "Explotacion Laboral infantil": "VIOLENCIA",
    "Caza ilegal": "ORDEN PÚBLICO",
    "Pesca ilegal": "ORDEN PÚBLICO",
    "Tala ilegal": "ORDEN PÚBLICO",
    "Trata de personas": "DELITOS CONTRA LA VIDA",
    "Violencia intrafamiliar": "VIOLENCIA",
    "Contaminacion Sonica": "ORDEN PÚBLICO",
    "Falta de oportunidades laborales.": "RIESGO SOCIAL",
    "Problemas Vecinales.": "RIESGO SOCIAL",
    "Usurpacion de terrenos (Precarios)": "ORDEN PÚBLICO",
    "Personas en situación de calle.": "ORDEN PÚBLICO",
    "Desvinculación escolar": "RIESGO SOCIAL",
    "Zona de prostitución": "RIESGO SOCIAL",
    "Consumo de alcohol en vía pública": "ALCOHOL",
    "Personas con exceso de tiempo de ocio": "RIESGO SOCIAL",
    "Falta de salubridad publica": "ORDEN PÚBLICO",
    "Deficiencias en el alumbrado publico": "ORDEN PÚBLICO",
    "Hospedajes ilegales (Cuarterías)": "ORDEN PÚBLICO",
    "Lotes baldíos.": "ORDEN PÚBLICO",
    "Ventas informales (Ambulantes)": "ORDEN PÚBLICO",
    "Pérdida de espacios públicos": "ORDEN PÚBLICO",
    "Falta de inversion social": "RIESGO SOCIAL",
    "Consumo de drogas": "DROGAS",
    "Deficiencia en la infraestructura vial": "ORDEN PÚBLICO",
    "Bunker (Puntos de venta y consumo de drogas)": "DROGAS",
}

# Sinónimos (para texto y para matchear encabezados)
# -> Claves normalizadas (sin tildes, minúsculas)
SINONIMOS: Dict[str, List[str]] = {
    "Disturbios(Riñas)": ["disturbios", "riñas", "riña", "peleas", "riñas callejeras"],
    "Daños/Vandalismo": ["danos", "vandalismo", "grafiti", "daño a la propiedad", "destruccion"],
    "Extorsión": ["extorsion", "cobro de piso", "vacuna"],
    "Hurto": ["hurto", "sustraccion"],
    "Receptación": ["receptacion", "compra de robado", "reduccion"],
    "Contrabando": ["contrabando"],
    "Maltrato animal": ["maltrato animal", "crueldad animal"],
    "Tráfico ilegal de personas": ["trafico de personas", "trata de personas"],
    "Venta de drogas": ["venta de drogas", "punto de venta", "narcomenudeo"],
    "Homicidios": ["homicidio", "homicidios"],
    "Lesiones": ["lesiones", "lesionados", "golpiza"],
    "Delitos sexuales": ["delitos sexuales", "abuso sexual", "violacion", "acoso sexual"],
    "Acoso sexual callejero": ["acoso sexual callejero", "acoso en la calle"],
    "Robo a personas": ["robo a personas", "asalto a persona", "atraco a persona"],
    "Robo a comercio (Intimidación)": ["robo a comercio intimidacion", "asalto a comercio", "intimidacion comercio"],
    "Robo a vivienda (Intimidación)": ["robo a vivienda intimidacion", "asalto a vivienda"],
    "Robo a transporte público con Intimidación": ["robo a transporte publico", "asalto bus"],
    "Estafas o defraudación": ["estafas", "defraudacion", "estafa"],
    "Estafa informática": ["estafa informatica", "phishing"],
    "Robo a comercio (Tacha)": ["robo a comercio tacha", "tacha comercio"],
    "Robo a edificacion (Tacha)": ["robo a edificacion tacha", "tacha edificacion"],
    "Robo a vivienda (Tacha)": ["robo a vivienda tacha", "tacha vivienda"],
    "Robo a vehículos (Tacha)": ["robo a vehiculos tacha", "tacha vehiculos"],
    "Abigeato (Robo y destace de ganado)": ["abigeato", "robo de ganado"],
    "Robo de bienes agrícola": ["robo de bienes agricola", "robo finca agricola"],
    "Robo de vehículos": ["robo de vehiculos", "robo carro", "robo moto"],
    "Robo de cable": ["robo de cable"],
    "Robo de combustible": ["robo de combustible"],
    "Abandono de personas (Menor de edad, adulto mayor o capacidades diferentes)": ["abandono de personas", "abandono de menor", "abandono adulto mayor"],
    "Explotacion Sexual infantil": ["explotacion sexual infantil"],
    "Explotacion Laboral infantil": ["explotacion laboral infantil", "trabajo infantil"],
    "Caza ilegal": ["caza ilegal"],
    "Pesca ilegal": ["pesca ilegal"],
    "Tala ilegal": ["tala ilegal"],
    "Trata de personas": ["trata de personas"],
    "Violencia intrafamiliar": ["violencia intrafamiliar", "violencia domestica"],
    "Contaminacion Sonica": ["contaminacion sonora", "ruido", "musica alta", "bulla"],
    "Falta de oportunidades laborales.": ["falta de oportunidades laborales", "desempleo"],
    "Problemas Vecinales.": ["problemas vecinales", "conflictos vecinales"],
    "Usurpacion de terrenos (Precarios)": ["usurpacion de terrenos", "precarios"],
    "Personas en situación de calle.": ["personas en situacion de calle", "indigencia", "habitantes de calle"],
    "Desvinculación escolar": ["desvinculacion escolar", "abandono escolar"],
    "Zona de prostitución": ["zona de prostitucion", "prostitucion"],
    "Consumo de alcohol en vía pública": ["consumo de alcohol en via publica", "licores en via publica"],
    "Personas con exceso de tiempo de ocio": ["exceso de tiempo de ocio", "ocio juvenil"],
    "Falta de salubridad publica": ["falta de salubridad publica", "insalubridad"],
    "Deficiencias en el alumbrado publico": ["deficiencias en el alumbrado publico", "alumbrado deficiente", "falta de alumbrado"],
    "Hospedajes ilegales (Cuarterías)": ["hospedajes ilegales", "cuarterias"],
    "Lotes baldíos.": ["lotes baldios", "lote baldío"],
    "Ventas informales (Ambulantes)": ["ventas informales", "ambulantes"],
    "Pérdida de espacios públicos": ["perdida de espacios publicos"],
    "Falta de inversion social": ["falta de inversion social"],
    "Consumo de drogas": ["consumo de drogas", "consumen drogas"],
    "Deficiencia en la infraestructura vial": ["deficiencia en la infraestructura vial", "huecos", "baches"],
    "Bunker (Puntos de venta y consumo de drogas)": ["bunker", "bunquer", "búnker", "punto de venta y consumo"],
}

# ---------------------------
# Lectura (todas las filas) y protección de duplicados
# ---------------------------
@st.cache_data(show_spinner=False)
def read_matriz(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes), sheet_name="matriz", engine="openpyxl")
    return df

# ---------------------------
# Detección
# ---------------------------
def build_regex_by_desc() -> Dict[str, re.Pattern]:
    compiled: Dict[str, re.Pattern] = {}
    for desc, keys in SINONIMOS.items():
        tokens = [re.escape(norm_text(k)) for k in keys if norm_text(k)]
        if not tokens:
            continue
        # bordes "suaves" para español
        pat = r"(?:(?<=\s)|^)(" + "|".join(tokens) + r")(?:(?=\s)|$)"
        compiled[desc] = re.compile(pat)
    return compiled

def header_marked_series(s: pd.Series) -> pd.Series:
    # Verdadero si la celda es numérica != 0 o texto no vacío distinto de ("no","0",...)
    num = pd.to_numeric(s, errors="coerce").fillna(0) != 0
    txt = s.astype(str).apply(norm_text)
    mask = ~txt.isin(["", "no", "0", "nan", "none", "false"])
    return num | mask

def detect_by_headers(df_norm: pd.DataFrame, regex_by_desc: Dict[str, re.Pattern]) -> Dict[str, int]:
    """
    Para cada descriptor: si el NOMBRE de alguna columna contiene el descriptor o sus sinónimos,
    cuenta filas "marcadas" en esas columnas (OR por fila para no sobrecontar).
    """
    counts: Dict[str, int] = {d: 0 for d in CATEGORIA_POR_DESCRIPTOR.keys()}
    # pre-normalizamos encabezados
    cols = list(df_norm.columns)
    for desc, pat in regex_by_desc.items():
        # columnas cuyo nombre matchea el descriptor/sinónimos
        hit_cols = [c for c in cols if re.search(pat, " " + c + " ") is not None]
        if not hit_cols:
            continue
        # OR entre todas esas columnas
        mask_any = None
        for c in hit_cols:
            try:
                s = df_norm[c]
                # si por duplicado de encabezado c agrupa varias, ya lo resolvimos haciendo columnas únicas
                m = header_marked_series(s)
                mask_any = m if mask_any is None else (mask_any | m)
            except Exception:
                continue
        if mask_any is not None:
            counts[desc] += int(mask_any.sum())
    return counts

def guess_text_columns(df_norm: pd.DataFrame) -> List[str]:
    """
    Heurística segura para columnas de texto:
    - dtype object o muchos strings
    - nombres que sugieren campo abierto (observación, descripción, comentario, por qué, problema, etc.)
    """
    hints = ["observ", "descr", "coment", "suger", "porque", "por que", "por qué", "detalle", "problema", "actividad", "insegur"]
    text_cols: List[str] = []
    for col in df_norm.columns:
        s = df_norm[col]
        # si por duplicado vino como DataFrame (no debería por make_unique), lo saltamos
        if not hasattr(s, "dtype"):
            continue
        looks_text = (getattr(s, "dtype", None) == object) or any(h in col for h in hints)
        if looks_text:
            # si hay suficientes no vacíos
            sample = s.astype(str).head(200).apply(norm_text)
            if (sample != "").mean() > 0.05 or any(h in col for h in hints):
                text_cols.append(col)
    return text_cols

def detect_in_text(df_norm: pd.DataFrame, regex_by_desc: Dict[str, re.Pattern]) -> Dict[str, int]:
    """
    Para cada descriptor: OR de coincidencias en TODAS las columnas de texto (1 por fila máx).
    """
    counts: Dict[str, int] = {d: 0 for d in CATEGORIA_POR_DESCRIPTOR.keys()}
    text_cols = guess_text_columns(df_norm)
    if not text_cols:
        return counts
    for desc, pat in regex_by_desc.items():
        mask_any = None
        for c in text_cols:
            s = df_norm[c].astype(str).apply(norm_text)
            m = s.str.contains(pat, na=False)
            mask_any = m if mask_any is None else (mask_any | m)
        if mask_any is not None:
            counts[desc] += int(mask_any.sum())
    return counts

# ---------------------------
# Copilado, Pareto y Excel
# ---------------------------
def build_copilado(counts_header: Dict[str, int], counts_text: Dict[str, int]) -> pd.DataFrame:
    # sumamos conteos de encabezados y de texto
    total_counts: Dict[str, int] = {}
    for d in CATEGORIA_POR_DESCRIPTOR.keys():
        total_counts[d] = counts_header.get(d, 0) + counts_text.get(d, 0)
    # quitamos ceros
    rows = [(d, f) for d, f in total_counts.items() if f > 0]
    if not rows:
        return pd.DataFrame({"Descriptor": [], "Frecuencia": []})
    df = pd.DataFrame(rows, columns=["Descriptor", "Frecuencia"])
    return df.sort_values(["Frecuencia", "Descriptor"], ascending=[False, True], ignore_index=True)

def build_pareto(copilado: pd.DataFrame) -> pd.DataFrame:
    if copilado.empty:
        return pd.DataFrame(columns=["Categoría","Descriptor","Frecuencia","Porcentaje","% Acumulado","Acumulado","80/20"])
    df = copilado.copy()
    df["Categoría"] = df["Descriptor"].map(CATEGORIA_POR_DESCRIPTOR).fillna("")
    total = df["Frecuencia"].sum()
    df["Porcentaje"] = (df["Frecuencia"] / total) * 100.0
    df = df.sort_values(["Frecuencia","Descriptor"], ascending=[False, True], ignore_index=True)
    df["% Acumulado"] = df["Porcentaje"].cumsum()
    df["Acumulado"] = df["Frecuencia"].cumsum()
    df["80/20"] = np.where(df["% Acumulado"] <= 80.0, "≤80%", ">80%")
    return df[["Categoría","Descriptor","Frecuencia","Porcentaje","% Acumulado","Acumulado","80/20"]]

def export_excel(copilado: pd.DataFrame, pareto: pd.DataFrame) -> bytes:
    from pandas import ExcelWriter
    import xlsxwriter  # noqa: F401
    output = BytesIO()
    with ExcelWriter(output, engine="xlsxwriter") as writer:
        copilado.to_excel(writer, index=False, sheet_name="Copilado Comunidad")
        pareto.to_excel(writer, index=False, sheet_name="Pareto Comunidad")
        # gráfico
        wb = writer.book
        ws = writer.sheets["Pareto Comunidad"]
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
            ws.insert_chart(1, 9, chart, {'x_scale': 1.2, 'y_scale': 1.2})
    return output.getvalue()

# ---------------------------
# UI mínima
# ---------------------------
st.title("Pareto Comunidad – MSP (automático, 1 archivo)")
archivo = st.file_uploader("📄 Subí la Plantilla (XLSX) – debe tener la hoja `matriz`", type=["xlsx"])

if not archivo:
    st.info("Subí la Plantilla para procesar.")
    st.stop()

# Lee TODAS las filas (571+), normaliza y desduplica encabezados
try:
    df_raw = read_matriz(archivo.getvalue())
except Exception as e:
    st.error(f"Error leyendo hoja `matriz`: {e}")
    st.stop()

df = normalize_columns(df_raw)

st.caption(f"Vista previa (primeras 20 de {len(df)} filas) – columnas normalizadas y únicas")
st.dataframe(df.head(20), use_container_width=True)

with st.spinner("Procesando descriptores (encabezados + texto abierto)…"):
    regex_by_desc = {desc: re.compile(r"(?:(?<=\s)|^)" + "|".join([re.escape(norm_text(k)) for k in SINONIMOS.get(desc, [desc]) if norm_text(k)]) + r"(?:(?=\s)|$)")
                     for desc in CATEGORIA_POR_DESCRIPTOR.keys()}
    # Conteo por encabezados
    counts_headers = detect_by_headers(df, regex_by_desc)
    # Conteo por texto
    counts_text = detect_in_text(df, regex_by_desc)

    copilado = build_copilado(counts_headers, counts_text)
    pareto = build_pareto(copilado)

if copilado.empty:
    st.warning("No se detectaron descriptores. Revisa que las columnas/textos contengan los nombres o sinónimos habituales.")
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




