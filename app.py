# Pareto Comunidad – MSP (1 archivo, formato idéntico a tu ejemplo)

import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Pareto Comunidad – MSP", layout="wide", initial_sidebar_state="collapsed")
TOP_N_GRAFICO = 60  # solo para la vista rápida en la web

# ---------- Normalización ----------
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
    df2 = df.copy()
    df2.columns = make_unique_columns([str(c) for c in df2.columns])
    return df2

@st.cache_data(show_spinner=False)
def read_matriz(file_bytes: bytes) -> pd.DataFrame:
    # Usa TODAS las filas
    return pd.read_excel(BytesIO(file_bytes), sheet_name="matriz", engine="openpyxl")

# ---------- Catálogo + sinónimos (ajústalo si ocupas) ----------
CATEGORIA_POR_DESCRIPTOR: Dict[str, str] = {
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
    "Usurpacion de terrenos (Precarios)": "ORDEN PÚBLICO",
    "Pérdida de espacios públicos": "ORDEN PÚBLICO",
    "Deficiencias en el alumbrado publico": "ORDEN PÚBLICO",
    "Violencia intrafamiliar": "VIOLENCIA",
    "Robo de cable": "DELITOS CONTRA LA PROPIEDAD",
    "Acoso sexual callejero": "RIESGO SOCIAL",
    "Hospedajes ilegales (Cuarterías)": "ORDEN PÚBLICO",
    "Desvinculación escolar": "RIESGO SOCIAL",
    "Robo a transporte público con Intimidación": "DELITOS CONTRA LA PROPIEDAD",
    "Ventas informales (Ambulantes)": "ORDEN PÚBLICO",
    "Maltrato animal": "ORDEN PÚBLICO",
    "Extorsión": "DELITOS CONTRA LA PROPIEDAD",
    "Homicidios": "DELITOS CONTRA LA VIDA",
    "Abandono de personas (Menor de edad, adulto mayor o capacidades diferentes)": "RIESGO SOCIAL",
    "Robo a edificacion (Tacha)": "DELITOS CONTRA LA PROPIEDAD",
    "Estafa informática": "DELITOS CONTRA LA PROPIEDAD",
    "Delitos sexuales": "VIOLENCIA",
    "Robo de bienes agrícola": "DELITOS CONTRA LA PROPIEDAD",
    "Robo de combustible": "DELITOS CONTRA LA PROPIEDAD",
    "Tala ilegal": "ORDEN PÚBLICO",
    "Trata de personas": "DELITOS CONTRA LA VIDA",
    "Explotacion Laboral infantil": "VIOLENCIA",
    "Caza ilegal": "ORDEN PÚBLICO",
    "Abigeato (Robo y destace de ganado)": "DELITOS CONTRA LA PROPIEDAD",
    "Zona de prostitución": "RIESGO SOCIAL",
    "Explotacion Sexual infantil": "VIOLENCIA",
    "Pesca ilegal": "ORDEN PÚBLICO",
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
    "Usurpacion de terrenos (Precarios)": ["usurpacion de terrenos", "precarios"],
    "Pérdida de espacios públicos": ["perdida de espacios publicos"],
    "Deficiencias en el alumbrado publico": ["deficiencias en el alumbrado publico", "falta de alumbrado"],
    "Violencia intrafamiliar": ["violencia intrafamiliar", "violencia domestica"],
    "Robo de cable": ["robo de cable"],
    "Acoso sexual callejero": ["acoso sexual callejero", "acoso en la calle"],
    "Hospedajes ilegales (Cuarterías)": ["hospedajes ilegales", "cuarterias"],
    "Desvinculación escolar": ["desvinculacion escolar", "abandono escolar"],
    "Robo a transporte público con Intimidación": ["robo a transporte publico", "asalto bus"],
    "Ventas informales (Ambulantes)": ["ventas informales", "ambulantes"],
    "Maltrato animal": ["maltrato animal", "crueldad animal"],
    "Extorsión": ["extorsion", "cobro de piso", "vacuna"],
    "Homicidios": ["homicidio", "homicidios"],
    "Abandono de personas (Menor de edad, adulto mayor o capacidades diferentes)": ["abandono de personas", "abandono de menor", "abandono adulto mayor"],
    "Robo a edificacion (Tacha)": ["robo a edificacion tacha"],
    "Estafa informática": ["estafa informatica", "phishing"],
    "Delitos sexuales": ["delitos sexuales", "abuso sexual", "violacion", "acoso sexual"],
    "Robo de bienes agrícola": ["robo de bienes agricola", "robo finca agricola"],
    "Robo de combustible": ["robo de combustible"],
    "Tala ilegal": ["tala ilegal"],
    "Trata de personas": ["trata de personas"],
    "Explotacion Laboral infantil": ["explotacion laboral infantil", "trabajo infantil"],
    "Caza ilegal": ["caza ilegal"],
    "Abigeato (Robo y destace de ganado)": ["abigeato", "robo de ganado"],
    "Zona de prostitución": ["zona de prostitucion", "prostitucion"],
    "Explotacion Sexual infantil": ["explotacion sexual infantil"],
    "Pesca ilegal": ["pesca ilegal"],
}

# ---------- Detección ----------
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
    counts = {d: 0 for d in CATEGORIA_POR_DESCRIPTOR.keys()}
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
    counts = {d: 0 for d in CATEGORIA_POR_DESCRIPTOR.keys()}
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
    for d in CATEGORIA_POR_DESCRIPTOR:
        total[d] = counts_headers.get(d, 0) + counts_text.get(d, 0)
    rows = [(d, f) for d, f in total.items() if f > 0]
    if not rows:
        return pd.DataFrame({"Descriptor": [], "Frecuencia": []})
    df = pd.DataFrame(rows, columns=["Descriptor", "Frecuencia"])
    return df.sort_values(["Frecuencia", "Descriptor"], ascending=[False, True], ignore_index=True)

def build_pareto(copilado: pd.DataFrame) -> pd.DataFrame:
    if copilado.empty:
        return pd.DataFrame(columns=["Categoría","Descriptor","Frecuencia","Porcentaje","% acumulado","Acumulado","80/20"])
    df = copilado.copy()
    df["Categoría"] = df["Descriptor"].map(CATEGORIA_POR_DESCRIPTOR).fillna("")
    total = df["Frecuencia"].sum()
    df["Porcentaje"] = (df["Frecuencia"] / total) * 100.0
    df = df.sort_values(["Frecuencia","Descriptor"], ascending=[False, True], ignore_index=True)
    df["% acumulado"] = df["Porcentaje"].cumsum()
    df["Acumulado"]   = df["Frecuencia"].cumsum()
    df["80/20"]       = np.where(df["% acumulado"] <= 80.0, "80%", "80%")
    return df[["Categoría","Descriptor","Frecuencia","Porcentaje","% acumulado","Acumulado","80/20"]]

# ---------- Export Excel (estilo idéntico) ----------
def export_excel(pareto: pd.DataFrame, titulo: str = "PARETO COMUNIDAD") -> bytes:
    from pandas import ExcelWriter
    import xlsxwriter  # noqa: F401

    out = BytesIO()
    with ExcelWriter(out, engine="xlsxwriter") as writer:
        # hoja principal
        pareto.to_excel(writer, index=False, sheet_name="Pareto Comunidad")
        wb = writer.book
        ws = writer.sheets["Pareto Comunidad"]

        n = len(pareto)
        if not n:
            return out.getvalue()

        # formatos
        fmt_head = wb.add_format({"bold": True, "align": "center", "bg_color": "#D9E1F2", "border": 1})
        fmt_pct  = wb.add_format({"num_format": "0.00%", "align": "center", "border": 1})
        fmt_int  = wb.add_format({"num_format": "0", "align": "center", "border": 1})
        fmt_txt  = wb.add_format({"align": "left", "border": 1})
        fmt_yel  = wb.add_format({"bg_color": "#FFF2CC"})
        fmt_center = wb.add_format({"align": "center"})

        # anchos + encabezados
        ws.set_column("A:A", 22, fmt_txt)   # Categoría
        ws.set_column("B:B", 50, fmt_txt)   # Descriptor
        ws.set_column("C:C", 12, fmt_int)   # Frecuencia
        ws.set_column("D:D", 12, fmt_pct)   # Porcentaje
        ws.set_column("E:E", 12, fmt_pct)   # % acumulado
        ws.set_column("F:F", 12, fmt_int)   # Acumulado
        ws.set_column("G:G", 8,  fmt_center)# 80/20

        # encabezado con fondo
        ws.set_row(0, None, fmt_head)

        # pintar <=80% en amarillo
        cutoff_idx = int((pareto["% acumulado"] <= 80).sum())
        if cutoff_idx > 0:
            ws.conditional_format(1, 0, cutoff_idx, 6, {"type": "no_blanks", "format": fmt_yel})

        # columnas auxiliares para líneas
        # J: constante 80% (para línea horizontal gris)
        ws.write(0, 9, "80/20")
        for i in range(n):
            ws.write_number(i+1, 9, 0.80)

        # K-L: dos puntos para línea vertical roja
        ws.write(0,10,"CorteX"); ws.write(0,11,"%")
        corte_row = max(1, cutoff_idx)  # al menos 1
        xcat = pareto.iloc[corte_row-1]["Descriptor"]
        ws.write(1,10, xcat); ws.write(2,10, xcat)
        ws.write_number(1,11, 0.0); ws.write_number(2,11, 1.0)

        # gráfico tipo Excel
        chart = wb.add_chart({'type': 'column'})
        # barras coloreadas: azules hasta corte, grises después
        points = [{"fill": {"color": "#5B9BD5"}} for _ in range(n)]
        for i in range(cutoff_idx, n):
            points[i] = {"fill": {"color": "#A6A6A6"}}

        chart.add_series({
            'name': 'Frecuencia',
            'categories': ['Pareto Comunidad', 1, 1, n, 1],  # B
            'values':     ['Pareto Comunidad', 1, 2, n, 2],  # C
            'points': points,
        })

        linea = wb.add_chart({'type': 'line'})
        linea.add_series({
            'name': '% acumulado',
            'categories': ['Pareto Comunidad', 1, 1, n, 1],
            'values':     ['Pareto Comunidad', 1, 4, n, 4],  # E
            'y2_axis': True,
        })
        chart.combine(linea)

        horiz = wb.add_chart({'type': 'line'})
        horiz.add_series({
            'name': '80/20',
            'categories': ['Pareto Comunidad', 1, 1, n, 1],
            'values':     ['Pareto Comunidad', 1, 9, n, 9],  # J
            'y2_axis': True,
        })
        chart.combine(horiz)

        vline = wb.add_chart({'type': 'line'})
        vline.add_series({
            'name': '',
            'categories': ['Pareto Comunidad', 1, 10, 2, 10],  # K2:K3
            'values':     ['Pareto Comunidad', 1, 11, 2, 11],  # L2:L3
            'y2_axis': True,
            'line': {'color': '#C00000', 'width': 2.25},     # rojo
        })
        chart.combine(vline)

        chart.set_title({'name': titulo.upper()})
        chart.set_x_axis({'name': '', 'num_font': {'rotation': -45}})
        chart.set_y_axis({'name': '', 'major_gridlines': {'visible': False}})
        chart.set_y2_axis({'name': '', 'min': 0, 'max': 1, 'major_unit': 0.1, 'num_format': '0%'})
        chart.set_legend({'position': 'bottom'})

        ws.insert_chart(1, 9, chart, {'x_scale': 1.35, 'y_scale': 1.35})

    return out.getvalue()

# ---------- UI mínima ----------
st.title("Pareto Comunidad – MSP (idéntico al formato)")
archivo = st.file_uploader("📄 Subí la Plantilla (XLSX) – hoja `matriz`", type=["xlsx"])
if not archivo:
    st.info("Subí la Plantilla para procesar.")
    st.stop()

# Leer y normalizar columnas (todas las filas)
try:
    df_raw = read_matriz(archivo.getvalue())
except Exception as e:
    st.error(f"Error leyendo 'matriz': {e}")
    st.stop()

df = normalize_columns(df_raw)
st.caption(f"Vista previa (primeras 20 de {len(df)} filas)")
st.dataframe(df.head(20), use_container_width=True)

# Detección combinada (encabezados + texto)
def build_regex_all():
    compiled = {}
    for d in CATEGORIA_POR_DESCRIPTOR:
        keys = SINONIMOS.get(d, []) + [d]
        toks = [re.escape(norm_text(k)) for k in keys if norm_text(k)]
        compiled[d] = re.compile(r"(?:(?<=\s)|^)(" + "|".join(toks) + r")(?:(?=\s)|$)")
    return compiled

with st.spinner("Procesando (encabezados + texto abierto)…"):
    regex = build_regex_all()
    counts_h = detect_by_headers(df, regex)
    counts_t = detect_in_text(df, regex)
    copilado = build_copilado(counts_h, counts_t)
    pareto   = build_pareto(copilado)

if copilado.empty:
    st.warning("No se detectaron descriptores con el catálogo actual.")
    st.stop()

st.subheader("Pareto Comunidad (tabla)")
st.dataframe(pareto, use_container_width=True)

# Vista rápida (web)
plot_df = pareto.head(TOP_N_GRAFICO).copy()
c1, c2 = st.columns([1,1.2])
with c1:
    st.subheader("Top (vista rápida)")
    st.dataframe(plot_df, use_container_width=True)
with c2:
    st.subheader("Gráfico (rápido)")
    st.bar_chart(plot_df.set_index("Descriptor")["Frecuencia"])
    st.line_chart(plot_df.set_index("Descriptor")["% acumulado"])

# Descarga Excel con gráfico idéntico
st.subheader("Descargar Excel final")
st.download_button(
    "⬇️ Pareto Comunidad (Excel con formato y gráfico)",
    data=export_excel(pareto, titulo="PARETO COMUNIDAD"),
    file_name="Pareto_Comunidad.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)



