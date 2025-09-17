# Pareto Comunidad – MSP (automático, 1 archivo, export con línea vertical 80%)

import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Optional

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Pareto Comunidad – MSP", layout="wide", initial_sidebar_state="collapsed")
TOP_N_GRAFICO = 50

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
    return pd.read_excel(BytesIO(file_bytes), sheet_name="matriz", engine="openpyxl")

# ---------- Catálogo + sinónimos (puedes ampliar) ----------
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
    "Consumo de drogas": ["consumo de drogas", "consumen drogas", "fumando piedra", "consumo marihuana"],
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
    "Robo a vehiculos": ["robo a vehiculos", "robo carro", "robo moto"],
    "Robo a comercio (Intimidación)": ["robo a comercio intimidacion", "asalto a comercio"],
    "Daños/Vandalismo": ["danos", "vandalismo", "grafiti", "daño a la propiedad", "destruccion"],
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
        tokens = [re.escape(norm_text(k)) for k in keys if norm_text(k)]
        if not tokens:
            tokens = [re.escape(norm_text(desc))]
        pat = r"(?:(?<=\s)|^)(" + "|".join(tokens) + r")(?:(?=\s)|$)"
        compiled[desc] = re.compile(pat)
    return compiled

def detect_by_headers(df_norm: pd.DataFrame, regex_by_desc: Dict[str, re.Pattern]) -> Dict[str, int]:
    counts = {d: 0 for d in CATEGORIA_POR_DESCRIPTOR.keys()}
    cols = list(df_norm.columns)
    for desc, pat in regex_by_desc.items():
        hit_cols = [c for c in cols if re.search(pat, " " + c + " ") is not None]
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

# ---------- Export Excel (con línea vertical 80%) ----------
def export_excel(copilado: pd.DataFrame, pareto: pd.DataFrame) -> bytes:
    from pandas import ExcelWriter
    import xlsxwriter  # noqa: F401

    out = BytesIO()
    with ExcelWriter(out, engine="xlsxwriter") as writer:
        # Hojas
        copilado.to_excel(writer, index=False, sheet_name="Copilado Comunidad")
        pareto.to_excel(writer, index=False, sheet_name="Pareto Comunidad")

        wb  = writer.book
        ws  = writer.sheets["Pareto Comunidad"]

        n = len(pareto)
        if not n:
            return out.getvalue()

        # Formatos
        fmt_pct   = wb.add_format({"num_format": "0.00%", "align": "center"})
        fmt_int   = wb.add_format({"num_format": "0", "align": "center"})
        fmt_head  = wb.add_format({"bold": True, "align": "center", "bg_color": "#D9E1F2"})
        fmt_yellow= wb.add_format({"bg_color": "#FFF2CC"})
        fmt_center= wb.add_format({"align": "center"})

        # Encabezados centrados y porcentajes
        for c in range(8):  # A..H
            ws.set_row(0, None, fmt_head)
        ws.set_column("A:A", 22)
        ws.set_column("B:B", 42)
        ws.set_column("C:C", 12, fmt_int)
        ws.set_column("D:D", 12, fmt_pct)  # Porcentaje
        ws.set_column("E:E", 12, fmt_pct)  # % Acumulado
        ws.set_column("F:F", 12, fmt_int)
        ws.set_column("G:G", 8,  fmt_center)

        # Pintar <=80% en amarillo
        cutoff_idx = int((pareto["% Acumulado"] <= 80).sum())  # filas (1..k) cumplen
        if cutoff_idx > 0:
            ws.conditional_format(1, 0, cutoff_idx, 7, {"type": "no_blanks", "format": fmt_yellow})

        # --- Datos auxiliares para líneas en el gráfico ---
        # Horizontal 80%: una serie constante al 80
        ws.write(0, 9, "Const 80%")  # J1
        for i in range(n):
            ws.write(i+1, 9, 0.8)

        # Vertical 80%: dos puntos (0% y 100%) en la categoría de corte
        # los ponemos en K y L (categoría y valor)
        ws.write(0, 10, "Corte X")   # K1
        ws.write(0, 11, "Corte Y")   # L1
        corte_cat_row = cutoff_idx if cutoff_idx >= 1 else 1
        # categorías iguales en dos filas -> línea vertical
        ws.write(1, 10, pareto.iloc[corte_cat_row-1, 1])  # Descriptor
        ws.write(2, 10, pareto.iloc[corte_cat_row-1, 1])  # Descriptor
        ws.write(1, 11, 0.0)
        ws.write(2, 11, 1.0)

        # --- Gráfico ---
        chart = wb.add_chart({'type': 'column'})
        # Colorear barras por punto (amarillas hasta corte)
        points = [{"fill": {"color": "#5B9BD5"}} for _ in range(n)]
        for i in range(cutoff_idx, n):
            points[i] = {"fill": {"color": "#A6A6A6"}}

        chart.add_series({
            'name': 'Frecuencia',
            'categories': ['Pareto Comunidad', 1, 1, n, 1],  # B
            'values':     ['Pareto Comunidad', 1, 2, n, 2],  # C
            'points': points,
        })

        # Línea % acumulado (eje secundario)
        line = wb.add_chart({'type': 'line'})
        line.add_series({
            'name': '% Acumulado',
            'categories': ['Pareto Comunidad', 1, 1, n, 1],
            'values':     ['Pareto Comunidad', 1, 4, n, 4],  # E
            'y2_axis': True,
        })
        chart.combine(line)

        # Línea horizontal 80%
        line80 = wb.add_chart({'type': 'line'})
        line80.add_series({
            'name': '80%',
            'categories': ['Pareto Comunidad', 1, 1, n, 1],
            'values':     ['Pareto Comunidad', 1, 9, n, 9],  # J
            'y2_axis': True,
        })
        chart.combine(line80)

        # Línea vertical (dos puntos con la misma categoría)
        vline = wb.add_chart({'type': 'line'})
        vline.add_series({
            'name': 'Corte 80%',
            'categories': ['Pareto Comunidad', 1, 10, 2, 10],  # K2:K3 (misma categoría)
            'values':     ['Pareto Comunidad', 1, 11, 2, 11],  # L2:L3 -> 0 a 1
            'y2_axis': True,
        })
        chart.combine(vline)

        chart.set_title({'name': 'Pareto Comunidad'})
        chart.set_x_axis({'name': 'Descriptor'})
        chart.set_y_axis({'name': 'Frecuencia'})
        chart.set_y2_axis({'name': '% Acumulado', 'min': 0, 'max': 1, 'major_unit': 0.1, 'num_format': '0%'})

        ws.insert_chart(1, 9, chart, {'x_scale': 1.35, 'y_scale': 1.35})

    return out.getvalue()

# ---------- UI ----------
st.title("Pareto Comunidad – MSP (automático, con corte 80% marcado)")
archivo = st.file_uploader("📄 Subí la Plantilla (XLSX) – hoja `matriz`", type=["xlsx"])
if not archivo:
    st.info("Subí la Plantilla para procesar.")
    st.stop()

# Leer todas las filas y normalizar columnas
try:
    df_raw = read_matriz(archivo.getvalue())
except Exception as e:
    st.error(f"Error leyendo 'matriz': {e}")
    st.stop()

df = normalize_columns(df_raw)
st.caption(f"Vista previa (primeras 20 de {len(df)} filas)")
st.dataframe(df.head(20), use_container_width=True)

# Detección combinada
with st.spinner("Procesando (encabezados + texto abierto)…"):
    regex_by_desc = build_regex_by_desc()
    counts_headers = detect_by_headers(df, regex_by_desc)
    counts_text    = detect_in_text(df, regex_by_desc)
    copilado = build_copilado(counts_headers, counts_text)
    pareto   = build_pareto(copilado)

if copilado.empty:
    st.warning("No se detectaron descriptores. Ajusta/añade sinónimos si tu formulario usa otros términos.")
    st.stop()

st.subheader("Copilado Comunidad")
st.dataframe(copilado, use_container_width=True)

st.subheader("Pareto Comunidad")
st.dataframe(pareto, use_container_width=True)

st.subheader("Gráfico (vista rápida)")
plot_df = pareto.head(TOP_N_GRAFICO).copy()
st.bar_chart(plot_df.set_index("Descriptor")["Frecuencia"])
st.line_chart(plot_df.set_index("Descriptor")["% Acumulado"])

st.subheader("Descargar Excel (con gráfico y corte 80%)")
st.download_button(
    "⬇️ Copilado + Pareto + gráfico",
    data=export_excel(copilado, pareto),
    file_name="Pareto_Comunidad.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


