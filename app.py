# =========================
# 📊 Pareto Comunidad – MSP (1 solo archivo, catálogo embebido)
# =========================
# Flujo:
# 1) Subir Plantilla de Comunidad (hoja 'matriz').
# 2) Generar 'Copilado Comunidad' (Descriptor, Frecuencia) con catálogo embebido.
# 3) Generar 'Pareto Comunidad' (Categoría, Descriptor, Frecuencia, Porcentaje, % Acumulado, Acumulado, 80/20).
# 4) Ver gráfico (barras + línea) con corte al 80% y descargar Excel con todo y gráfico.
#
# Optimización:
# - UI pinta de inmediato; no procesa hasta que subes el archivo.
# - Preview limitada (para no colgar el frontend).
# - Matching por encabezado rápido; escaneo profundo de texto solo si lo activas (toggle).
# - Escaneo profundo vectorizado con regex y por bloques.
# - Gráfico limitado a TOP_N_GRAFICO (descarga incluye TODOS).
#
# Catálogo:
# - Incluye un catálogo de ejemplo embebido basado en descriptores típicos (puedes ampliarlo en la UI).
# - Formato de pegado: NOMBRE CORTO,CATEGORÍA,DESCRIPTOR
#   (una fila por línea, sin encabezados).

import io
from io import BytesIO
import re
import unicodedata
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

st.set_page_config(page_title="Pareto Comunidad – MSP", layout="wide")
TOP_N_GRAFICO = 40  # limitar gráfico para no hacerlo pesado (la descarga incluye todos)

# -------------------------
# Catálogo embebido (ejemplo ampliable)
# -------------------------
CATALOGO_BASE = [
    # NOMBRE CORTO, CATEGORÍA, DESCRIPTOR
    ("HURTO", "DELITOS CONTRA LA PROPIEDAD", "HURTO"),
    ("ROBO", "DELITOS CONTRA LA PROPIEDAD", "ROBO"),
    ("DAÑOS A LA PROPIEDAD", "DELITOS CONTRA LA PROPIEDAD", "DAÑOS A LA PROPIEDAD"),
    ("ASALTO", "DELITOS CONTRA LA PROPIEDAD", "ASALTO"),
    ("TENTATIVA DE ROBO", "DELITOS CONTRA LA PROPIEDAD", "TENTATIVA DE ROBO"),

    ("VENTA DE DROGAS", "DROGAS", "VENTA DE DROGAS"),
    ("TRÁFICO DE DROGAS", "DROGAS", "TRÁFICO DE DROGAS"),
    ("MICROTRÁFICO", "DROGAS", "MICROTRÁFICO"),
    ("CONSUMO DE DROGAS", "DROGAS", "CONSUMO DE DROGAS"),
    ("BÚNKER", "DROGAS", "BÚNKER"),
    ("PUNTO DE VENTA", "DROGAS", "PUNTO DE VENTA"),
    ("CONSUMO DE ALCOHOL", "ALCOHOL", "CONSUMO DE ALCOHOL EN VÍA PÚBLICA"),
    ("LICORES", "ALCOHOL", "CONSUMO DE ALCOHOL EN VÍA PÚBLICA"),

    ("HOMICIDIOS", "DELITOS CONTRA LA VIDA", "HOMICIDIOS"),
    ("HERIDOS", "DELITOS CONTRA LA VIDA", "HERIDOS"),
    ("TENTATIVA DE HOMICIDIO", "DELITOS CONTRA LA VIDA", "TENTATIVA DE HOMICIDIO"),

    ("VIOLENCIA DOMÉSTICA", "VIOLENCIA", "VIOLENCIA DOMÉSTICA"),
    ("VIOLENCIA INTRAFAMILIAR", "VIOLENCIA", "VIOLENCIA DOMÉSTICA"),
    ("AGRESIÓN", "VIOLENCIA", "AGRESIÓN"),
    ("ABUSO SEXUAL", "VIOLENCIA", "ABUSO SEXUAL"),
    ("VIOLACIÓN", "VIOLENCIA", "VIOLACIÓN"),
    ("ACOSO SEXUAL CALLEJERO", "RIESGO SOCIAL", "ACOSO SEXUAL CALLEJERO"),
    ("ACOSO ESCOLAR", "RIESGO SOCIAL", "ACOSO ESCOLAR (BULLYING)"),
    ("ACTOS OBSCENOS", "RIESGO SOCIAL", "ACTOS OBSCENOS EN VIA PUBLICA"),

    ("PANDILLAS", "ORDEN PÚBLICO", "PANDILLAS"),
    ("VAGANCIA", "ORDEN PÚBLICO", "VAGANCIA"),
    ("INDIGENCIA", "ORDEN PÚBLICO", "INDIGENCIA"),
    ("RUIDO", "ORDEN PÚBLICO", "CONTAMINACIÓN SONORA"),
    ("CARRERA ILEGAL", "ORDEN PÚBLICO", "CARRERAS ILEGALES"),
    ("ARMAS BLANCAS", "ORDEN PÚBLICO", "PORTACIÓN DE ARMA BLANCA"),

    # Puedes seguir ampliando...
]

# -------------------------
# Utilidades de normalización
# -------------------------
def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def norm(s: Optional[str]) -> str:
    """Normaliza strings: quita tildes, baja a minúsculas y colapsa espacios."""
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = strip_accents(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def split_tokens(txt: str) -> List[str]:
    """Divide celdas con listas tipo '1,2,3' o textos con separadores comunes."""
    if not isinstance(txt, str):
        return []
    parts = re.split(r"[;,/|]+", txt)
    return [p.strip() for p in parts if p and p.strip()]

def preview_df(df: pd.DataFrame, n: int = 200) -> pd.DataFrame:
    """Solo primeras N para evitar carga pesada en el frontend."""
    return df.head(n).copy()

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza encabezados una sola vez para matching por header."""
    df2 = df.copy()
    df2.columns = [
        re.sub(r"\s+", " ", strip_accents(str(c)).strip().lower())
        for c in df2.columns
    ]
    return df2

# -------------------------
# Carga de archivos + UI
# -------------------------
st.title("Pareto Comunidad (MSP) – 1 archivo")

with st.expander("Instrucciones", expanded=True):
    st.markdown(f"""
1. **Sube la Plantilla de Comunidad** (hoja `matriz`).
2. (Opcional) **Amplía el catálogo** pegando filas en el apartado de abajo (formato: `NOMBRE CORTO,CATEGORÍA,DESCRIPTOR`).
3. Genera el **Copilado** y luego el **Pareto**.  
4. Descarga el **Excel** con hojas y **gráfico embebido**.

> El gráfico muestra **top {TOP_N_GRAFICO}** para rendimiento; la descarga incluye **todos**.
""")

plantilla_file = st.file_uploader("📄 Plantilla de Comunidad (XLSX) – hoja 'matriz'", type=["xlsx"], key="plantilla")
st.divider()

# --------- Catálogo embebido + pegado opcional ----------
with st.expander("➕ Ampliar/actualizar catálogo (pegar líneas CSV)", expanded=False):
    st.markdown("Formato por línea: `NOMBRE CORTO,CATEGORÍA,DESCRIPTOR` (sin encabezados).")
    pasted = st.text_area("Pega aquí (opcional):", height=160, placeholder="EJEMPLO:\nHURTO,DELITOS CONTRA LA PROPIEDAD,HURTO\nVENTA DE DROGAS,DROGAS,VENTA DE DROGAS")
    def parse_pasted(text: str) -> List[Tuple[str,str,str]]:
        rows = []
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            parts = [p.strip() for p in line.split(",")]
            if len(parts) < 3:
                continue
            nc, cat, desc = parts[0], parts[1], ",".join(parts[2:]).strip()
            rows.append((nc, cat, desc))
        return rows
    pasted_rows = parse_pasted(pasted) if pasted else []

# -------------------------
# Lectura segura de la plantilla
# -------------------------
@st.cache_data(show_spinner=False)
def read_plantilla_matriz(file) -> pd.DataFrame:
    # Leemos completo; preview se hace limitada en la UI
    return pd.read_excel(file, sheet_name="matriz", engine="openpyxl")

# -------------------------
# Construcción del catálogo (embebido + opcional pegado)
# -------------------------
def build_catalogo_df() -> pd.DataFrame:
    base = pd.DataFrame(CATALOGO_BASE, columns=["NOMBRE CORTO", "CATEGORÍA", "DESCRIPTOR"])
    if pasted_rows:
        extra = pd.DataFrame(pasted_rows, columns=["NOMBRE CORTO", "CATEGORÍA", "DESCRIPTOR"])
        # Concatenar y deduplicar por DESCRIPTOR (priorizamos últimas filas pegadas)
        cat_df = pd.concat([base, extra], ignore_index=True)
        cat_df = cat_df.dropna(subset=["DESCRIPTOR"])
        cat_df = cat_df.drop_duplicates(subset=["DESCRIPTOR"], keep="last")
        return cat_df
    else:
        return base

def build_keyword_maps(desc_df: pd.DataFrame) -> Tuple[Dict[str, str], Dict[str, str], Dict[str, str]]:
    """
    Devuelve tres mapas normalizados:
    - by_descriptor_norm: norm(DESCRIPTOR) -> DESCRIPTOR
    - by_nombre_corto_norm: norm(NOMBRE CORTO) -> DESCRIPTOR
    - cat_by_descriptor: DESCRIPTOR -> CATEGORÍA
    """
    by_desc_norm: Dict[str, str] = {}
    by_nc_norm: Dict[str, str] = {}
    cat_by_desc: Dict[str, str] = {}
    for _, r in desc_df.iterrows():
        desc = str(r["DESCRIPTOR"]).strip()
        cat = str(r["CATEGORÍA"]).strip() if "CATEGORÍA" in r and not pd.isna(r["CATEGORÍA"]) else ""
        nc  = str(r["NOMBRE CORTO"]).strip() if "NOMBRE CORTO" in desc_df.columns and not pd.isna(r["NOMBRE CORTO"]) else ""
        if desc:
            by_desc_norm[norm(desc)] = desc
            cat_by_desc[desc] = cat
        if nc:
            by_nc_norm[norm(nc)] = desc
    return by_desc_norm, by_nc_norm, cat_by_desc

# -------------------------
# Extracción de descriptores (rápido + profundo opcional)
# -------------------------
def detect_descriptors_from_dataframe(df_raw: pd.DataFrame,
                                      by_desc_norm: Dict[str, str],
                                      by_nc_norm: Dict[str, str],
                                      deep_scan: bool = False) -> List[str]:
    """
    Modo rápido: solo por encabezado y celdas marcadas (muy veloz).
    Modo profundo (opcional): añade búsqueda en texto, vectorizada con regex y por bloques.
    """
    df = normalize_columns(df_raw)

    # 1) mapeo por encabezado (rápido)
    keys_desc = list(by_desc_norm.keys())
    keys_nc   = list(by_nc_norm.keys())

    def header_to_descriptor(ncol: str) -> Optional[str]:
        for k in keys_desc:
            if k and (k == ncol or k in ncol or ncol in k):
                return by_desc_norm[k]
        for k in keys_nc:
            if k and (k == ncol or k in ncol or ncol in k):
                return by_nc_norm[k]
        return None

    col_map: Dict[str, str] = {}  # col_name_normalized -> descriptor oficial
    for col in df.columns:
        d = header_to_descriptor(col)
        if d:
            col_map[col] = d

    def is_marked_series(s: pd.Series) -> pd.Series:
        # True si hay algo distinto de vacío/0/no
        s2 = s.copy()
        # numéricos
        numeric_mask = pd.to_numeric(s2, errors="coerce").fillna(0) != 0
        # texto
        txt = s2.astype(str).str.strip().str.lower()
        txt = txt.apply(strip_accents)
        text_mask = ~txt.isin(["", "no", "0", "nan", "none", "false"])
        return numeric_mask | text_mask

    hits: List[str] = []

    # 2) contar por columnas mapeadas (rápido y vectorizado)
    for col, desc in col_map.items():
        mask = is_marked_series(df[col])
        count = int(mask.sum())
        if count > 0:
            hits.extend([desc] * count)

    if not deep_scan:
        return hits

    # 3) escaneo profundo en texto (vectorizado por regex)
    all_keys = list(set(keys_desc + keys_nc))
    all_keys.sort(key=len, reverse=True)  # más largos primero
    pattern = "|".join([re.escape(k) for k in all_keys if k])
    if not pattern:
        return hits

    regex = re.compile(pattern)

    # columnas con texto (object) que valen la pena
    text_cols = []
    for col in df.columns:
        col_series = df[col]
        if col_series.dtype == object:
            sample = col_series.head(200).astype(str)
            if (sample.str.strip() != "").mean() > 0.05:
                text_cols.append(col)

    def norm_text_cell(x: str) -> str:
        x = strip_accents(x)
        x = re.sub(r"\s+", " ", x.strip().lower())
        return x

    block_size = 5000
    for col in text_cols:
        col_norm = df[col].astype(str).apply(norm_text_cell)
        for i in range(0, len(col_norm), block_size):
            part = col_norm.iloc[i:i+block_size]
            matched_idx = part[part.str.contains(regex, na=False)].index
            if len(matched_idx) == 0:
                continue
            for ridx in matched_idx:
                txt = part.loc[ridx]
                m = regex.search(txt)
                if not m:
                    continue
                key = m.group(0)
                desc = by_desc_norm.get(key) or by_nc_norm.get(key)
                if desc:
                    hits.append(desc)

    return hits

# -------------------------
# Agregaciones y Pareto
# -------------------------
def make_copilado(hits: List[str]) -> pd.DataFrame:
    """Devuelve Copilado Comunidad: Descriptor, Frecuencia."""
    if not hits:
        return pd.DataFrame({"Descriptor": [], "Frecuencia": []})
    s = pd.Series(hits, name="Descriptor")
    df = s.value_counts(dropna=False).rename_axis("Descriptor").reset_index(name="Frecuencia")
    df = df.sort_values(["Frecuencia", "Descriptor"], ascending=[False, True], ignore_index=True)
    return df

def make_pareto(copilado_df: pd.DataFrame, cat_by_desc: Dict[str, str]) -> pd.DataFrame:
    """Calcula tabla de Pareto uniendo Categoría y métricas (% y acumulados)."""
    if copilado_df.empty:
        return pd.DataFrame(columns=[
            "Categoría", "Descriptor", "Frecuencia", "Porcentaje", "% Acumulado", "Acumulado", "80/20"
        ])

    df = copilado_df.copy()
    df["Categoría"] = df["Descriptor"].map(cat_by_desc).fillna("")
    total = df["Frecuencia"].sum()
    df["Porcentaje"] = (df["Frecuencia"] / total) * 100.0
    df = df.sort_values(["Frecuencia", "Descriptor"], ascending=[False, True], ignore_index=True)
    df["% Acumulado"] = df["Porcentaje"].cumsum()
    df["Acumulado"] = df["Frecuencia"].cumsum()
    df["80/20"] = np.where(df["% Acumulado"] <= 80.0, "≤80%", ">80%")
    df = df[["Categoría", "Descriptor", "Frecuencia", "Porcentaje", "% Acumulado", "Acumulado", "80/20"]]
    return df

def plot_pareto(df_pareto: pd.DataFrame):
    """Gráfico Pareto con barras (Frecuencia) + línea (% Acumulado) y corte al 80%."""
    if df_pareto.empty:
        return go.Figure()

    x = df_pareto["Descriptor"].astype(str).tolist()
    y = df_pareto["Frecuencia"].tolist()
    cum = df_pareto["% Acumulado"].tolist()

    fig = go.Figure()
    # Barras
    fig.add_bar(x=x, y=y, name="Frecuencia", yaxis="y1")
    # Línea acumulada
    fig.add_trace(go.Scatter(x=x, y=cum, name="% Acumulado", yaxis="y2", mode="lines+markers"))
    # Línea de corte 80%
    fig.add_trace(go.Scatter(x=x, y=[80.0]*len(x), name="Corte 80%", yaxis="y2", mode="lines", line=dict(dash="dash")))

    fig.update_layout(
        title="Pareto Comunidad",
        xaxis_title="Descriptor",
        yaxis=dict(title="Frecuencia"),
        yaxis2=dict(title="% Acumulado", overlaying="y", side="right", range=[0, 100]),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        margin=dict(l=20, r=20, t=50, b=80),
        hovermode="x unified"
    )
    return fig

# -------------------------
# Exportar a Excel con gráfico embebido
# -------------------------
def build_excel_bytes(copilado_df: pd.DataFrame, pareto_df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        # Hojas
        copilado_df.to_excel(writer, index=False, sheet_name="Copilado Comunidad")
        pareto_df.to_excel(writer, index=False, sheet_name="Pareto Comunidad")

        wb  = writer.book
        wsP = writer.sheets["Pareto Comunidad"]

        # Insertar gráfico de Pareto (barras + línea) en Excel
        n = len(pareto_df)
        if n >= 1:
            # Columnas en hoja Pareto:
            # (A) Categoría, (B) Descriptor, (C) Frecuencia, (D) Porcentaje, (E) % Acumulado, (F) Acumulado, (G) 80/20
            chart = wb.add_chart({'type': 'column'})

            # Series Frecuencia (barras)
            chart.add_series({
                'name':       'Frecuencia',
                'categories': ['Pareto Comunidad', 1, 1, n, 1],  # B2:B(n+1)
                'values':     ['Pareto Comunidad', 1, 2, n, 2],  # C2:C(n+1)
                'y2_axis':    False
            })

            # Línea % Acumulado
            line_chart = wb.add_chart({'type': 'line'})
            line_chart.add_series({
                'name':       '% Acumulado',
                'categories': ['Pareto Comunidad', 1, 1, n, 1],  # B
                'values':     ['Pareto Comunidad', 1, 4, n, 4],  # E
                'y2_axis':    True,
                'marker':     {'type': 'automatic'}
            })

            chart.combine(line_chart)
            chart.set_title({'name': 'Pareto Comunidad'})
            chart.set_x_axis({'name': 'Descriptor'})
            chart.set_y_axis({'name': 'Frecuencia'})
            chart.set_y2_axis({'name': '% Acumulado', 'major_gridlines': {'visible': False}, 'min': 0, 'max': 100})

            # Insertar gráfico en la hoja (fila 2, col 9 ≈ J3)
            wsP.insert_chart(1, 9, chart, {'x_scale': 1.3, 'y_scale': 1.3})

    return output.getvalue()

# -------------------------
# UI principal
# -------------------------
if plantilla_file:
    with st.spinner("Leyendo archivo…"):
        try:
            df_matriz = read_plantilla_matriz(plantilla_file)
        except Exception as e:
            st.error(f"Error al leer la Plantilla: {e}")
            st.stop()

    st.subheader("1) Previsualización")
    st.caption("Primeras filas de la hoja `matriz` (preview limitada)")
    st.dataframe(preview_df(df_matriz), use_container_width=True)

    # Construir catálogo (embebido + pegado)
    cat_df = build_catalogo_df()
    st.caption("Catálogo activo (primeras filas)")
    st.dataframe(cat_df.head(20), use_container_width=True)

    # Construir mapas de palabras clave
    by_desc_norm, by_nc_norm, cat_by_desc = build_keyword_maps(cat_df)

    st.subheader("2) Generar 'Copilado Comunidad'")
    deep_scan = st.toggle(
        "Activar escaneo profundo de texto (más lento)",
        value=False,
        help="Si está apagado, solo se detecta por encabezados y celdas marcadas. Enciéndelo si necesitas raspar menciones en texto libre."
    )

    with st.spinner("Procesando…"):
        hits = detect_descriptors_from_dataframe(df_matriz, by_desc_norm, by_nc_norm, deep_scan=deep_scan)
        copilado_df = make_copilado(hits)

    if copilado_df.empty:
        st.warning("No se detectaron descriptores con el catálogo actual. Puedes ampliar el catálogo pegando filas en el expander superior.")
        st.stop()

    st.success("Copilado generado.")
    st.dataframe(copilado_df, use_container_width=True)

    st.subheader("3) Generar 'Pareto Comunidad'")
    pareto_df = make_pareto(copilado_df, cat_by_desc)
    st.dataframe(pareto_df, use_container_width=True)

    st.subheader("4) Gráfico Pareto (con corte al 80%)")
    pareto_df_plot = pareto_df.head(TOP_N_GRAFICO).copy()
    fig = plot_pareto(pareto_df_plot)
    st.plotly_chart(fig, use_container_width=True)
    st.caption(f"Mostrando top {len(pareto_df_plot)} por frecuencia. La descarga incluye **todos** los descriptores.")

    st.subheader("5) Descarga")
    xls_bytes = build_excel_bytes(copilado_df, pareto_df)
    st.download_button(
        label="⬇️ Descargar Excel (Copilado + Pareto + Gráfico)",
        data=xls_bytes,
        file_name="Pareto_Comunidad.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Sube la **Plantilla de Comunidad** (XLSX, hoja `matriz`) para comenzar. Puedes ampliar el catálogo en el expander si lo necesitas.")



