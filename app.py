# =========================
# 📊 Pareto Comunidad – MSP
# =========================
# Flujo:
# 1) Subir Plantilla de Comunidad (hoja 'matriz') y Descriptores (hoja 'Descriptores (ACTUALIZADOS)')
# 2) Generar 'Copilado Comunidad' (Descriptor, Frecuencia)
# 3) Generar 'Pareto Comunidad' (Categoría, Descriptor, Frecuencia, Porcentaje, % Acumulado, Acumulado, 80/20)
# 4) Ver gráfico (barras + línea) con corte al 80% y descargar Excel con todo y gráfico

import io
from io import BytesIO
import re
import unicodedata
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from unidecode import unidecode
import plotly.graph_objects as go

st.set_page_config(page_title="Pareto Comunidad – MSP", layout="wide")

# -------------------------
# Utilidades de normalización
# -------------------------
def norm(s: str) -> str:
    if s is None:
        return ""
    if not isinstance(s, str):
        s = str(s)
    s = unidecode(s.strip().lower())
    s = re.sub(r"\s+", " ", s)
    return s

def split_tokens(txt: str) -> List[str]:
    """Divide celdas con listas tipo '1,2,3' o textos separados por comas/puntos y coma."""
    if not isinstance(txt, str):
        return []
    parts = re.split(r"[;,/|]+", txt)
    return [p.strip() for p in parts if p and p.strip()]

# -------------------------
# Carga de archivos
# -------------------------
st.title("Pareto Comunidad (MSP)")

st.markdown("""
Sube los archivos requeridos:

- **Plantilla de Comunidad** (hoja: `matriz`)
- **Descriptores (ACTUALIZADOS)** (hoja: `Descriptores (ACTUALIZADOS)`)
""")

col_up1, col_up2 = st.columns(2)
with col_up1:
    plantilla_file = st.file_uploader("📄 Plantilla de Comunidad (XLSX)", type=["xlsx"], key="plantilla")
with col_up2:
    desc_file = st.file_uploader("📚 Descriptores (ACTUALIZADOS) (XLSX)", type=["xlsx"], key="desc")

st.divider()

# -------------------------
# Lectura segura de hojas
# -------------------------
@st.cache_data(show_spinner=False)
def read_plantilla_matriz(file) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name="matriz")

@st.cache_data(show_spinner=False)
def read_desc(file) -> pd.DataFrame:
    # Esperado: NOMBRE CORTO, CATEGORÍA, DESCRIPTOR, DESCRIPCIÓN
    df = pd.read_excel(file, sheet_name="Descriptores (ACTUALIZADOS)")
    # Limpieza de encabezados (quitar espacios y normalizar)
    df.columns = [c.strip() for c in df.columns]
    # Renombrar variantes comunes si vinieran con diferencias
    ren = {
        "CATEGORIA": "CATEGORÍA",
        "DESCRIPTORES": "DESCRIPTOR",
        "NOMBRE CORTO": "NOMBRE CORTO",
    }
    for k, v in ren.items():
        if k in df.columns and v not in df.columns:
            df = df.rename(columns={k: v})
    # Chequeo mínimo
    needed = {"NOMBRE CORTO", "CATEGORÍA", "DESCRIPTOR"}
    miss = [c for c in needed if c not in df.columns]
    if miss:
        raise ValueError(f"Faltan columnas en Descriptores (ACTUALIZADOS): {miss}")
    # Drop filas vacías de DESCRIPTOR
    df = df.dropna(subset=["DESCRIPTOR"]).copy()
    return df

def build_keyword_maps(desc_df: pd.DataFrame) -> Tuple[Dict[str, str], Dict[str, str], Dict[str, str]]:
    """
    Devuelve tres mapas normalizados:
    - by_descriptor_norm: norm(DESCRIPTOR) -> DESCRIPTOR
    - by_nombre_corto_norm: norm(NOMBRE CORTO) -> DESCRIPTOR
    - cat_by_descriptor: DESCRIPTOR -> CATEGORÍA
    """
    by_desc_norm = {}
    by_nc_norm = {}
    cat_by_desc = {}
    for _, r in desc_df.iterrows():
        desc = str(r["DESCRIPTOR"]).strip()
        cat = str(r["CATEGORÍA"]).strip() if not pd.isna(r["CATEGORÍA"]) else ""
        nc = str(r["NOMBRE CORTO"]).strip() if "NOMBRE CORTO" in desc_df.columns and not pd.isna(r["NOMBRE CORTO"]) else ""
        if desc:
            by_desc_norm[norm(desc)] = desc
            cat_by_desc[desc] = cat
        if nc:
            by_nc_norm[norm(nc)] = desc
    return by_desc_norm, by_nc_norm, cat_by_desc

# -------------------------
# Extracción de descriptores desde la plantilla
# -------------------------
def detect_descriptors_from_dataframe(df: pd.DataFrame,
                                      by_desc_norm: Dict[str, str],
                                      by_nc_norm: Dict[str, str]) -> List[str]:
    """
    Heurística flexible:
    - Si el nombre de la columna coincide (por substring) con algún DESCRIPTOR o NOMBRE CORTO,
      y la celda está marcada (no vacía/True/"si"/"sí"/"1"), cuenta 1 mención.
    - Además, examina texto en celdas (strings) y busca substrings que empaten con DESCRIPTOR/NOMBRE CORTO.
    """
    # Set de claves normalizadas para búsqueda rápida
    keys_desc = list(by_desc_norm.keys())
    keys_nc = list(by_nc_norm.keys())

    hits: List[str] = []

    # Mapa rápido de columnas candidatas (por encabezado)
    col_map_descriptor = {}  # col -> DESCRIPTOR (si el header matchea)
    for col in df.columns:
        ncol = norm(str(col))
        # match exacto/substring con descriptor
        matched_desc = None
        for k in keys_desc:
            if k and (k == ncol or k in ncol or ncol in k):
                matched_desc = by_desc_norm[k]
                break
        if matched_desc is None:
            for k in keys_nc:
                if k and (k == ncol or k in ncol or ncol in k):
                    matched_desc = by_nc_norm[k]
                    break
        if matched_desc:
            col_map_descriptor[col] = matched_desc

    def is_marked(val) -> bool:
        if pd.isna(val):
            return False
        if isinstance(val, (int, float)):
            return val != 0
        v = norm(str(val))
        return v not in ("", "no", "0", "nan", "ninguno") and v != "false"

    # 1) Recorremos columnas "mapeadas por encabezado"
    for col, desc in col_map_descriptor.items():
        for _, v in df[col].items():
            if is_marked(v):
                hits.append(desc)

    # 2) Búsqueda en texto de todas las celdas tipo string
    #    - Divide por separadores comunes y revisa cada token
    for col in df.columns:
        series = df[col]
        for _, v in series.items():
            if isinstance(v, str) and v.strip():
                # tokens directos
                tokens = split_tokens(v)
                if not tokens:
                    tokens = [v]
                for tk in tokens:
                    ntk = norm(tk)
                    if not ntk:
                        continue
                    # match descriptor exacto o substring
                    found = False
                    for k in keys_desc:
                        if k and (k == ntk or k in ntk or ntk in k):
                            hits.append(by_desc_norm[k])
                            found = True
                            break
                    if found:
                        continue
                    for k in keys_nc:
                        if k and (k == ntk or k in ntk or ntk in k):
                            hits.append(by_nc_norm[k])
                            break

    return hits

def make_copilado(hits: List[str]) -> pd.DataFrame:
    """Devuelve Copilado Comunidad: Descriptor, Frecuencia."""
    if not hits:
        return pd.DataFrame({"Descriptor": [], "Frecuencia": []})
    s = pd.Series(hits, name="Descriptor")
    df = s.value_counts(dropna=False).rename_axis("Descriptor").reset_index(name="Frecuencia")
    # Orden desc por Frecuencia (y por Descriptor para estabilidad)
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
    # Reordenar columnas al formato solicitado
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

    # Línea horizontal 80%
    fig.add_hline(y=80, line_dash="dash", line_color="red", annotation_text="Corte 80%", annotation_position="top left", secondary_y=True)

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
        # Construimos los rangos a partir del tamaño real de la tabla
        n = len(pareto_df)
        # Columnas: Descriptor(B), Frecuencia(C), % Acumulado(E)
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

        # Insertar gráfico en la hoja (fila 2, col 9 aprox. = J3)
        wsP.insert_chart(1, 9, chart, {'x_scale': 1.3, 'y_scale': 1.3})

        # Agregar una línea de referencia al 80% en la hoja (no hay objeto "línea horizontal" en xlsxwriter para y2,
        # pero queda cubierta por el gráfico interactivo y por la columna "80/20" en la tabla).
    return output.getvalue()

# -------------------------
# UI principal
# -------------------------
if plantilla_file and desc_file:
    with st.spinner("Leyendo archivos..."):
        try:
            df_matriz = read_plantilla_matriz(plantilla_file)
            df_desc   = read_desc(desc_file)
        except Exception as e:
            st.error(f"Error al leer archivos: {e}")
            st.stop()

    st.subheader("1) Previsualización")
    st.caption("Primeras filas de la hoja `matriz` de la Plantilla")
    st.dataframe(df_matriz.head(10), use_container_width=True)

    st.caption("Catálogo de descriptores (columnas clave)")
    st.write(df_desc[["NOMBRE CORTO", "CATEGORÍA", "DESCRIPTOR"]].head(10))

    # Construir mapas de palabras clave
    by_desc_norm, by_nc_norm, cat_by_desc = build_keyword_maps(df_desc)

    st.subheader("2) Generar 'Copilado Comunidad'")
    st.caption("Detectando descriptores en columnas y texto...")

    hits = detect_descriptors_from_dataframe(df_matriz, by_desc_norm, by_nc_norm)
    copilado_df = make_copilado(hits)

    if copilado_df.empty:
        st.warning("No se detectaron descriptores. Revisa que los encabezados/celdas contengan valores coincidentes con el catálogo.")
        st.stop()

    st.success("Copilado generado.")
    st.dataframe(copilado_df, use_container_width=True)

    st.subheader("3) Generar 'Pareto Comunidad'")
    pareto_df = make_pareto(copilado_df, cat_by_desc)
    st.dataframe(pareto_df, use_container_width=True)

    st.subheader("4) Gráfico Pareto (con corte al 80%)")
    fig = plot_pareto(pareto_df)
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("5) Descarga")
    xls_bytes = build_excel_bytes(copilado_df, pareto_df)
    st.download_button(
        label="⬇️ Descargar Excel (Copilado + Pareto + Gráfico)",
        data=xls_bytes,
        file_name="Pareto_Comunidad.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Sube la **Plantilla de Comunidad** y el **catálogo Descriptores (ACTUALIZADOS)** para comenzar.")


