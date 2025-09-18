# app.py — Generador de Pareto (Delitos, Riesgos Sociales y Otros Factores)
# ─────────────────────────────────────────────────────────────────────────────
# ✅ Hecho para tus matrices grandes (115+ columnas, 500–600 filas)
# ✅ Lee la matriz .XLSX tal cual
# ✅ Detecta encabezados AMARILLOS (rango AI–ET por defecto, configurable)
# ✅ (NUEVO) Filtro por CATEGORÍAS con 3 preseleccionadas: "Delitos",
#    "Riesgos sociales", "Otros factores" usando un catálogo de descriptores
#    opcional (archivo "DESCRIPTORES ACTUALIZADOS 2024 v2.xlsx").
# ✅ Cuenta TODAS las menciones (no vacío, ≠ 0) — sin omitir filas
# ✅ Calcula % y % acumulado (el acumulado cierra en el total)
# ✅ Pareto: barras + línea acumulada + línea 80% y corte vertical
# ✅ Descargas: PNG y Excel (tabla + gráfico embebido)
#
# Requisitos (requirements.txt):
#   streamlit
#   pandas
#   numpy
#   openpyxl
#   matplotlib
#
# Ejecuta:  streamlit run app.py

from __future__ import annotations
from io import BytesIO
from typing import List, Optional, Tuple

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.drawing.image import Image as XLImage

# ─────────────────────────────────────────────────────────────────────────────
# Parámetros
# ─────────────────────────────────────────────────────────────────────────────
CATEGORIAS_OBJETIVO = ["Delitos", "Riesgos sociales", "Otros factores"]
YELLOW_HEXES = {"FFFF00", "FFEB9C", "FFF2CC", "FFE699", "FFD966", "FFF4CC"}

# ─────────────────────────────────────────────────────────────────────────────
# Utilidades
# ─────────────────────────────────────────────────────────────────────────────

def _strip(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def detectar_columnas_amarillas(
    xlsx_file: BytesIO,
    sheet_name: Optional[str],
    header_row: int = 1,
    limit_from_col: Optional[str] = None,
    limit_to_col: Optional[str] = None,
) -> List[str]:
    """Regresa los encabezados pintados de AMARILLO en la fila header_row.
    Si se indican columnas Desde/Hasta (letras Excel), se limita a ese rango.
    """
    xlsx_file.seek(0)
    wb = load_workbook(filename=xlsx_file, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    start_col_idx = 1
    end_col_idx = ws.max_column
    if limit_from_col and limit_to_col:
        start_col_idx = column_index_from_string(limit_from_col)
        end_col_idx = column_index_from_string(limit_to_col)

    headers = []
    for col_idx in range(start_col_idx, end_col_idx + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        is_yellow = False
        fill = cell.fill
        if fill and getattr(fill, "patternType", None):
            fg = getattr(fill, "fgColor", None)
            rgb = getattr(fg, "rgb", None) if fg is not None else None
            if rgb:
                rgb = rgb.upper().replace("#", "")
                if len(rgb) == 8:
                    rgb = rgb[2:]  # ARGB → RGB
                if rgb in YELLOW_HEXES:
                    is_yellow = True
        if is_yellow:
            header_text = str(cell.value) if cell.value is not None else get_column_letter(col_idx)
            if _strip(header_text):
                headers.append(header_text)
    return headers


def cargar_dataframe(xlsx_file: BytesIO, sheet_name: Optional[str], header_row: int) -> pd.DataFrame:
    xlsx_file.seek(0)
    df = pd.read_excel(xlsx_file, sheet_name=sheet_name, header=header_row - 1, engine="openpyxl")
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def leer_catalogo_descriptores(
    catalogo_file: BytesIO,
    nombre_hoja: Optional[str],
    col_desc: str,
    col_cat: str,
) -> pd.DataFrame:
    """Lee el catálogo con al menos columnas: DESCRIPTOR y CATEGORIA.
    Devuelve un DF normalizado con esas dos columnas (str.strip()).
    """
    catalogo_file.seek(0)
    dfc = pd.read_excel(catalogo_file, sheet_name=nombre_hoja, engine="openpyxl")
    # Normaliza nombres
    cols = {c: c.strip().upper() for c in dfc.columns}
    dfc.columns = [cols[c] for c in dfc.columns]
    cdesc = col_desc.strip().upper()
    ccat = col_cat.strip().upper()
    if cdesc not in dfc.columns or ccat not in dfc.columns:
        raise ValueError(
            f"El catálogo debe tener columnas '{col_desc}' y '{col_cat}'. Encontradas: {list(dfc.columns)}"
        )
    out = dfc[[cdesc, ccat]].copy()
    out.columns = ["DESCRIPTOR", "CATEGORIA"]
    out["DESCRIPTOR"] = out["DESCRIPTOR"].astype(str).str.strip()
    out["CATEGORIA"] = out["CATEGORIA"].astype(str).str.strip()
    out = out.dropna(how="any")
    return out


def filtrar_por_categorias(
    headers: List[str],
    catalogo_df: Optional[pd.DataFrame],
    categorias_seleccionadas: List[str],
) -> List[str]:
    """Filtra una lista de headers usando un catálogo (DESCRIPTOR→CATEGORIA).
    Si no hay catálogo, regresa headers tal cual.
    """
    if catalogo_df is None or catalogo_df.empty:
        return headers
    mapa = {r.DESCRIPTOR: r.CATEGORIA for r in catalogo_df.itertuples(index=False)}
    keep = []
    cats = set([_strip(x).lower() for x in categorias_seleccionadas])
    for h in headers:
        cat = mapa.get(h)
        if cat and _strip(cat).lower() in cats:
            keep.append(h)
    return keep


def contar_frecuencias(df: pd.DataFrame, columnas: List[str]) -> pd.DataFrame:
    """Cuenta menciones (no vacío y ≠ 0) por columna/descriptor."""
    cols_ok = [c for c in columnas if c in df.columns]
    freqs = []
    for c in cols_ok:
        s = df[c]
        mask = ~s.isna()
        mask &= s.astype(str).str.strip().ne("")
        num = pd.to_numeric(s, errors="coerce")
        mask &= ~((~num.isna()) & (num == 0))
        freqs.append({"DESCRIPTOR": c, "frecuencia": int(mask.sum())})
    out = pd.DataFrame(freqs)
    out = out[out["frecuencia"] > 0].sort_values("frecuencia", ascending=False).reset_index(drop=True)
    return out


def construir_pareto(df_freqs: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    total = int(df_freqs["frecuencia"].sum())
    if total == 0:
        return (
            pd.DataFrame(columns=["#", "DESCRIPTOR", "frecuencia", "%", "acumul", "acumul%", "dentro_80"]),
            0,
        )
    df = df_freqs.copy()
    df["%"] = (df["frecuencia"] / total) * 100.0
    df["acumul"] = df["frecuencia"].cumsum()
    df["acumul%"] = df["%"].cumsum()
    df["#"] = np.arange(1, len(df) + 1)
    df["dentro_80"] = df["acumul%"] <= 80.0
    df = df[["#", "DESCRIPTOR", "frecuencia", "%", "acumul", "acumul%", "dentro_80"]]
    return df, total


def graficar_pareto(df_pareto: pd.DataFrame, titulo: str) -> plt.Figure:
    if df_pareto.empty:
        fig, ax = plt.subplots(figsize=(14, 6))
        ax.text(0.5, 0.5, "Sin datos", ha="center", va="center", fontsize=16)
        ax.axis("off")
        return fig

    x = np.arange(len(df_pareto))
    frec = df_pareto["frecuencia"].values
    acum_pct = df_pareto["acumul%"].values

    fig, ax1 = plt.subplots(figsize=(18, 7), dpi=130)
    ax1.bar(x, frec, label="Frecuencia", color="#4E79A7")
    ax1.set_ylabel("Frecuencia")
    ax1.set_xticks(x)
    ax1.set_xticklabels(df_pareto["DESCRIPTOR"].values, rotation=65, ha="right")
    ax1.grid(axis="y", alpha=0.25)

    ax2 = ax1.twinx()
    ax2.plot(x, acum_pct, marker="o", linewidth=2.2, label="Acumulado", color="#F28E2B")
    ax2.set_ylabel("Porcentaje acumulado")
    ax2.set_ylim(0, 105)
    ax2.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f"{v:.0f}%"))

    ax2.axhline(80, color="#707070", linestyle="--", linewidth=1.8, label="80/20")
    idx80 = int(np.argmax(acum_pct >= 80)) if (acum_pct >= 80).any() else len(acum_pct) - 1
    ax1.axvline(idx80, color="#D62728", linestyle="-", linewidth=1.6)

    fig.suptitle(titulo, fontsize=16, y=0.98)
    lines, labels = [], []
    for ax in (ax1, ax2):
        L = ax.get_legend_handles_labels()
        lines += L[0]
        labels += L[1]
    ax1.legend(lines, labels, loc="upper right")
    fig.tight_layout(rect=[0, 0.03, 1, 0.95])
    return fig


def formatear_para_mostrar(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    out["%"] = out["%"].map(lambda v: f"{v:,.2f}%")
    out["acumul%"] = out["acumul%"].map(lambda v: f"{v:,.2f}%")
    out["dentro_80"] = out["dentro_80"].map(lambda b: "Sí" if bool(b) else "No")
    return out


def exportar_excel_con_grafico(df_pareto: pd.DataFrame, fig: plt.Figure, hoja="PARETO") -> BytesIO:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_pareto.to_excel(writer, sheet_name=hoja, index=False)
        wb = writer.book
        ws = wb[hoja]
        img_io = BytesIO()
        fig.savefig(img_io, format="png", dpi=220, bbox_inches="tight")
        img_io.seek(0)
        img = XLImage(img_io)
        ws.add_image(img, f"B{len(df_pareto) + 4}")
        writer._save()
    bio.seek(0)
    return bio

# ─────────────────────────────────────────────────────────────────────────────
# UI — Streamlit
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Generador de Pareto – Comunidad", layout="wide")
st.title("Generador de Pareto – Comunidad")

with st.sidebar:
    st.header("⚙️ Configuración")
    comunidad = st.text_input("Nombre de la comunidad (para el título)", value="DESAMPARADOS NORTE")
    archivo = st.file_uploader("Matriz (.xlsx)", type=["xlsx"])
    hoja = st.text_input("Hoja (vacío = activa)", value="")
    header_row = st.number_input("Fila del encabezado", 1, 999, 1)

    st.markdown("**Rango de columnas (letras Excel)**")
    c1, c2 = st.columns(2)
    with c1:
        from_col = st.text_input("Desde", value="AI")
    with c2:
        to_col = st.text_input("Hasta", value="ET")
    usar_rango = st.checkbox("Limitar a este rango", value=True)
    detectar_amarillas = st.checkbox("Detectar encabezados AMARILLOS", value=True)

    st.divider()
    st.subheader("Catálogo de descriptores (opcional)")
    st.caption("Para filtrar por categorías exactas: usa tu archivo 'DESCRIPTORES ACTUALIZADOS 2024 v2.xlsx'")
    catalogo = st.file_uploader("Catálogo .xlsx", type=["xlsx"], help="Debe incluir columnas DESCRIPTOR y CATEGORIA (o cambia nombres abajo)")
    hoja_catalogo = st.text_input("Hoja del catálogo (opcional)", value="")
    col_desc = st.text_input("Columna descriptor", value="DESCRIPTOR")
    col_cat = st.text_input("Columna categoría", value="CATEGORIA")
    cats = st.multiselect("Categorías a incluir", options=CATEGORIAS_OBJETIVO + ["(todas)"] , default=CATEGORIAS_OBJETIVO)
    incluir_categorias = st.checkbox("Forzar SOLO categorías seleccionadas (si hay catálogo)", value=True)

    procesar = st.button("Procesar Pareto", type="primary")

if not archivo or not procesar:
    st.info("Sube tu matriz, deja AI–ET y AMARILLOS activos, y pulsa **Procesar Pareto**. Si agregas catálogo, se filtrarán solo **Delitos / Riesgos sociales / Otros factores** por defecto.")
    st.stop()

sheetname = hoja.strip() or None
rango_desde = from_col.strip().upper() if (usar_rango and from_col.strip()) else None
rango_hasta = to_col.strip().upper() if (usar_rango and to_col.strip()) else None

# 1) Detectar amarillas (pre-selección)
try:
    pre_cols = detectar_columnas_amarillas(
        archivo, sheetname, header_row=header_row,
        limit_from_col=rango_desde, limit_to_col=rango_hasta,
    ) if detectar_amarillas else []
except Exception as e:
    st.error(f"Error detectando encabezados amarillos: {e}")
    pre_cols = []

# 2) Cargar la matriz completa
try:
    df = cargar_dataframe(archivo, sheetname, header_row)
except Exception as e:
    st.error(f"No se pudo leer el Excel: {e}")
    st.stop()

# 3) Encabezados candidatos según rango
if usar_rango and rango_desde and rango_hasta:
    try:
        archivo.seek(0)
        wb = load_workbook(filename=archivo, data_only=True)
        ws = wb[sheetname] if sheetname else wb.active
        idx_from = column_index_from_string(rango_desde)
        idx_to = column_index_from_string(rango_hasta)
        headers_rango = []
        for col_idx in range(idx_from, idx_to + 1):
            v = ws.cell(row=header_row, column=col_idx).value
            if v is not None and str(v).strip() != "":
                headers_rango.append(str(v))
        candidatos = [h for h in headers_rango if h in df.columns]
    except Exception:
        candidatos = list(df.columns)
else:
    candidatos = list(df.columns)

# 4) Leer catálogo (si hay) y filtrar por categorías objetivo (pre-cargado)
catalogo_df = None
if catalogo is not None:
    try:
        catalogo_df = leer_catalogo_descriptores(catalogo, hoja_catalogo or None, col_desc, col_cat)
    except Exception as e:
        st.warning(f"No se pudo leer el catálogo: {e}")

if incluir_categorias and catalogo_df is not None:
    categorias_elegidas = [c for c in cats if c != "(todas)"] or CATEGORIAS_OBJETIVO
    candidatos = filtrar_por_categorias(candidatos, catalogo_df, categorias_elegidas)
    pre_cols = filtrar_por_categorias(pre_cols, catalogo_df, categorias_elegidas)

# 5) Selector manual (con preselección)
st.subheader("Columnas a considerar (Delitos / Riesgos sociales / Otros factores)")
seleccion = st.multiselect(
    "Encabezados a contar",
    options=candidatos,
    default=(pre_cols if pre_cols else candidatos),
)
if not seleccion:
    st.warning("Debes seleccionar al menos una columna.")
    st.stop()

# 6) Conteo y Pareto
freqs = contar_frecuencias(df, seleccion)
if freqs.empty:
    st.warning("No hay menciones en las columnas seleccionadas.")
    st.stop()

df_pareto, total = construir_pareto(freqs)

# 7) KPIs
st.markdown(
    f"**Filas:** {len(df)}  •  **Columnas:** {len(seleccion)}  •  **Total (Σ frecuencia):** {total}"
)

# 8) Tabla
st.subheader("Tabla del Pareto")
st.dataframe(formatear_para_mostrar(df_pareto), use_container_width=True)

# 9) Gráfico
titulo = f"PARETO COMUNIDAD {(_strip(comunidad).upper() or 'SIN NOMBRE')}"
fig = graficar_pareto(df_pareto, titulo)
st.pyplot(fig, use_container_width=True)

# 10) Descargas
c1, c2 = st.columns(2)
with c1:
    png_io = BytesIO()
    fig.savefig(png_io, format="png", dpi=220, bbox_inches="tight")
    png_io.seek(0)
    st.download_button("⬇️ PNG del gráfico", data=png_io, file_name=f"pareto_{_strip(comunidad).replace(' ', '_').lower()}.png", mime="image/png")
with c2:
    xio = exportar_excel_con_grafico(df_pareto, fig, hoja="PARETO")
    st.download_button("⬇️ Excel (tabla + gráfico)", data=xio, file_name=f"pareto_{_strip(comunidad).replace(' ', '_').lower()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")






