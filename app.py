# app.py — Pareto AUTOMÁTICO (solo subir matriz)
# ─────────────────────────────────────────────────────────────────────────────
# ✔ Sube TU matriz (.xlsx) y listo → se genera el Pareto automáticamente.
# ✔ No requiere catálogo ni selección manual.
# ✔ Lee la hoja ACTIVA, encabezado en fila 1.
# ✔ Toma las columnas de encabezado AMARILLO en el rango AI–ET (config interno).
# ✔ Si no detecta amarillos, usa TODOS los encabezados NO vacíos del rango AI–ET.
# ✔ Cuenta TODAS las menciones (no vacío y ≠ 0) por descriptor (encabezado).
# ✔ Calcula % y % acumulado (cierra en el total) y grafica barras + línea 80%.
# ✔ Descargas: PNG y Excel (tabla + gráfico pegado).
#
# Requisitos: streamlit, pandas, numpy, openpyxl, matplotlib

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
# Parámetros fijos que pediste
# ─────────────────────────────────────────────────────────────────────────────
RANGO_DESDE = "AI"   # letras Excel
RANGO_HASTA = "ET"
HEADER_ROW = 1       # encabezados en fila 1
TITULO_COMUNIDAD_DEF = "DESAMPARADOS NORTE"  # puedes cambiarlo en el panel

# Colores amarillos comunes en Excel
YELLOW_HEXES = {"FFFF00", "FFEB9C", "FFF2CC", "FFE699", "FFD966", "FFF4CC"}

# ─────────────────────────────────────────────────────────────────────────────
# Utilidades
# ─────────────────────────────────────────────────────────────────────────────

def _strip(x):
    if pd.isna(x):
        return ""
    return str(x).strip()


def hoja_activa_nombre(xlsx_file: BytesIO) -> str:
    xlsx_file.seek(0)
    wb = load_workbook(filename=xlsx_file, data_only=True)
    return wb.active.title


def leer_df_hoja(xlsx_file: BytesIO, sheet_name: Optional[str]) -> pd.DataFrame:
    """Lee una ÚNICA hoja como DataFrame (evita dict)."""
    xlsx_file.seek(0)
    hoja = sheet_name or hoja_activa_nombre(xlsx_file)
    xlsx_file.seek(0)
    df = pd.read_excel(xlsx_file, sheet_name=hoja, header=HEADER_ROW - 1, engine="openpyxl")
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def headers_en_rango(xlsx_file: BytesIO, sheet_name: Optional[str], desde: str, hasta: str) -> List[str]:
    """Devuelve TODOS los encabezados no vacíos encontrados en el rango (por posición)."""
    xlsx_file.seek(0)
    wb = load_workbook(filename=xlsx_file, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    idx_from = column_index_from_string(desde)
    idx_to = column_index_from_string(hasta)
    out = []
    for c in range(idx_from, idx_to + 1):
        v = ws.cell(row=HEADER_ROW, column=c).value
        if v is not None and str(v).strip() != "":
            out.append(str(v))
    return out


def headers_amarillos_en_rango(xlsx_file: BytesIO, sheet_name: Optional[str], desde: str, hasta: str) -> List[str]:
    xlsx_file.seek(0)
    wb = load_workbook(filename=xlsx_file, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    idx_from = column_index_from_string(desde)
    idx_to = column_index_from_string(hasta)
    amarillos = []
    for c in range(idx_from, idx_to + 1):
        cell = ws.cell(row=HEADER_ROW, column=c)
        fill = cell.fill
        is_yellow = False
        if fill and getattr(fill, "patternType", None):
            fg = getattr(fill, "fgColor", None)
            rgb = getattr(fg, "rgb", None) if fg is not None else None
            if rgb:
                rgb = rgb.upper().replace("#", "")
                if len(rgb) == 8:  # ARGB
                    rgb = rgb[2:]
                if rgb in YELLOW_HEXES:
                    is_yellow = True
        if is_yellow:
            header_text = str(cell.value) if cell.value is not None else get_column_letter(c)
            if _strip(header_text):
                amarillos.append(header_text)
    return amarillos


def contar_frecuencias(df: pd.DataFrame, columnas: List[str]) -> pd.DataFrame:
    """Cuenta menciones (no vacío y ≠ 0) por columna/descriptor.
    Devuelve SIEMPRE un DF con columnas ["DESCRIPTOR", "frecuencia"].
    """
    cols = [c for c in columnas if c in df.columns]
    if not cols:
        return pd.DataFrame(columns=["DESCRIPTOR", "frecuencia"])  # vacío seguro

    freqs = []
    for c in cols:
        s = df[c]
        mask = ~s.isna()
        mask &= s.astype(str).str.strip().ne("")
        num = pd.to_numeric(s, errors="coerce")
        mask &= ~((~num.isna()) & (num == 0))
        freqs.append({"DESCRIPTOR": c, "frecuencia": int(mask.sum())})

    out = pd.DataFrame(freqs, columns=["DESCRIPTOR", "frecuencia"])  # garantiza columnas
    if out.empty:
        return pd.DataFrame(columns=["DESCRIPTOR", "frecuencia"])  # vacío seguro

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


def exportar_excel_con_grafico(df_pareto: pd.DataFrame, fig: plt.Figure, hoja="PARETO") -> BytesIO:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_pareto.to_excel(writer, sheet_name=hoja, index=False)
        wb = writer.book; ws = wb[hoja]
        img_io = BytesIO()
        fig.savefig(img_io, format="png", dpi=220, bbox_inches="tight")
        img_io.seek(0)
        img = XLImage(img_io)
        ws.add_image(img, f"B{len(df_pareto) + 4}")
        writer._save()
    bio.seek(0)
    return bio

# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Generador de Pareto – Comunidad", layout="wide")
st.title("Generador de Pareto – Comunidad")

comunidad = st.text_input("Nombre de la comunidad (título)", value=TITULO_COMUNIDAD_DEF)
archivo = st.file_uploader("Sube la MATRIZ (.xlsx) — solo esto", type=["xlsx"])

if not archivo:
    st.info("Sube tu matriz y el Pareto se genera automáticamente con los encabezados AMARILLOS del rango AI–ET.")
    st.stop()

# 1) Leer DF único
try:
    df = leer_df_hoja(archivo, sheet_name=None)
except Exception as e:
    st.error(f"No se pudo leer la matriz: {e}")
    st.stop()

# 2) Encabezados
amarillos = []
try:
    amarillos = headers_amarillos_en_rango(archivo, sheet_name=None, desde=RANGO_DESDE, hasta=RANGO_HASTA)
except Exception:
    amarillos = []

if not amarillos:
    # Fallback: todos los encabezados no vacíos del rango
    try:
        amarillos = headers_en_rango(archivo, sheet_name=None, desde=RANGO_DESDE, hasta=RANGO_HASTA)
    except Exception:
        st.error("No se pudieron determinar los encabezados en el rango AI–ET.")
        st.stop()

# 3) Conteo y Pareto
freqs = contar_frecuencias(df, amarillos)
if freqs.empty:
    # Mensaje útil: mostrar por qué quedó vacío
    faltantes = [h for h in amarillos if h not in df.columns]
    if faltantes:
        st.error("No se encontraron estas columnas en la hoja (revisa el rango AI–ET o el nombre exacto del encabezado):\n- " + "\n- ".join(faltantes))
    else:
        st.warning("No hay menciones (todas las celdas vacías o 0) en las columnas detectadas del rango AI–ET.")
    st.stop()

df_pareto, total = construir_pareto(freqs)

# 4) KPIs
st.markdown(f"**Filas:** {len(df)} • **Columnas (detectadas):** {len(amarillos)} • **Total (Σ frecuencia):** {total}")

# 5) Tabla y gráfico
st.subheader("Tabla de Pareto")
out = df_pareto.copy()
out["%"] = out["%"].map(lambda v: f"{v:,.2f}%")
out["acumul%"] = out["acumul%"].map(lambda v: f"{v:,.2f}%")
out["dentro_80"] = out["dentro_80"].map(lambda b: "Sí" if bool(b) else "No")
st.dataframe(out, use_container_width=True)

fig = graficar_pareto(df_pareto, f"PARETO COMUNIDAD {(_strip(comunidad).upper() or 'SIN NOMBRE')}")
st.pyplot(fig, use_container_width=True)

# 6) Descargas
c1, c2 = st.columns(2)
with c1:
    buf = BytesIO(); fig.savefig(buf, format="png", dpi=220, bbox_inches="tight"); buf.seek(0)
    st.download_button("⬇️ PNG del gráfico", data=buf, file_name=f"pareto_{_strip(comunidad).replace(' ', '_').lower()}.png", mime="image/png")
with c2:
    xio = exportar_excel_con_grafico(df_pareto, fig, hoja="PARETO")
    st.download_button("⬇️ Excel (tabla + gráfico)", data=xio, file_name=f"pareto_{_strip(comunidad).replace(' ', '_').lower()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

