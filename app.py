# app.py — Generador de Pareto AUTOMÁTICO (solo subir matriz)
# Requisitos: streamlit, openpyxl, pandas, numpy, matplotlib
# Ejecuta: streamlit run app.py

from __future__ import annotations
from io import BytesIO
from typing import List, Tuple, Optional

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.drawing.image import Image as XLImage

# --------------------- PARÁMETROS POR DEFECTO ---------------------
TITULO_COMUNIDAD_DEF = "DESAMPARADOS NORTE"
RANGO_DEF_DESDE = "AI"     # rango preferido (como tus plantillas)
RANGO_DEF_HASTA = "ET"
PROBAR_FILAS_HEADER = 8    # buscar encabezado en las primeras N filas
COLOR_AMARILLOS = {"FFFF00", "FFEB9C", "FFF2CC", "FFE699", "FFD966", "FFF4CC"}

# --------------------- UTILIDADES BASE ---------------------
def _s(x):
    if x is None:
        return ""
    if isinstance(x, float) and np.isnan(x):
        return ""
    return str(x).strip()

def _wb_ws(file: BytesIO):
    file.seek(0)
    wb = load_workbook(filename=file, data_only=True)
    return wb, wb.active  # hoja activa por defecto

def _col_idx(letra: str) -> int:
    return column_index_from_string(letra)

def _rango_idx(desde: str, hasta: str) -> Tuple[int, int]:
    return _col_idx(desde), _col_idx(hasta)

def detectar_fila_encabezado(ws, idx_from: int, idx_to: int, probar_n: int) -> int:
    """Devuelve la fila (1-based) con más celdas no vacías dentro del rango."""
    best_row, best_cnt = 1, -1
    top = min(probar_n, ws.max_row)
    for r in range(1, top + 1):
        cnt = 0
        for c in range(idx_from, idx_to + 1):
            if _s(ws.cell(r, c).value) != "":
                cnt += 1
        if cnt > best_cnt:
            best_row, best_cnt = r, cnt
    return best_row

def headers_en_rango(ws, header_row: int, idx_from: int, idx_to: int) -> List[str]:
    headers = []
    for c in range(idx_from, idx_to + 1):
        v = _s(ws.cell(header_row, c).value)
        if v != "":
            headers.append(v)
    return headers

def headers_amarillos(ws, header_row: int, idx_from: int, idx_to: int) -> List[str]:
    amarillos = []
    for c in range(idx_from, idx_to + 1):
        cell = ws.cell(header_row, c)
        fill = cell.fill
        is_yellow = False
        if fill and getattr(fill, "patternType", None):
            fg = getattr(fill, "fgColor", None)
            rgb = getattr(fg, "rgb", None) if fg is not None else None
            if rgb:
                rgb = rgb.upper().replace("#", "")
                if len(rgb) == 8:  # ARGB -> RGB
                    rgb = rgb[2:]
                if rgb in COLOR_AMARILLOS:
                    is_yellow = True
        if is_yellow:
            v = _s(cell.value) or get_column_letter(c)
            if v:
                amarillos.append(v)
    return amarillos

def construir_mapa_columna(ws, header_row: int) -> dict:
    """Devuelve {nombre_columna_normalizado -> índice_columna} para TODA la hoja."""
    mapa = {}
    for c in range(1, ws.max_column + 1):
        v = _s(ws.cell(header_row, c).value)
        if v != "":
            mapa[v] = c
    return mapa

def contar_frecuencias_por_posicion(ws, fila_inicio: int, columnas_idx: List[int]) -> pd.DataFrame:
    """
    Cuenta menciones por columna usando ÍNDICES de columna (robusto ante encabezados raros).
    Se considera mención si la celda NO está vacía y NO es 0 (numérico).
    """
    freqs = []
    max_r = ws.max_row
    for c in columnas_idx:
        count = 0
        for r in range(fila_inicio, max_r + 1):
            val = ws.cell(r, c).value
            if val is None:
                continue
            s = _s(val)
            if s == "":
                continue
            # descartar cero explícito
            try:
                num = float(s.replace(",", ".")) if isinstance(s, str) else float(val)
                if num == 0:
                    continue
            except Exception:
                pass
            count += 1
        nombre = _s(ws.cell(fila_inicio - 1, c).value)  # encabezado original
        if nombre == "":
            nombre = get_column_letter(c)
        freqs.append({"DESCRIPTOR": nombre, "frecuencia": int(count)})
    out = pd.DataFrame(freqs, columns=["DESCRIPTOR", "frecuencia"])
    out = out[out["frecuencia"] > 0].sort_values("frecuencia", ascending=False).reset_index(drop=True)
    return out

def construir_pareto(df_freqs: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    total = int(df_freqs["frecuencia"].sum()) if not df_freqs.empty else 0
    if total == 0:
        return pd.DataFrame(columns=["#", "DESCRIPTOR", "frecuencia", "%", "acumul", "acumul%", "dentro_80"]), 0
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
        lines += L[0]; labels += L[1]
    ax1.legend(lines, labels, loc="upper right")
    fig.tight_layout(rect=[0, 0.03, 1, 0.95])
    return fig

def exportar_excel_con_grafico(df_pareto: pd.DataFrame, fig: plt.Figure) -> BytesIO:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        df_pareto.to_excel(writer, sheet_name="PARETO", index=False)
        wb = writer.book
        ws = wb["PARETO"]
        img_io = BytesIO()
        fig.savefig(img_io, format="png", dpi=220, bbox_inches="tight")
        img_io.seek(0)
        ws.add_image(XLImage(img_io), f"B{len(df_pareto) + 4}")
        writer._save()
    bio.seek(0)
    return bio

# --------------------- UI ---------------------
st.set_page_config(page_title="Generador de Pareto – Comunidad", layout="wide")
st.title("Generador de Pareto – Comunidad")

comunidad = st.text_input("Nombre de la comunidad", value=TITULO_COMUNIDAD_DEF)
archivo = st.file_uploader("Sube la MATRIZ (.xlsx)", type=["xlsx"])

# Opciones avanzadas (por si tu rango no es AI–ET)
with st.expander("Opciones avanzadas (si tu plantilla usa otro rango/hoja)"):
    hoja_manual = st.text_input("Nombre de la hoja (vacío = activa)", value="")
    rango_desde = st.text_input("Rango desde (letra Excel)", value=RANGO_DEF_DESDE).strip().upper()
    rango_hasta = st.text_input("Rango hasta (letra Excel)", value=RANGO_DEF_HASTA).strip().upper()

if not archivo:
    st.info("Sube tu matriz. El sistema detecta la fila de encabezados y cuenta por posición.")
    st.stop()

# --------------------- LECTURA Y DETECCIÓN ROBUSTA ---------------------
try:
    archivo.seek(0)
    wb = load_workbook(filename=archivo, data_only=True)
    ws = wb[hoja_manual] if hoja_manual.strip() in wb.sheetnames else wb.active
except Exception as e:
    st.error(f"No se pudo abrir el Excel: {e}")
    st.stop()

# 1) Rango preferido (AI–ET, editable) → si no da, se expande a todo el ancho de la hoja
try:
    idx_from, idx_to = _rango_idx(rango_desde or RANGO_DEF_DESDE, rango_hasta or RANGO_DEF_HASTA)
except Exception:
    st.error("Rango inválido. Usa letras de columnas (p. ej., AI y ET).")
    st.stop()

header_row = detectar_fila_encabezado(ws, idx_from, idx_to, PROBAR_FILAS_HEADER)
amarillos = headers_amarillos(ws, header_row, idx_from, idx_to)
if not amarillos:
    # Fallback 1: todos los encabezados no vacíos del rango preferido
    amarillos = headers_en_rango(ws, header_row, idx_from, idx_to)

# Si aun así no hay encabezados útiles, expandir al ancho completo de la hoja
if not amarillos:
    idx_from, idx_to = 1, ws.max_column
    header_row = detectar_fila_encabezado(ws, idx_from, idx_to, PROBAR_FILAS_HEADER)
    amarillos = headers_amarillos(ws, header_row, idx_from, idx_to) or headers_en_rango(ws, header_row, idx_from, idx_to)

if not amarillos:
    st.error("No se ubicaron encabezados en la fila detectada. Revisa que tu plantilla tenga títulos en una sola fila.")
    st.stop()

# 2) Mapear nombres → índices reales (porque contamos por posición)
mapa = construir_mapa_columna(ws, header_row)
cols_idx = [mapa[n] for n in amarillos if n in mapa]

if not cols_idx:
    st.error("Los encabezados detectados no coinciden con columnas reales. Verifica celdas combinadas o espacios extraños.")
    st.stop()

# 3) Conteo de menciones (no vacío y ≠ 0) desde la fila siguiente al encabezado
freqs = contar_frecuencias_por_posicion(ws, header_row + 1, cols_idx)
if freqs.empty:
    st.error("No se encontraron datos en las columnas detectadas (todas vacías o 0).")
    st.stop()

# --------------------- PARETO + DESCARGAS ---------------------
df_pareto, total = construir_pareto(freqs)
st.markdown(f"**Filas analizadas:** {ws.max_row - header_row}  •  **Columnas consideradas:** {len(cols_idx)}  •  **Total (Σ frecuencia):** {total}")

# Mostrar con formato %
mostrar = df_pareto.copy()
mostrar["%"] = mostrar["%"].map(lambda v: f"{v:,.2f}%")
mostrar["acumul%"] = mostrar["acumul%"].map(lambda v: f"{v:,.2f}%")
mostrar["dentro_80"] = mostrar["dentro_80"].map(lambda b: "Sí" if bool(b) else "No")
st.subheader("Tabla de Pareto")
st.dataframe(mostrar, use_container_width=True)

fig = graficar_pareto(df_pareto, f"PARETO COMUNIDAD {(_s(comunidad).upper() or 'SIN NOMBRE')}")
st.pyplot(fig, use_container_width=True)

c1, c2 = st.columns(2)
with c1:
    png_io = BytesIO(); fig.savefig(png_io, format="png", dpi=220, bbox_inches="tight"); png_io.seek(0)
    st.download_button("⬇️ Descargar gráfico (PNG)", data=png_io,
                       file_name=f"pareto_{_s(comunidad).replace(' ','_').lower()}.png",
                       mime="image/png")
with c2:
    xio = exportar_excel_con_grafico(df_pareto, fig)
    st.download_button("⬇️ Descargar Excel (tabla + gráfico)", data=xio,
                       file_name=f"pareto_{_s(comunidad).replace(' ','_').lower()}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
