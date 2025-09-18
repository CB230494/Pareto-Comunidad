# app.py — Generador de Pareto (delitos, riesgos sociales y otros factores)
# ─────────────────────────────────────────────────────────────────────────────
# ✓ Lee tu "matriz" (xlsx) tal y como la usas hoy (115+ columnas, 500+ filas)
# ✓ Detecta automáticamente las columnas con encabezado AMARILLO (entre un rango)
# ✓ También puedes indicar un rango de columnas (por letras, ej. AI–ET)
# ✓ Cuenta TODAS las menciones (no vacíos, no "0") por descriptor
# ✓ Calcula % y % acumulado (el total del acumulado es el total de datos)
# ✓ Dibuja el Pareto igual al ejemplo (barras + línea acumulada + 80/20)
# ✓ Permite descargar: PNG del gráfico y Excel con la tabla + el gráfico embebido
#
# Requisitos (requirements.txt sugerido):
#   streamlit
#   pandas
#   openpyxl
#   matplotlib
#   numpy
#
# Uso rápido en Streamlit Cloud/Local:
#   streamlit run app.py
#
# Notas:
# - Si tu encabezado no está en la fila 1, cámbialo en el panel lateral.
# - Por defecto se detectan encabezados AMARILLOS típicos (FFEB9C, FFF2CC, FFD966, FFFF00, etc.).
# - El título del Pareto se arma con "PARETO COMUNIDAD <nombre>".
# - El rango de columnas (por letras) es opcional; se usa para limitar la búsqueda.

from __future__ import annotations
import io
from io import BytesIO
from typing import List, Tuple, Optional

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.drawing.image import Image as XLImage

# ─────────────────────────────────────────────────────────────────────────────
# Utilidades
# ─────────────────────────────────────────────────────────────────────────────

YELLOW_HEXES = {
    "FFFF00",  # amarillo puro
    "FFEB9C",  # amarillo pastel (Excel Theme)
    "FFF2CC",  # amarillo claro (Excel Theme)
    "FFE699",
    "FFD966",
    "FFF4CC",
}


def _strip(s):
    if pd.isna(s):
        return ""
    return str(s).strip()


def detectar_columnas_amarillas(
    xlsx_file: BytesIO,
    sheet_name: Optional[str],
    header_row: int = 1,
    limit_from_col: Optional[str] = None,
    limit_to_col: Optional[str] = None,
) -> List[str]:
    """Devuelve una lista de NOMBRES DE COLUMNA (texto de encabezado)
    que están pintadas de amarillo en la fila de encabezado.

    - Si se dan limit_from_col y limit_to_col (letras Excel, ej. AI y ET),
      se restringe la búsqueda a ese rango.
    """
    xlsx_file.seek(0)
    wb = load_workbook(filename=xlsx_file, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    start_col_idx = 1
    end_col_idx = ws.max_column
    if limit_from_col and limit_to_col:
        start_col_idx = column_index_from_string(limit_from_col)
        end_col_idx = column_index_from_string(limit_to_col)

    headers_yellow: List[str] = []
    for col_idx in range(start_col_idx, end_col_idx + 1):
        cell = ws.cell(row=header_row, column=col_idx)
        fill = cell.fill
        is_yellow = False
        if fill and fill.patternType and fill.fgColor is not None:
            rgb = None
            # openpyxl puede traer .rgb o .indexed/theme; priorizamos rgb
            if hasattr(fill.fgColor, "rgb") and fill.fgColor.rgb:
                rgb = fill.fgColor.rgb
            elif hasattr(fill.fgColor, "indexed") and fill.fgColor.indexed is not None:
                # no fiable para color exacto; lo ignoramos
                rgb = None
            if rgb:
                rgb = rgb.upper().replace("0x", "").replace("#", "")
                # Excel guarda ARGB (8 chars). Conservamos los últimos 6 (RGB)
                if len(rgb) == 8:
                    rgb = rgb[2:]
                if rgb in YELLOW_HEXES:
                    is_yellow = True
        if is_yellow:
            headers_yellow.append(str(cell.value) if cell.value is not None else get_column_letter(col_idx))

    return [h for h in headers_yellow if _strip(h) != ""]


def cargar_dataframe(
    xlsx_file: BytesIO,
    sheet_name: Optional[str],
    header_row: int = 1,
) -> pd.DataFrame:
    """Carga toda la hoja como DataFrame usando pandas, tomando header_row como encabezado."""
    xlsx_file.seek(0)
    df = pd.read_excel(xlsx_file, sheet_name=sheet_name, header=header_row - 1, engine="openpyxl")
    # Elimina filas totalmente vacías
    df = df.dropna(how="all").reset_index(drop=True)
    return df


def contar_frecuencias(df: pd.DataFrame, columnas: List[str]) -> pd.DataFrame:
    """Cuenta las menciones por columna. Considera como mención todo valor NO vacío y diferente de 0.
    Devuelve un DataFrame con columnas: DESCRIPTOR, frecuencia.
    """
    cols_ok = [c for c in columnas if c in df.columns]
    if not cols_ok:
        return pd.DataFrame(columns=["DESCRIPTOR", "frecuencia"])  # vacío

    # Conteo por columna: no vacíos y distintos de 0
    freqs = {}
    for c in cols_ok:
        serie = df[c]
        # True si hay valor: no NaN, no "", no 0 (numérico o texto "0")
        mask = ~serie.isna()
        mask &= serie.astype(str).str.strip().ne("")
        # Quitar ceros explícitos
        try:
            # intenta convertir a numérico; los no numéricos quedan NaN
            num = pd.to_numeric(serie, errors="coerce")
            mask &= ~((~num.isna()) & (num == 0))
        except Exception:
            pass
        freqs[c] = int(mask.sum())

    out = (
        pd.DataFrame({"DESCRIPTOR": list(freqs.keys()), "frecuencia": list(freqs.values())})
        .query("frecuencia > 0")
        .sort_values("frecuencia", ascending=False)
        .reset_index(drop=True)
    )
    return out


def tabla_pareto(df_freqs: pd.DataFrame) -> pd.DataFrame:
    """Calcula % y acumulados para Pareto."""
    total = int(df_freqs["frecuencia"].sum())
    if total == 0:
        # estructura vacía amigable
        return pd.DataFrame(
            columns=["#", "DESCRIPTOR", "frecuencia", "%", "acumul", "acumul%", "dentro_80"]
        )

    df = df_freqs.copy()
    df["%"] = (df["frecuencia"] / total) * 100.0
    df["acumul"] = df["frecuencia"].cumsum()
    df["acumul%"] = df["%"]].cumsum()
    df["#"] = np.arange(1, len(df) + 1)
    df["dentro_80"] = df["acumul%"] <= 80.0
    # Reordenamos columnas para que coincidan con el ejemplo
    df = df[["#", "DESCRIPTOR", "frecuencia", "%", "acumul", "acumul%", "dentro_80"]]
    return df, total


def graficar_pareto(df_pareto: pd.DataFrame, titulo: str = "PARETO") -> plt.Figure:
    """Devuelve la figura Matplotlib del Pareto (barras + acumulado + 80/20)."""
    if df_pareto.empty:
        fig, ax = plt.subplots(figsize=(14, 6))
        ax.text(0.5, 0.5, "Sin datos", ha="center", va="center", fontsize=16)
        ax.axis("off")
        return fig

    x = np.arange(len(df_pareto))
    frec = df_pareto["frecuencia"].values
    acum_pct = df_pareto["acumul%"].values

    fig, ax1 = plt.subplots(figsize=(18, 7), dpi=130)
    bars = ax1.bar(x, frec, label="Frecuencia", color="#4E79A7")
    ax1.set_ylabel("Frecuencia")
    ax1.set_xticks(x)
    ax1.set_xticklabels(df_pareto["DESCRIPTOR"].values, rotation=65, ha="right")
    ax1.grid(axis="y", alpha=0.25)

    # Eje 2: % acumulado
    ax2 = ax1.twinx()
    ax2.plot(x, acum_pct, marker="o", linewidth=2.2, label="Acumulado", color="#F28E2B")
    ax2.set_ylabel("Porcentaje acumulado")
    ax2.set_ylim(0, 105)
    ax2.yaxis.set_major_formatter(FuncFormatter(lambda v, pos: f"{v:.0f}%"))

    # Línea 80/20 y corte vertical
    ax2.axhline(80, color="#707070", linestyle="--", linewidth=1.8, label="80/20")
    # índice donde se cruza el 80%
    idx80 = int(np.argmax(acum_pct >= 80)) if (acum_pct >= 80).any() else len(acum_pct) - 1
    ax1.axvline(idx80, color="#D62728", linestyle="-", linewidth=1.6)

    # Título y leyenda unificada
    fig.suptitle(titulo, fontsize=16, y=0.98)
    lines, labels = [], []
    for ax in (ax1, ax2):
        L = ax.get_legend_handles_labels()
        lines += L[0]
        labels += L[1]
    ax1.legend(lines, labels, loc="upper right")
    fig.tight_layout(rect=[0, 0.03, 1, 0.95])
    return fig


def formatear_tabla_mostrar(df: pd.DataFrame) -> pd.DataFrame:
    """Devuelve una copia con formatos para presentar en pantalla (no cambia los tipos)."""
    if df.empty:
        return df
    out = df.copy()
    out["%"] = out["%"].map(lambda v: f"{v:,.2f}%")
    out["acumul%"] = out["acumul%"].map(lambda v: f"{v:,.2f}%")
    out["dentro_80"] = out["dentro_80"].map(lambda b: "Sí" if bool(b) else "No")
    return out


def exportar_excel_con_grafico(
    df_pareto: pd.DataFrame,
    fig: plt.Figure,
    nombre_hoja: str = "PARETO",
) -> BytesIO:
    """Genera un Excel con la tabla y pega el gráfico como imagen en la misma hoja."""
    from pandas import ExcelWriter

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        # Guardar la tabla (sin formateo de strings, para que Excel pueda calcular)
        df_x = df_pareto.copy()
        df_x.to_excel(writer, sheet_name=nombre_hoja, index=False)
        wb = writer.book
        ws = wb[nombre_hoja]
        # Añadir imagen del gráfico
        img_io = BytesIO()
        fig.savefig(img_io, format="png", dpi=220, bbox_inches="tight")
        img_io.seek(0)
        img = XLImage(img_io)
        # Colocar el gráfico debajo de la tabla (fila = len(df) + 4)
        fila_img = len(df_x) + 4
        celda = f"B{fila_img}"
        ws.add_image(img, celda)
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

    archivo = st.file_uploader("Matriz (Excel .xlsx)", type=["xlsx"], accept_multiple_files=False)

    hoja = st.text_input("Nombre de la hoja (opcional; deja vacío para la activa)", value="")
    header_row = st.number_input("Fila del encabezado", min_value=1, value=1, step=1)

    st.markdown("**Rango de columnas (opcional, por letras Excel)**")
    col1, col2 = st.columns(2)
    with col1:
        from_col = st.text_input("Desde", value="AI")
    with col2:
        to_col = st.text_input("Hasta", value="ET")

    usar_rango = st.checkbox("Limitar a este rango (recomendado)", value=True)

    detectar_amarillas = st.checkbox("Detectar encabezados AMARILLOS y preseleccionar", value=True)

    st.markdown("— *Si no quieres la detección por color, desmárcalo y elige manualmente luego.*")

    procesar = st.button("Procesar Pareto", type="primary")

if archivo and procesar:
    # 1) Detectar columnas amarillas (opcional)
    try:
        sheetname = hoja if hoja.strip() else None
        rango_desde = from_col.strip().upper() if (usar_rango and from_col.strip()) else None
        rango_hasta = to_col.strip().upper() if (usar_rango and to_col.strip()) else None

        preselect_cols: List[str] = []
        if detectar_amarillas:
            preselect_cols = detectar_columnas_amarillas(
                archivo, sheetname, header_row=header_row,
                limit_from_col=rango_desde, limit_to_col=rango_hasta,
            )
    except Exception as e:
        st.error(f"Error detectando encabezados amarillos: {e}")
        preselect_cols = []

    # 2) Cargar DataFrame completo
    try:
        df = cargar_dataframe(archivo, sheetname, header_row)
    except Exception as e:
        st.error(f"No se pudo leer el Excel: {e}")
        st.stop()

    # 3) Determinar las columnas candidatas para conteo
    if usar_rango and from_col.strip() and to_col.strip():
        try:
            # Restringimos las columnas del DataFrame por posición (Excel-like)
            archivo.seek(0)
            wb = load_workbook(filename=archivo, data_only=True)
            ws = wb[sheetname] if sheetname else wb.active
            idx_from = column_index_from_string(from_col.strip().upper())
            idx_to = column_index_from_string(to_col.strip().upper())
            # Sacamos los encabezados reales en ese rango
            headers_rango = []
            for col_idx in range(idx_from, idx_to + 1):
                val = ws.cell(row=header_row, column=col_idx).value
                if val is not None and str(val).strip() != "":
                    headers_rango.append(str(val))
            candidatos = [h for h in headers_rango if h in df.columns]
        except Exception:
            candidatos = list(df.columns)
    else:
        candidatos = list(df.columns)

    # 4) Selector manual para el usuario (preselecciona amarillas si hay)
    st.subheader("Selección de columnas a considerar (delitos, riesgos sociales y otros factores)")
    seleccion = st.multiselect(
        "Elige las columnas (encabezados) a contar como descriptores",
        options=candidatos,
        default=[c for c in preselect_cols if c in candidatos] or candidatos,
        help="Se cuentan todas las celdas NO vacías y distintas de 0 por cada columna."
    )

    if not seleccion:
        st.warning("Debes seleccionar al menos una columna.")
        st.stop()

    # 5) Conteo y tabla Pareto
    df_freqs = contar_frecuencias(df, seleccion)
    if df_freqs.empty:
        st.warning("No hay menciones en las columnas seleccionadas.")
        st.stop()

    df_pareto, total = tabla_pareto(df_freqs)

    # 6) Mostrar KPIs
    st.markdown(
        f"**Total de filas analizadas:** {len(df)}  •  **Columnas consideradas:** {len(seleccion)}  •  **Total de datos (∑ frecuencia):** {total}"
    )

    # 7) Tabla formateada
    st.subheader("Tabla de Pareto")
    st.dataframe(formatear_tabla_mostrar(df_pareto), use_container_width=True)

    # 8) Gráfico
    titulo = f"PARETO COMUNIDAD {comunidad.strip().upper()}"
    fig = graficar_pareto(df_pareto, titulo)
    st.pyplot(fig, use_container_width=True)

    # 9) Descargas
    colA, colB = st.columns(2)
    with colA:
        # PNG
        buf_png = BytesIO()
        fig.savefig(buf_png, format="png", dpi=220, bbox_inches="tight")
        buf_png.seek(0)
        st.download_button(
            "⬇️ Descargar gráfico (PNG)", data=buf_png, file_name=f"pareto_{comunidad.replace(' ', '_').lower()}.png",
            mime="image/png"
        )
    with colB:
        # Excel con tabla + gráfico
        excel_io = exportar_excel_con_grafico(df_pareto, fig, nombre_hoja="PARETO")
        st.download_button(
            "⬇️ Descargar Excel (tabla + gráfico)", data=excel_io,
            file_name=f"pareto_{comunidad.replace(' ', '_').lower()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

else:
    st.info(
        "Sube tu matriz y presiona **Procesar Pareto**. Por defecto detecto los **encabezados amarillos** (AI–ET) y cuento todas las menciones."
    )






