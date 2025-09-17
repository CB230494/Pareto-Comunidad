# =========================
# Pareto Comunidad – MSP
# (Lee Plantilla 'matriz' + Diccionario de Códigos → Copilado + Pareto + Excel con gráfico)
# =========================

import re
import unicodedata
from io import BytesIO
from typing import Dict, List, Optional, Tuple

import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Pareto Comunidad – MSP", layout="wide", initial_sidebar_state="collapsed")

TOP_N_GRAFICO = 40

# -------------------------
# Utils de normalización
# -------------------------
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

def norm_cols(cols: List[str]) -> List[str]:
    return [norm_text(c) for c in cols]

def to_numeric_or_str(x):
    """Intenta convertir a número, si no, vuelve string normalizado."""
    try:
        if pd.isna(x):
            return np.nan
    except Exception:
        pass
    # Primero probamos como número
    try:
        # aceptar ints floats (pero 1.0 -> 1)
        n = pd.to_numeric(x)
        if isinstance(n, (np.floating, float)):
            if n.is_integer():
                return int(n)
            return float(n)
        return int(n)
    except Exception:
        # Dejar como texto normalizado (pero sin perder mayúsculas del diccionario que se usa en display)
        return norm_text(str(x))

def split_multi(value: str) -> List[str]:
    """
    Divide valores compuestos en una celda: "1,3;5/6|7" -> ["1","3","5","6","7"]
    """
    if value is None:
        return []
    s = str(value)
    if s.strip() == "":
        return []
    parts = re.split(r"[;,/|]+", s)
    return [p.strip() for p in parts if p.strip() != ""]

# -------------------------
# Lectura de archivos
# -------------------------
@st.cache_data(show_spinner=False)
def read_matriz(file) -> pd.DataFrame:
    return pd.read_excel(file, sheet_name="matriz", engine="openpyxl")

@st.cache_data(show_spinner=False)
def read_mapping(file) -> pd.DataFrame:
    """
    Lee el diccionario/mapeo desde la primera hoja que tenga las columnas requeridas.
    Columnas esperadas (nombres flexibles):
      - Columna  -> nombre de la columna en 'matriz' (alias: campo, variable)
      - Código   -> valor a buscar (alias: codigo, valor)
      - Descriptor
      - Categoría (alias: categoria)
    Opcional: Nombre Corto
    """
    xls = pd.ExcelFile(file)
    best = None
    for sh in xls.sheet_names:
        df = pd.read_excel(file, sheet_name=sh, engine="openpyxl")
        # normalizar encabezados
        original_cols = list(df.columns)
        lower_cols = [norm_text(c) for c in df.columns]
        df.columns = lower_cols

        # detectar columnas
        col_col = None
        for c in ["columna", "campo", "variable"]:
            if c in df.columns:
                col_col = c
                break
        code_col = None
        for c in ["codigo", "código", "valor", "code"]:
            if c in df.columns:
                code_col = c
                break
        desc_col = None
        for c in ["descriptor", "descriptores"]:
            if c in df.columns:
                desc_col = c
                break
        cat_col = None
        for c in ["categoria", "categoría", "category"]:
            if c in df.columns:
                cat_col = c
                break
        nc_col = None
        for c in ["nombre corto", "nombrecorto", "alias"]:
            if c in df.columns:
                nc_col = c
                break

        needed = [col_col, code_col, desc_col]
        if all(v is not None for v in needed):
            # reconstruir con nombres estándar
            out = pd.DataFrame({
                "Columna": df[col_col],
                "Codigo": df[code_col],
                "Descriptor": df[desc_col],
                "Categoria": df[cat_col] if cat_col in df.columns else "",
            })
            if nc_col:
                out["NombreCorto"] = df[nc_col]
            else:
                out["NombreCorto"] = ""
            best = out.dropna(subset=["Columna", "Codigo", "Descriptor"]).copy()
            # limpiar espacios
            for c in ["Columna", "Descriptor", "Categoria", "NombreCorto"]:
                best[c] = best[c].astype(str).str.strip()
            return best

    raise ValueError("No se encontró ninguna hoja en el diccionario con columnas reconocibles (Columna, Código, Descriptor).")

# -------------------------
# Conteo a partir del Mapeo
# -------------------------
def build_header_index(df_matriz: pd.DataFrame) -> Dict[str, str]:
    """
    Retorna un mapa: nombre_normalizado -> nombre_real_columna
    """
    norm2real = {}
    for c in df_matriz.columns:
        norm2real[norm_text(str(c))] = c
    return norm2real

def match_value(cell, code) -> bool:
    """
    Evalúa si la celda 'cell' cumple el 'code' indicado en el diccionario.
    - Soporta: números exactos, texto exacto normalizado, y listas en la celda.
    - Si el código es lista (separada por comas, /, |, ;) se toma como 'cualquiera de'.
    """
    # Preparar lista de posibles códigos desde 'code'
    code_items = []
    if isinstance(code, (int, float)) and not pd.isna(code):
        code_items = [code]
    else:
        # puede llegar como "1" o "1,3,5" o "si"
        txt = str(code)
        for piece in split_multi(txt):
            # cada pieza intentar a num
            try:
                n = pd.to_numeric(piece)
                if float(n).is_integer():
                    code_items.append(int(n))
                else:
                    code_items.append(float(n))
            except Exception:
                code_items.append(norm_text(piece))

    # Revisar celda
    if pd.isna(cell):
        return False

    # Si la celda es numérica, comparamos contra números
    try:
        ncell = pd.to_numeric(cell)
        # exact match con cualquiera
        for ci in code_items:
            if isinstance(ci, (int, float)) and ncell == ci:
                return True
        # si el código es texto pero la celda es número, no cuenta
        return False
    except Exception:
        pass

    # Si la celda es texto, tratamos listas
    text = str(cell)
    parts = split_multi(text)
    if parts:
        # comparar cada parte
        for p in parts:
            # intentar número
            try:
                pn = pd.to_numeric(p)
                for ci in code_items:
                    if isinstance(ci, (int, float)) and pn == ci:
                        return True
            except Exception:
                # comparar texto normalizado
                pn = norm_text(p)
                for ci in code_items:
                    if not isinstance(ci, (int, float)) and pn == ci:
                        return True
        return False
    else:
        # texto plano
        nt = norm_text(text)
        for ci in code_items:
            if not isinstance(ci, (int, float)) and nt == ci:
                return True
        return False

def compute_hits_from_mapping(df_matriz: pd.DataFrame, map_df: pd.DataFrame) -> pd.DataFrame:
    """
    Recorre el diccionario: para cada fila (Columna, Codigo, Descriptor, Categoria)
    busca coincidencias en df_matriz y cuenta filas donde se cumple.
    Devuelve una tabla agregada por Descriptor (suma de frecuencias).
    """
    # Índice de columnas disponibles (normalizadas)
    header_index = build_header_index(df_matriz)

    # Normalizar mapeo (columna de destino, y preparar código)
    rows = []
    for _, r in map_df.iterrows():
        col_key_norm = norm_text(r["Columna"])
        if col_key_norm not in header_index:
            # columna no existe en matriz; la saltamos
            continue
        col_real = header_index[col_key_norm]
        code = r["Codigo"]
        desc = str(r["Descriptor"]).strip()
        cat  = str(r["Categoria"]).strip() if "Categoria" in map_df.columns else ""
        rows.append((col_real, code, desc, cat))

    if not rows:
        return pd.DataFrame({"Descriptor": [], "Frecuencia": []})

    # Conteo eficiente por columna
    counts: Dict[str, int] = {}
    for col_real, code, desc, cat in rows:
        serie = df_matriz[col_real]
        # Más eficiente: aplicar función booleana vectorizada con .apply (celda por celda)
        mask = serie.apply(lambda x: match_value(x, code))
        freq = int(mask.sum())
        if freq > 0:
            counts[desc] = counts.get(desc, 0) + freq

    # Construir Copilado
    if not counts:
        return pd.DataFrame({"Descriptor": [], "Frecuencia": []})
    copilado = pd.DataFrame(
        [(d, f) for d, f in counts.items()],
        columns=["Descriptor", "Frecuencia"]
    ).sort_values(["Frecuencia", "Descriptor"], ascending=[False, True], ignore_index=True)
    return copilado

# -------------------------
# Pareto
# -------------------------
def make_pareto(copilado_df: pd.DataFrame, map_df: pd.DataFrame) -> pd.DataFrame:
    if copilado_df.empty:
        return pd.DataFrame(columns=["Categoría","Descriptor","Frecuencia","Porcentaje","% Acumulado","Acumulado","80/20"])

    # Crear mapa Descriptor -> Categoria usando el diccionario
    # Si un descriptor aparece con múltiples categorías en el diccionario, preferimos la más frecuente (última ocurrencia).
    cat_map = {}
    for _, r in map_df.iterrows():
        d = str(r["Descriptor"]).strip()
        c = str(r["Categoria"]).strip() if "Categoria" in map_df.columns else ""
        if d:
            cat_map[d] = c

    df = copilado_df.copy()
    df["Categoría"] = df["Descriptor"].map(cat_map).fillna("")
    total = df["Frecuencia"].sum()
    df["Porcentaje"] = (df["Frecuencia"] / total) * 100.0
    df = df.sort_values(["Frecuencia", "Descriptor"], ascending=[False, True], ignore_index=True)
    df["% Acumulado"] = df["Porcentaje"].cumsum()
    df["Acumulado"] = df["Frecuencia"].cumsum()
    df["80/20"] = np.where(df["% Acumulado"] <= 80.0, "≤80%", ">80%")
    return df[["Categoría","Descriptor","Frecuencia","Porcentaje","% Acumulado","Acumulado","80/20"]]

# -------------------------
# Exportar Excel con gráfico
# -------------------------
def export_excel(copilado: pd.DataFrame, pareto: pd.DataFrame) -> bytes:
    from pandas import ExcelWriter
    import xlsxwriter  # noqa: F401
    out = BytesIO()
    with ExcelWriter(out, engine="xlsxwriter") as writer:
        copilado.to_excel(writer, index=False, sheet_name="Copilado Comunidad")
        pareto.to_excel(writer, index=False, sheet_name="Pareto Comunidad")

        wb  = writer.book
        wsP = writer.sheets["Pareto Comunidad"]
        n = len(pareto)
        if n >= 1:
            # (A) Categoría, (B) Descriptor, (C) Frecuencia, (D) Porcentaje, (E) % Acumulado, (F) Acumulado, (G) 80/20
            chart = wb.add_chart({'type': 'column'})
            chart.add_series({
                'name':       'Frecuencia',
                'categories': ['Pareto Comunidad', 1, 1, n, 1],  # B
                'values':     ['Pareto Comunidad', 1, 2, n, 2],  # C
            })
            line = wb.add_chart({'type': 'line'})
            line.add_series({
                'name':       '% Acumulado',
                'categories': ['Pareto Comunidad', 1, 1, n, 1],  # B
                'values':     ['Pareto Comunidad', 1, 4, n, 4],  # E
                'y2_axis':    True,
            })
            chart.combine(line)
            chart.set_title({'name': 'Pareto Comunidad'})
            chart.set_x_axis({'name': 'Descriptor'})
            chart.set_y_axis({'name': 'Frecuencia'})
            chart.set_y2_axis({'name': '% Acumulado', 'min': 0, 'max': 100})
            wsP.insert_chart(1, 9, chart, {'x_scale': 1.2, 'y_scale': 1.2})
    return out.getvalue()

# -------------------------
# UI
# -------------------------
st.title("Pareto Comunidad (MSP) – con Diccionario de Códigos")
st.caption("1) Sube la **Plantilla** (hoja `matriz`). 2) Sube el **Diccionario/Mapeo** (códigos → descriptor/categoría). 3) Procesa y descarga.")

col1, col2 = st.columns([1,1])
with col1:
    plantilla_file = st.file_uploader("📄 Plantilla de Comunidad (XLSX, hoja `matriz`)", type=["xlsx"], key="plantilla")
with col2:
    mapping_file = st.file_uploader("🧭 Diccionario / Mapeo (XLSX)", type=["xlsx"], key="mapeo")

demo = st.checkbox("Probar DEMO sin subir archivos", value=False, help="Genera datos de prueba para comprobar el flujo.")

st.divider()

if demo:
    # DEMO: dataframe simple con columnas y códigos
    st.info("DEMO activa: usando datos de ejemplo.")
    df_demo = pd.DataFrame({
        "problema_seguridad": [1, 2, 2, 1, 3, 2, 1, 1, np.nan, 3, 2],
        "consumo_drogas":     [0, 1, 0, 1, 1, 0, 1, 0, np.nan, 1, 0],
        "dano_propiedad":     ["1,3", "3", "", "1", "0", "1/3", "1", None, "3", "", "3"],
    })
    map_demo = pd.DataFrame({
        "Columna":   ["problema_seguridad","problema_seguridad","problema_seguridad","consumo_drogas","dano_propiedad","dano_propiedad"],
        "Codigo":    [1, 2, 3, 1, 1, 3],
        "Descriptor":["ASALTO","ROBO","HURTO","CONSUMO DE DROGAS","DAÑOS A LA PROPIEDAD","VANDALISMO"],
        "Categoria": ["DELITOS CONTRA LA PROPIEDAD","DELITOS CONTRA LA PROPIEDAD","DELITOS CONTRA LA PROPIEDAD","DROGAS","DELITOS CONTRA LA PROPIEDAD","DELITOS CONTRA LA PROPIEDAD"]
    })
    st.write("**Vista previa Plantilla (DEMO)**")
    st.dataframe(df_demo.head(15), use_container_width=True)
    st.write("**Vista previa Diccionario (DEMO)**")
    st.dataframe(map_demo, use_container_width=True)

    copilado = compute_hits_from_mapping(df_demo, map_demo)
    st.subheader("Copilado Comunidad (DEMO)")
    st.dataframe(copilado, use_container_width=True)

    pareto = make_pareto(copilado, map_demo)
    st.subheader("Pareto Comunidad (DEMO)")
    st.dataframe(pareto, use_container_width=True)

    st.subheader("Gráficos (DEMO)")
    plot_df = pareto.head(TOP_N_GRAFICO).copy()
    st.bar_chart(plot_df.set_index("Descriptor")["Frecuencia"])
    st.line_chart(plot_df.set_index("Descriptor")["% Acumulado"])

    st.subheader("Descargar")
    st.download_button(
        "⬇️ Excel DEMO (Copilado + Pareto + Gráfico)",
        data=export_excel(copilado, pareto),
        file_name="Pareto_Comunidad_DEMO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.stop()

# ---- Flujo real con archivos ----
if not plantilla_file or not mapping_file:
    st.info("Sube **ambos archivos** para procesar.")
    st.stop()

try:
    with st.spinner("Leyendo Plantilla (hoja 'matriz')…"):
        df_matriz = read_matriz(plantilla_file)
    st.caption("Vista previa Plantilla (primeras filas)")
    st.dataframe(df_matriz.head(20), use_container_width=True)
except Exception as e:
    st.error(f"Error al leer Plantilla: {e}")
    st.stop()

try:
    with st.spinner("Leyendo Diccionario / Mapeo…"):
        df_map = read_mapping(mapping_file)
    st.caption("Vista previa Diccionario detectado")
    st.dataframe(df_map.head(30), use_container_width=True)
except Exception as e:
    st.error(f"Error al leer Diccionario/Mapeo: {e}")
    st.stop()

with st.spinner("Procesando (conteo por códigos)…"):
    copilado_df = compute_hits_from_mapping(df_matriz, df_map)

if copilado_df.empty:
    st.warning("No se detectaron coincidencias usando el diccionario. Revisa que los nombres de columna y códigos del mapeo correspondan exactamente a la hoja 'matriz'.")
    st.stop()

st.subheader("Copilado Comunidad")
st.dataframe(copilado_df, use_container_width=True)

pareto_df = make_pareto(copilado_df, df_map)
st.subheader("Pareto Comunidad")
st.dataframe(pareto_df, use_container_width=True)

st.subheader("Gráficos")
plot_df = pareto_df.head(TOP_N_GRAFICO).copy()
st.bar_chart(plot_df.set_index("Descriptor")["Frecuencia"])
st.line_chart(plot_df.set_index("Descriptor")["% Acumulado"])

st.subheader("Descargar Excel")
st.download_button(
    "⬇️ Copilado + Pareto + Gráfico",
    data=export_excel(copilado_df, pareto_df),
    file_name="Pareto_Comunidad.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

