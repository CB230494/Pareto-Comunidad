# =========================
# Pareto Comunidad – MSP (Plantilla + Diccionario, fix cache serializable)
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

# =========================
# Utils
# =========================
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

def split_multi(value: str) -> List[str]:
    """Divide '1,3;5/6|7' -> ['1','3','5','6','7']"""
    if value is None:
        return []
    s = str(value)
    if s.strip() == "":
        return []
    parts = re.split(r"[;,/|]+", s)
    return [p.strip() for p in parts if p.strip() != ""]

def build_header_index(df_matriz: pd.DataFrame) -> Dict[str, str]:
    """nombre_normalizado -> nombre_real_columna"""
    return {norm_text(str(c)): c for c in df_matriz.columns}

def match_value(cell, code) -> bool:
    """Compara celda vs código (número o texto). Soporta celdas con múltiples valores."""
    # preparar lista de códigos candidatos
    code_items = []
    if isinstance(code, (int, float)) and not pd.isna(code):
        code_items = [code]
    else:
        for piece in split_multi(str(code)):
            try:
                n = pd.to_numeric(piece)
                code_items.append(int(n) if float(n).is_integer() else float(n))
            except Exception:
                code_items.append(norm_text(piece))

    if pd.isna(cell):
        return False

    # numérico puro
    try:
        ncell = pd.to_numeric(cell)
        return any((isinstance(ci, (int, float)) and ncell == ci) for ci in code_items)
    except Exception:
        pass

    # texto (posible lista)
    parts = split_multi(str(cell))
    if parts:
        for p in parts:
            # comparar número
            try:
                pn = pd.to_numeric(p)
                if any((isinstance(ci, (int, float)) and pn == ci) for ci in code_items):
                    return True
            except Exception:
                # comparar texto normalizado
                pn = norm_text(p)
                if any((not isinstance(ci, (int, float)) and pn == ci) for ci in code_items):
                    return True
        return False
    # texto simple
    nt = norm_text(str(cell))
    return any((not isinstance(ci, (int, float)) and nt == ci) for ci in code_items)

# =========================
# Lectura (cacheando SOLO datos serializables)
# =========================
@st.cache_data(show_spinner=False)
def read_matriz_bytes(file_bytes: bytes) -> pd.DataFrame:
    # Usa TODAS las filas de 'matriz'
    return pd.read_excel(BytesIO(file_bytes), sheet_name="matriz", engine="openpyxl")

@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes) -> List[str]:
    # Devuelvo SOLO la lista de hojas (serializable)
    xls = pd.ExcelFile(BytesIO(file_bytes))
    return xls.sheet_names

@st.cache_data(show_spinner=False)
def read_mapping_sheet(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    # Leo una hoja específica del diccionario desde bytes
    return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, engine="openpyxl")

def autodetect_mapping(df: pd.DataFrame) -> Optional[Dict[str, str]]:
    """
    Intenta detectar automáticamente las columnas del diccionario:
    retorna {'columna': <name>, 'codigo': <name>, 'descriptor': <name>, 'categoria': <name or None>}
    o None si no pudo.
    """
    cols = [c for c in df.columns]
    norm_cols = [norm_text(c) for c in cols]
    col_map = dict(zip(norm_cols, cols))

    def pick(cands):
        for k in cands:
            if k in col_map:
                return col_map[k]
        return None

    col_col = pick(["columna", "campo", "variable"])
    code_col = pick(["codigo", "código", "valor", "code"])
    desc_col = pick(["descriptor", "descriptores", "descripcion", "descripción"])
    cat_col  = pick(["categoria", "categoría", "category"])

    if col_col and code_col and desc_col:
        return {"columna": col_col, "codigo": code_col, "descriptor": desc_col, "categoria": cat_col}
    return None

# =========================
# Cómputo
# =========================
def compute_hits_from_mapping(df_matriz: pd.DataFrame, df_map: pd.DataFrame,
                              col_col: str, code_col: str, desc_col: str, cat_col: Optional[str]) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Recorre df_map y cuenta coincidencias en df_matriz.
    Devuelve (copilado, pareto)
    """
    header_index = build_header_index(df_matriz)

    # Normalizamos los nombres de columna destino
    rows = []
    for _, r in df_map.iterrows():
        col_key_norm = norm_text(r[col_col])
        if col_key_norm not in header_index:
            continue
        col_real = header_index[col_key_norm]
        code = r[code_col]
        desc = str(r[desc_col]).strip()
        cat  = str(r[cat_col]).strip() if cat_col and (cat_col in df_map.columns) else ""
        if desc == "":
            continue
        rows.append((col_real, code, desc, cat))

    counts: Dict[str, int] = {}
    cat_map: Dict[str, str] = {}
    for col_real, code, desc, cat in rows:
        serie = df_matriz[col_real]
        mask = serie.apply(lambda x: match_value(x, code))
        freq = int(mask.sum())
        if freq > 0:
            counts[desc] = counts.get(desc, 0) + freq
            if desc not in cat_map:
                cat_map[desc] = cat

    # Copilado
    if not counts:
        copilado = pd.DataFrame({"Descriptor": [], "Frecuencia": []})
    else:
        copilado = (pd.DataFrame(list(counts.items()), columns=["Descriptor", "Frecuencia"])
                    .sort_values(["Frecuencia","Descriptor"], ascending=[False, True], ignore_index=True))

    # Pareto
    if copilado.empty:
        pareto = pd.DataFrame(columns=["Categoría","Descriptor","Frecuencia","Porcentaje","% Acumulado","Acumulado","80/20"])
    else:
        dfp = copilado.copy()
        dfp["Categoría"] = dfp["Descriptor"].map(cat_map).fillna("")
        total = dfp["Frecuencia"].sum()
        dfp["Porcentaje"] = (dfp["Frecuencia"] / total) * 100.0
        dfp = dfp.sort_values(["Frecuencia","Descriptor"], ascending=[False, True], ignore_index=True)
        dfp["% Acumulado"] = dfp["Porcentaje"].cumsum()
        dfp["Acumulado"] = dfp["Frecuencia"].cumsum()
        dfp["80/20"] = np.where(dfp["% Acumulado"] <= 80.0, "≤80%", ">80%")
        pareto = dfp[["Categoría","Descriptor","Frecuencia","Porcentaje","% Acumulado","Acumulado","80/20"]]

    return copilado, pareto

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
            wsP.insert_chart(1, 9, chart, {'x_scale': 1.2, 'y_scale': 1.2})
    return out.getvalue()

# =========================
# UI
# =========================
st.title("Pareto Comunidad (MSP) – Plantilla + Diccionario de Códigos")
st.caption("La app lee **todas** las filas de `matriz`. Si el diccionario no usa encabezados esperados, puedes elegirlos manualmente.")

col1, col2 = st.columns([1,1])
with col1:
    plantilla_file = st.file_uploader("📄 Plantilla de Comunidad (XLSX, hoja `matriz`)", type=["xlsx"], key="plantilla")
with col2:
    mapping_file = st.file_uploader("🧭 Diccionario / Mapeo (XLSX)", type=["xlsx"], key="mapeo")

if not plantilla_file or not mapping_file:
    st.info("Sube **ambos archivos** para continuar.")
    st.stop()

# Convertimos ambos a BYTES para cachear de forma segura y reabrir múltiples veces
plantilla_bytes = plantilla_file.getvalue()
mapping_bytes   = mapping_file.getvalue()

# --- Plantilla ---
try:
    with st.spinner("Leyendo Plantilla (hoja 'matriz')…"):
        df_matriz = read_matriz_bytes(plantilla_bytes)  # TODAS las filas
    st.caption(f"Vista previa Plantilla (primeras 20 de {len(df_matriz)} filas)")
    st.dataframe(df_matriz.head(20), use_container_width=True)
except Exception as e:
    st.error(f"Error al leer Plantilla: {e}")
    st.stop()

# --- Diccionario: obtener hojas y permitir seleccionar ---
sheet_names = get_sheet_names(mapping_bytes)
sel_sheet = st.selectbox("Hoja del diccionario", options=sheet_names, index=0)

df_map_raw = read_mapping_sheet(mapping_bytes, sel_sheet)
st.caption("Vista previa Diccionario (primeras 30 filas)")
st.dataframe(df_map_raw.head(30), use_container_width=True)

auto = autodetect_mapping(df_map_raw)

st.markdown("### Selección de columnas del diccionario")
c1, c2, c3, c4 = st.columns(4)
with c1:
    col_col = st.selectbox("Columna (en 'matriz')", options=list(df_map_raw.columns),
                           index=(list(df_map_raw.columns).index(auto["columna"]) if auto else 0))
with c2:
    code_col = st.selectbox("Código/Valor", options=list(df_map_raw.columns),
                            index=(list(df_map_raw.columns).index(auto["codigo"]) if auto else 1))
with c3:
    desc_col = st.selectbox("Descriptor", options=list(df_map_raw.columns),
                            index=(list(df_map_raw.columns).index(auto["descriptor"]) if auto else 2))
with c4:
    cat_col  = st.selectbox("Categoría (opcional)", options=["(ninguna)"] + list(df_map_raw.columns),
                            index=((["(ninguna)"] + list(df_map_raw.columns)).index(auto["categoria"])
                                   if (auto and auto["categoria"] in df_map_raw.columns) else 0))
cat_col_real = None if cat_col == "(ninguna)" else cat_col

st.divider()
go_btn = st.button("Procesar")

if not go_btn:
    st.stop()

with st.spinner("Procesando (conteo por códigos sobre todas las filas)…"):
    copilado_df, pareto_df = compute_hits_from_mapping(df_matriz, df_map_raw, col_col, code_col, desc_col, cat_col_real)

if copilado_df.empty:
    st.warning("No se detectaron coincidencias. Verifica que 'Columna' coincida con un encabezado de `matriz` y que 'Código' corresponda a los valores reales.")
    st.stop()

st.subheader("Copilado Comunidad")
st.dataframe(copilado_df, use_container_width=True)

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



