# ============================================================================
# ============================== PARTE 1/10 =================================
# ============ Encabezado, imports, config y estilos Matplotlib =============
# ============================================================================

# app.py — Pareto 80/20 + Portafolio + Unificado + Sheets + Informe PDF (con PÍLDORA)
# ---------------------------------------------------------------------------------
# Ejecuta con:
#   streamlit run app.py
# ---------------------------------------------------------------------------------

import io
from textwrap import wrap
from typing import List, Dict, Tuple

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# ====== Google Sheets (DB) ======
import gspread
from google.oauth2.service_account import Credentials

# ====== PDF (ReportLab/Platypus) ======
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame,
    Paragraph, Spacer, Image as RLImage, Table, TableStyle,
    PageBreak, NextPageTemplate
)
from reportlab.platypus.flowables import KeepTogether
from datetime import datetime

# ----------------- CONFIG (actualizado) -----------------
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1XZjXQfLb5Jiptp_BXuCfg9QZNEz6ZWh9hbtp0rRAGpM/edit?usp=sharing"
WS_PARETOS = "paretos"  # si tu pestaña se llama distinto, cámbiala aquí

st.set_page_config(page_title="Pareto de Descriptores", layout="wide")

# Paleta (verde/azul)
VERDE = "#1B9E77"
AZUL  = "#2C7FB8"
TEXTO = "#124559"
GRIS  = "#6B7280"

# ============== AJUSTES MATPLOTLIB (legibilidad) ==============
plt.rcParams.update({
    "figure.dpi": 180,
    "savefig.dpi": 180,
    "axes.titlesize": 18,
    "axes.labelsize": 13,
    "xtick.labelsize": 10,
    "ytick.labelsize": 11,
    "axes.grid": True,
    "grid.alpha": 0.25,
})

# ============================================================================
# ============================== PARTE 2/10 =================================
# ========================= Catálogo embebido (CSV) =========================
# ============================================================================

CATALOGO: List[Dict[str, str]] = [
    {"categoria": "Delito", "descriptor": "Abandono de personas (menor de edad, adulto mayor o con capacidades diferentes)"},
    {"categoria": "Delito", "descriptor": "Abigeato (robo y destace de ganado)"},
    {"categoria": "Delito", "descriptor": "Aborto"},
    {"categoria": "Delito", "descriptor": "Abuso de autoridad"},
    {"categoria": "Riesgo social", "descriptor": "Accidentes de tránsito"},
    {"categoria": "Delito", "descriptor": "Accionamiento de arma de fuego (balaceras)"},
    {"categoria": "Riesgo social", "descriptor": "Acoso escolar (bullying)"},
    {"categoria": "Riesgo social", "descriptor": "Acoso laboral (mobbing)"},
    {"categoria": "Riesgo social", "descriptor": "Acoso sexual callejero"},
    {"categoria": "Riesgo social", "descriptor": "Actos obscenos en vía pública"},
    {"categoria": "Delito", "descriptor": "Administración fraudulenta, apropiaciones indebidas o enriquecimiento ilícito"},
    {"categoria": "Delito", "descriptor": "Agresión con armas"},
    {"categoria": "Riesgo social", "descriptor": "Agrupaciones delincuenciales no organizadas"},
    {"categoria": "Delito", "descriptor": "Alteración de datos y sabotaje informático"},
    {"categoria": "Otros factores", "descriptor": "Ambiente laboral inadecuado"},
    {"categoria": "Delito", "descriptor": "Amenazas"},
    {"categoria": "Riesgo social", "descriptor": "Analfabetismo"},
    {"categoria": "Riesgo social", "descriptor": "Bajos salarios"},
    {"categoria": "Riesgo social", "descriptor": "Barras de fútbol"},
    {"categoria": "Riesgo social", "descriptor": "Búnker (eje de expendio de drogas)"},
    {"categoria": "Delito", "descriptor": "Calumnia"},
    {"categoria": "Delito", "descriptor": "Caza ilegal"},
    {"categoria": "Delito", "descriptor": "Conducción temeraria"},
    {"categoria": "Riesgo social", "descriptor": "Consumo de alcohol en vía pública"},
    {"categoria": "Riesgo social", "descriptor": "Consumo de drogas"},
    {"categoria": "Riesgo social", "descriptor": "Contaminación sónica"},
    {"categoria": "Delito", "descriptor": "Contrabando"},
    {"categoria": "Delito", "descriptor": "Corrupción"},
    {"categoria": "Delito", "descriptor": "Corrupción policial"},
    {"categoria": "Delito", "descriptor": "Cultivo de droga (marihuana)"},
    {"categoria": "Delito", "descriptor": "Daño ambiental"},
    {"categoria": "Delito", "descriptor": "Daños/vandalismo"},
    {"categoria": "Riesgo social", "descriptor": "Deficiencia en la infraestructura vial"},
    {"categoria": "Otros factores", "descriptor": "Deficiencia en la línea 9-1-1"},
    {"categoria": "Riesgo social", "descriptor": "Deficiencias en el alumbrado público"},
    {"categoria": "Delito", "descriptor": "Delincuencia organizada"},
    {"categoria": "Delito", "descriptor": "Delitos contra el ámbito de intimidad (violación de secretos, correspondencia y comunicaciones electrónicas)"},
    {"categoria": "Delito", "descriptor": "Delitos sexuales"},
    {"categoria": "Riesgo social", "descriptor": "Desaparición de personas"},
    {"categoria": "Riesgo social", "descriptor": "Desarticulación interinstitucional"},
    {"categoria": "Riesgo social", "descriptor": "Desempleo"},
    {"categoria": "Riesgo social", "descriptor": "Desvinculación estudiantil"},
    {"categoria": "Delito", "descriptor": "Desobediencia"},
    {"categoria": "Delito", "descriptor": "Desórdenes en vía pública"},
    {"categoria": "Delito", "descriptor": "Disturbios (riñas)"},
    {"categoria": "Riesgo social", "descriptor": "Enfrentamientos estudiantiles"},
    {"categoria": "Delito", "descriptor": "Estafa o defraudación"},
    {"categoria": "Delito", "descriptor": "Estupro (delitos sexuales contra menor de edad)"},
    {"categoria": "Delito", "descriptor": "Evasión y quebrantamiento de pena"},
    {"categoria": "Delito", "descriptor": "Explosivos"},
    {"categoria": "Delito", "descriptor": "Extorsión"},
    {"categoria": "Delito", "descriptor": "Fabricación, producción o reproducción de pornografía"},
    {"categoria": "Riesgo social", "descriptor": "Facilismo económico"},
    {"categoria": "Delito", "descriptor": "Falsificación de moneda y otros valores"},
    {"categoria": "Riesgo social", "descriptor": "Falta de cámaras de seguridad"},
    {"categoria": "Otros factores", "descriptor": "Falta de capacitación policial"},
    {"categoria": "Riesgo social", "descriptor": "Falta de control a patentes"},
    {"categoria": "Riesgo social", "descriptor": "Falta de control fronterizo"},
    {"categoria": "Riesgo social", "descriptor": "Falta de corresponsabilidad en seguridad"},
    {"categoria": "Riesgo social", "descriptor": "Falta de cultura vial"},
    {"categoria": "Riesgo social", "descriptor": "Falta de cultura y compromiso ciudadano"},
    {"categoria": "Riesgo social", "descriptor": "Falta de educación familiar"},
    {"categoria": "Otros factores", "descriptor": "Falta de incentivos"},
    {"categoria": "Riesgo social", "descriptor": "Falta de inversión social"},
    {"categoria": "Riesgo social", "descriptor": "Falta de legislación de extinción de dominio"},
    {"categoria": "Otros factores", "descriptor": "Falta de personal administrativo"},
    {"categoria": "Otros factores", "descriptor": "Falta de personal policial"},
    {"categoria": "Otros factores", "descriptor": "Falta de policías de tránsito"},
    {"categoria": "Riesgo social", "descriptor": "Falta de políticas públicas en seguridad"},
    {"categoria": "Riesgo social", "descriptor": "Falta de presencia policial"},
    {"categoria": "Riesgo social", "descriptor": "Falta de salubridad pública"},
    {"categoria": "Riesgo social", "descriptor": "Familias disfuncionales"},
    {"categoria": "Delito", "descriptor": "Fraude informático"},
    {"categoria": "Delito", "descriptor": "Grooming"},
    {"categoria": "Riesgo social", "descriptor": "Hacinamiento carcelario"},
    {"categoria": "Riesgo social", "descriptor": "Hacinamiento policial"},
    {"categoria": "Delito", "descriptor": "Homicidio"},
    {"categoria": "Riesgo social", "descriptor": "Hospedajes ilegales (cuarterías)"},
    {"categoria": "Delito", "descriptor": "Hurto"},
    {"categoria": "Otros factores", "descriptor": "Inadecuado uso del recurso policial"},
    {"categoria": "Riesgo social", "descriptor": "Incumplimiento al plan regulador de la municipalidad"},
    {"categoria": "Delito", "descriptor": "Incumplimiento del deber alimentario"},
    {"categoria": "Riesgo social", "descriptor": "Indiferencia social"},
    {"categoria": "Otros factores", "descriptor": "Inefectividad en el servicio de policía"},
    {"categoria": "Riesgo social", "descriptor": "Ineficiencia en la administración de justicia"},
    {"categoria": "Otros factores", "descriptor": "Infraestructura inadecuada"},
    {"categoria": "Riesgo social", "descriptor": "Intolerancia social"},
    {"categoria": "Otros factores", "descriptor": "Irrespeto a la jefatura"},
    {"categoria": "Otros factores", "descriptor": "Irrespeto al subalterno"},
    {"categoria": "Otros factores", "descriptor": "Jornadas laborales extensas"},
    {"categoria": "Delito", "descriptor": "Lavado de activos"},
    {"categoria": "Delito", "descriptor": "Lesiones"},
    {"categoria": "Delito", "descriptor": "Ley de armas y explosivos N° 7530"},
    {"categoria": "Riesgo social", "descriptor": "Ley de control de tabaco (Ley 9028)"},
    {"categoria": "Riesgo social", "descriptor": "Lotes baldíos"},
    {"categoria": "Delito", "descriptor": "Maltrato animal"},
    {"categoria": "Delito", "descriptor": "Narcotráfico"},
    {"categoria": "Riesgo social", "descriptor": "Necesidades básicas insatisfechas"},
    {"categoria": "Riesgo social", "descriptor": "Percepción de inseguridad"},
    {"categoria": "Riesgo social", "descriptor": "Pérdida de espacios públicos"},
    {"categoria": "Riesgo social", "descriptor": "Personas con exceso de tiempo de ocio"},
    {"categoria": "Riesgo social", "descriptor": "Personas en estado migratorio irregular"},
    {"categoria": "Riesgo social", "descriptor": "Personas en situación de calle"},
    {"categoria": "Delito", "descriptor": "Menores en vulnerabilidad"},
    {"categoria": "Delito", "descriptor": "Pesca ilegal"},
    {"categoria": "Delito", "descriptor": "Portación ilegal de armas"},
    {"categoria": "Riesgo social", "descriptor": "Presencia multicultural"},
    {"categoria": "Otros factores", "descriptor": "Presión por resultados operativos"},
    {"categoria": "Delito", "descriptor": "Privación de libertad sin ánimo de lucro"},
    {"categoria": "Riesgo social", "descriptor": "Problemas vecinales"},
    {"categoria": "Delito", "descriptor": "Receptación"},
    {"categoria": "Delito", "descriptor": "Relaciones impropias"},
    {"categoria": "Delito", "descriptor": "Resistencia (irrespeto a la autoridad)"},
    {"categoria": "Delito", "descriptor": "Robo a comercio (intimidación)"},
    {"categoria": "Delito", "descriptor": "Robo a comercio (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo a edificación (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo a personas"},
    {"categoria": "Delito", "descriptor": "Robo a transporte comercial"},
    {"categoria": "Delito", "descriptor": "Robo a vehículos (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo a vivienda (intimidación)"},
    {"categoria": "Delito", "descriptor": "Robo a vivienda (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo de bicicleta"},
    {"categoria": "Delito", "descriptor": "Robo de cultivos"},
    {"categoria": "Delito", "descriptor": "Robo de motocicletas/vehículos (bajonazo)"},
    {"categoria": "Delito", "descriptor": "Robo de vehículos"},
    {"categoria": "Delito", "descriptor": "Secuestro"},
    {"categoria": "Delito", "descriptor": "Simulación de delito"},
    {"categoria": "Riesgo social", "descriptor": "Sistema jurídico desactualizado"},
    {"categoria": "Riesgo social", "descriptor": "Suicidio"},
    {"categoria": "Delito", "descriptor": "Sustracción de una persona menor de edad o incapaz"},
    {"categoria": "Delito", "descriptor": "Tala ilegal"},
    {"categoria": "Riesgo social", "descriptor": "Tendencia social hacia el delito (pautas de crianza violenta)"},
    {"categoria": "Riesgo social", "descriptor": "Tenencia de droga"},
    {"categoria": "Delito", "descriptor": "Tentativa de homicidio"},
    {"categoria": "Delito", "descriptor": "Terrorismo"},
    {"categoria": "Riesgo social", "descriptor": "Trabajo informal"},
    {"categoria": "Delito", "descriptor": "Tráfico de armas"},
    {"categoria": "Delito", "descriptor": "Tráfico de influencias"},
    {"categoria": "Riesgo social", "descriptor": "Transporte informal (Uber, porteadores, piratas)"},
    {"categoria": "Delito", "descriptor": "Trata de personas"},
    {"categoria": "Delito", "descriptor": "Turbación de actos religiosos y profanaciones"},
    {"categoria": "Delito", "descriptor": "Uso ilegal de uniformes, insignias o dispositivos policiales"},
    {"categoria": "Delito", "descriptor": "Usurpación de terrenos (precarios)"},
    {"categoria": "Delito", "descriptor": "Venta de drogas"},
    {"categoria": "Riesgo social", "descriptor": "Ventas informales (ambulantes)"},
    {"categoria": "Riesgo social", "descriptor": "Vigilancia informal"},
    {"categoria": "Delito", "descriptor": "Violación de domicilio"},
    {"categoria": "Delito", "descriptor": "Violación de la custodia de las cosas"},
    {"categoria": "Delito", "descriptor": "Violación de sellos"},
    {"categoria": "Delito", "descriptor": "Violencia de género"},
    {"categoria": "Delito", "descriptor": "Violencia intrafamiliar"},
    {"categoria": "Riesgo social", "descriptor": "Xenofobia"},
    {"categoria": "Riesgo social", "descriptor": "Zonas de prostitución"},
    {"categoria": "Riesgo social", "descriptor": "Zonas vulnerables"},
    {"categoria": "Delito", "descriptor": "Robo a transporte público con intimidación"},
    {"categoria": "Delito", "descriptor": "Robo de cable"},
    {"categoria": "Delito", "descriptor": "Explotación sexual infantil"},
    {"categoria": "Delito", "descriptor": "Explotación laboral infantil"},
    {"categoria": "Delito", "descriptor": "Tráfico ilegal de personas"},
    {"categoria": "Riesgo social", "descriptor": "Bares clandestinos"},
    {"categoria": "Delito", "descriptor": "Robo de combustible"},
    {"categoria": "Delito", "descriptor": "Femicidio"},
    {"categoria": "Delito", "descriptor": "Delitos contra la vida (homicidios, heridos)"},
    {"categoria": "Delito", "descriptor": "Venta y consumo de drogas en vía pública"},
    {"categoria": "Delito", "descriptor": "Asalto (a personas, comercio, vivienda, transporte público)"},
    {"categoria": "Delito", "descriptor": "Robo de ganado y agrícola"},
    {"categoria": "Delito", "descriptor": "Robo de equipo agrícola"},
]
# ============================================================================
# ============================== PARTE 3/10 =================================
# ====================== Utilidades base y cálculo Pareto ====================
# ============================================================================

def _map_descriptor_a_categoria() -> Dict[str, str]:
    df = pd.DataFrame(CATALOGO)
    return dict(zip(df["descriptor"], df["categoria"]))
DESC2CAT = _map_descriptor_a_categoria()

def normalizar_freq_map(freq_map: Dict[str, int]) -> Dict[str, int]:
    out = {}
    for d, v in (freq_map or {}).items():
        try:
            vv = int(pd.to_numeric(v, errors="coerce"))
            if vv > 0:
                out[d] = vv
        except Exception:
            continue
    return out

def df_desde_freq_map(freq_map: Dict[str, int]) -> pd.DataFrame:
    items = []
    for d, f in normalizar_freq_map(freq_map).items():
        items.append({"descriptor": d, "categoria": DESC2CAT.get(d, "—"), "frecuencia": int(f)})
    df = pd.DataFrame(items)
    if df.empty:
        return pd.DataFrame(columns=["descriptor", "categoria", "frecuencia"])
    return df

def combinar_maps(maps: List[Dict[str, int]]) -> Dict[str, int]:
    total = {}
    for m in maps:
        for d, f in normalizar_freq_map(m).items():
            total[d] = total.get(d, 0) + int(f)
    return total

def info_pareto(freq_map: Dict[str, int]) -> Dict[str, int]:
    d = normalizar_freq_map(freq_map)
    return {"descriptores": len(d), "total": int(sum(d.values()))}

# --- Cálculo Pareto ---
def calcular_pareto(df_in: pd.DataFrame) -> pd.DataFrame:
    df = df_in.copy()
    df["frecuencia"] = pd.to_numeric(df["frecuencia"], errors="coerce").fillna(0).astype(int)
    df = df[df["frecuencia"] > 0]
    if df.empty:
        return df.assign(porcentaje=0.0, acumulado=0, pct_acum=0.0,
                         segmento_real="20%", segmento="80%")
    df = df.sort_values("frecuencia", ascending=False)
    total = int(df["frecuencia"].sum())
    df["porcentaje"] = (df["frecuencia"] / total * 100).round(2)
    df["acumulado"]  = df["frecuencia"].cumsum()
    df["pct_acum"]   = (df["acumulado"] / total * 100).round(2)
    df["segmento_real"] = np.where(df["pct_acum"] <= 80.00, "80%", "20%")
    df["segmento"] = "80%"
    return df.reset_index(drop=True)

def _colors_for_segments(segments: List[str]) -> List[str]:
    return [VERDE if s == "80%" else AZUL for s in segments]

def _wrap_labels(labels: List[str], width: int = 22) -> List[str]:
    return ["\n".join(wrap(str(t), width=width)) for t in labels]
# ============================================================================
# ============================== PARTE 4/10 =================================
# ====== Gráfico Pareto (UI) + Exportación Excel con gráfico combinado ======
# ============================================================================

def dibujar_pareto(df_par: pd.DataFrame, titulo: str):
    if df_par.empty:
        st.info("Ingresa frecuencias (>0) para ver el gráfico.")
        return
    x        = np.arange(len(df_par))
    freqs    = df_par["frecuencia"].to_numpy()
    pct_acum = df_par["pct_acum"].to_numpy()
    colors_b = _colors_for_segments(df_par["segmento_real"].tolist())

    fig, ax1 = plt.subplots(figsize=(14, 5.8))
    ax1.bar(x, freqs, color=colors_b)
    ax1.set_ylabel("Frecuencia")
    ax1.set_xticks(x)
    # --- CAMBIO: etiquetas diagonales 45° para mejor lectura ---
    ax1.set_xticklabels(_wrap_labels(df_par["descriptor"].tolist(), 24), rotation=45, ha="right")
    ax1.set_title(titulo if titulo.strip() else "Diagrama de Pareto", color=TEXTO)

    ax2 = ax1.twinx()
    ax2.plot(x, pct_acum, marker="o", linewidth=2, color=TEXTO)
    ax2.set_ylabel("% acumulado"); ax2.set_ylim(0, 110)
    if (df_par["segmento_real"] == "80%").any():
        cut_idx = np.where(df_par["segmento_real"].to_numpy() == "80%")[0].max()
        ax1.axvline(cut_idx + 0.5, linestyle=":", color="k")
    ax2.axhline(80, linestyle="--", linewidth=1, color="#666666")
    st.pyplot(fig)

# --- Excel export ---
def exportar_excel_con_grafico(df_par: pd.DataFrame, titulo: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        hoja = "Pareto"
        df_x = df_par.copy()
        df_x["porcentaje"] = (df_x["porcentaje"] / 100.0).round(4)
        df_x["pct_acum"]   = (df_x["pct_acum"] / 100.0).round(4)
        df_x = df_x[["categoria", "descriptor", "frecuencia",
                     "porcentaje", "pct_acum", "acumulado", "segmento"]]
        df_x.to_excel(writer, sheet_name=hoja, index=False, startrow=0, startcol=0)
        wb = writer.book; ws = writer.sheets[hoja]
        pct_fmt = wb.add_format({"num_format": "0.00%"})
        total_fmt = wb.add_format({"bold": True})
        ws.set_column("A:A", 18); ws.set_column("B:B", 55); ws.set_column("C:C", 12)
        ws.set_column("D:D", 12, pct_fmt); ws.set_column("E:E", 18, pct_fmt)
        ws.set_column("F:F", 12); ws.set_column("G:G", 10)
        n = len(df_x)
        cats = f"=Pareto!$B$2:$B${n+1}"; vals = f"=Pareto!$C$2:$C${n+1}"; pcts = f"=Pareto!$E$2:$E${n+1}"
        total = int(df_par["frecuencia"].sum())
        ws.write(n + 2, 1, "TOTAL:", total_fmt); ws.write(n + 2, 2, total, total_fmt)
        chart = wb.add_chart({"type": "column"})
        points = [{"fill": {"color": (VERDE if s == "80%" else AZUL)}} for s in df_par["segmento_real"]]
        chart.add_series({"name": "Frecuencia", "categories": cats, "values": vals, "points": points})
        line = wb.add_chart({"type": "line"})
        line.add_series({"name": "% acumulado", "categories": cats, "values": pcts,
                         "y2_axis": True, "marker": {"type": "circle"}})
        chart.combine(line)
        chart.set_y_axis({"name": "Frecuencia"})
        chart.set_y2_axis({"name": "Porcentaje acumulado",
                           "min": 0, "max": 1.10, "major_unit": 0.10, "num_format": "0%"})
        title_text = titulo.strip() if titulo else ""
        chart.set_title({"name": title_text or "Diagrama de Pareto"})
        chart.set_legend({"position": "bottom"}); chart.set_size({"width": 1180, "height": 420})
        ws.insert_chart("I2", chart)
    return output.getvalue()
# ============================================================================
# ============================== PARTE 5/10 =================================
# ======================== Conectores Google Sheets (gspread) ================
# ============================================================================

SCOPES = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]

def _gc():
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=SCOPES)
    return gspread.authorize(creds)

def _open_sheet():
    gc = _gc(); return gc.open_by_url(SPREADSHEET_URL)

def _ensure_ws(sh, title: str, header: List[str]):
    try:
        ws = sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=10)
        ws.append_row(header); return ws
    values = ws.get_all_values()
    if not values:
        ws.append_row(header)
    else:
        first = values[0]
        if [c.strip().lower() for c in first] != [c.strip().lower() for c in header]:
            ws.clear(); ws.append_row(header)
    return ws

def sheets_cargar_portafolio() -> Dict[str, Dict[str, int]]:
    try:
        sh = _open_sheet(); ws = _ensure_ws(sh, WS_PARETOS, ["nombre","descriptor","frecuencia"])
        rows = ws.get_all_records()
        port: Dict[str, Dict[str, int]] = {}
        for r in rows:
            nom = str(r.get("nombre","")).strip()
            desc = str(r.get("descriptor","")).strip()
            freq = int(pd.to_numeric(r.get("frecuencia",0), errors="coerce") or 0)
            if not nom or not desc or freq <= 0: continue
            bucket = port.setdefault(nom, {}); bucket[desc] = bucket.get(desc, 0) + freq
        return port
    except Exception:
        return {}

def sheets_guardar_pareto(nombre: str, freq_map: Dict[str, int], sobrescribir: bool = True):
    sh = _open_sheet()
    ws = _ensure_ws(sh, WS_PARETOS, ["nombre","descriptor","frecuencia"])
    if sobrescribir:
        vals = ws.get_all_values()
        header = vals[0] if vals else ["nombre","descriptor","frecuencia"]
        others = [r for r in vals[1:] if (len(r)>0 and r[0].strip().lower()!=nombre.strip().lower())]
        ws.clear(); ws.update("A1", [header])
        if others: ws.append_rows(others, value_input_option="RAW")
    rows_new = [[nombre, d, int(f)] for d, f in normalizar_freq_map(freq_map).items()]
    if rows_new: ws.append_rows(rows_new, value_input_option="RAW")
# ============================================================================
# ============================== PARTE 6/10 =================================
# =================== Estado de sesión + Estilos básicos PDF =================
# ============================================================================

# ---- Estado de sesión ----
st.session_state.setdefault("freq_map", {})
st.session_state.setdefault("portafolio", {})
st.session_state.setdefault("msel", [])
st.session_state.setdefault("reset_after_save", False)

if not st.session_state["portafolio"]:
    loaded = sheets_cargar_portafolio()
    if loaded: st.session_state["portafolio"].update(loaded)

if st.session_state.get("reset_after_save", False):
    st.session_state["freq_map"] = {}
    st.session_state["msel"] = []
    st.session_state.pop("editor_freq", None)
    st.session_state["reset_after_save"] = False

# ---- Estilos PDF / páginas ----
PAGE_W, PAGE_H = A4

def _styles():
    ss = getSampleStyleSheet()
    ss.add(ParagraphStyle(
        name="CoverTitle", fontName="Helvetica-Bold",
        fontSize=30, leading=36, textColor=TEXTO, alignment=1, spaceAfter=10
    ))
    ss.add(ParagraphStyle(
        name="CoverSubtitle", parent=ss["Normal"], fontSize=12,
        leading=16, textColor=GRIS, alignment=1, spaceAfter=10
    ))
    ss.add(ParagraphStyle(
        name="CoverDate", parent=ss["Normal"], fontSize=15,
        leading=18, textColor=TEXTO, alignment=1, spaceBefore=8   # fecha centrada
    ))
    ss.add(ParagraphStyle(
        name="TitleBig", parent=ss["Title"], fontSize=24,
        leading=28, textColor=TEXTO, alignment=0, spaceAfter=10
    ))
    ss.add(ParagraphStyle(
        name="TitleBigCenter", parent=ss["Title"], fontSize=24,
        leading=28, textColor=TEXTO, alignment=1, spaceAfter=10
    ))
    ss.add(ParagraphStyle(name="H1", parent=ss["Heading1"], fontSize=18, leading=22, textColor=TEXTO, spaceAfter=8))
    ss.add(ParagraphStyle(name="H1Center", parent=ss["Heading1"], fontSize=18, leading=22, textColor=TEXTO, spaceAfter=8, alignment=1))
    ss.add(ParagraphStyle(name="Body", parent=ss["Normal"], fontSize=11, leading=14, textColor="#111"))
    ss.add(ParagraphStyle(name="Small", parent=ss["Normal"], fontSize=9.6, leading=12, textColor=GRIS))
    ss.add(ParagraphStyle(name="TableHead", parent=ss["Normal"], fontSize=11, leading=13, textColor=colors.white))
    return ss

def _page_cover(canv, doc):
    canv.setFillColor(colors.HexColor(TEXTO))
    canv.rect(0, PAGE_H - 0.9*cm, PAGE_W, 0.9*cm, fill=1, stroke=0)

def _page_normal(_canv, _doc):
    pass

def _page_last(canv, _doc):
    canv.setFillColor(colors.HexColor(TEXTO))
    canv.rect(0, 0, PAGE_W, 0.9*cm, fill=1, stroke=0)
# ============================================================================
# ============================== PARTE 7/10 =================================
# ============ Imágenes para PDF (Pareto/Modalidades) y textos helper =======
# ============================================================================

def _pareto_png(df_par: pd.DataFrame, titulo: str) -> bytes:
    x        = np.arange(len(df_par))
    freqs    = df_par["frecuencia"].to_numpy()
    pct_acum = df_par["pct_acum"].to_numpy()
    colors_b = _colors_for_segments(df_par["segmento_real"].tolist())
    fig, ax1 = plt.subplots(figsize=(12.5, 5.0))
    ax1.bar(x, freqs, color=colors_b)
    ax1.set_ylabel("Frecuencia")
    ax1.set_xticks(x)
    # --- CAMBIO: etiquetas diagonales 45° para el PDF ---
    ax1.set_xticklabels(_wrap_labels(df_par["descriptor"].tolist(), 24), rotation=45, ha="right")
    ax1.set_title(titulo if titulo.strip() else "Diagrama de Pareto", color=TEXTO)
    ax2 = ax1.twinx()
    ax2.plot(x, pct_acum, marker="o", linewidth=2, color=TEXTO)
    ax2.set_ylabel("% acumulado"); ax2.set_ylim(0, 110)
    if (df_par["segmento_real"] == "80%").any():
        cut_idx = np.where(df_par["segmento_real"].to_numpy() == "80%")[0].max()
        ax1.axvline(cut_idx + 0.5, linestyle=":", color="k")
    ax2.axhline(80, linestyle="--", linewidth=1, color="#666666")
    buf = io.BytesIO(); fig.tight_layout(); fig.savefig(buf, format="PNG"); plt.close(fig)
    return buf.getvalue()

# ===== MODALIDADES (incluye 'pill') =====
def _modalidades_png(title: str, data_pairs: List[Tuple[str, float]], kind: str = "barh") -> bytes:
    """
    kind: 'barh' (por defecto), 'bar', 'lollipop', 'donut', 'comp100', 'pill'
    """
    labels = [l for l, p in data_pairs if str(l).strip()]
    vals   = [float(p or 0) for l, p in data_pairs if str(l).strip()]
    if not labels:
        labels, vals = ["Sin datos"], [100.0]

    order = np.argsort(vals)[::-1]
    labels = [labels[i] for i in order]
    vals   = [vals[i]   for i in order]
    n = len(labels)

    import matplotlib as mpl
    from matplotlib.patches import FancyBboxPatch, Circle

    cmap = mpl.cm.get_cmap("Blues")
    colors_seq = [cmap(0.35 + 0.5*(i/max(1, n-1))) for i in range(n)]

    if kind == "donut":
        fig, ax = plt.subplots(figsize=(7.8, 5.4))
        wedges, _, _ = ax.pie(
            vals, labels=None, autopct=lambda p: f"{p:.1f}%",
            startangle=90, pctdistance=0.8,
            wedgeprops=dict(width=0.4, edgecolor="white"),
            colors=colors_seq
        )
        ax.legend(wedges, [f"{l} ({v:.1f}%)" for l, v in zip(labels, vals)],
                  title="Modalidades", loc="center left", bbox_to_anchor=(1.02, 0.5), fontsize=9)
        ax.set_title(title, color=TEXTO)

    elif kind == "lollipop":
        fig, ax = plt.subplots(figsize=(11.5, 5.4))
        y = np.arange(n)
        ax.hlines(y=y, xmin=0, xmax=vals, color="#94a3b8", linewidth=2)
        ax.plot(vals, y, "o", markersize=8, color=AZUL)
        ax.set_yticks(y)
        ax.set_yticklabels(_wrap_labels(labels, 35))
        ax.invert_yaxis()
        ax.set_xlabel("Porcentaje"); ax.set_xlim(0, max(100, max(vals)*1.05))
        for i, v in enumerate(vals):
            ax.text(v + 1, i, f"{v:.1f}%", va="center", fontsize=10)
        ax.set_title(title, color=TEXTO)

    elif kind == "bar":
        fig, ax = plt.subplots(figsize=(11.5, 5.4))
        x = np.arange(n)
        ax.bar(x, vals, color=colors_seq)
        ax.set_xticks(x)
        ax.set_xticklabels(_wrap_labels(labels, 20), rotation=0)
        ax.set_ylabel("Porcentaje"); ax.set_ylim(0, max(100, max(vals)*1.15))
        for i, v in enumerate(vals):
            ax.text(i, v + max(vals)*0.03, f"{v:.1f}%", ha="center", fontsize=10)
        ax.set_title(title, color=TEXTO)

    elif kind == "comp100":
        fig, ax = plt.subplots(figsize=(11.5, 3.0))
        left = 0.0
        for i, (lab, v) in enumerate(zip(labels, vals)):
            w = max(0.0, float(v))
            ax.barh(0, w, left=left, color=colors_seq[i])
            if w >= 7:
                ax.text(left + w/2, 0, f"{lab}\n{v:.1f}%", va="center", ha="center", fontsize=9, color="white")
            left += w
        ax.set_xlim(0, max(100, sum(vals)))
        ax.set_yticks([]); ax.set_xlabel("Porcentaje (composición)")
        ax.set_title(title, color=TEXTO)
        ax.grid(False)

    elif kind == "pill":
        fig_height = 0.9 + n*0.85
        fig, ax = plt.subplots(figsize=(10.8, fig_height))
        ax.set_xlim(0, 100); ax.set_ylim(0, n)
        ax.axis("off")
        track_h = 0.72
        round_r = track_h/2

        for i, (lab, v) in enumerate(zip(labels, vals)):
            y = n - 1 - i + (1 - track_h)/2

            track = FancyBboxPatch(
                (0.8, y), 98.4, track_h,
                boxstyle=f"round,pad=0,rounding_size={round_r}",
                linewidth=1, edgecolor="#9dbbd6", facecolor="#e6f0fb"
            )
            ax.add_patch(track)

            prog_w = max(0.001, min(98.4, float(v)))
            prog = FancyBboxPatch(
                (0.8, y), prog_w, track_h,
                boxstyle=f"round,pad=0,rounding_size={round_r}",
                linewidth=0, facecolor=AZUL, alpha=0.35
            )
            ax.add_patch(prog)

            ax.add_patch(Circle((0.8 + round_r*0.8, y + track_h/2), round_r*0.9, color="#a3a3a3", alpha=0.6))
            ax.add_patch(Circle((0.8 + round_r*0.6, y + track_h/2), round_r*0.9, color=AZUL, alpha=0.9))

            badge_w = 12.0; badge_h = track_h*0.8
            badge_x = 5.0;  badge_y = y + (track_h - badge_h)/2
            badge = FancyBboxPatch(
                (badge_x, badge_y), badge_w, badge_h,
                boxstyle=f"round,pad=0.25,rounding_size={badge_h/2}",
                linewidth=1, edgecolor="#cfd8e3", facecolor="white"
            )
            ax.add_patch(badge)
            ax.text(badge_x + badge_w/2, y + track_h/2, f"{v:.1f}%", ha="center", va="center", fontsize=10)

            ax.text(badge_x + badge_w + 3.0, y + track_h/2, lab, va="center", ha="left",
                    fontsize=12, color="#0f172a")

        ax.set_title(title, color=TEXTO)

    else:  # 'barh'
        fig, ax = plt.subplots(figsize=(11.5, 5.4))
        y = np.arange(n)
        ax.barh(y, vals, color=colors_seq)
        ax.set_yticks(y)
        ax.set_yticklabels(_wrap_labels(labels, 35))
        ax.invert_yaxis()
        ax.set_xlabel("Porcentaje")
        ax.set_xlim(0, max(100, max(vals)*1.05))
        for i, v in enumerate(vals):
            ax.text(v + 1, i, f"{v:.1f}%", va="center", fontsize=10)
        ax.set_title(title, color=TEXTO)

    buf = io.BytesIO(); fig.tight_layout(); fig.savefig(buf, format="PNG"); plt.close(fig)
    return buf.getvalue()

def _tema_descriptor(descriptor: str) -> str:
    d = descriptor.lower()
    if "droga" in d or "búnker" in d or "bunker" in d or "narco" in d or "venta de drogas" in d:
        return "drogas"
    if "robo" in d or "hurto" in d or "asalto" in d or "vehícul" in d or "comercio" in d:
        return "delitos contra la propiedad"
    if "violencia" in d or "lesion" in d or "homicidio" in d:
        return "violencia"
    if "infraestructura" in d or "alumbrado" in d or "lotes" in d:
        return "condiciones urbanas / entorno"
    return "seguridad y convivencia"

def _resumen_texto(df_par: pd.DataFrame) -> str:
    if df_par.empty:
        return "Sin datos disponibles."
    total = int(df_par["frecuencia"].sum())
    n = len(df_par)
    top = df_par.iloc[0]
    idx80 = int(np.where(df_par["segmento_real"].to_numpy() == "80%")[0].max() + 1) if (df_par["segmento_real"]=="80%").any() else 0
    tema = _tema_descriptor(str(top["descriptor"]))
    return (f"Se registran <b>{total}</b> hechos distribuidos en <b>{n}</b> descriptores. "
            f"El descriptor de mayor incidencia pertenece al ámbito de <b>{tema}</b>: "
            f"<b>{top['descriptor']}</b>, con <b>{int(top['frecuencia'])}</b> casos "
            f"({float(top['porcentaje']):.2f}%). El punto de corte del <b>80%</b> se alcanza con "
            f"<b>{idx80}</b> descriptores, útiles para la priorización operativa.")

def _texto_modalidades(descriptor: str, pares: List[Tuple[str, float]]) -> str:
    pares_filtrados = [(l, p) for l, p in pares if str(l).strip() and (p or 0) > 0]
    pares_orden = sorted(pares_filtrados, key=lambda x: x[1], reverse=True)
    tema = _tema_descriptor(descriptor)
    if not pares_orden:
        return (f"Para <b>{descriptor}</b> (ámbito: <b>{tema}</b>) no se reportaron modalidades con porcentaje. "
                "Se sugiere recolectar esta información para focalizar acciones.")
    top_txt = "; ".join([f"<b>{l}</b> ({p:.1f}%)" for l, p in pares_orden[:2]])
    return (f"En <b>{descriptor}</b> (ámbito: <b>{tema}</b>) destacan: {top_txt}. "
            "Esto orienta intervenciones específicas sobre las variantes de mayor peso.")
# ============================================================================
# ============================== PARTE 8/10 =================================
# ========= Tabla PDF, generador de Informe PDF y UI de desgloses ===========
# ============================================================================

def _tabla_resultados_flowable(df_par: pd.DataFrame, doc_width: float) -> Table:
    """
    Cuadro simplificado: solo Descriptor, Frecuencia y %.
    Incluye una fila final con el 'Total de respuestas tratadas'.
    """
    # Anchos: Descriptor más amplio para permitir wraps
    fracs = [0.62, 0.20, 0.18]  # Descriptor, Frecuencia, %
    col_widths = [f * doc_width for f in fracs]

    stys = _styles()

    # Estilo de celda con ajuste de línea
    from reportlab.lib.styles import ParagraphStyle
    cell_style = ParagraphStyle(
        name="CellWrap",
        parent=stys["Normal"],
        fontSize=9.6,
        leading=12,
        textColor="#111111",
        wordWrap="CJK",
        spaceBefore=0,
        spaceAfter=0,
    )

    # Encabezado (3 columnas)
    head = [
        Paragraph("Descriptor", stys["TableHead"]),
        Paragraph("Frecuencia", stys["TableHead"]),
        Paragraph("%", stys["TableHead"]),
    ]
    data = [head]

    # Filas de datos
    total_respuestas = int(df_par["frecuencia"].sum()) if not df_par.empty else 0
    for _, r in df_par.iterrows():
        descriptor = Paragraph(str(r["descriptor"]), cell_style)
        frecuencia = int(r["frecuencia"])
        pct = f'{float(r["porcentaje"]):.2f}%'
        data.append([descriptor, frecuencia, pct])

    # Fila de total
    total_row = [Paragraph("<b>Total de respuestas tratadas</b>", cell_style), total_respuestas, ""]
    data.append(total_row)
    total_index = len(data) - 1

    # Construcción de la tabla
    t = Table(data, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    style = TableStyle([
        # Header
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor(TEXTO)),
        ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,0), 11),

        # Cuerpo
        ("FONTSIZE",   (0,1), (-1,-1), 9.6),
        ("ALIGN",      (0,1), (0,-1), "LEFT"),   # Descriptor
        ("ALIGN",      (1,1), (-1,-2), "RIGHT"), # Frecuencia y %
        ("LEFTPADDING",(0,0), (-1,-1), 6),
        ("RIGHTPADDING",(0,0), (-1,-1), 6),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),

        # Zebrados + grid
        ("ROWBACKGROUNDS", (0,1), (-1,-2), [colors.whitesmoke, colors.Color(0.97,0.97,0.97)]),
        ("GRID", (0,0), (-1,-1), 0.25, colors.lightgrey),

        # Fila de total
        ("BACKGROUND", (0,total_index), (-1,total_index), colors.Color(0.93, 0.96, 0.99)),
        ("FONTNAME",   (0,total_index), (-1,total_index), "Helvetica-Bold"),
        ("ALIGN",      (1,total_index), (1,total_index), "RIGHT"),
        ("ALIGN",      (2,total_index), (2,total_index), "RIGHT"),
    ])
    t.setStyle(style)
    return t


def generar_pdf_informe(nombre_informe: str,
                        df_par: pd.DataFrame,
                        desgloses: List[Dict]) -> bytes:
    buf = io.BytesIO()
    doc = BaseDocTemplate(
        buf, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm
    )
    frame_std  = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
    frame_last = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="last")

    doc.addPageTemplates([
        PageTemplate(id="Cover",  frames=[frame_std], onPage=_page_cover),
        PageTemplate(id="Normal", frames=[frame_std], onPage=_page_normal),
        PageTemplate(id="Last",   frames=[frame_last], onPage=_page_last),
    ])
    stys = _styles()
    story: List = []

    # ---------- PORTADA ----------
    story += [NextPageTemplate("Normal")]
    story += [Spacer(1, 2.2*cm)]
    story += [Paragraph(f"Informe de Resultados Diagrama de Pareto - {nombre_informe}", stys["CoverTitle"])]
    story += [Paragraph("Estrategia Sembremos Seguridad", stys["CoverSubtitle"])]
    story += [Paragraph(datetime.now().strftime("Fecha: %d/%m/%Y"), stys["CoverDate"])]
    story += [PageBreak()]

    # ---------- INTRODUCCIÓN (permanece aquí si ya la venías usando) ----------
    story += [Paragraph("Introducción", stys["TitleBig"]), Spacer(1, 0.2*cm)]
    story += [Paragraph(
        "Este informe presenta un análisis tipo <b>Pareto (80/20)</b> sobre los descriptores seleccionados. "
        "El objetivo es identificar los elementos que concentran la mayor parte de los hechos reportados para apoyar la "
        "priorización operativa y la toma de decisiones. El documento incluye el gráfico de Pareto, un cuadro "
        "resumido con frecuencia y porcentaje por descriptor, y al final una sección de conclusiones y recomendaciones.",
        stys["Body"]
    ), Spacer(1, 0.35*cm)]

    # ---------- RESULTADOS ----------
    story += [Paragraph("Resultados generales", stys["TitleBig"]), Spacer(1, 0.2*cm)]
    story += [Paragraph(_resumen_texto(df_par), stys["Body"]), Spacer(1, 0.3*cm)]

    pareto_png = _pareto_png(df_par, "Diagrama de Pareto")
    story += [RLImage(io.BytesIO(pareto_png), width=doc.width, height=8.6*cm)]
    story += [Spacer(1, 0.25*cm)]
    story += [Paragraph(
        "El diagrama muestra la frecuencia por descriptor (barras en verde/azul) y el <b>porcentaje acumulado</b> (línea). "
        "La línea punteada del 80% indica el <b>punto de corte</b> para priorización.",
        stys["Small"]
    ), Spacer(1, 0.25*cm)]
    story.append(_tabla_resultados_flowable(df_par, doc.width))
    story += [Spacer(1, 0.25*cm)]

    # ---------- MODALIDADES (título + texto + gráfico SIEMPRE juntos) ----------
    for sec in desgloses:
        descriptor = sec.get("descriptor", "").strip()
        rows = sec.get("rows", [])
        chart_kind = sec.get("chart", "barh")
        pares = [(r.get("Etiqueta",""), float(r.get("%", 0) or 0)) for r in rows]

        bloque = [
            Spacer(1, 0.4*cm),
            Paragraph(f"Modalidades de la problemática — {descriptor}", stys["TitleBig"]),
            Spacer(1, 0.1*cm),
            Paragraph(_texto_modalidades(descriptor, pares), stys["Small"]),
            Spacer(1, 0.2*cm),
            RLImage(io.BytesIO(_modalidades_png(descriptor or 'Modalidades', pares, kind=chart_kind)),
                    width=doc.width, height=8.5*cm),
        ]
        story += [KeepTogether(bloque)]

    # ---------- CIERRE ----------
    story += [PageBreak(), NextPageTemplate("Last")]
    story += [
        Paragraph("Conclusiones y recomendaciones", stys["TitleBigCenter"]),
        Spacer(1, 0.2*cm),
        Paragraph(
            "• Priorizar intervenciones sobre los descriptores que conforman el <b>80% acumulado</b>. "
            "• Coordinar acciones interinstitucionales enfocadas en las <b>modalidades</b> con mayor porcentaje. "
            "• Fortalecer la participación comunitaria y el control territorial en puntos críticos. "
            "• Monitorear indicadores mensualmente para evaluar la efectividad de las acciones.",
            stys["Body"]
        ),
        Spacer(1, 0.8*cm),
        Paragraph("Dirección de Programas Policiales Preventivos – MSP", stys["H1Center"]),
        Paragraph("Sembremos Seguridad", stys["H1Center"]),
    ]

    doc.build(story)
    return buf.getvalue()


# === Helpers UI formulario de desgloses (con selector de tipo de gráfico) ===
def ui_desgloses(descriptor_list: List[str], key_prefix: str) -> List[Dict]:
    st.caption("Opcional: agrega secciones de ‘Modalidades’. Cada sección admite hasta 10 filas (Etiqueta + %).")
    max_secs = max(0, len(descriptor_list))
    default_val = 1 if max_secs > 0 else 0
    n_secs = st.number_input("Cantidad de secciones de Modalidades",
                             min_value=0, max_value=max_secs, value=default_val, step=1,
                             key=f"{key_prefix}_nsecs")
    desgloses: List[Dict] = []
    for i in range(n_secs):
        with st.expander(f"Sección Modalidades #{i+1}", expanded=(i == 0)):
            dsel = st.selectbox(f"Descriptor para la sección #{i+1}",
                                options=["(elegir)"] + descriptor_list, index=0, key=f"{key_prefix}_desc_{i}")

            chart_kind = st.selectbox(
                "Tipo de gráfico",
                options=[("Barras horizontales", "barh"),
                         ("Barras verticales", "bar"),
                         ("Lollipop (palo+punto)", "lollipop"),
                         ("Dona / Pie", "donut"),
                         ("Barra 100% (composición)", "comp100"),
                         ("Píldora (progreso redondeado)", "pill")],
                index=0, format_func=lambda x: x[0], key=f"{key_prefix}_chart_{i}"
            )[1]

            rows = [{"Etiqueta":"", "%":0.0} for _ in range(10)]
            df_rows = pd.DataFrame(rows)
            de = st.data_editor(
                df_rows, key=f"{key_prefix}_rows_{i}", use_container_width=True,
                column_config={
                    "Etiqueta": st.column_config.TextColumn("Etiqueta / Modalidad", width="large"),
                    "%": st.column_config.NumberColumn("Porcentaje", min_value=0.0, max_value=100.0, step=0.1)
                },
                num_rows="fixed"
            )
            total_pct = float(pd.to_numeric(de["%"], errors="coerce").fillna(0).sum())
            st.caption(f"Suma actual: {total_pct:.1f}% (recomendado ≈100%)")

            if dsel != "(elegir)":
                desgloses.append({"descriptor": dsel,
                                  "rows": de.to_dict(orient="records"),
                                  "chart": chart_kind})
    return desgloses

# ============================================================================
# ============================== PARTE 9/10 =================================
# ======================= UI principal (Editor Pareto/Guardado) ==============
# ============================================================================

st.title("Pareto de Descriptores")

c_t1, c_t2, c_t3 = st.columns([2,1,1])
with c_t1:
    titulo = st.text_input("Título del Pareto (opcional)", value="Pareto Comunidad")
with c_t2:
    nombre_para_guardar = st.text_input("Nombre para guardar este Pareto", value="Comunidad")
with c_t3:
    if st.button("🔄 Recargar portafolio desde Sheets"):
        st.session_state["portafolio"] = sheets_cargar_portafolio()
        st.success("Portafolio recargado desde Google Sheets.")
        st.rerun()

cat_df = pd.DataFrame(CATALOGO).sort_values(["categoria","descriptor"]).reset_index(drop=True)
opciones = cat_df["descriptor"].tolist()
seleccion = st.multiselect("1) Escoge uno o varios descriptores", options=opciones,
                           default=st.session_state["msel"], key="msel")

st.subheader("2) Asigna la frecuencia")
if seleccion:
    base = cat_df[cat_df["descriptor"].isin(seleccion)].copy()
    base["frecuencia"] = [st.session_state["freq_map"].get(d, 0) for d in base["descriptor"]]

    edit = st.data_editor(
        base, key="editor_freq", num_rows="fixed", use_container_width=True,
        column_config={
            "descriptor": st.column_config.TextColumn("DESCRIPTOR", width="large"),
            "categoria": st.column_config.TextColumn("CATEGORÍA", width="small"),
            "frecuencia": st.column_config.NumberColumn("Frecuencia", min_value=0, step=1),
        },
    )
    for _, row in edit.iterrows():
        st.session_state["freq_map"][row["descriptor"]] = int(row["frecuencia"])

    df_in = edit[["descriptor","categoria"]].copy()
    df_in["frecuencia"] = df_in["descriptor"].map(st.session_state["freq_map"]).fillna(0).astype(int)

    st.subheader("3) Pareto (en edición)")
    tabla = calcular_pareto(df_in)

    mostrar = tabla.copy()[["categoria","descriptor","frecuencia","porcentaje","pct_acum","acumulado","segmento"]]
    mostrar = mostrar.rename(columns={"pct_acum": "porcentaje acumulado"})
    mostrar["porcentaje"] = mostrar["porcentaje"].map(lambda x: f"{x:.2f}%")
    mostrar["porcentaje acumulado"] = mostrar["porcentaje acumulado"].map(lambda x: f"{x:.2f}%")

    c1, c2 = st.columns([1,1], gap="large")
    with c1:
        st.markdown("**Tabla de Pareto**")
        if not tabla.empty:
            st.dataframe(mostrar, use_container_width=True, hide_index=True)
        else:
            st.info("Ingresa frecuencias (>0) para ver la tabla.")
    with c2:
        st.markdown("**Gráfico de Pareto**"); dibujar_pareto(tabla, titulo)

    st.subheader("4) Guardar / Descargar")
    col_g1, col_g2, _ = st.columns([1,1,2])
    with col_g1:
        sobrescribir = st.checkbox("Sobrescribir si existe", value=True)
        if st.button("💾 Guardar este Pareto"):
            nombre = nombre_para_guardar.strip()
            if not nombre:
                st.warning("Indica un nombre para guardar el Pareto.")
            else:
                st.session_state["portafolio"][nombre] = normalizar_freq_map(st.session_state["freq_map"])
                try:
                    sheets_guardar_pareto(nombre, st.session_state["freq_map"], sobrescribir=sobrescribir)
                    st.success(f"Pareto '{nombre}' guardado en Google Sheets y en la sesión.")
                except Exception as e:
                    st.warning(f"Se guardó en la sesión, pero hubo un problema con Sheets: {e}")
                st.session_state["reset_after_save"] = True
                st.rerun()
    with col_g2:
        if not tabla.empty:
            st.download_button(
                "⬇️ Excel del Pareto (edición)",
                data=exportar_excel_con_grafico(tabla, titulo),
                file_name=f"pareto_{(nombre_para_guardar or 'edicion').lower().replace(' ','_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ===== Informe PDF desde el editor =====
    st.markdown("---")
    st.header("🧾 Elaborar Informe PDF (desde el editor)")
    nombre_informe = st.text_input(
        "Nombre del informe (portada)",
        value=(nombre_para_guardar.strip() or "Pareto Comunidad"),
        key="inf_editor_nombre"
    )
    desgloses = ui_desgloses(tabla["descriptor"].tolist(), key_prefix="editor")
    col_inf1, col_inf2 = st.columns([1,3])
    with col_inf1:
        gen = st.button("📄 Generar Informe PDF (editor)", type="primary", use_container_width=True, key="btn_inf_editor")
    with col_inf2:
        st.caption("Incluye: Portada, Resumen + Gráfico + Tabla, Modalidades (si agregas) y Cierre.")

    if gen:
        if tabla.empty:
            st.warning("No hay datos para el informe. Asigna frecuencias primero.")
        else:
            pdf_bytes = generar_pdf_informe(nombre_informe, tabla, desgloses)
            st.success("Informe generado.")
            st.download_button(
                "⬇️ Descargar Informe PDF",
                data=pdf_bytes,
                file_name=f"informe_{nombre_informe.lower().replace(' ','_')}.pdf",
                mime="application/pdf",
                key="dl_inf_editor",
            )

else:
    st.info("Selecciona al menos un descriptor para continuar. Tus frecuencias se conservarán si luego agregas más descriptores.")
# ============================================================================ #
# 🧩 PARTE 10 — PORTAFOLIO, UNIFICADO Y DESCARGAS
# ============================================================================ #

st.markdown("---")
st.header("📁 Portafolio de Paretos (guardados)")

if not st.session_state["portafolio"]:
    st.info("Aún no hay paretos guardados. Guarda el primero desde la sección anterior.")
else:
    st.subheader("Selecciona paretos para Unificar")
    nombres = sorted(st.session_state["portafolio"].keys())
    sel_unif = st.multiselect(
        "Elige 2 o más paretos para combinar (o usa el botón de 'Unificar todos')",
        options=nombres, default=[], key="sel_unif"
    )

    c_unif1, c_unif2 = st.columns([1, 1])
    with c_unif1:
        unificar_todos = st.button("🔗 Unificar TODOS los paretos guardados")
    with c_unif2:
        st.caption(f"Total de paretos guardados: **{len(nombres)}**")

    st.markdown("### Paretos guardados")

    for nom in nombres:
        freq_map = st.session_state["portafolio"][nom]
        meta = info_pareto(freq_map)
        with st.expander(f"🔹 {nom} — {meta['descriptores']} descriptores | Total: {meta['total']}"):
            df_base = df_desde_freq_map(freq_map)
            tabla_g = calcular_pareto(df_base)

            mostrar_g = tabla_g.copy()[["categoria", "descriptor", "frecuencia",
                                        "porcentaje", "pct_acum", "acumulado", "segmento"]]
            mostrar_g = mostrar_g.rename(columns={"pct_acum": "porcentaje acumulado"})
            if not mostrar_g.empty:
                mostrar_g["porcentaje"] = mostrar_g["porcentaje"].map(lambda x: f"{x:.2f}%")
                mostrar_g["porcentaje acumulado"] = mostrar_g["porcentaje acumulado"].map(lambda x: f"{x:.2f}%")

            cc1, cc2, cc3 = st.columns([1, 1, 1])
            with cc1:
                if not mostrar_g.empty:
                    st.dataframe(mostrar_g, use_container_width=True, hide_index=True)
                else:
                    st.info("Este pareto no tiene frecuencias > 0.")
            with cc2:
                st.markdown("**Gráfico**")
                dibujar_pareto(tabla_g, f"Pareto — {nom}")
            with cc3:
                st.markdown("**Acciones**")
                if not tabla_g.empty:
                    st.download_button(
                        "⬇️ Excel de este Pareto",
                        data=exportar_excel_con_grafico(tabla_g, f"Pareto — {nom}"),
                        file_name=f"pareto_{nom.lower().replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{nom}",
                    )
                if st.button("📥 Cargar este Pareto al editor", key=f"load_{nom}"):
                    st.session_state["freq_map"] = dict(freq_map)
                    st.session_state["msel"] = list(freq_map.keys())
                    st.success(f"Pareto '{nom}' cargado al editor (arriba). Desplázate para editar.")

                try:
                    pop = st.popover("📄 Informe PDF de este Pareto")
                except Exception:
                    pop = st.expander("📄 Informe PDF de este Pareto", expanded=False)

                with pop:
                    nombre_inf_ind = st.text_input("Nombre del informe", value=f"{nom}", key=f"inf_nom_{nom}")
                    desgloses_ind = ui_desgloses(tabla_g["descriptor"].tolist(), key_prefix=f"inf_{nom}")
                    if st.button("Generar PDF", key=f"btn_inf_{nom}"):
                        pdf_bytes = generar_pdf_informe(nombre_inf_ind, tabla_g, desgloses_ind)
                        st.download_button(
                            "⬇️ Descargar PDF",
                            data=pdf_bytes,
                            file_name=f"informe_{nom.lower().replace(' ', '_')}.pdf",
                            mime="application/pdf",
                            key=f"dl_inf_{nom}",
                        )

                if st.button("🗑️ Eliminar de la sesión", key=f"del_{nom}"):
                    try:
                        del st.session_state["portafolio"][nom]
                        st.warning(f"Pareto '{nom}' eliminado del portafolio de la sesión.")
                        st.rerun()
                    except Exception:
                        st.error("No se pudo eliminar. Intenta de nuevo.")

    # ============================= UNIFICADO =================================
    st.markdown("---")
    st.header("🔗 Pareto Unificado (por filtro o general)")
    maps_a_unir = []
    titulo_unif = ""

    if unificar_todos and nombres:
        maps_a_unir = [st.session_state["portafolio"][n] for n in nombres]
        titulo_unif = "Pareto General (todos los paretos)"
    elif len(st.session_state.get("sel_unif", [])) >= 2:
        maps_a_unir = [st.session_state["portafolio"][n] for n in st.session_state["sel_unif"]]
        titulo_unif = f"Unificado: {', '.join(st.session_state['sel_unif'])}"

    if maps_a_unir:
        combinado = combinar_maps(maps_a_unir)
        df_unif = df_desde_freq_map(combinado)
        tabla_unif = calcular_pareto(df_unif)

        mostrar_u = tabla_unif.copy()[["categoria", "descriptor", "frecuencia",
                                       "porcentaje", "pct_acum", "acumulado", "segmento"]]
        mostrar_u = mostrar_u.rename(columns={"pct_acum": "porcentaje acumulado"})
        if not mostrar_u.empty:
            mostrar_u["porcentaje"] = mostrar_u["porcentaje"].map(lambda x: f"{x:.2f}%")
            mostrar_u["porcentaje acumulado"] = mostrar_u["porcentaje acumulado"].map(lambda x: f"{x:.2f}%")

        cu1, cu2 = st.columns([1, 1], gap="large")
        with cu1:
            st.markdown("**Tabla Unificada**")
            if not mostrar_u.empty:
                st.dataframe(mostrar_u, use_container_width=True, hide_index=True)
            else:
                st.info("Sin datos > 0 en la combinación seleccionada.")
        with cu2:
            st.markdown("**Gráfico Unificado**")
            dibujar_pareto(tabla_unif, titulo_unif or "Pareto Unificado")

        if not tabla_unif.empty:
            st.download_button(
                "⬇️ Descargar Excel del Pareto Unificado",
                data=exportar_excel_con_grafico(tabla_unif, titulo_unif or "Pareto Unificado"),
                file_name="pareto_unificado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_unificado",
            )

        # ===== Informe PDF Unificado =====
        st.markdown("### 🧾 Elaborar Informe PDF (Unificado / General)")
        nombre_inf_unif = st.text_input(
            "Nombre del informe",
            value=(titulo_unif or "Pareto Unificado"),
            key="inf_unif_nombre"
        )
        desgloses_unif = ui_desgloses(tabla_unif["descriptor"].tolist(), key_prefix="unif")

        if st.button("📄 Generar Informe PDF (unificado)", key="btn_inf_unif"):
            pdf_bytes = generar_pdf_informe(nombre_inf_unif, tabla_unif, desgloses_unif)
            st.download_button(
                "⬇️ Descargar Informe PDF",
                data=pdf_bytes,
                file_name=f"informe_{(titulo_unif or 'Pareto Unificado').lower().replace(' ', '_')}.pdf",
                mime="application/pdf",
                key="dl_inf_unif",
            )

    else:
        st.info(
            "Selecciona 2+ paretos en el multiselect o usa el botón 'Unificar TODOS' "
            "para habilitar el unificado."
        )




