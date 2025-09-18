# app.py — Pareto automático robusto
# Solo sube la matriz .xlsx y listo.

from __future__ import annotations
from io import BytesIO
from typing import List, Tuple

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.drawing.image import Image as XLImage

# Parámetros fijos
RANGO_DESDE = "AI"
RANGO_HASTA = "ET"
TRY_HEADER_ROWS = 6   # busca encabezado en las primeras 6 filas
TITULO_COMUNIDAD_DEF = "DESAMPARADOS NORTE"
YELLOW_HEXES = {"FFFF00","FFEB9C","FFF2CC","FFE699","FFD966","FFF4CC"}

# ───────────────────────────
def _strip(x):
    if x is None: return ""
    if isinstance(x, float) and np.isnan(x): return ""
    return str(x).strip()

def hoja_activa(xfile: BytesIO):
    xfile.seek(0)
    wb = load_workbook(filename=xfile, data_only=True)
    return wb, wb.active

def rango_indices(desde: str, hasta: str) -> Tuple[int,int]:
    return column_index_from_string(desde), column_index_from_string(hasta)

def autodetect_header_row(ws, idx_from, idx_to, try_n=TRY_HEADER_ROWS) -> int:
    best_row, best_cnt = 1, -1
    top = min(try_n, ws.max_row)
    for r in range(1, top+1):
        cnt = sum(1 for c in range(idx_from, idx_to+1)
                  if _strip(ws.cell(r,c).value) != "")
        if cnt > best_cnt:
            best_row, best_cnt = r, cnt
    return best_row

def headers_yellow(ws, header_row, idx_from, idx_to) -> List[str]:
    amarillos=[]
    for c in range(idx_from, idx_to+1):
        cell = ws.cell(header_row, c)
        fill = cell.fill
        is_yellow=False
        if fill and getattr(fill,"patternType",None):
            fg=getattr(fill,"fgColor",None)
            rgb=getattr(fg,"rgb",None) if fg else None
            if rgb:
                rgb=rgb.upper().replace("#","")
                if len(rgb)==8: rgb=rgb[2:]
                if rgb in YELLOW_HEXES: is_yellow=True
        if is_yellow:
            v=_strip(cell.value) or get_column_letter(c)
            amarillos.append(v)
    return amarillos

def leer_dataframe(xfile: BytesIO, header_row:int) -> pd.DataFrame:
    xfile.seek(0)
    df = pd.read_excel(xfile, header=header_row-1, engine="openpyxl")
    return df.dropna(how="all").reset_index(drop=True)

def contar_frecuencias(df: pd.DataFrame, columnas: List[str]) -> pd.DataFrame:
    cols=[c for c in columnas if c in df.columns]
    if not cols: return pd.DataFrame(columns=["DESCRIPTOR","frecuencia"])
    freqs=[]
    for c in cols:
        s=df[c]
        mask=~s.isna()
        mask&=s.astype(str).str.strip().ne("")
        num=pd.to_numeric(s,errors="coerce")
        mask&=~((~num.isna())&(num==0))
        freqs.append({"DESCRIPTOR":c,"frecuencia":int(mask.sum())})
    out=pd.DataFrame(freqs,columns=["DESCRIPTOR","frecuencia"])
    return out[out["frecuencia"]>0].sort_values("frecuencia",ascending=False).reset_index(drop=True)

def construir_pareto(df_freqs):
    total=int(df_freqs["frecuencia"].sum())
    if total==0:
        return pd.DataFrame(columns=["#","DESCRIPTOR","frecuencia","%","acumul","acumul%","dentro_80"]),0
    df=df_freqs.copy()
    df["%"]=(df["frecuencia"]/total)*100
    df["acumul"]=df["frecuencia"].cumsum()
    df["acumul%"]=df["%"].cumsum()
    df["#"]=np.arange(1,len(df)+1)
    df["dentro_80"]=df["acumul%"]<=80
    return df[["#","DESCRIPTOR","frecuencia","%","acumul","acumul%","dentro_80"]],total

def graficar_pareto(df_pareto, titulo):
    if df_pareto.empty:
        fig,ax=plt.subplots(); ax.text(0.5,0.5,"Sin datos",ha="center",va="center")
        return fig
    x=np.arange(len(df_pareto)); frec=df_pareto["frecuencia"].values; acum=df_pareto["acumul%"].values
    fig,ax1=plt.subplots(figsize=(18,7),dpi=130)
    ax1.bar(x,frec,color="#4E79A7"); ax1.set_ylabel("Frecuencia")
    ax1.set_xticks(x); ax1.set_xticklabels(df_pareto["DESCRIPTOR"].values,rotation=65,ha="right")
    ax2=ax1.twinx(); ax2.plot(x,acum,marker="o",color="#F28E2B")
    ax2.set_ylim(0,105); ax2.yaxis.set_major_formatter(FuncFormatter(lambda v,p:f"{v:.0f}%"))
    ax2.axhline(80,color="#707070",linestyle="--"); idx80=int(np.argmax(acum>=80)) if (acum>=80).any() else len(acum)-1
    ax1.axvline(idx80,color="#D62728")
    fig.suptitle(titulo,fontsize=16); fig.tight_layout(); return fig

def exportar_excel(df_pareto,fig) -> BytesIO:
    bio=BytesIO()
    with pd.ExcelWriter(bio,engine="openpyxl") as writer:
        df_pareto.to_excel(writer,sheet_name="PARETO",index=False)
        wb=writer.book; ws=wb["PARETO"]
        buf=BytesIO(); fig.savefig(buf,format="png",dpi=220,bbox_inches="tight"); buf.seek(0)
        ws.add_image(XLImage(buf),f"B{len(df_pareto)+4}"); writer._save()
    bio.seek(0); return bio

# ───────────────────────────
st.set_page_config(page_title="Pareto Comunidad",layout="wide")
st.title("Generador de Pareto – Comunidad")

comunidad=st.text_input("Nombre de la comunidad",value=TITULO_COMUNIDAD_DEF)
archivo=st.file_uploader("Sube la MATRIZ (.xlsx)",type=["xlsx"])

if not archivo:
    st.info("Sube tu matriz. El sistema detectará el encabezado automáticamente (AI–ET).")
    st.stop()

# 1) hoja activa + detectar fila encabezado
wb,ws=hoja_activa(archivo)
idx_from,idx_to=rango_indices(RANGO_DESDE,RANGO_HASTA)
header_row=autodetect_header_row(ws,idx_from,idx_to)

# 2) encabezados amarillos o fallback
amarillos=headers_yellow(ws,header_row,idx_from,idx_to)
if not amarillos:
    amarillos=[_strip(ws.cell(header_row,c).value) for c in range(idx_from,idx_to+1) if _strip(ws.cell(header_row,c).value)!=""]

# 3) dataframe completo
df=leer_dataframe(archivo,header_row)

# 4) conteo
freqs=contar_frecuencias(df,amarillos)
if freqs.empty:
    st.error("No se encontraron datos en las columnas detectadas del rango AI–ET.")
    st.stop()

df_pareto,total=construir_pareto(freqs)

st.markdown(f"**Filas:** {len(df)} • **Columnas:** {len(amarillos)} • **Total:** {total}")
st.dataframe(df_pareto,use_container_width=True)

fig=graficar_pareto(df_pareto,f"PARETO COMUNIDAD {comunidad.upper()}")
st.pyplot(fig,use_container_width=True)

c1,c2=st.columns(2)
with c1:
    buf=BytesIO(); fig.savefig(buf,format="png",dpi=220,bbox_inches="tight"); buf.seek(0)
    st.download_button("⬇️ PNG",data=buf,file_name=f"pareto_{comunidad.lower().replace(' ','_')}.png",mime="image/png")
with c2:
    xio=exportar_excel(df_pareto,fig)
    st.download_button("⬇️ Excel",data=xio,file_name=f"pareto_{comunidad.lower().replace(' ','_')}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
