# -*- coding: utf-8 -*-
"""
Generador de Pedido SUOLMEX (B1–B4) — Versión Refactorizada

Cambios clave vs. tu versión original:
- Estructura modular con funciones claras y constantes en un solo lugar.
- Login/roles, caché de fichas y parseo B1–B4 intactos.
- Cálculo y visualización con MERMA aplicándose a Poliol/ISO/Mezcla en el UI y en PDF.
- Editor de pedido, resumen por código/colores, Excel para compras y PDF profesional.
- Estilo visual consolidado y helpers reutilizables.

Pega este archivo como tu app principal de Streamlit.
"""

# ===================== Imports =====================
import os
import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime
import hashlib
import io
import re
import time
import json
import uuid
import hashlib
import sqlite3
import unicodedata
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
from fpdf import FPDF
import time
import re
import unicodedata

# ===================== Constantes =====================
APP_TITLE = "Generador de Pedido SUOLMEX (B1–B4)"
DB_PATH    = "usuarios.db"
FICHAS_PATH = "FICHAS2.xlsx"
LOGO_PATH   = "logo_suolmex.jpg"

# Config PDF/UI
MERMA_FRAC = 0.03            # 3% merma
UNIDADES    = "kg"
DEC         = 2
DETALLADO   = True           # PDF con detalle por modelo/talla

# Hojas válidas en fichas (ajusta a tu realidad)
FICHAS_HOJAS = ['6001', '2066', '2060', '4098', 'PLANTILLAS']

# ===================== Utilidades Generales =====================
def clean_key(s):
    if pd.isna(s):
        return ""
    
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return " ".join(s.split())


_num_pat = re.compile(r"[-+]?\d*\.?\d+")


def es_numero_valido(s: str) -> bool:
    if s is None:
        return False

    s = str(s).strip()

    if s.upper() in ["", "0"]:
        return False

    return bool(_num_pat.fullmatch(s))

# Formateadores comunes
fmt_num = lambda x, dec=DEC: (f"{float(x):,.{dec}f}" if pd.notna(x) else "-")
fmt_ent = lambda x: (f"{int(x):,}" if pd.notna(x) else "-")

# Merma
con_merma = lambda x, frac=MERMA_FRAC: (x / (1.0 - frac))

# ===================== DB / Sesión =====================
def init_db():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    c = conn.cursor()
    c.execute(
        """
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo TEXT UNIQUE NOT NULL,
            contrasena TEXT NOT NULL,
            rol TEXT NOT NULL
        )
        """
    )
    conn.commit()
    return conn, c

def calcular_resumen_bandas(resumen_df, merma_frac=0.03):
    """
    Agrupa por código (la columna 'Hoja' que en tu caso es el código de ficha)
    y devuelve pares, poliol/ISO y mezclas con merma.
    """
    resumen_df = resumen_df.copy()
    resumen_df["Poliol_kg"] = resumen_df["Poliol (g)"] / 1000
    resumen_df["ISO_kg"] = resumen_df["ISO (g)"] / 1000
hash_password = lambda s: hashlib.sha256(s.encode()).hexdigest()

    agg = resumen_df.groupby("Hoja").agg(
        pares_total=("Cantidad pares", "sum"),
        poliol_necesario_kg=("Poliol_kg", "sum"),
        iso_necesario_kg=("ISO_kg", "sum"),
    ).reset_index().rename(columns={"Hoja": "codigo"})
def ensure_admin(c, conn):
    if not c.execute("SELECT 1 FROM usuarios WHERE codigo='admin'").fetchone():
        c.execute("INSERT INTO usuarios (codigo, contrasena, rol) VALUES (?, ?, ?)",
                  ('admin', hash_password('admin123'), 'admin'))
        conn.commit()

    agg["poliol_con_merma_kg"] = agg["poliol_necesario_kg"] / (1 - merma_frac)
    agg["iso_con_merma_kg"] = agg["iso_necesario_kg"] / (1 - merma_frac)
    agg["mezcla_sin_merma_kg"] = agg["poliol_necesario_kg"] + agg["iso_necesario_kg"]
    agg["mezcla_total_con_merma_kg"] = agg["poliol_con_merma_kg"] + agg["iso_con_merma_kg"]
# Persistencia local ligera por navegador

    totales = {
        "pares_total": agg["pares_total"].sum(),
        "poliol_necesario_kg": agg["poliol_necesario_kg"].sum(),
        "iso_necesario_kg": agg["iso_necesario_kg"].sum(),
        "poliol_con_merma_kg": agg["poliol_con_merma_kg"].sum(),
        "iso_con_merma_kg": agg["iso_con_merma_kg"].sum(),
        "mezcla_sin_merma_kg": agg["mezcla_sin_merma_kg"].sum(),
        "mezcla_total_con_merma_kg": agg["mezcla_total_con_merma_kg"].sum(),
    }
def obtener_session_id():
    if "session_id" not in st.session_state:
        st.session_state.session_id = uuid.uuid4().hex
    return st.session_state.session_id

    return agg, totales
def _session_path():
    return f"session_{obtener_session_id()}.json"

def guardar_sesion():
    with open(_session_path(), "w", encoding="utf-8") as f:
        json.dump({
            "logueado": st.session_state.get("logueado", False),
            "usuario":  st.session_state.get("usuario", None),
            "rol":      st.session_state.get("rol", None),
        }, f)

def cargar_sesion():
    try:
        with open(_session_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
            st.session_state.logueado = data.get("logueado", False)
            st.session_state.usuario  = data.get("usuario")
            st.session_state.rol      = data.get("rol")
    except Exception:
        st.session_state.logueado = False

# ===================== Fichas =====================
@st.cache_data(show_spinner=False)
def cargar_fichas(mtime: float):
    if not os.path.exists(FICHAS_PATH):
        return pd.DataFrame()
    xl = pd.ExcelFile(FICHAS_PATH)
    dfs = []
    for hoja in FICHAS_HOJAS:
        if hoja in xl.sheet_names:
            df = xl.parse(hoja)
            df['Hoja'] = hoja
            dfs.append(df)
    if not dfs:
        return pd.DataFrame()
    df = pd.concat(dfs, ignore_index=True)

    # Normalizaciones
    df["Codigo del Producto"] = df["Codigo del Producto"].astype(str).str.strip()
    df["Linea"]   = df["Linea"].astype(str).str.strip().str.upper()
    df["Corrida"] = df["Corrida"].astype(str).str.strip()
    df["Peso/Pie"] = pd.to_numeric(df["Peso/Pie"], errors="coerce")

    # Claves limpias para matching
    df["__Linea_clean"]   = df["Linea"].apply(clean_key)
    df["__Corrida_clean"] = df["Corrida"].apply(clean_key)

    # Filtro de filas válidas
    return df.dropna(subset=["Peso/Pie", "Relacion Poliol:ISO"]).copy()

# ===================== Parseo programación B1–B4 =====================
def extract_programacion_estatica_B1_B4(file_like, codigo_producto_override=None, hojas_objetivo=None, debug=False):
if hojas_objetivo is None:
hojas_objetivo = ["B1", "B2", "B3", "B4"]
try:
xl = pd.ExcelFile(file_like)
except Exception as e:
st.error(f"No se pudo abrir el Excel: {e}")
        return pd.DataFrame(columns=["Código del Producto","Modelo","Talla","Cantidad pares","Hoja"])
        return pd.DataFrame(columns=["Código del Producto","Modelo","Talla","Cantidad pares","Hoja"]) 

resultados = []
bloques_inicio = [3, 9, 15, 21, 27, 33, 39, 45]
@@ -106,17 +199,14 @@ def es_columna_total_dinamica(df, col, fila_tallas):
modelo = str(df_raw.iat[fila_mod, 0]).strip()
if not (codigo and modelo):
continue
    
            fila_color = inicio + 3

            fila_color = inicio + 3
color = str(df_raw.iat[fila_color, 0]).strip()
    

if hoja == "B1":
extra = str(df_raw.iat[fila_color, 28]).strip()

if extra and any(ch.isalpha() for ch in extra):
color = extra
                    
if not color:
color = modelo

@@ -126,613 +216,823 @@ def es_columna_total_dinamica(df, col, fila_tallas):
excl.add(c)

for c in range(2,12):
                if c in excl: continue
                if c in excl: 
                    continue
talla = str(df_raw.iat[fila_code, c]).strip()
cnt   = str(df_raw.iat[fila_cnt,  c]).strip()
if talla and es_numero_valido(cnt):
resultados.append({
"Código del Producto": codigo,
                        "Color": color, 
                        "Color": color,
"Modelo": modelo,
"Talla": talla,
"Cantidad pares": float(cnt),
                        "Hoja": hoja
                        "Hoja": hoja,
})

            # Bloque lateral B1 (cols 28–39)
if hoja == "B1":
                for inicio in bloques_idx:

                    fila_c2     = inicio      
                    fila_m2     = inicio + 1   
                    fila_color2 = inicio + 3   
                    fila_t2     = inicio       
                    fila_p2     = inicio + 3   
                for inicio2 in bloques_idx:
                    fila_c2     = inicio2
                    fila_m2     = inicio2 + 1
                    fila_color2 = inicio2 + 3
                    fila_t2     = inicio2
                    fila_p2     = inicio2 + 3

cod2 = str(df_raw.iat[fila_c2, 28]).strip()
mod2 = str(df_raw.iat[fila_m2, 28]).strip()
if not cod2 or not mod2:
continue
                    color2 = str(df_raw.iat[fila_color2, 28]).strip() or mod2

                    color2 = str(df_raw.iat[fila_color2, 28]).strip()
                    if not color2:
                        color2 = mod2

                    for col in range(30, 38):
                    for col in range(30, 40):
talla2 = str(df_raw.iat[fila_t2, col]).strip()
cr2    = str(df_raw.iat[fila_p2, col]).strip()
if talla2.startswith("#") and es_numero_valido(cr2):
resultados.append({
"Código del Producto": cod2,
                                "Color":               color2,
                                "Modelo":              mod2,
                                "Talla":               talla2,
                                "Cantidad pares":      float(cr2),
                                "Hoja":                hoja
                                "Color": color2,
                                "Modelo": mod2,
                                "Talla": talla2,
                                "Cantidad pares": float(cr2),
                                "Hoja": hoja,
})


if not resultados:
        return pd.DataFrame(columns=["Código del Producto","Modelo","Talla","Cantidad pares","Hoja"])
        return pd.DataFrame(columns=["Código del Producto","Modelo","Talla","Cantidad pares","Hoja"]) 

df_res = pd.DataFrame(resultados).drop_duplicates(
subset=["Código del Producto","Modelo","Talla","Cantidad pares","Hoja"]
)
df_res["Modelo_norm"] = df_res["Modelo"].str.strip().str.upper()
    df_res["Talla_norm"] = df_res["Talla"].str.strip().str.upper()
    df_res["Talla_norm"]  = df_res["Talla"].str.strip().str.upper()
return df_res

DB_PATH = "usuarios.db"
FICHAS_PATH = "FICHAS2.xlsx"
LOGO_PATH = "logo_suolmex.jpg"
st.set_page_config(page_title="Generador de Pedido SUOLMEX (B1-B4)", layout="wide")
# ===================== Cálculo de totales por código (con merma) =====================
def calcular_resumen_bandas(resumen_df, merma_frac=MERMA_FRAC):
    resumen_df = resumen_df.copy()
    resumen_df["Poliol_kg"] = resumen_df["Poliol (g)"] / 1000
    resumen_df["ISO_kg"]    = resumen_df["ISO (g)"]    / 1000

    agg = (
        resumen_df.groupby("Hoja")
        .agg(
            pares_total=("Cantidad pares", "sum"),
            poliol_necesario_kg=("Poliol_kg", "sum"),
            iso_necesario_kg=("ISO_kg", "sum"),
        )
        .reset_index()
        .rename(columns={"Hoja": "codigo"})
    )

st.markdown("""
    <style>
        body { background-color: #f5f7fa; }
        .stApp { font-family: 'Segoe UI', sans-serif; }
        h1, h2, h3 { color: #00264d; }
        .stButton>button {
            background-color: #00264d;
            color: white;
            font-weight: bold;
            padding: 8px 20px;
            border-radius: 6px;
            border: none;
        }
        .stButton>button:hover { background-color: #003366; }
        .stDataFrame thead tr th {
            background-color: #e6e9ef;
            color: #00264d;
        }
    </style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.image(LOGO_PATH, width=200)
    st.markdown("### Instrucciones")
    st.markdown("""
    1. Sube tu Excel de programación con hojas B1-B4.
    2. Revisa el pedido extraído o ingrésalo manualmente.
    3. Genera el pedido consolidado y descarga el PDF.
    """)
    debug_prog = st.checkbox("Mostrar debug de programación")
    debug_fichas = st.checkbox("Mostrar debug de fichas")
    recargar = st.checkbox("Recargar Excel de fichas (forzar)", key="force_reparse")
    agg["poliol_con_merma_kg"] = agg["poliol_necesario_kg"].apply(con_merma)
    agg["iso_con_merma_kg"]    = agg["iso_necesario_kg"].apply(con_merma)
    agg["mezcla_sin_merma_kg"] = agg["poliol_necesario_kg"] + agg["iso_necesario_kg"]
    agg["mezcla_total_con_merma_kg"] = agg["poliol_con_merma_kg"] + agg["iso_con_merma_kg"]

def obtener_session_id():
    if "session_id" not in st.session_state:
        st.session_state.session_id = uuid.uuid4().hex
    return st.session_state.session_id
    totales = {
        "pares_total": int(agg["pares_total"].sum()),
        "poliol_necesario_kg": agg["poliol_necesario_kg"].sum(),
        "iso_necesario_kg": agg["iso_necesario_kg"].sum(),
        "poliol_con_merma_kg": agg["poliol_con_merma_kg"].sum(),
        "iso_con_merma_kg": agg["iso_con_merma_kg"].sum(),
        "mezcla_sin_merma_kg": agg["mezcla_sin_merma_kg"].sum(),
        "mezcla_total_con_merma_kg": agg["mezcla_total_con_merma_kg"].sum(),
    }
    return agg, totales

def path_sesion_local():
    return f"session_{obtener_session_id()}.json"
# ===================== Generación de PDF =====================
class PDF(FPDF):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._dejavu = False
        self._is_portada = False

    def header(self):
        if self._is_portada:
            return
        try:
            self.image(LOGO_PATH, 10, 7, 20)
        except Exception:
            pass
        set_font(self, size=11, bold=True)
        self.cell(0, 10, "Resumen de Pedido SUOLMEX", 0, 0, 'C')
        self.ln(14)
        hr(self, ypad=0.5, color=(210,210,210), w=0.25)

    def footer(self):
        self.set_y(-17)
        hr(self, ypad=0.5, color=(220,220,220), w=0.2)
        set_font(self, size=8, bold=False)
        self.set_text_color(90,90,90)
        self.cell(0, 5, f"Página {self.page_no()}/{{nb}}", 0, 1, 'C')
        info = getattr(self, "_footer_info", "")
        if info:
            self.cell(0, 5, info, 0, 0, 'C')
        self.set_text_color(0,0,0)

# Helpers PDF (se reutilizan en header/footer)
def try_add_font(pdf, family, style, path):
    try:
        pdf.add_font(family, style, path, uni=True)
        return True
    except Exception:
        return False

def guardar_sesion():
    with open(path_sesion_local(), "w") as f:
        json.dump({
            "logueado": st.session_state.logueado,
            "usuario": st.session_state.usuario,
            "rol": st.session_state.rol
        }, f)
def set_font(pdf, size=10, bold=False):
    fam = "DejaVuLGCSans" if getattr(pdf, "_dejavu", False) else "Helvetica"
    pdf.set_font(fam, "B" if bold else "", size)

def cargar_sesion():
    try:
        with open(path_sesion_local(), "r") as f:
            data = json.load(f)
            st.session_state.logueado = data.get("logueado", False)
            st.session_state.usuario = data.get("usuario", None)
            st.session_state.rol = data.get("rol", None)
    except:
        st.session_state.logueado = False
def hr(pdf, ypad=2, color=(200,200,200), w=0.25):
    pdf.ln(ypad); pdf.set_draw_color(*color); pdf.set_line_width(w)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y()); pdf.ln(ypad)

def need_break(pdf, block_h=8, bottom=15):
    return pdf.get_y() + block_h > (pdf.h - bottom)

def ensure_space(pdf, block_h):
    if need_break(pdf, block_h=block_h):
        pdf.add_page()

def tabla(pdf, headers, rows, widths, aligns=None,
          header_fill=(60,60,60), header_text=(255,255,255),
          zebra_1=(245,245,245), zebra_2=(255,255,255),
          row_h=8, font_size=10):
    if aligns is None: aligns = ["L"] * len(headers)
    # Header
    set_font(pdf, size=font_size, bold=True)
    pdf.set_fill_color(*header_fill); pdf.set_text_color(*header_text)
    for t, w, a in zip(headers, widths, aligns):
        pdf.cell(w, row_h, str(t), border=1, ln=0, align=a, fill=True)
    pdf.ln(row_h)
    # Rows
    set_font(pdf, size=font_size, bold=False); pdf.set_text_color(0,0,0)
    fill = False
    for r in rows:
        if need_break(pdf, row_h):
            pdf.add_page()
            # redraw header
            set_font(pdf, size=font_size, bold=True)
            pdf.set_fill_color(*header_fill); pdf.set_text_color(*header_text)
            for t, w, a in zip(headers, widths, aligns):
                pdf.cell(w, row_h, str(t), border=1, ln=0, align=a, fill=True)
            pdf.ln(row_h)
            set_font(pdf, size=font_size, bold=False); pdf.set_text_color(0,0,0)
        pdf.set_fill_color(*(zebra_1 if fill else zebra_2))
        for val, w, a in zip(r, widths, aligns):
            pdf.cell(w, row_h, str(val), border=1, ln=0, align=a, fill=True)
        pdf.ln(row_h)
        fill = not fill

def cuadro_info(pdf, data, titulo=None):
    if titulo:
        set_font(pdf, size=11, bold=True)
        ensure_space(pdf, 10)
        pdf.cell(0, 8, titulo, ln=1)
    set_font(pdf, size=10)
    w_key, w_val, row_h = 85, 95, 8
    ensure_space(pdf, row_h*len(data)+2)
    pdf.set_fill_color(245,245,245)
    for k, v in data:
        pdf.cell(w_key, row_h, str(k), border=1, ln=0, align="L", fill=True)
        pdf.cell(w_val, row_h, str(v), border=1, ln=1, align="R", fill=True)
    pdf.ln(2)

def titulo_seccion(pdf, texto):
    set_font(pdf, size=14, bold=True)
    ensure_space(pdf, 14)
    pdf.cell(0, 10, texto, ln=1)
    hr(pdf, ypad=2, color=(210,210,210), w=0.25)

def agg_por_color(df):
    if df.empty: 
        return df
    out = (
        df.groupby("Color", dropna=False)
          .agg(
              pares_total=("Cantidad pares", "sum"),
              poliol_kg=("Poliol (g)", lambda s: s.sum()/1000),
              iso_kg=("ISO (g)",       lambda s: s.sum()/1000),
          ).reset_index()
    )
    out["Color"] = out["Color"].fillna("SIN COLOR")
    return out

def cols_modelo_talla(df):
    modelo_col = next((c for c in ["Modelo","Linea"] if c in df.columns), None)
    talla_col  = next((c for c in ["Talla","Corrida"] if c in df.columns), None)
    return modelo_col, talla_col

def detalle_por_color(pdf, df_section, color):
    modelo_col, talla_col = cols_modelo_talla(df_section)
    if not modelo_col or not talla_col:
        return
    df_color = df_section[df_section["Color"] == color].copy()
    if df_color.empty:
        return

    df_color["_poliol_kg"] = df_color["Poliol (g)"] / 1000.0
    df_color["_iso_kg"]    = df_color["ISO (g)"]    / 1000.0

    g = (
        df_color.groupby([modelo_col, talla_col], dropna=False)
                .agg(pares=("Cantidad pares", "sum"),
                     pol=("_poliol_kg", "sum"),
                     iso=("_iso_kg",    "sum"))
                .reset_index()
    )
    g[modelo_col] = g[modelo_col].fillna("SIN MODELO")
    g[talla_col]  = g[talla_col].fillna("SIN TALLA")

    rows = []
    for _, rr in g.iterrows():
        pol_m = con_merma(float(rr["pol"]))
        iso_m = con_merma(float(rr["iso"]))
        rows.append([
            str(rr[modelo_col]), str(rr[talla_col]),
            fmt_ent(rr["pares"]), fmt_num(pol_m), fmt_num(iso_m), fmt_num(pol_m+iso_m)
        ])

    set_font(pdf, size=11, bold=True)
    ensure_space(pdf, 10)
    pdf.cell(0, 7, f"Detalle por modelo/talla — {color}", ln=1)

    tabla(
        pdf,
        headers=["Modelo", "Talla", "Pares", "Poliol", "ISO", "Mezcla"],
        rows=rows,
        widths=[62, 22, 18, 28, 28, 28],
        aligns=["L","C","R","R","R","R"],
        row_h=7,
        font_size=9,
    )
    pdf.ln(1)

def seccion_pdf(pdf, titulo, df_grouped, df_raw, global_totales):
    # Blindaje por si llega algo que no es DataFrame (evita el error 'tuple'.empty)
    if not isinstance(df_grouped, pd.DataFrame):
        df_grouped = pd.DataFrame()

    titulo_seccion(pdf, titulo)
    if df_grouped.empty:
        set_font(pdf)
        pdf.cell(0, 8, "Sin registros", ln=1)
        pdf.ln(3)
        return

    rows, tpares, tpol, tiso = [], 0, 0.0, 0.0
    for _, r in df_grouped.iterrows():
        color = str(r["Color"]) 
        pares = int(r["pares_total"]) 
        pol   = float(r["poliol_kg"]) 
        iso   = float(r["iso_kg"]) 
        pol_m, iso_m = con_merma(pol), con_merma(iso)
        rows.append([color, fmt_ent(pares), fmt_num(pol_m), fmt_num(iso_m), fmt_num(pol_m+iso_m)])
        tpares += pares; tpol += pol; tiso += iso

    ensure_space(pdf, 10 + 8*(len(rows)+1))
    tabla(
        pdf,
        headers=["Color", "Pares", "Poliol", "ISO", "Mezcla"],
        rows=rows,
        widths=[60,25,35,35,35],
        aligns=['L','R','R','R','R']
    )

conn = sqlite3.connect(DB_PATH, check_same_thread=False)
c = conn.cursor()
c.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT UNIQUE NOT NULL,
        contrasena TEXT NOT NULL,
        rol TEXT NOT NULL
    if DETALLADO:
        for _, r in df_grouped.iterrows():
            detalle_por_color(pdf, df_raw, str(r["Color"]))

    pol_m, iso_m = con_merma(tpol), con_merma(tiso)
    cuadro_info(pdf, [
        ("Pares", fmt_ent(tpares)),
        ("Poliol (c/merma)", f"{fmt_num(pol_m)} {UNIDADES}"),
        ("ISO (c/merma)",    f"{fmt_num(iso_m)} {UNIDADES}"),
        ("Mezcla (c/merma)", f"{fmt_num(pol_m+iso_m)} {UNIDADES}"),
    ], f"Totales {titulo}")

    global_totales["pares"]    += tpares
    global_totales["poliol"]   += tpol
    global_totales["iso"]      += tiso
    global_totales["poliol_m"] += pol_m
    global_totales["iso_m"]    += iso_m

def generar_pdf(resumen_df, usuario, fecha_hora, nombre_archivo):
    pdf = PDF(orientation='P', unit='mm', format='A4')
    pdf._dejavu = (
        try_add_font(pdf, "DejaVuLGCSans", "",  "DejaVuLGCSans.ttf") and
        try_add_font(pdf, "DejaVuLGCSans", "B", "DejaVuLGCSans-Bold.ttf")
)
""")
conn.commit()
    pdf.alias_nb_pages()

    # Portada
    pdf._is_portada = True
    pdf.add_page()
    try:
        pdf.image(LOGO_PATH, x=10, y=14, w=28)
    except Exception:
        pass
    set_font(pdf, size=18, bold=True)
    pdf.set_xy(0, 22); pdf.cell(0, 12, "Resumen de Pedido SUOLMEX", ln=1, align='C')

    cuadro_info(pdf, [
        ("Usuario", str(usuario)),
        ("Fecha y hora", fecha_hora.replace("_"," ")),
        ("Merma", f"{MERMA_FRAC*100:.0f}%"),
        ("Unidades", UNIDADES),
    ])

    set_font(pdf, size=9)
    pdf.multi_cell(0, 6, "Desglose de pares y requerimientos de Poliol/ISO por sección y color. "
                      "Incluye totales y detalle por modelo/talla (si aplica).")
    hr(pdf, ypad=3)
    pdf._is_portada = False

    # --- Datos base
    df_suelas = resumen_df[resumen_df["Hoja"].isin(['6001','2066','2060','4098'])].copy()
    df_plant  = resumen_df[resumen_df["Hoja"].astype(str) == "PLANTILLAS"].copy()

    # --- CORRECCIÓN: esta función local ahora SIEMPRE devuelve (df_raw, df_grouped) como DataFrames
    def preparar(df):
        if df.empty:
            return df, df
        else:
            return df, agg_por_color(df)

def encriptar_contra(contra):
    return hashlib.sha256(contra.encode()).hexdigest()
    df_suelas_raw, g_suelas = preparar(df_suelas)
    df_plant_raw,  g_plant  = preparar(df_plant)

if not c.execute("SELECT * FROM usuarios WHERE codigo = 'admin'").fetchone():
    c.execute("INSERT INTO usuarios (codigo, contrasena, rol) VALUES (?, ?, ?)",
              ('admin', encriptar_contra('admin123'), 'admin'))
    conn.commit()
    # --- Acumuladores globales
    tot_global = dict(pares=0, poliol=0.0, iso=0.0, poliol_m=0.0, iso_m=0.0)

    # --- Secciones
    seccion_pdf(pdf, "Suelas", g_suelas, df_suelas_raw, tot_global)
    seccion_pdf(pdf, "Plantillas", g_plant, df_plant_raw, tot_global)

    # --- Resumen Global
    if need_break(pdf, block_h=55):
        pdf.add_page()

    set_font(pdf, size=14, bold=True)
    pdf.cell(0, 10, "Resumen Global", ln=1)
    hr(pdf, ypad=2)

if "logueado" not in st.session_state:
    cargar_sesion()
    cuadro_info(pdf, [
        ("Pares totales", fmt_ent(tot_global["pares"])),
        ("Poliol (c/merma)",  f"{fmt_num(tot_global['poliol_m'])} {UNIDADES}"),
        ("ISO (c/merma)",     f"{fmt_num(tot_global['iso_m'])} {UNIDADES}"),
        ("Mezcla (c/merma)",  f"{fmt_num(tot_global['poliol_m']+tot_global['iso_m'])} {UNIDADES}"),
    ])

st.session_state.setdefault("logueado", False)
st.session_state.setdefault("usuario", "")
st.session_state.setdefault("rol", "")
    set_font(pdf, size=9)
    pdf.set_text_color(90,90,90)
    pdf.cell(0, 6, f"Merma del {MERMA_FRAC*100:.0f}% aplicada a Poliol e ISO.", ln=1)
    pdf.set_text_color(0,0,0)

if not st.session_state.get("logueado", False):
    pdf._footer_info = f"Usuario: {usuario}   •   Fecha: {fecha_hora.replace('_',' ')}"

    pdf.output(nombre_archivo)


# ===================== UI Helpers =====================
STYLES = """
    <style>
        body { background-color: #f5f7fa; }
        .stApp { font-family: 'Segoe UI', sans-serif; }
        h1, h2, h3 { color: #00264d; }
        .stButton>button {
            background-color: #00264d; color: white; font-weight: bold;
            padding: 8px 20px; border-radius: 6px; border: none;
        }
        .stButton>button:hover { background-color: #003366; }
        .stDataFrame thead tr th { background-color: #e6e9ef; color: #00264d; }
        .metric-card { background:#fff; border:1px solid #e6e9ef; border-radius:10px; padding:12px; }
        .pill { background:#eef2f7; border-radius:999px; padding:2px 10px; }
    </style>
"""

def header_sidebar():
    with st.sidebar:
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=200)
        st.markdown("### Instrucciones")
        st.markdown("""
        1. Sube tu Excel de programación con hojas B1–B4.
        2. Revisa el pedido extraído o ingrésalo manualmente.
        3. Genera el pedido consolidado y descarga el PDF.
        """)
        st.session_state.setdefault("debug_prog", False)
        st.session_state.setdefault("debug_fichas", False)
        st.checkbox("Mostrar debug de programación", key="debug_prog")
        st.checkbox("Mostrar debug de fichas", key="debug_fichas")
        st.checkbox("Recargar Excel de fichas (forzar)", key="force_reparse")

# ===================== Vistas =====================
def login_view(c, conn):
st.subheader("Iniciar sesión")
with st.form("login_form"):
codigo = st.text_input("Código de usuario")
contrasena = st.text_input("Contraseña", type="password")
if st.form_submit_button("Entrar"):
            user = c.execute("SELECT contrasena, rol FROM usuarios WHERE codigo = ?", (codigo,)).fetchone()
            if user and encriptar_contra(contrasena) == user[0]:
            row = c.execute("SELECT contrasena, rol FROM usuarios WHERE codigo=?", (codigo,)).fetchone()
            if row and hash_password(contrasena) == row[0]:
st.session_state.logueado = True
                st.session_state.usuario = codigo
                st.session_state.rol = user[1]
                st.session_state.usuario  = codigo
                st.session_state.rol      = row[1]
guardar_sesion()
st.success("Sesión iniciada correctamente.")
else:
st.error("Credenciales incorrectas.")
st.stop()

if st.button("Cerrar sesión"):
    ruta = path_sesion_local()
    if Path(ruta).exists():
        Path(ruta).unlink()

    st.session_state["logueado"] = False
    st.session_state["usuario"] = ""
    st.session_state["rol"] = ""
    st.success("Sesión cerrada.")


    def refresh_url():
        if hasattr(st, "query_params"):
            params = dict(st.query_params)
            params["_r"] = str(uuid.uuid4())
            st.query_params.clear()
            st.query_params.update(params)
        else:
            try:
                params = st.experimental_get_query_params()
            except AttributeError:
                params = {}
            params["_r"] = str(uuid.uuid4())
            st.experimental_set_query_params(**params)

    refresh_url()


if st.session_state.get("logueado", False):
    st.success(f"Sesión iniciada como **{st.session_state.usuario}** ({st.session_state.rol})")

with st.expander("Historial de pedidos generados"):
    folder = Path("historial_pedidos")
    if folder.exists():
        archivos_pdf = sorted(folder.glob("*.pdf"), reverse=True)
        if archivos_pdf:
            for archivo in archivos_pdf:
                with open(archivo, "rb") as f:
                    st.download_button(
                        label=f" {archivo.name}",
                        data=f,
                        file_name=archivo.name,
                        mime="application/pdf",
                        key=archivo.name
                    )
def logout_button():
    if st.button("Cerrar sesión"):
        p = Path(_session_path())
        if p.exists():
            p.unlink(missing_ok=True)
        st.session_state["logueado"] = False
        st.session_state["usuario"]  = ""
        st.session_state["rol"]      = ""
        st.success("Sesión cerrada.")
        # Fuerza recarga de URL para limpiar estado
        try:
            params = st.query_params.to_dict() if hasattr(st, "query_params") else st.experimental_get_query_params()
            params["_r"] = uuid.uuid4().hex
            if hasattr(st, "query_params"):
                st.query_params.clear(); st.query_params.update(params)
            else:
                st.experimental_set_query_params(**params)
        except Exception:
            pass

# ===================== App =====================
def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown(STYLES, unsafe_allow_html=True)
    header_sidebar()

    # DB y sesión
    conn, c = init_db()
    ensure_admin(c, conn)

    if "logueado" not in st.session_state:
        cargar_sesion()
    st.session_state.setdefault("logueado", False)
    st.session_state.setdefault("usuario", "")
    st.session_state.setdefault("rol", "")

    if not st.session_state.get("logueado", False):
        login_view(c, conn)

    # Barra superior
    colA, colB = st.columns([0.7, 0.3])
    with colA:
        st.title("Explosión de producto SUOLMEX")
        st.success(f"Sesión: **{st.session_state.usuario}** ({st.session_state.rol})")
    with colB:
        logout_button()

    # Historial PDFs
    with st.expander("Historial de pedidos generados"):
        folder = Path("historial_pedidos")
        if folder.exists():
            archivos_pdf = sorted(folder.glob("*.pdf"), reverse=True)
            if archivos_pdf:
                for archivo in archivos_pdf:
                    with open(archivo, "rb") as f:
                        st.download_button(
                            label=f" {archivo.name}", data=f, file_name=archivo.name,
                            mime="application/pdf", key=f"dl_{archivo.name}"
                        )
            else:
                st.info("No hay PDFs generados todavía.")
else:
            st.info("No hay PDFs generados todavía.")
    else:
        st.info("No se ha generado ningún pedido aún.")

if st.session_state.get("rol") == "admin":
    st.markdown("---")
    st.subheader("Gestión de Usuarios")
    with st.expander("Crear nuevo usuario"):
        with st.form("crear_user"):
            nuevo = st.text_input("Nuevo código de usuario")
            contra = st.text_input("Contraseña", type="password")
            rol = st.selectbox("Rol", ["admin", "empleado"])
            if st.form_submit_button("Crear"):
                try:
                    c.execute("INSERT INTO usuarios (codigo, contrasena, rol) VALUES (?, ?, ?)",
                              (nuevo, encriptar_contra(contra), rol))
            st.info("No se ha generado ningún pedido aún.")

    # Gestión de usuarios (admin)
    if st.session_state.get("rol") == "admin":
        st.markdown("---")
        st.subheader("Gestión de Usuarios")
        with st.expander("Crear nuevo usuario"):
            with st.form("crear_user"):
                nuevo = st.text_input("Nuevo código de usuario")
                contra = st.text_input("Contraseña", type="password")
                rol = st.selectbox("Rol", ["admin", "empleado"])
                if st.form_submit_button("Crear"):
                    try:
                        c.execute("INSERT INTO usuarios (codigo, contrasena, rol) VALUES (?, ?, ?)",
                                  (nuevo, hash_password(contra), rol))
                        conn.commit()
                        st.success("Usuario creado.")
                    except Exception:
                        st.error("Ese código ya existe.")
        with st.expander("Editar o eliminar usuarios"):
            usuarios = pd.read_sql("SELECT codigo, rol FROM usuarios", conn)
            st.dataframe(usuarios, use_container_width=True)
            if not usuarios.empty:
                editar = st.selectbox("Selecciona usuario", usuarios["codigo"])
                nueva_contra = st.text_input("Nueva contraseña", type="password")
                cols = st.columns(2)
                if cols[0].button("Actualizar contraseña"):
                    c.execute("UPDATE usuarios SET contrasena=? WHERE codigo=?",
                              (hash_password(nueva_contra), editar))
conn.commit()
                    st.success("Usuario creado.")
                except:
                    st.error("Ese código ya existe.")
    with st.expander("Editar o eliminar usuarios"):
        usuarios = pd.read_sql("SELECT codigo, rol FROM usuarios", conn)
        st.dataframe(usuarios)
        editar = st.selectbox("Selecciona usuario", usuarios["codigo"])
        nueva_contra = st.text_input("Nueva contraseña", type="password")
        if st.button("Actualizar contraseña"):
            c.execute("UPDATE usuarios SET contrasena = ? WHERE codigo = ?",
                      (encriptar_contra(nueva_contra), editar))
            conn.commit()
            st.success("Contraseña actualizada.")
        if editar != "admin" and st.button("Eliminar usuario"):
            c.execute("DELETE FROM usuarios WHERE codigo = ?", (editar,))
            conn.commit()
            st.warning("Usuario eliminado.")

@st.cache_data
def cargar_fichas(mtime: float):
    excel_file = pd.ExcelFile(FICHAS_PATH)
    hojas_deseadas = ['6001', '2066', '2060', '4098', 'PLANTILLAS']
    dataframes = []
    for hoja in hojas_deseadas:
        if hoja in excel_file.sheet_names:
            df_temp = excel_file.parse(hoja)
            df_temp['Hoja'] = hoja
            dataframes.append(df_temp)
    if not dataframes:
        return pd.DataFrame()
    df = pd.concat(dataframes, ignore_index=True)
    df["Codigo del Producto"] = df["Codigo del Producto"].astype(str).str.strip()
    df["Linea"] = df["Linea"].astype(str).str.strip().str.upper()
    df["Corrida"] = df["Corrida"].astype(str).str.strip()
    df["Peso/Pie"] = pd.to_numeric(df["Peso/Pie"], errors="coerce")
    df["__Linea_clean"] = df["Linea"].apply(clean_key)
    df["__Corrida_clean"] = df["Corrida"].apply(clean_key)
    return df.dropna(subset=["Peso/Pie", "Relacion Poliol:ISO"])

if recargar:
    st.session_state["fichas_reload_flag"] = not st.session_state.get("fichas_reload_flag", False)

mtime = os.path.getmtime(FICHAS_PATH) if os.path.exists(FICHAS_PATH) else 0
fichas = cargar_fichas(mtime)

if debug_fichas:
    st.subheader("Debug: fichas cargadas (muestra)")
    st.dataframe(fichas[["Codigo del Producto", "Linea", "Corrida", "__Linea_clean", "__Corrida_clean"]].drop_duplicates().head(100))

if "pedido_total" not in st.session_state:
    st.session_state["pedido_total"] = []
if "corrida_seleccionada" not in st.session_state:
    st.session_state["corrida_seleccionada"] = None

st.markdown("---")
st.title("Generador de Pedido SUOLMEX (B1–B4 consolidado)")
tab_excel, tab_manual = st.tabs(["Desde Excel", "Manual"])

with tab_excel:
    uploaded_prog = st.file_uploader("Sube tu Excel de programación", type=["xlsx"], key="prog_excel_tab")
    codigo_producto_override = st.text_input("Código del Producto (override si no lo detecta)", value="", key="override_excel")

    pedido_programado = pd.DataFrame()
    if uploaded_prog:
        st.info("Interpretando archivo de programación...")
        t0 = time.time()
        pedido_programado = extract_programacion_estatica_B1_B4(
            uploaded_prog,
            codigo_producto_override or None,
            hojas_objetivo=["B1", "B2", "B3", "B4"],
            debug=debug_prog
                    st.success("Contraseña actualizada.")
                if editar != "admin" and cols[1].button("Eliminar usuario"):
                    c.execute("DELETE FROM usuarios WHERE codigo=?", (editar,))
                    conn.commit()
                    st.warning("Usuario eliminado.")

    # Fichas
    if st.session_state.get("force_reparse", False):
        st.session_state["fichas_reload_flag"] = not st.session_state.get("fichas_reload_flag", False)
    mtime = os.path.getmtime(FICHAS_PATH) if os.path.exists(FICHAS_PATH) else 0
    fichas = cargar_fichas(mtime)

    if st.session_state.get("debug_fichas"):
        st.subheader("Debug: fichas cargadas (muestra)")
        st.dataframe(
            fichas[["Codigo del Producto","Linea","Corrida","__Linea_clean","__Corrida_clean"]]
                  .drop_duplicates()
                  .head(120),
            use_container_width=True
)
        st.write(f"Tiempo de parseo: {time.time() - t0:.3f}s")

        st.subheader("Pedido extraído (debug)")
        st.dataframe(pedido_programado)
    # Estado de pedido
    st.session_state.setdefault("pedido_total", [])
    st.session_state.setdefault("corrida_seleccionada", None)

    # Tabs
    tab_excel, tab_manual = st.tabs(["Desde Excel", "Manual"])

    # -------- Tab Excel --------
    with tab_excel:
        uploaded_prog = st.file_uploader("Sube tu Excel de programación", type=["xlsx"], key="prog_excel_tab")
        codigo_override = st.text_input("Código del Producto (override si no lo detecta)", value="", key="override_excel")

        pedido_programado = pd.DataFrame()
        if uploaded_prog:
            st.info("Interpretando archivo de programación…")
            t0 = time.time()
            pedido_programado = extract_programacion_estatica_B1_B4(
                uploaded_prog, codigo_override or None,
                hojas_objetivo=["B1","B2","B3","B4"],
                debug=st.session_state.get("debug_prog")
            )
            st.caption(f"Tiempo de parseo: {time.time() - t0:.3f}s")

        if pedido_programado.empty:
            st.warning("No se extrajeron líneas válidas de las hojas objetivo. Verifica el archivo y la nomenclatura de hojas.")
        else:
            for _, row in pedido_programado.iterrows():
                codigo = str(row["Código del Producto"]).strip()
                modelo = str(row["Modelo"]).strip().upper()
                corrida = str(row["Talla"]).strip()
                color = str(row["Color"]).strip()
                try:
                    cantidad = int(float(row["Cantidad pares"]))
                except:
                    continue
                if cantidad <= 0:
                    continue
            if st.session_state.get("debug_prog"):
                st.subheader("Pedido extraído (debug)")
                st.dataframe(pedido_programado, use_container_width=True)

                modelo_norm = clean_key(modelo)
                corrida_norm = clean_key(corrida)
            if pedido_programado.empty:
                st.warning("No se extrajeron líneas válidas. Verifica las hojas y la nomenclatura.")
            else:
                # Merge con fichas -> Explosión
                for _, row in pedido_programado.iterrows():
                    codigo = str(row["Código del Producto"]).strip()
                    modelo = str(row["Modelo"]).strip().upper()
                    corrida = str(row["Talla"]).strip()
                    color   = str(row.get("Color", "")).strip()
                    try:
                        cantidad = int(float(row["Cantidad pares"]))
                    except Exception:
                        continue
                    if cantidad <= 0:
                        continue

                    modelo_norm  = clean_key(modelo)
                    corrida_norm = clean_key(corrida)

                ficha = fichas[
                    (fichas["Codigo del Producto"].astype(str).str.strip() == codigo) &
                    (fichas["__Linea_clean"] == modelo_norm) &
                    (fichas["__Corrida_clean"] == corrida_norm)
                    ficha = fichas[
                        (fichas["Codigo del Producto"].astype(str).str.strip() == codigo) &
                        (fichas["__Linea_clean"]   == modelo_norm) &
                        (fichas["__Corrida_clean"] == corrida_norm)
                    ]
                    if not ficha.empty:
                        f = ficha.iloc[0]
                        peso_total = f['Peso/Pie'] * cantidad * 2
                        try:
                            pol_str, iso_str = f['Relacion Poliol:ISO'].split(":")
                            pol, iso = float(pol_str), float(iso_str)
                        except Exception:
                            pol, iso = 0.0, 0.0
                        total = (pol + iso) if (pol + iso) != 0 else 1
                        cant_pol = peso_total * (pol / total)
                        cant_iso = peso_total * (iso / total)

                        nuevo = {
                            "uid": uuid.uuid4().hex,
                            "Código": codigo,
                            "Modelo": modelo,
                            "Color":  color or modelo,
                            "Talla":  corrida,
                            "Cantidad pares": cantidad,
                            "Peso Total (g)": peso_total,
                            "Poliol (g)": cant_pol,
                            "ISO (g)":    cant_iso,
                            "Hoja": f['Hoja'],
                        }

                        existe = any(
                            it.get("Código") == nuevo["Código"] and
                            it.get("Modelo") == nuevo["Modelo"] and
                            it.get("Talla")  == nuevo["Talla"]  and
                            it.get("Cantidad pares") == nuevo["Cantidad pares"]
                            for it in st.session_state["pedido_total"]
                        )
                        if not existe:
                            st.session_state["pedido_total"].append(nuevo)
                    else:
                        st.warning(f"Sin ficha para código='{codigo}', modelo='{modelo}', talla='{corrida}'.")
                st.success("Explosión automática agregada al pedido.")

    # -------- Tab Manual --------
    with tab_manual:
        st.subheader("Ingreso manual")
        if fichas.empty:
            st.warning("No se pudo cargar FICHAS2.xlsx")
        else:
            codigo_manual = st.selectbox("Código del Producto:", sorted(fichas["Codigo del Producto"].unique()), key="manual_codigo")
            modelos = fichas[fichas["Codigo del Producto"] == codigo_manual]["Linea"].unique()
            modelo_manual = st.selectbox("Modelo:", sorted(modelos), key="manual_modelo")
            cantidad = st.number_input("Cantidad de pares:", min_value=1, step=1, key="manual_cantidad")

            corridas = fichas[
                (fichas["Codigo del Producto"] == codigo_manual) &
                (fichas["Linea"] == modelo_manual)
            ]["Corrida"].unique()
            st.markdown("#### Selecciona una talla:")
            cols = st.columns(min(5, max(1, len(corridas))))
            for i, talla in enumerate(sorted(corridas)):
                if cols[i % 5].button(str(talla), key=f"manual_talla_{i}"):
                    st.session_state["corrida_seleccionada"] = talla

            if st.session_state.get("corrida_seleccionada"):
                corrida = st.session_state["corrida_seleccionada"]
                ficha_manual = fichas[
                    (fichas["Codigo del Producto"] == codigo_manual) &
                    (fichas["Linea"] == modelo_manual) &
                    (fichas["Corrida"] == corrida)
]
                if not ficha.empty:
                    ficha = ficha.iloc[0]
                    peso_total = ficha['Peso/Pie'] * cantidad * 2
                if not ficha_manual.empty:
                    f = ficha_manual.iloc[0]
                    peso_total = f['Peso/Pie'] * cantidad * 2
try:
                        poliol_str, iso_str = ficha['Relacion Poliol:ISO'].split(":")
                        poliol = float(poliol_str)
                        iso = float(iso_str)
                    except:
                        poliol, iso = 0.0, 0.0
                    total_partes = (poliol + iso) if (poliol + iso) != 0 else 1
                    cantidad_poliol = peso_total * (poliol / total_partes)
                    cantidad_iso = peso_total * (iso / total_partes)
                        pol_str, iso_str = f['Relacion Poliol:ISO'].split(":")
                        pol, iso = float(pol_str), float(iso_str)
                    except Exception:
                        pol, iso = 0.0, 0.0
                    total = (pol + iso) if (pol + iso) != 0 else 1
                    cant_pol = peso_total * (pol / total)
                    cant_iso = peso_total * (iso / total)

nuevo = {
"uid": uuid.uuid4().hex,
                        "Código": codigo,
                        "Modelo": modelo,
                        "Color": color, 
                        "Talla": corrida,
                        "Código": codigo_manual,
                        "Modelo": modelo_manual,
                        "Color":  modelo_manual,
                        "Talla":  corrida,
"Cantidad pares": cantidad,
"Peso Total (g)": peso_total,
                        "Poliol (g)": cantidad_poliol,
                        "ISO (g)": cantidad_iso,
                        "Hoja": ficha['Hoja']
                        "Poliol (g)": cant_pol,
                        "ISO (g)":    cant_iso,
                        "Hoja": f['Hoja'],
}

existe = any(
                        item.get("Código") == nuevo["Código"] and
                        item.get("Modelo") == nuevo["Modelo"] and
                        item.get("Talla") == nuevo["Talla"] and
                        item.get("Cantidad pares") == nuevo["Cantidad pares"]
                        for item in st.session_state["pedido_total"]
                        it.get("Código") == nuevo["Código"] and
                        it.get("Modelo") == nuevo["Modelo"] and
                        it.get("Talla")  == nuevo["Talla"]  and
                        it.get("Cantidad pares") == nuevo["Cantidad pares"]
                        for it in st.session_state["pedido_total"]
)
if not existe:
st.session_state["pedido_total"].append(nuevo)
                else:
                    st.warning(f"No se encontró ficha para código='{codigo}', modelo='{modelo}', talla='{corrida}'.")

            st.success("Explosión automática agregada al pedido.")

with tab_manual:
    st.subheader("Ingreso manual")
    codigo_manual = st.selectbox("Código del Producto:", sorted(fichas["Codigo del Producto"].unique()), key="manual_codigo")
    modelos = fichas[fichas["Codigo del Producto"] == codigo_manual]["Linea"].unique()
    modelo_manual = st.selectbox("Modelo:", sorted(modelos), key="manual_modelo")
    cantidad = st.number_input("Cantidad de pares:", min_value=1, step=1, key="manual_cantidad")
    corridas = fichas[
        (fichas["Codigo del Producto"] == codigo_manual) &
        (fichas["Linea"] == modelo_manual)
    ]["Corrida"].unique()
    st.markdown("#### Selecciona una talla:")
    cols = st.columns(min(5, len(corridas)))
    for i, talla in enumerate(sorted(corridas)):
        if cols[i % 5].button(talla, key=f"manual_talla_{i}"):
            st.session_state["corrida_seleccionada"] = talla

    if st.session_state.get("corrida_seleccionada"):
        corrida = st.session_state["corrida_seleccionada"]
        ficha_manual = fichas[
            (fichas["Codigo del Producto"] == codigo_manual) &
            (fichas["Linea"] == modelo_manual) &
            (fichas["Corrida"] == corrida)
        ]
        if not ficha_manual.empty:
            ficha_manual = ficha_manual.iloc[0]
            peso_total = ficha_manual['Peso/Pie'] * cantidad * 2
            try:
                poliol_str, iso_str = ficha_manual['Relacion Poliol:ISO'].split(":")
                poliol = float(poliol_str)
                iso = float(iso_str)
            except:
                poliol, iso = 0.0, 0.0
            total_partes = (poliol + iso) if (poliol + iso) != 0 else 1
            cantidad_poliol = peso_total * (poliol / total_partes)
            cantidad_iso = peso_total * (iso / total_partes)
            nuevo = {
                "uid": uuid.uuid4().hex,
                "Código": codigo_manual,
                "Color": modelo_manual,
                "Modelo": modelo_manual,
                "Talla": corrida,
                "Cantidad pares": cantidad,
                "Peso Total (g)": peso_total,
                "Poliol (g)": cantidad_poliol,
                "ISO (g)": cantidad_iso,
                "Hoja": ficha_manual['Hoja']
            }
            existe = any(
                item.get("Código") == nuevo["Código"] and
                item.get("Modelo") == nuevo["Modelo"] and
                item.get("Talla") == nuevo["Talla"] and
                item.get("Cantidad pares") == nuevo["Cantidad pares"]
                for item in st.session_state["pedido_total"]
            )
            if not existe:
                st.session_state["pedido_total"].append(nuevo)
            st.success(f"Agregado: {modelo_manual} - Talla {corrida} - {cantidad} pares.")
            st.session_state["corrida_seleccionada"] = None

if st.session_state["pedido_total"]:
    st.markdown("---")
    st.subheader("Resumen del Pedido")

    resumen_df = pd.DataFrame(st.session_state["pedido_total"])

    df_bandas, totales = calcular_resumen_bandas(resumen_df)

    for row in resumen_df.to_dict("records"):
        cols = st.columns([5, 1])
        with cols[0]:
            st.markdown(
                f"**Código:** {row['Código']} | **Modelo:** {row['Modelo']} | "
                f"**Talla:** {row['Talla']} | **Cantidad:** {row['Cantidad pares']} pares | "
                f"**Poliol:** {row['Poliol (g)']:.2f} g | **ISO:** {row['ISO (g)']:.2f} g | "
                f"**Banda:** {row.get('Hoja','')}"
            )
        with cols[1]:
            uid = row.get("uid")
            key = f"eliminar_{uid}" if uid else f"eliminar_{row.get('Código','')}_{row.get('Modelo','')}_{row.get('Talla','')}_{row.get('Cantidad pares','')}"
            if st.button("Eliminar", key=key):
                if uid:
                    st.session_state["pedido_total"] = [
                        item for item in st.session_state["pedido_total"] if item.get("uid") != uid
                    ]
                else:
                    to_remove = None
                    for item in st.session_state["pedido_total"]:
                        if (item.get("Código") == row.get("Código") and
                            item.get("Modelo") == row.get("Modelo") and
                            item.get("Talla") == row.get("Talla") and
                            item.get("Cantidad pares") == row.get("Cantidad pares")):
                            to_remove = item
                            break
                    if to_remove:
                        st.session_state["pedido_total"].remove(to_remove)
                st.success("Elemento eliminado.")
                st.experimental_rerun()

    st.markdown("---")
    st.subheader("Resumen por código")
    for _, row in df_bandas.iterrows():
        st.markdown(f"**Código {row['codigo']}**")
        st.markdown(f"- Pares totales: {int(row['pares_total'])} pares")
        st.markdown(f"- Poliol necesario: {row['poliol_necesario_kg']:.2f} kg")
        st.markdown(f"- ISO necesario: {row['iso_necesario_kg']:.2f} kg")
        st.markdown(f"- Poliol con merma (3%): {row['poliol_con_merma_kg']:.2f} kg")
        st.markdown(f"- ISO con merma (3%): {row['iso_con_merma_kg']:.2f} kg")
        st.markdown(f"- Mezcla sin merma: {row['mezcla_sin_merma_kg']:.2f} kg")
        st.markdown(f"- Mezcla total con merma: {row['mezcla_total_con_merma_kg']:.2f} kg")
        st.markdown("")


    st.subheader("Totales generales")
    st.markdown(f"- **Poliol total necesario:** {totales['poliol_necesario_kg']:.2f} kg")
    st.markdown(f"- **Pares total:** {int(totales['pares_total'])} pares")
    st.markdown(f"- **ISO total necesario:** {totales['iso_necesario_kg']:.2f} kg")
    st.markdown(f"- **Poliol total con merma (3%):** {totales['poliol_con_merma_kg']:.2f} kg")
    st.markdown(f"- **ISO total con merma (3%):** {totales['iso_con_merma_kg']:.2f} kg")
    st.markdown(f"- **Mezcla sin merma total:** {totales['mezcla_sin_merma_kg']:.2f} kg")
    st.markdown(f"- **Mezcla total con merma:** {totales['mezcla_total_con_merma_kg']:.2f} kg")

resumen_df = pd.DataFrame(st.session_state["pedido_total"])

col1, col2 = st.columns(2)

with col1:
    if st.button("Generar PDF", key="pdf_generar"):
    
        Path("historial_pedidos").mkdir(exist_ok=True)
        fecha_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        usuario = st.session_state.usuario
        nombre_archivo = f"historial_pedidos/pedido_{usuario}_{fecha_hora}.pdf"

        class PDF(FPDF):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.add_font('DejaVuLGCSans', '', 'DejaVuLGCSans.ttf', uni=True)
                self.add_font('DejaVuLGCSans', 'B', 'DejaVuLGCSans-Bold.ttf', uni=True)

            def header(self):
                self.image(LOGO_PATH, 10, 8, 33)
                self.set_font('DejaVuLGCSans', 'B', 12)
                self.cell(0, 10, "Resumen de Pedido SUOLMEX", 0, 1, 'C')
                self.ln(10)

        pdf = PDF()
        pdf.add_page()
        pdf.set_font('DejaVuLGCSans', '', 10)

        pdf.cell(0, 10, f"Usuario: {usuario}", ln=True)
        pdf.cell(0, 10, f"Fecha y hora: {fecha_hora.replace('_', ' ')}", ln=True)
        pdf.ln(5)

        merma_frac = 0.03

        pdf.set_font('DejaVuLGCSans', 'B', 11)
        pdf.cell(0, 8, "Suelas", ln=True)
        pdf.ln(2)

        df_suelas = resumen_df[resumen_df["Hoja"].isin(["6001","2066","2060","4098"])]
        df_suelas_color = (
            df_suelas
            .groupby("Color")
            .agg(
                pares_total=("Cantidad pares", "sum"),
                poliol_kg  =("Poliol (g)", lambda s: s.sum()/1000),
                iso_kg     =("ISO (g)",    lambda s: s.sum()/1000),
            )
            .reset_index()
                    st.success(f"Agregado: {modelo_manual} - Talla {corrida} - {cantidad} pares.")
                    st.session_state["corrida_seleccionada"] = None

    # -------- Editor / Resumen --------
    if st.session_state["pedido_total"]:
        df_intermedio = pd.DataFrame(st.session_state["pedido_total"]).copy()
        if "Calcular Poliol" not in df_intermedio:
            df_intermedio["Calcular Poliol"] = True
        if "Calcular ISO" not in df_intermedio:
            df_intermedio["Calcular ISO"] = True

        st.markdown("---")
        st.subheader("Revisa y ajusta componentes")

        editor = getattr(st, "data_editor", None) or getattr(st, "experimental_data_editor", None)
        if editor is None:
            st.error("Tu versión de Streamlit no soporta data_editor. Actualiza Streamlit.")
            df_edit = df_intermedio.copy()
        else:
            df_edit = editor(df_intermedio, num_rows="dynamic", use_container_width=True)

        # Respeta flags del editor
        df_edit.loc[~df_edit["Calcular Poliol"], "Poliol (g)"] = 0
        df_edit.loc[~df_edit["Calcular ISO"],    "ISO (g)"]    = 0

        # Persistir al estado para el resto del flujo
        st.session_state["pedido_total"] = (
            df_edit.drop(columns=["Calcular Poliol", "Calcular ISO"], errors="ignore").to_dict("records")
)

        for _, row in df_suelas_color.iterrows():
            color    = row["Color"]
            pares    = int(row["pares_total"])
            poliol   = row["poliol_kg"]
            iso      = row["iso_kg"]
            poliol_m = poliol / (1 - merma_frac)
            iso_m    = iso   / (1 - merma_frac)
            mezcla   = poliol + iso
            mezcla_m = poliol_m + iso_m

            pdf.set_font('DejaVuLGCSans', 'B', 9)
            pdf.cell(0, 6, f"{color} — {pares} pares", ln=True)
            pdf.set_font('DejaVuLGCSans', '', 9)
            pdf.cell(0, 5, f"  Poliol: {poliol:.2f} kg (c/ merma: {poliol_m:.2f} kg)", ln=True)
            pdf.cell(0, 5, f"  ISO:    {iso:.2f} kg (c/ merma: {iso_m:.2f} kg)", ln=True)
            pdf.cell(0, 5, f"  Mezcla sin merma: {mezcla:.2f} kg", ln=True)
            pdf.cell(0, 5, f"  Mezcla total c/ merma: {mezcla_m:.2f} kg", ln=True)
            pdf.ln(2)

        pdf.ln(4) 

        pdf.set_font('DejaVuLGCSans', 'B', 11)
        pdf.cell(0, 8, "Plantillas", ln=True)
        pdf.ln(2)

        df_plant = resumen_df[resumen_df["Hoja"] == "PLANTILLAS"]
        if not df_plant.empty:
            df_plant_color = (
                df_plant
                .groupby("Color")
                .agg(
                    pares_total=("Cantidad pares", "sum"),
                    poliol_kg  =("Poliol (g)", lambda s: s.sum()/1000),
                    iso_kg     =("ISO (g)",    lambda s: s.sum()/1000),
        st.markdown("---")
        st.subheader("Resumen del Pedido (valores con merma)")

        resumen_df = pd.DataFrame(st.session_state["pedido_total"]).copy()

        # Mostrar cada línea formateada con merma aplicada
        for row in resumen_df.to_dict("records"):
            # Pasar a kg con merma
            pol_kg_m = con_merma(row.get("Poliol (g)", 0) / 1000.0)
            iso_kg_m = con_merma(row.get("ISO (g)",    0) / 1000.0)
            mez_kg_m = pol_kg_m + iso_kg_m
            cols = st.columns([5, 1])
            with cols[0]:
                st.markdown(
                    f"**Código:** {row['Código']} | **Modelo:** {row['Modelo']} | "
                    f"**Talla:** {row['Talla']} | **Cantidad:** {row['Cantidad pares']} pares | "
                    f"**Poliol (c/merma):** {fmt_num(pol_kg_m)} {UNIDADES} | "
                    f"**ISO (c/merma):** {fmt_num(iso_kg_m)} {UNIDADES} | "
                    f"**Mezcla:** {fmt_num(mez_kg_m)} {UNIDADES} | "
                    f"**Banda:** {row.get('Hoja','')}"
)
                .reset_index()
            )
            with cols[1]:
                uid = row.get("uid")
                key = f"eliminar_{uid}" if uid else f"eliminar_{row['Código']}_{row['Modelo']}_{row['Talla']}_{row['Cantidad pares']}"
                if st.button("Eliminar", key=key):
                    if uid:
                        st.session_state["pedido_total"] = [it for it in st.session_state["pedido_total"] if it.get("uid") != uid]
                    else:
                        st.session_state["pedido_total"] = [it for it in st.session_state["pedido_total"]
                            if not (it.get("Código")==row["Código"] and it.get("Modelo")==row["Modelo"] and
                                    it.get("Talla")==row["Talla"] and it.get("Cantidad pares")==row["Cantidad pares"]) ]
                    st.success("Elemento eliminado.")
                    st.experimental_rerun()

        # Resumen por código (usando función con merma)
        df_bandas, totales = calcular_resumen_bandas(resumen_df)

        st.markdown("---")
        st.subheader("Resumen por código (con merma)")
        for _, r in df_bandas.iterrows():
            st.markdown(f"**Código {r['codigo']}**")
            st.markdown(f"- Pares totales: {int(r['pares_total'])} pares")
            st.markdown(f"- Poliol (c/merma): {fmt_num(r['poliol_con_merma_kg'])} {UNIDADES}")
            st.markdown(f"- ISO (c/merma): {fmt_num(r['iso_con_merma_kg'])} {UNIDADES}")
            st.markdown(f"- Mezcla (c/merma): {fmt_num(r['mezcla_total_con_merma_kg'])} {UNIDADES}")
            st.caption(f"(Sin merma) Poliol: {fmt_num(r['poliol_necesario_kg'])} {UNIDADES} • ISO: {fmt_num(r['iso_necesario_kg'])} {UNIDADES}")
            st.markdown("")

        st.subheader("Totales generales")
        st.markdown(f"- **Pares total:** {fmt_ent(totales['pares_total'])} pares")
        st.markdown(f"- **Poliol (c/merma):** {fmt_num(totales['poliol_con_merma_kg'])} {UNIDADES}")
        st.markdown(f"- **ISO (c/merma):** {fmt_num(totales['iso_con_merma_kg'])} {UNIDADES}")
        st.markdown(f"- **Mezcla (c/merma):** {fmt_num(totales['mezcla_total_con_merma_kg'])} {UNIDADES}")
        st.caption(f"(Sin merma) Poliol: {fmt_num(totales['poliol_necesario_kg'])} {UNIDADES} • ISO: {fmt_num(totales['iso_necesario_kg'])} {UNIDADES}")

        # Excel para compras
        buffer = io.BytesIO()
        resumen_df.to_excel(buffer, index=False, sheet_name="Pedido para Compras")
        buffer.seek(0)
        st.download_button(
            "Descargar Excel para Compras",
            data=buffer,
            file_name="pedido_compras.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

            for _, row in df_plant_color.iterrows():
                color    = row["Color"]
                pares    = int(row["pares_total"])
                poliol   = row["poliol_kg"]
                iso      = row["iso_kg"]
                poliol_m = poliol / (1 - merma_frac)
                iso_m    = iso   / (1 - merma_frac)
                mezcla   = poliol + iso
                mezcla_m = poliol_m + iso_m

                pdf.set_font('DejaVuLGCSans', 'B', 9)
                pdf.cell(0, 6, f"{color} — {pares} pares", ln=True)
                pdf.set_font('DejaVuLGCSans', '', 9)
                pdf.cell(0, 5, f"  Poliol: {poliol:.2f} kg (c/ merma: {poliol_m:.2f} kg)", ln=True)
                pdf.cell(0, 5, f"  ISO:    {iso:.2f} kg (c/ merma: {iso_m:.2f} kg)", ln=True)
                pdf.cell(0, 5, f"  Mezcla sin merma: {mezcla:.2f} kg", ln=True)
                pdf.cell(0, 5, f"  Mezcla total c/ merma: {mezcla_m:.2f} kg", ln=True)
                pdf.ln(2)


        pdf.output(nombre_archivo)
        st.success(f"PDF generado: {Path(nombre_archivo).name}")
        with open(nombre_archivo, "rb") as f:
            st.download_button(
                "Descargar PDF",
                data=f,
                file_name=Path(nombre_archivo).name,
                key="pdf_descargar"
            )
        # Acciones
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Generar PDF", key="pdf_generar"):
                Path("historial_pedidos").mkdir(exist_ok=True)
                fecha_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                usuario = st.session_state.usuario
                nombre_archivo = f"historial_pedidos/pedido_{usuario}_{fecha_hora}.pdf"
                try:
                    generar_pdf(resumen_df, usuario, fecha_hora, nombre_archivo)
                    st.success(f"PDF generado: {Path(nombre_archivo).name}")
                    with open(nombre_archivo, "rb") as f:
                        st.download_button(
                            "Descargar PDF", data=f, file_name=Path(nombre_archivo).name,
                            key="pdf_descargar"
                        )
                except Exception as e:
                    st.error(f"No se pudo generar el PDF: {e}")

        with col2:
            if st.button("Reiniciar Pedido", key="pedido_reiniciar"):
                st.session_state["pedido_total"] = []
                st.success("Pedido reiniciado.")

    else:
        st.info("Agrega líneas de pedido desde Excel o manualmente.")


# ===================== Entrypoint =====================
if __name__ == "__main__":
    main()

with col2:
    if st.button("Reiniciar Pedido", key="pedido_reiniciar_unico"):
        st.session_state["pedido_total"] = []
        st.success("Pedido reiniciado.")
        
