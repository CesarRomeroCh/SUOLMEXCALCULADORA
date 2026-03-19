# ===================== Imports =====================
import os
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

# ===================== CONSTANTES =====================

APP_TITLE = "Generador de Pedido SUOLMEX (B1–B4)"

DB_PATH = "usuarios.db"

FICHAS_PATH = "FICHAS2.xlsx"

LOGO_PATH = "logo_suolmex.jpg"

# Configuración general
DEFAULT_MERMA = 0.03
UNIDADES = "kg"
DEC = 2
DETALLADO = True

# Hojas válidas en fichas
FICHAS_HOJAS = ['6001', '2066', '2060', '4098', 'PLANTILLAS']

# ===================== UTILIDADES GENERALES =====================

def guardar_historial_produccion(df_resumen):

    historial_path = "historial_produccion.csv"

    if df_resumen.empty:
        return

    df = df_resumen.copy()

    df["Poliol_kg"] = df["Poliol (g)"] / 1000
    df["ISO_kg"] = df["ISO (g)"] / 1000
    df["Mezcla_kg"] = df["Poliol_kg"] + df["ISO_kg"]

    pares_total = int(df["Cantidad pares"].sum())
    poliol_total = df["Poliol_kg"].sum()
    iso_total = df["ISO_kg"].sum()
    mezcla_total = df["Mezcla_kg"].sum()

    merma_frac = st.session_state.get("merma_frac", DEFAULT_MERMA)
    mezcla_con_merma = mezcla_total * (1 + merma_frac)

    registro = pd.DataFrame([{
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "usuario": st.session_state.get("usuario", ""),
        "pares": pares_total,
        "poliol_kg": round(poliol_total, 2),
        "iso_kg": round(iso_total, 2),
        "mezcla_kg": round(mezcla_total, 2),
        "mezcla_con_merma_kg": round(mezcla_con_merma, 2),
        "merma_frac": merma_frac
    }])

    if os.path.exists(historial_path):

        historial = pd.read_csv(historial_path)
        historial = pd.concat([historial, registro], ignore_index=True)

    else:

        historial = registro

    historial.to_csv(historial_path, index=False)
    
def recalcular_explosion(item, fichas):

    codigo = item["Código"]
    modelo = item["Modelo"]
    talla  = item["Talla"]
    pares  = item["Cantidad pares"]

    ficha = fichas[
        (fichas["Codigo del Producto"].astype(str).str.strip() == codigo) &
        (fichas["Linea"].astype(str).str.strip().str.upper() == modelo) &
        (fichas["Corrida"].astype(str).str.strip() == talla)
    ]

    if ficha.empty:
        return

    f = ficha.iloc[0]

    peso_total = f["Peso/Pie"] * pares * 2

    try:
        relacion = str(f["Relacion Poliol:ISO"]).replace(" ", "")

        if ":" in relacion:
            pol_str, iso_str = relacion.split(":")
            pol = float(pol_str)
            iso = float(iso_str)
        else:
            pol, iso = 0, 0

    except:
        pol, iso = 0, 0

    total = pol + iso if (pol + iso) != 0 else 1

    item["Peso Total (g)"] = peso_total
    item["Poliol (g)"] = peso_total * (pol / total)
    item["ISO (g)"] = peso_total * (iso / total)

def clean_key(s):

    if pd.isna(s):
        return ""

    s = str(s).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))

    return " ".join(s.split())


_num_pat = re.compile(r"[-+]?\d*\.?\d+")

def detectar_duplicados_fichas(df):

    df_check = df.copy()

    # Normalizar datos
    df_check["Codigo_norm"] = df_check["Codigo del Producto"].astype(str).str.strip()
    df_check["Modelo_norm"] = df_check["Linea"].apply(clean_key)
    df_check["Talla_norm"]  = df_check["Corrida"].apply(clean_key)

    # Clave única
    df_check["key"] = (
        df_check["Codigo_norm"] + "|" +
        df_check["Modelo_norm"] + "|" +
        df_check["Talla_norm"]
    )

    duplicados = df_check[df_check.duplicated("key", keep=False)]

    return duplicados

def es_numero_valido(s: str) -> bool:

    if s is None:
        return False

    s = str(s).strip()

    if s.upper() in ["", "0"]:
        return False

    return bool(_num_pat.fullmatch(s))

# ===================== VALIDACIÓN PROGRAMACIÓN =====================

def detectar_modelos_sin_ficha(programacion_df, fichas_df):

    errores = []

    if programacion_df.empty:
        return ["No se detectó programación válida en el Excel."]

    modelos_programados = set(programacion_df["Modelo_norm"].unique())

    modelos_fichas = set(
        fichas_df["Linea"]
        .astype(str)
        .str.strip()
        .str.upper()
        .unique()
    )

    for modelo in modelos_programados:

        if modelo not in modelos_fichas:
            errores.append(f"Modelo sin ficha técnica: {modelo}")

    return errores

# ===================== FORMATEADORES =====================

fmt_num = lambda x, dec=DEC: (
    f"{float(x):,.{dec}f}" if pd.notna(x) else "-"
)

fmt_ent = lambda x: (
    f"{int(x):,}" if pd.notna(x) else "-"
)

# ===================== MERMA =====================

con_merma = lambda x, frac=None: (
    x / (
        1.0
        - (
            frac
            if frac is not None
            else st.session_state.get("merma_frac", DEFAULT_MERMA)
        )
    )
)

# ===================== BASE DE DATOS =====================

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


hash_password = lambda s: hashlib.sha256(s.encode()).hexdigest()


def ensure_admin(c, conn):

    if not c.execute(
        "SELECT 1 FROM usuarios WHERE codigo='admin'"
    ).fetchone():

        c.execute(
            "INSERT INTO usuarios (codigo, contrasena, rol) VALUES (?, ?, ?)",
            ('admin', hash_password('admin123'), 'admin')
        )

        conn.commit()
        
# ===================== SESIÓN =====================

def obtener_session_id():

    if "session_id" not in st.session_state:
        st.session_state.session_id = uuid.uuid4().hex

    return st.session_state.session_id


def _session_path():

    return f"session_{obtener_session_id()}.json"


def guardar_sesion():

    with open(_session_path(), "w", encoding="utf-8") as f:

        json.dump(
            {
                "logueado": st.session_state.get("logueado", False),
                "usuario":  st.session_state.get("usuario", None),
                "rol":      st.session_state.get("rol", None),
            },
            f,
        )


def cargar_sesion():

    try:

        with open(_session_path(), "r", encoding="utf-8") as f:

            data = json.load(f)

            st.session_state.logueado = data.get("logueado", False)
            st.session_state.usuario  = data.get("usuario")
            st.session_state.rol      = data.get("rol")

    except Exception:

        st.session_state.logueado = False
        
# ===================== CARGA DE FICHAS =====================

@st.cache_data(show_spinner=False)
def cargar_fichas(mtime: float):

    if not os.path.exists(FICHAS_PATH):
        return pd.DataFrame()

    xl = pd.ExcelFile(FICHAS_PATH)

    dfs = []

    for hoja in FICHAS_HOJAS:

        if hoja in xl.sheet_names:

            df = xl.parse(hoja)

            df["Hoja"] = hoja

            dfs.append(df)

    if not dfs:
        return pd.DataFrame()

    df = pd.concat(dfs, ignore_index=True)

    # Normalizaciones

    df["Codigo del Producto"] = df["Codigo del Producto"].astype(str).str.strip()

    df["Linea"] = df["Linea"].astype(str).str.strip().str.upper()

    df["Corrida"] = df["Corrida"].astype(str).str.strip()

    df["Peso/Pie"] = pd.to_numeric(df["Peso/Pie"], errors="coerce")

    # Claves limpias para matching

    df["__Linea_clean"] = df["Linea"].apply(clean_key)

    df["__Corrida_clean"] = df["Corrida"].apply(clean_key)

    # Filtro de filas válidas

    return df.dropna(subset=["Peso/Pie", "Relacion Poliol:ISO"]).copy()

def guardar_fichas_excel(df):

    hojas = {}

    for hoja in df["Hoja"].unique():

        hojas[hoja] = df[df["Hoja"] == hoja].drop(
            columns=["Hoja","__Linea_clean","__Corrida_clean"],
            errors="ignore"
        )

    with pd.ExcelWriter(FICHAS_PATH, engine="openpyxl") as writer:

        for hoja, data in hojas.items():

            data.to_excel(writer, sheet_name=hoja, index=False)
            
# ===================== Parseo programación B1–B4 =====================

def extract_programacion_estatica_B1_B4(file_like, codigo_producto_override=None, hojas_objetivo=None, debug=False):
    if hojas_objetivo is None:
        hojas_objetivo = ["B1", "B2", "B3", "B4"]
    try:
        xl = pd.ExcelFile(file_like)
    except Exception as e:
        st.error(f"No se pudo abrir el Excel: {e}")
        return pd.DataFrame(columns=["Código del Producto","Modelo","Talla","Cantidad pares","Hoja"]) 

    resultados = []
    bloques_inicio = [3, 9, 15, 21, 27, 33, 39, 45]
    bloques_idx    = [b-1 for b in bloques_inicio]
    total_col_map  = {"B1":12, "B2":10, "B3":14, "B4":13}

    def es_columna_total_dinamica(df, col, fila_tallas):
        for f in (fila_tallas, fila_tallas-1):
            if 0 <= f < len(df):
                txt = str(df.iat[f, col]).strip().upper()
                if "TOTAL" in txt or ("PARES" in txt and "PROG" in txt):
                    return True
        return False

    for hoja in hojas_objetivo:
        if hoja not in xl.sheet_names:
            if debug: st.warning(f"No existe hoja {hoja}.")
            continue
        df_raw = xl.parse(hoja, header=None, dtype=str).fillna("")

        if debug:
            st.subheader(f"DEBUG: hoja {hoja} cruda (columnas 0–29)")
            st.dataframe(df_raw.iloc[:60, :30])

        for inicio in bloques_idx:
            fila_code = inicio
            fila_mod  = inicio + 1
            fila_cnt  = inicio + 3

            raw_code = str(df_raw.iat[fila_code, 0]).strip()
            if codigo_producto_override and codigo_producto_override.strip():
                codigo = codigo_producto_override.strip()
            else:
                codigo = raw_code if raw_code not in ("","nan","None") else None

            modelo = str(df_raw.iat[fila_mod, 0]).strip()
            if not (codigo and modelo):
                continue

            fila_color = inicio + 3
            color = str(df_raw.iat[fila_color, 0]).strip()

            if hoja == "B1":
                extra = str(df_raw.iat[fila_color, 28]).strip()
                if extra and any(ch.isalpha() for ch in extra):
                    color = extra
            if not color:
                color = modelo

            excl = { total_col_map.get(hoja) }
            for c in range(2,12):
                if es_columna_total_dinamica(df_raw, c, fila_code):
                    excl.add(c)

            for c in range(2,12):
                if c in excl: 
                    continue
                talla = str(df_raw.iat[fila_code, c]).strip()
                cnt   = str(df_raw.iat[fila_cnt,  c]).strip()
                if talla and es_numero_valido(cnt):
                    resultados.append({
                        "Código del Producto": codigo,
                        "Color": color,
                        "Modelo": modelo,
                        "Talla": talla,
                        "Cantidad pares": float(cnt),
                        "Hoja": hoja,
                    })

            # Bloque lateral B1 (cols 28–39)
            if hoja == "B1":
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

                    for col in range(30, 40):
                        talla2 = str(df_raw.iat[fila_t2, col]).strip()
                        cr2    = str(df_raw.iat[fila_p2, col]).strip()
                        if talla2.startswith("#") and es_numero_valido(cr2):
                            resultados.append({
                                "Código del Producto": cod2,
                                "Color": color2,
                                "Modelo": mod2,
                                "Talla": talla2,
                                "Cantidad pares": float(cr2),
                                "Hoja": hoja,
                            })

    if not resultados:
        return pd.DataFrame(columns=["Código del Producto","Modelo","Talla","Cantidad pares","Hoja"]) 

    df_res = pd.DataFrame(resultados).drop_duplicates(
        subset=["Código del Producto","Modelo","Talla","Cantidad pares","Hoja"]
    ).reset_index(drop=True)
    
    df_res["Modelo_norm"] = df_res["Modelo"].str.strip().str.upper()
    df_res["Talla_norm"]  = df_res["Talla"].str.strip().str.upper()
    return df_res

# ===================== Cálculo de totales por código (con merma) =====================

def calcular_resumen_bandas(resumen_df, merma_frac):
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

    agg["poliol_con_merma_kg"] = agg["poliol_necesario_kg"] * (1 + merma_frac)
    agg["iso_con_merma_kg"]    = agg["iso_necesario_kg"] * (1 + merma_frac)
    agg["mezcla_sin_merma_kg"] = agg["poliol_necesario_kg"] + agg["iso_necesario_kg"]
    agg["mezcla_total_con_merma_kg"] = agg["poliol_con_merma_kg"] + agg["iso_con_merma_kg"]

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

def set_font(pdf, size=10, bold=False):
    fam = "DejaVuLGCSans" if getattr(pdf, "_dejavu", False) else "Helvetica"
    pdf.set_font(fam, "B" if bold else "", size)

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

# ---------- Pedido de químicos (agrupado por código y color) ----------

    pedido_quimicos = (
        resumen_df
        .groupby(["Código","Color"], as_index=False)
        .agg({
            "Cantidad pares":"sum",
            "Poliol (g)":"sum",
            "ISO (g)":"sum"
        })
    )

    pedido_quimicos["Poliol (kg)"] = pedido_quimicos["Poliol (g)"] / 1000
    pedido_quimicos["ISO (kg)"] = pedido_quimicos["ISO (g)"] / 1000

    pedido_quimicos = pedido_quimicos.sort_values(["Código","Color"])

    pdf = PDF(orientation='P', unit='mm', format='A4')
    pdf._dejavu = (
        try_add_font(pdf, "DejaVuLGCSans", "",  "DejaVuLGCSans.ttf") and
        try_add_font(pdf, "DejaVuLGCSans", "B", "DejaVuLGCSans-Bold.ttf")
    )
    pdf.alias_nb_pages()

    #Portada
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
        ("Merma", f"{st.session_state.get('merma_frac', DEFAULT_MERMA)*100:.0f}%"),
        ("Unidades", UNIDADES),
    ])

    set_font(pdf, size=9)
    pdf.multi_cell(0, 6, "Desglose de pares y requerimientos de Poliol/ISO por sección y color. "
                      "Incluye totales y detalle por modelo/talla (si aplica).")
    hr(pdf, ypad=3)
    pdf._is_portada = False

    # ---------- Tabla pedido químicos ----------

    titulo_seccion(pdf, "Pedido de Químicos")

    rows = []

    for _, r in pedido_quimicos.iterrows():

        rows.append([
            r["Código"],
            r["Color"],
            fmt_ent(r["Cantidad pares"]),
            fmt_num(r["Poliol (kg)"]),
            fmt_num(r["ISO (kg)"])
        ])

    tabla(
        pdf,
        headers=["Código","Color","Pares","Poliol (kg)","ISO (kg)"],
        rows=rows,
        widths=[35,60,25,35,35],
        aligns=["C","L","R","R","R"]
    )
    
    total_poliol = pedido_quimicos["Poliol (kg)"].sum()
    total_iso = pedido_quimicos["ISO (kg)"].sum()

    pdf.ln(5)

    cuadro_info(pdf,[
        ("Total Poliol", f"{fmt_num(total_poliol)} kg"),
        ("Total ISO", f"{fmt_num(total_iso)} kg")
    ])

    set_font(pdf, size=9)
    pdf.set_text_color(90,90,90)
    pdf.cell(0, 6, f"Merma del {st.session_state.get('merma_frac', DEFAULT_MERMA)*100:.0f}% aplicada a Poliol e ISO.", ln=1)
    pdf.set_text_color(0,0,0)

    pdf._footer_info = f"Usuario: {usuario}   •   Fecha: {fecha_hora.replace('_',' ')}"

    pdf.output(nombre_archivo, "F")

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

        # LOGO
        if os.path.exists(LOGO_PATH):
            st.image(LOGO_PATH, width=200)

        st.markdown("---")

        # NAVEGACIÓN PRINCIPAL
        st.markdown("### Navegación")

        pagina = st.radio(
            "Ir a:",
            [
                "Explosión",
                "Fichas / Modelos",
                "KPIs",
                "Usuarios"
            ]
        )

        st.markdown("---")
        st.caption(f"👤 {st.session_state.get('usuario','')}")
        st.markdown("---")

        # BOTÓN CERRAR SESIÓN (AHORA EN SIDEBAR)
        if st.button("Cerrar sesión", use_container_width=True):
            p = Path(_session_path())
            if p.exists():
                p.unlink(missing_ok=True)

            st.session_state["logueado"] = False
            st.session_state["usuario"]  = ""
            st.session_state["rol"]      = ""

            st.success("Sesión cerrada.")

            try:
                params = st.query_params.to_dict() if hasattr(st, "query_params") else st.experimental_get_query_params()
                params["_r"] = uuid.uuid4().hex
                if hasattr(st, "query_params"):
                    st.query_params.clear()
                    st.query_params.update(params)
                else:
                    st.experimental_set_query_params(**params)
            except Exception:
                pass

            st.rerun()

    return pagina

# ===================== Vistas =====================

def login_view(c, conn):
    st.subheader("Iniciar sesión")
    with st.form("login_form"):
        codigo = st.text_input("Código de usuario")
        contrasena = st.text_input("Contraseña", type="password")
        if st.form_submit_button("Entrar"):
            row = c.execute("SELECT contrasena, rol FROM usuarios WHERE codigo=?", (codigo,)).fetchone()
            if row and hash_password(contrasena) == row[0]:
                st.session_state.logueado = True
                st.session_state.usuario  = codigo
                st.session_state.rol      = row[1]
                guardar_sesion()
                st.success("Sesión iniciada correctamente.")
                st.rerun()
            else:
                st.error("Credenciales incorrectas.")
    st.stop()

# ===================== PANTALLAS =====================

def pantalla_explosion(fichas):

    st.header("Explosión de producción")
    
# ---------- Inicializar controles ----------

    if "calc_poliol_global" not in st.session_state:
        st.session_state["calc_poliol_global"] = True

    if "calc_iso_global" not in st.session_state:
        st.session_state["calc_iso_global"] = True

    if "control_modelos" not in st.session_state:
        st.session_state["control_modelos"] = {}
    
    st.info("Edita el porcentaje de merma.")

    merma_input = st.number_input(
        "Merma %",
        min_value=0.0,
        max_value=10.0,
        value=st.session_state.get("merma_frac", 0.03) * 100,
        step=0.1
    )

    st.session_state.merma_frac = merma_input / 100
    
        # Estado de pedido
    st.session_state.setdefault("pedido_total", [])
    st.session_state.setdefault("corrida_seleccionada", None)
    
    # Tabs principales
    tab_prog, tab_manual, tab_control, tab_resumen = st.tabs(
        ["Programación", "Manual", "Control", "Resumen"]
    )
    
        # -------- Tab Excel --------
    with tab_prog:
        uploaded_prog = st.file_uploader(
            "Sube el Excel de programación del día",
            type=["xlsx"],
            help="Debe contener las hojas B1, B2, B3 y B4 con la programación de producción",
            key="prog_excel_tab"
        )
        
            
        with st.expander("Opciones avanzadas"):
            codigo_override = st.text_input("Código del Producto (override si no lo detecta)", value="", key="override_excel")

            st.caption("Usar solo si el Excel no detecta el código automáticamente.")
            
            st.markdown("---")

            st.session_state.setdefault("debug_prog", False)
            st.session_state.setdefault("debug_fichas", False)
            st.session_state.setdefault("force_reparse", False)

            st.checkbox("Mostrar debug de programación", key="debug_prog")
            st.checkbox("Mostrar debug de fichas", key="debug_fichas")
            st.checkbox("Forzar recarga de fichas", key="force_reparse")
            
        pedido_programado = pd.DataFrame()
        if uploaded_prog:

            #RESET CONTROLADO DEL PEDIDO
            if "pedido_total" not in st.session_state:
                st.session_state["pedido_total"] = []
            else:
                st.session_state["pedido_total"].clear()

            with st.spinner("Procesando programación..."):

                t0 = time.time()

                pedido_programado = extract_programacion_estatica_B1_B4(
                    uploaded_prog,
                    codigo_override or None,
                    hojas_objetivo=["B1","B2","B3","B4"],
                    debug=st.session_state.get("debug_prog")
                )
                if "Color" not in pedido_programado.columns:
                    pedido_programado["Color"] = pedido_programado["Modelo"]

                pedido_programado = (
                    pedido_programado
                    .groupby(
                        ["Código del Producto","Modelo","Talla","Color","Hoja"],
                        as_index=False
                    )["Cantidad pares"]
                    .sum()
                )
                pedido_programado["Modelo_norm"] = pedido_programado["Modelo"].str.strip().str.upper()
                pedido_programado["Talla_norm"]  = pedido_programado["Talla"].str.strip().str.upper()
                errores_modelos = detectar_modelos_sin_ficha(pedido_programado, fichas)

                if errores_modelos:
                    st.error("Se detectaron modelos sin ficha técnica")

                    for e in errores_modelos:
                        st.warning(e)
                        
                    st.stop()
                    
                st.caption(f"Tiempo de parseo: {time.time() - t0:.3f}s")

            if st.session_state.get("debug_prog"):
                st.subheader("Pedido extraído (debug)")
                st.dataframe(pedido_programado, use_container_width=True)

            if pedido_programado.empty:
                st.warning("No se extrajeron líneas válidas. Verifica las hojas y la nomenclatura.")
            else:
                st.session_state["pedido_total"] = []
                # ---- Preparar claves para merge ----

                pedido_programado["Modelo_clean"] = pedido_programado["Modelo"].apply(clean_key)
                pedido_programado["Talla_clean"]  = pedido_programado["Talla"].apply(clean_key)

                # ---- Merge con fichas ----
                merge = pedido_programado.merge(
                    fichas,
                    left_on=["Código del Producto", "Modelo_clean", "Talla_clean"],
                    right_on=["Codigo del Producto", "__Linea_clean", "__Corrida_clean"],
                    how="left",
                    suffixes=("_prog", "_ficha")
                )
                
                avisos_ficha = {}
                
                # ---- Recorrer resultados ----
                for _, r in merge.iterrows():

                    codigo = str(r["Código del Producto"]).strip()
                    modelo = str(r["Modelo"]).strip().upper()
                    corrida = str(r["Talla"]).strip()
                    color = str(r.get("Color", "")).strip() or modelo

                    try:
                        cantidad = int(float(r["Cantidad pares"]))
                    except:
                        continue

                    if cantidad <= 0:
                        continue

                    if not modelo or not corrida:
                        continue

                    # Validar ficha existente
                    if pd.isna(r["Peso/Pie"]):

                        key = (codigo, modelo)

                        avisos_ficha[key] = True

                        continue

                    # ---- Cálculo explosión ----

                    peso_total = r["Peso/Pie"] * cantidad * 2

                    try:
                        relacion = str(r["Relacion Poliol:ISO"]).replace(" ", "")

                        if ":" in relacion:
                            pol_str, iso_str = relacion.split(":")
                            pol = float(pol_str)
                            iso = float(iso_str)
                        else:
                            pol, iso = 0, 0

                    except:
                        pol, iso = 0, 0

                    total = (pol + iso) if (pol + iso) != 0 else 1

                    cant_pol = peso_total * (pol / total)
                    cant_iso = peso_total * (iso / total)

                    nuevo = {
                        "uid": uuid.uuid4().hex,
                        "Código": codigo,
                        "Modelo": modelo,
                        "Color": color,
                        "Talla": corrida,
                        "Cantidad pares": cantidad,
                        "Peso Total (g)": peso_total,
                        "Poliol (g)": cant_pol,
                        "ISO (g)": cant_iso,
                        "Hoja": r.get("Hoja_ficha") if pd.notna(r.get("Hoja_ficha")) else r["Hoja"],
                    }

                    existe = any(
                        it["Código"] == nuevo["Código"]
                        and it["Modelo"] == nuevo["Modelo"]
                        and it["Talla"] == nuevo["Talla"]
                        and it["Hoja"] == nuevo["Hoja"]
                        for it in st.session_state["pedido_total"]
                    )

                    if not existe:
                        st.session_state["pedido_total"].append(nuevo)

                    # ---- Mostrar avisos de explosión ----
                if avisos_ficha:

                    st.markdown("### Avisos de explosión")

                for (codigo, modelo) in avisos_ficha.keys():

                    st.warning(
                        f"{modelo} | Código {codigo}\n\n"
                        f"No existe ficha técnica para este modelo."
                    )

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
                if not ficha_manual.empty:
                    f = ficha_manual.iloc[0]
                    peso_total = f['Peso/Pie'] * cantidad * 2
                    
                    try:
                        relacion = str(f["Relacion Poliol:ISO"]).replace(" ", "")

                        if ":" in relacion:
                            pol_str, iso_str = relacion.split(":")
                            pol = float(pol_str)
                            iso = float(iso_str)
                        else:
                            pol, iso = 0.0, 0.0

                    except:
                        pol, iso = 0.0, 0.0
                        
                    total = (pol + iso) if (pol + iso) != 0 else 1
                    cant_pol = peso_total * (pol / total)
                    cant_iso = peso_total * (iso / total)

                    nuevo = {
                        "uid": uuid.uuid4().hex,
                        "Código": codigo_manual,
                        "Modelo": modelo_manual,
                        "Color":  modelo_manual,
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
                    st.success(f"Agregado: {modelo_manual} - Talla {corrida} - {cantidad} pares.")
                    st.session_state["corrida_seleccionada"] = None
                    
                    # Estado de control de explosión
                    st.session_state.setdefault("calc_poliol_global", True)
                    st.session_state.setdefault("calc_iso_global", True)
                    st.session_state.setdefault("control_modelos", {})

    # -------- Editor / Resumen --------
    if st.session_state["pedido_total"]:
        df_intermedio = pd.DataFrame(st.session_state["pedido_total"]).copy()
        if "Calcular Poliol" not in df_intermedio:
            df_intermedio["Calcular Poliol"] = True
        if "Calcular ISO" not in df_intermedio:
            df_intermedio["Calcular ISO"] = True

        st.markdown("---")
        with tab_control:
            
            st.subheader("Control de explosión")

            # -------- GLOBAL --------

            col1, col2 = st.columns(2)

            with col1:
                st.session_state["calc_poliol_global"] = st.checkbox(
                    "Calcular Poliol",
                    value=st.session_state["calc_poliol_global"]
                )

            with col2:
                st.session_state["calc_iso_global"] = st.checkbox(
                    "Calcular ISO",
                    value=st.session_state["calc_iso_global"]
                )

            if st.session_state.get("modelo_eliminar"):

                codigo_del, modelo_del = st.session_state["modelo_eliminar"]

                st.session_state["pedido_total"] = [
                    it for it in st.session_state["pedido_total"]
                    if not (it["Código"] == codigo_del and it["Modelo"] == modelo_del)
                ]

                st.session_state["modelo_eliminar"] = None
                
            df_control = pd.DataFrame(st.session_state["pedido_total"])

            if not df_control.empty:

                modelos = df_control.groupby(["Código","Modelo"])

                st.markdown("### Modelos en explosión")

                for (codigo, modelo), df_modelo in modelos:

                    pares_modelo = df_modelo["Cantidad pares"].sum()

                    key_modelo = f"{codigo}_{modelo}"

                    ctrl_modelo = st.session_state["control_modelos"].setdefault(
                        key_modelo,
                        {"poliol": True, "iso": True}
                    )

                    with st.expander(f"{modelo} | Código {codigo} | {pares_modelo} pares"):

                        col1, col2, col3 = st.columns([1,1,2])

                        with col1:
                            ctrl_modelo["poliol"] = st.checkbox(
                                "Poliol",
                                value=ctrl_modelo["poliol"],
                                key=f"pol_model_{key_modelo}"
                            )

                        with col2:
                            ctrl_modelo["iso"] = st.checkbox(
                                "ISO",
                                value=ctrl_modelo["iso"],
                                key=f"iso_model_{key_modelo}"
                            )

                        with col3:
                            if st.button("Eliminar modelo", key=f"del_model_{key_modelo}"):

                                st.session_state["modelo_eliminar"] = (codigo, modelo)

                        # Encabezados de tabla
                        h1, h2, h3 = st.columns([2,3,1])

                        with h1:
                            st.markdown("### Talla")

                        with h2:
                            st.markdown("### Pares")

                        with h3:
                            st.markdown("### Accion")

                        #st.divider()

                        for _, row in df_modelo.iterrows():

                            uid = row["uid"]

                            col1, col2, col3 = st.columns([2,3,1])

                            # ---- TALLA ----
                            with col1:
                                st.markdown(f"**{row['Talla']}**")

                            # ---- INPUT PARES ----
                            with col2:
                                new_pares = st.number_input(
                                    "pares",
                                    min_value=0,
                                    value=int(row["Cantidad pares"]),
                                    key=f"pares_{uid}",
                                    label_visibility="collapsed"
                                )

                                if new_pares != row["Cantidad pares"]:
                                    for item in st.session_state["pedido_total"]:
                                        if item["uid"] == uid:
                                            item["Cantidad pares"] = new_pares
                                            recalcular_explosion(item, fichas)

                            # ---- BOTON ELIMINAR ----
                            with col3:
                                if st.button("BORRAR", key=f"del_{uid}", help="Eliminar talla"):
                                    st.session_state["pedido_total"] = [
                                        it for it in st.session_state["pedido_total"]
                                        if it.get("uid") != uid
                                    ]

                                    st.rerun()


            resumen_df = pd.DataFrame(st.session_state["pedido_total"]).copy()

            if not st.session_state["calc_poliol_global"]:
                resumen_df["Poliol (g)"] = 0

            if not st.session_state["calc_iso_global"]:
                resumen_df["ISO (g)"] = 0
        
            # Resumen por código (usando función con merma)
            df_bandas, totales = calcular_resumen_bandas(
                resumen_df,
                st.session_state.get("merma_frac", DEFAULT_MERMA)
            )

            st.markdown("---")
        with tab_resumen:
            st.subheader("Resumen por código (con merma)")
            for _, r in df_bandas.iterrows():

                with st.expander(f"Código {r['codigo']}"):

                    col1, col2, col3, col4 = st.columns(4)

                    with col1:
                        st.metric(
                            "Pares Totales",
                            f"{fmt_ent(r['pares_total'])}"
                        )

                    with col2:
                        st.metric(
                            "Poliol Total",
                            f"{fmt_num(r['poliol_con_merma_kg'])} {UNIDADES}"
                        )

                    with col3:
                        st.metric(
                            "ISO Total",
                            f"{fmt_num(r['iso_con_merma_kg'])} {UNIDADES}"
                        )

                    with col4:
                        st.metric(
                            "Mezcla Total",
                            f"{fmt_num(r['mezcla_total_con_merma_kg'])} {UNIDADES}"
                        )

                    st.caption(
                        f"Sin merma → Poliol: {fmt_num(r['poliol_necesario_kg'])} {UNIDADES} | "
                        f"ISO: {fmt_num(r['iso_necesario_kg'])} {UNIDADES}"
                    )
            st.subheader("Totales generales")
            col1, col2, col3 = st.columns(3)

            col1.metric(
                "Pares Totales",
                f"{fmt_ent(totales['pares_total'])}"
            )

            col2.metric(
                "Poliol Total",
                f"{fmt_num(totales['poliol_con_merma_kg'])} {UNIDADES}"
            )

            col3.metric(
                "ISO Total",
                f"{fmt_num(totales['iso_con_merma_kg'])} {UNIDADES}"
            )

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

            # Acciones
            col1, col2 = st.columns(2)
            with col1:
                if st.button("Generar PDF", key="pdf_generar"):
                    guardar_historial_produccion(resumen_df)
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

def generar_corrida_modelo(codigo, modelo, talla_min, talla_max, peso_min, peso_max, relacion):

    talla_min = int(talla_min)
    talla_max = int(talla_max)

    if talla_max <= talla_min:
        return None

    incremento = (peso_max - peso_min) / (talla_max - talla_min)

    filas = []

    for i, talla in enumerate(range(talla_min, talla_max + 1)):

        peso = peso_min + incremento * i

        filas.append({
            "Codigo del Producto": codigo,
            "Linea": modelo.upper().strip(),
            "Corrida": str(talla),
            "Peso/Pie": round(peso, 2),
            "Relacion Poliol:ISO": relacion,
            "Hoja": codigo
        })

    return pd.DataFrame(filas)

def guardar_corrida_excel(df_corrida):

    xl = pd.ExcelFile(FICHAS_PATH)

    hojas = {hoja: xl.parse(hoja) for hoja in xl.sheet_names}

    codigo = df_corrida["Codigo del Producto"].iloc[0]

    if codigo not in hojas:

        hojas[codigo] = pd.DataFrame(columns=df_corrida.columns)

    hojas[codigo] = pd.concat(
        [hojas[codigo], df_corrida.drop(columns=["Hoja"], errors="ignore")],
        ignore_index=True
    )

    with pd.ExcelWriter(FICHAS_PATH, engine="openpyxl") as writer:

        for hoja, df in hojas.items():
            df.to_excel(writer, sheet_name=hoja, index=False)
            
# ===================== 

def pantalla_fichas():

    st.header("Gestión de Fichas Técnicas")

    if not os.path.exists(FICHAS_PATH):
        st.error("No se encontró el archivo de fichas.")
        return

    mtime = os.path.getmtime(FICHAS_PATH)
    fichas = cargar_fichas(mtime)

    if fichas.empty:
        st.warning("No hay fichas cargadas.")
        return
    
    duplicados_global = detectar_duplicados_fichas(fichas)

    if not duplicados_global.empty:

        with st.expander("Se detectaron duplicados en fichas existentes"):

            st.dataframe(
                duplicados_global[
                    ["Codigo del Producto", "Linea", "Corrida"]
                ].drop_duplicates(),
                use_container_width=True
            )

    # ===============================
    # SELECCIÓN DE CÓDIGO
    # ===============================

    codigos = sorted(fichas["Codigo del Producto"].unique())

    codigo_sel = st.selectbox(
        "Selecciona el código de producto",
        codigos
    )

    df_codigo = fichas[
        fichas["Codigo del Producto"] == codigo_sel
    ].copy()

    # ===============================
    # EDITOR DE TABLA
    # ===============================

    st.subheader(f"Fichas del código {codigo_sel}")

    columnas_mostrar = [
        "Codigo del Producto",
        "Linea",
        "Corrida",
        "Peso/Pie",
        "Relacion Poliol:ISO",
        "Hoja"
    ]

    df_edit = st.data_editor(
        df_codigo[columnas_mostrar],
        num_rows="dynamic",
        use_container_width=True,
        key="editor_fichas"
    )

    # ===============================

    if st.button("Guardar cambios"):

        try:

            df_restante = fichas[
                fichas["Codigo del Producto"] != codigo_sel
            ]

            df_final = pd.concat([df_restante, df_edit])
            
            duplicados = detectar_duplicados_fichas(df_final)

            if not duplicados.empty:

                st.error(" Hay fichas duplicadas. Corrige antes de guardar.")

                st.dataframe(
                    duplicados[
                        ["Codigo del Producto", "Linea", "Corrida"]
                    ].drop_duplicates(),
                    use_container_width=True
                )

                return

            guardar_fichas_excel(df_final)

            st.cache_data.clear()

            st.success("Fichas guardadas correctamente.")

            st.rerun()

        except Exception as e:

            st.error(f"No se pudo guardar: {e}")

    st.markdown("---")

    st.subheader("Crear modelo con corrida automática")

    with st.form("crear_modelo_corrida"):

        col1, col2 = st.columns(2)

        with col1:

            codigo_new = st.selectbox(
                "Código del producto",
                codigos,
                key="codigo_corrida"
            )

            modelo_new = st.text_input(
                "Nombre del modelo"
            )

            relacion_new = st.text_input(
                "Relación Poliol:ISO",
                value="100:45"
            )

        with col2:

            talla_min = st.number_input(
                "Talla mínima",
                step=1
            )

            peso_min = st.number_input(
                "Peso talla mínima (g)",
                step=0.1
            )

            talla_max = st.number_input(
                "Talla máxima",
                step=1
            )

            peso_max = st.number_input(
                "Peso talla máxima (g)",
                step=0.1
            )

        crear = st.form_submit_button("Crear modelo")

        if crear:

            if not modelo_new:

                st.warning("Ingresa el nombre del modelo.")

            else:

                df_corrida = generar_corrida_modelo(
                    codigo_new,
                    modelo_new,
                    talla_min,
                    talla_max,
                    peso_min,
                    peso_max,
                    relacion_new
                )

                if df_corrida is None:
                    st.error("La talla máxima debe ser mayor.")
                else:
                    
                    fichas_actuales = cargar_fichas(os.path.getmtime(FICHAS_PATH))

                    df_test = pd.concat([fichas_actuales, df_corrida])

                    duplicados = detectar_duplicados_fichas(df_test)

                    if not duplicados.empty:

                        st.error("El modelo genera duplicados de talla.")

                        st.dataframe(
                            duplicados[
                                ["Codigo del Producto", "Linea", "Corrida"]
                            ].drop_duplicates(),
                            use_container_width=True
                        )     
                    else:

                        guardar_corrida_excel(df_corrida)

                        st.cache_data.clear()

                        st.success("Modelo creado con toda la corrida.")

                        st.rerun()

    # ===============================
    # CREAR NUEVO CÓDIGO
    # ===============================

    st.subheader("Crear nuevo código")

    nuevo_codigo = st.text_input("Nuevo código")

    if st.button("Crear hoja"):

        if not nuevo_codigo:
            st.warning("Ingresa un código.")
            return

        if nuevo_codigo in codigos:
            st.warning("Ese código ya existe.")
            return

        try:

            with pd.ExcelWriter(
                FICHAS_PATH,
                engine="openpyxl",
                mode="a"
            ) as writer:

                df_vacio = pd.DataFrame(columns=[
                    "Codigo del Producto",
                    "Linea",
                    "Corrida",
                    "Peso/Pie",
                    "Relacion Poliol:ISO"
                ])

                df_vacio.to_excel(
                    writer,
                    sheet_name=nuevo_codigo,
                    index=False
                )

            st.success(f"Código {nuevo_codigo} creado.")

            st.rerun()

        except Exception as e:

            st.error(f"No se pudo crear: {e}")


def pantalla_kpis():

    st.header("Indicadores de Producción")

    if "pedido_total" not in st.session_state:
        st.info("No hay datos de producción cargados.")
        return

    df = pd.DataFrame(st.session_state["pedido_total"])

    if df.empty:
        st.info("No hay producción registrada.")
        return

    # =========================
    # PREPARACIÓN DE DATOS
    # =========================

    df["Poliol_kg"] = df["Poliol (g)"] / 1000
    df["ISO_kg"] = df["ISO (g)"] / 1000
    df["Mezcla_kg"] = df["Poliol_kg"] + df["ISO_kg"]

    pares_total = int(df["Cantidad pares"].sum())
    poliol_total = df["Poliol_kg"].sum()
    iso_total = df["ISO_kg"].sum()
    mezcla_total = df["Mezcla_kg"].sum()

    consumo_par = (mezcla_total * 1000 / pares_total) if pares_total else 0

    merma_frac = st.session_state.get("merma_frac", DEFAULT_MERMA)

    mezcla_con_merma = mezcla_total * (1 + merma_frac)
    merma_kg = mezcla_con_merma - mezcla_total

    # =========================
    # KPIs PRINCIPALES
    # =========================

    st.subheader("Indicadores clave")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Producción total",
        f"{pares_total:,} pares"
    )

    col2.metric(
        "Poliol total",
        f"{poliol_total:,.1f} kg"
    )

    col3.metric(
        "ISO total",
        f"{iso_total:,.1f} kg"
    )

    col4.metric(
        "Consumo por par",
        f"{consumo_par:.1f} g"
    )

    st.markdown("---")

    # =========================
    # MERMA
    # =========================

    st.subheader("Indicador de merma")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric(
        "Mezcla sin merma",
        f"{mezcla_total:,.1f} kg"
    )

    col2.metric(
        "Mezcla con merma",
        f"{mezcla_con_merma:,.1f} kg"
    )

    col3.metric(
        "Merma aplicada",
        f"{merma_frac*100:.1f} %"
    )

    col4.metric(
        "Material extra",
        f"{merma_kg:,.1f} kg"
    )

    st.progress(min(merma_frac, 1.0))

    st.caption(
        f"La merma representa {merma_kg:,.1f} kg adicionales de mezcla."
    )

    st.markdown("---")

    # =========================
    # PRODUCCIÓN POR LÍNEA
    # =========================

    st.subheader("Producción por línea")

    linea = (
        df.groupby("Hoja")["Cantidad pares"]
        .sum()
        .reset_index()
    )

    st.bar_chart(
        linea.set_index("Hoja")
    )

    st.markdown("---")

    # =========================
    # TOP MODELOS
    # =========================

    st.subheader("Modelos con mayor producción")

    top_modelos = (
        df.groupby("Modelo")["Cantidad pares"]
        .sum()
        .sort_values(ascending=False)
        .head(5)
        .reset_index()
    )

    st.bar_chart(
        top_modelos.set_index("Modelo")
    )

    st.markdown("---")

# =========================

    st.markdown("---")
    st.subheader("Merma histórica")

    historial_path = "historial_produccion.csv"

    if os.path.exists(historial_path):

        hist = pd.read_csv(historial_path)

        hist["fecha"] = pd.to_datetime(hist["fecha"])

        ultimos = hist.sort_values("fecha").tail(10)

        st.line_chart(
            ultimos.set_index("fecha")["merma_frac"]
        )

        st.caption("Evolución de la merma en el tiempo.")

    st.markdown("---")
    st.subheader("Producción reciente")

    historial_path = "historial_produccion.csv"

    if os.path.exists(historial_path):

        hist = pd.read_csv(historial_path)

        hist["fecha"] = pd.to_datetime(hist["fecha"])

        ultimos = hist.sort_values("fecha").tail(10)

        st.line_chart(
            ultimos.set_index("fecha")["pares"]
        )

        st.caption("Producción (pares) en las últimas ejecuciones.")
    
def pantalla_usuarios():
    st.header("Administración de usuarios")
    st.info("Gestión de usuarios del sistema.")
    
# ===================== APP =====================

def main():

    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown(STYLES, unsafe_allow_html=True)

    # -------- CONFIGURACIÓN --------

    if "merma_frac" not in st.session_state:
        st.session_state.merma_frac = DEFAULT_MERMA

    # -------- BASE DE DATOS --------

    conn, c = init_db()
    ensure_admin(c, conn)

    # -------- SESIÓN --------

    if "logueado" not in st.session_state:
        cargar_sesion()

    st.session_state.setdefault("logueado", False)
    st.session_state.setdefault("usuario", "")
    st.session_state.setdefault("rol", "")

    # -------- LOGIN --------

    if not st.session_state.get("logueado", False):
        login_view(c, conn)

    # -------- BARRA SUPERIOR --------

    col1, col2 = st.columns([0.7, 0.3])

    with col1:
        st.title(APP_TITLE)

    # -------- SIDEBAR --------

    pagina = header_sidebar()

    # -------- CARGAR FICHAS --------

    if os.path.exists(FICHAS_PATH):
        mtime = os.path.getmtime(FICHAS_PATH)
        fichas = cargar_fichas(mtime)
        
    else:
        fichas = pd.DataFrame()
        
    # -------- NAVEGACIÓN --------

    if pagina == "Explosión":
        pantalla_explosion(fichas)

    elif pagina == "Fichas / Modelos":
        pantalla_fichas()

    elif pagina == "KPIs":
        pantalla_kpis()

    elif pagina == "Usuarios":
        pantalla_usuarios()
        
if __name__ == "__main__":
    main()

