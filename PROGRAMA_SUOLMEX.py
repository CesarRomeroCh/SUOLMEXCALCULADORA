import os
import streamlit as st
import sqlite3
import pandas as pd
from datetime import datetime
import hashlib
import json
import uuid
from pathlib import Path
from fpdf import FPDF
import time
import re
import unicodedata

def clean_key(s):
    if pd.isna(s):
        return ""
    s = str(s).strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return " ".join(s.split())

def es_numero_valido(s: str) -> bool:
    if s is None:
        return False
    s = str(s).strip()
    if s.upper() in ["", "0"]:
        return False
    return bool(re.fullmatch(r"[-+]?\d*\.?\d+", s))

def calcular_resumen_bandas(resumen_df, merma_frac=0.03):
    """
    Agrupa por código (la columna 'Hoja' que en tu caso es el código de ficha)
    y devuelve pares, poliol/ISO y mezclas con merma.
    """
    resumen_df = resumen_df.copy()
    resumen_df["Poliol_kg"] = resumen_df["Poliol (g)"] / 1000
    resumen_df["ISO_kg"] = resumen_df["ISO (g)"] / 1000

    agg = resumen_df.groupby("Hoja").agg(
        pares_total=("Cantidad pares", "sum"),
        poliol_necesario_kg=("Poliol_kg", "sum"),
        iso_necesario_kg=("ISO_kg", "sum"),
    ).reset_index().rename(columns={"Hoja": "codigo"})

    agg["poliol_con_merma_kg"] = agg["poliol_necesario_kg"] / (1 - merma_frac)
    agg["iso_con_merma_kg"] = agg["iso_necesario_kg"] / (1 - merma_frac)
    agg["mezcla_sin_merma_kg"] = agg["poliol_necesario_kg"] + agg["iso_necesario_kg"]
    agg["mezcla_total_con_merma_kg"] = agg["poliol_con_merma_kg"] + agg["iso_con_merma_kg"]

    totales = {
        "pares_total": agg["pares_total"].sum(),
        "poliol_necesario_kg": agg["poliol_necesario_kg"].sum(),
        "iso_necesario_kg": agg["iso_necesario_kg"].sum(),
        "poliol_con_merma_kg": agg["poliol_con_merma_kg"].sum(),
        "iso_con_merma_kg": agg["iso_con_merma_kg"].sum(),
        "mezcla_sin_merma_kg": agg["mezcla_sin_merma_kg"].sum(),
        "mezcla_total_con_merma_kg": agg["mezcla_total_con_merma_kg"].sum(),
    }

    return agg, totales

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
                if c in excl: continue
                talla = str(df_raw.iat[fila_code, c]).strip()
                cnt   = str(df_raw.iat[fila_cnt,  c]).strip()
                if talla and es_numero_valido(cnt):
                    resultados.append({
                        "Código del Producto": codigo,
                        "Color": color, 
                        "Modelo": modelo,
                        "Talla": talla,
                        "Cantidad pares": float(cnt),
                        "Hoja": hoja
                    })

            if hoja == "B1":
                for inicio in bloques_idx:

                    fila_c2     = inicio      
                    fila_m2     = inicio + 1   
                    fila_color2 = inicio + 3   
                    fila_t2     = inicio       
                    fila_p2     = inicio + 3   

                    cod2 = str(df_raw.iat[fila_c2, 28]).strip()
                    mod2 = str(df_raw.iat[fila_m2, 28]).strip()
                    if not cod2 or not mod2:
                        continue

                    color2 = str(df_raw.iat[fila_color2, 28]).strip()
                    if not color2:
                        color2 = mod2

                    for col in range(30, 38):
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
                            })


    if not resultados:
        return pd.DataFrame(columns=["Código del Producto","Modelo","Talla","Cantidad pares","Hoja"])

    df_res = pd.DataFrame(resultados).drop_duplicates(
        subset=["Código del Producto","Modelo","Talla","Cantidad pares","Hoja"]
    )
    df_res["Modelo_norm"] = df_res["Modelo"].str.strip().str.upper()
    df_res["Talla_norm"] = df_res["Talla"].str.strip().str.upper()
    return df_res

DB_PATH = "usuarios.db"
FICHAS_PATH = "FICHAS2.xlsx"
LOGO_PATH = "logo_suolmex.jpg"
st.set_page_config(page_title="Generador de Pedido SUOLMEX (B1-B4)", layout="wide")

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

def obtener_session_id():
    if "session_id" not in st.session_state:
        st.session_state.session_id = uuid.uuid4().hex
    return st.session_state.session_id

def path_sesion_local():
    return f"session_{obtener_session_id()}.json"

def guardar_sesion():
    with open(path_sesion_local(), "w") as f:
        json.dump({
            "logueado": st.session_state.logueado,
            "usuario": st.session_state.usuario,
            "rol": st.session_state.rol
        }, f)

def cargar_sesion():
    try:
        with open(path_sesion_local(), "r") as f:
            data = json.load(f)
            st.session_state.logueado = data.get("logueado", False)
            st.session_state.usuario = data.get("usuario", None)
            st.session_state.rol = data.get("rol", None)
    except:
        st.session_state.logueado = False

conn = sqlite3.connect(DB_PATH, check_same_thread=False)
c = conn.cursor()
c.execute("""
    CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT UNIQUE NOT NULL,
        contrasena TEXT NOT NULL,
        rol TEXT NOT NULL
    )
""")
conn.commit()

def encriptar_contra(contra):
    return hashlib.sha256(contra.encode()).hexdigest()

if not c.execute("SELECT * FROM usuarios WHERE codigo = 'admin'").fetchone():
    c.execute("INSERT INTO usuarios (codigo, contrasena, rol) VALUES (?, ?, ?)",
              ('admin', encriptar_contra('admin123'), 'admin'))
    conn.commit()

if "logueado" not in st.session_state:
    cargar_sesion()

st.session_state.setdefault("logueado", False)
st.session_state.setdefault("usuario", "")
st.session_state.setdefault("rol", "")

if not st.session_state.get("logueado", False):
    st.subheader("Iniciar sesión")
    with st.form("login_form"):
        codigo = st.text_input("Código de usuario")
        contrasena = st.text_input("Contraseña", type="password")
        if st.form_submit_button("Entrar"):
            user = c.execute("SELECT contrasena, rol FROM usuarios WHERE codigo = ?", (codigo,)).fetchone()
            if user and encriptar_contra(contrasena) == user[0]:
                st.session_state.logueado = True
                st.session_state.usuario = codigo
                st.session_state.rol = user[1]
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
        )
        st.write(f"Tiempo de parseo: {time.time() - t0:.3f}s")

        st.subheader("Pedido extraído (debug)")
        st.dataframe(pedido_programado)

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

                modelo_norm = clean_key(modelo)
                corrida_norm = clean_key(corrida)

                ficha = fichas[
                    (fichas["Codigo del Producto"].astype(str).str.strip() == codigo) &
                    (fichas["__Linea_clean"] == modelo_norm) &
                    (fichas["__Corrida_clean"] == corrida_norm)
                ]
                if not ficha.empty:
                    ficha = ficha.iloc[0]
                    peso_total = ficha['Peso/Pie'] * cantidad * 2
                    try:
                        poliol_str, iso_str = ficha['Relacion Poliol:ISO'].split(":")
                        poliol = float(poliol_str)
                        iso = float(iso_str)
                    except:
                        poliol, iso = 0.0, 0.0
                    total_partes = (poliol + iso) if (poliol + iso) != 0 else 1
                    cantidad_poliol = peso_total * (poliol / total_partes)
                    cantidad_iso = peso_total * (iso / total_partes)

                    nuevo = {
                        "uid": uuid.uuid4().hex,
                        "Código": codigo,
                        "Modelo": modelo,
                        "Color": color, 
                        "Talla": corrida,
                        "Cantidad pares": cantidad,
                        "Peso Total (g)": peso_total,
                        "Poliol (g)": cantidad_poliol,
                        "ISO (g)": cantidad_iso,
                        "Hoja": ficha['Hoja']
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
                )
                .reset_index()
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

with col2:
    if st.button("Reiniciar Pedido", key="pedido_reiniciar_unico"):
        st.session_state["pedido_total"] = []
        st.success("Pedido reiniciado.")
        
