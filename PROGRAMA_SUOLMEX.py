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

# ---------- utilidades ----------
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

# ---------- extractor fijo para B1-B4 ----------
def extract_programacion_estatica_B1_B4(file_like, codigo_producto_override=None, hojas_objetivo=None, debug=False):
    if hojas_objetivo is None:
        hojas_objetivo = ["B1", "B2", "B3", "B4"]
    try:
        xl = pd.ExcelFile(file_like)
    except Exception as e:
        st.error(f"No se pudo abrir el Excel: {e}")
        return pd.DataFrame(columns=["Código del Producto", "Modelo", "Talla", "Cantidad pares", "Hoja"])

    resultados = []
    bloques_inicio = [3, 9, 15, 21, 27, 33, 39, 45]
    bloques_idx = [b - 1 for b in bloques_inicio]

    for hoja in hojas_objetivo:
        if hoja not in xl.sheet_names:
            if debug:
                st.warning(f"No existe hoja {hoja}.")
            continue
        df_raw = xl.parse(hoja, header=None, dtype=str).fillna("")
        if debug:
            st.subheader(f"DEBUG: hoja {hoja} cruda")
            st.dataframe(df_raw.iloc[:60, :15])

        for inicio in bloques_idx:
            fila_codigo = inicio
            fila_modelo = inicio + 1
            fila_tallas = inicio
            fila_cantidades = inicio + 3

            if fila_modelo >= len(df_raw) or fila_cantidades >= len(df_raw):
                continue

            codigo_raw = str(df_raw.iat[fila_codigo, 0]).strip()
            if codigo_producto_override and str(codigo_producto_override).strip():
                codigo = str(codigo_producto_override).strip()
            else:
                codigo = codigo_raw if codigo_raw not in ["", "nan", "None"] else None

            modelo = str(df_raw.iat[fila_modelo, 0]).strip()
            if not modelo:
                continue

            # tallas columnas C..L -> índices 2..11
            tallas = []
            for col in range(2, 12):
                raw_talla = str(df_raw.iat[fila_tallas, col]).strip()
                if raw_talla.startswith("#"):
                    tallas.append((col, raw_talla))
                else:
                    tallas.append((col, None))

            for col_idx, talla_val in tallas:
                talla_final = talla_val
                if talla_final is None:
                    izquierda = col_idx - 1
                    derecha = col_idx + 1
                    while (izquierda >= 2) or (derecha <= 11):
                        if izquierda >= 2:
                            t = str(df_raw.iat[fila_tallas, izquierda]).strip()
                            if t.startswith("#"):
                                talla_final = t
                                break
                            izquierda -= 1
                        if derecha <= 11:
                            t = str(df_raw.iat[fila_tallas, derecha]).strip()
                            if t.startswith("#"):
                                talla_final = t
                                break
                            derecha += 1
                        if izquierda < 2 and derecha > 11:
                            break
                if not talla_final:
                    continue

                cantidad_raw = str(df_raw.iat[fila_cantidades, col_idx]).strip()
                if not es_numero_valido(cantidad_raw):
                    if cantidad_raw and cantidad_raw.upper() not in ["", "0"] and debug:
                        st.warning(f"[{hoja}] modelo '{modelo}' cantidad no estándar en col {col_idx+1}: '{cantidad_raw}'")
                    continue
                cantidad = float(cantidad_raw)

                if not codigo:
                    if debug:
                        st.warning(f"[{hoja}] modelo '{modelo}' sin código en bloque iniciado en fila {fila_codigo+1}")
                    continue

                resultados.append({
                    "Código del Producto": codigo,
                    "Modelo": modelo,
                    "Talla": talla_final,
                    "Cantidad pares": cantidad,
                    "Hoja": hoja
                })

    if not resultados:
        return pd.DataFrame(columns=["Código del Producto", "Modelo", "Talla", "Cantidad pares", "Hoja"])

    df_res = pd.DataFrame(resultados)
    df_res = df_res.drop_duplicates(subset=["Código del Producto", "Modelo", "Talla", "Cantidad pares", "Hoja"])
    df_res["Modelo_norm"] = df_res["Modelo"].apply(lambda s: str(s).strip().upper())
    df_res["Talla_norm"] = df_res["Talla"].apply(lambda s: str(s).strip().upper())
    return df_res

# ---------- configuración general ----------
DB_PATH = "usuarios.db"
FICHAS_PATH = "FICHAS2.xlsx"
LOGO_PATH = "logo_suolmex.jpg"
st.set_page_config(page_title="Generador de Pedido SUOLMEX (B1-B4)", layout="wide")

# estilo
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

# sidebar
with st.sidebar:
    st.image(LOGO_PATH, width=200)
    st.markdown("### Instrucciones")
    st.markdown("""
    1. Sube tu Excel de programación con hojas B1-B4.
    2. Revisa el pedido extraído.
    3. Genera el pedido consolidado y descarga el PDF.
    """)
    debug_prog = st.checkbox("Mostrar debug de programación")
    debug_fichas = st.checkbox("Mostrar debug de fichas")
    recargar = st.checkbox("Recargar Excel de fichas (forzar)", key="force_reparse")

# sesión y usuarios
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

# db
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

# admin default
if not c.execute("SELECT * FROM usuarios WHERE codigo = 'admin'").fetchone():
    c.execute("INSERT INTO usuarios (codigo, contrasena, rol) VALUES (?, ?, ?)",
              ('admin', encriptar_contra('admin123'), 'admin'))
    conn.commit()

# cargar sesión
if "logueado" not in st.session_state:
    cargar_sesion()

# login
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

# cerrar sesión
if st.button("Cerrar sesión"):
    ruta = path_sesion_local()
    if Path(ruta).exists():
        Path(ruta).unlink()
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    st.success("Sesión cerrada.")
    st.experimental_set_query_params()  # forzar reevaluación mínima

st.success(f"Sesión iniciada como **{st.session_state.usuario}** ({st.session_state.rol})")

# historial
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

# gestión usuarios (admin)
if st.session_state.rol == "admin":
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

# carga de fichas
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

# controlar recarga sin usar experimental_rerun
if recargar:
    st.session_state["fichas_reload_flag"] = not st.session_state.get("fichas_reload_flag", False)

mtime = os.path.getmtime(FICHAS_PATH) if os.path.exists(FICHAS_PATH) else 0
# Si cambió el flag, invalida el cache automáticamente por dependencia
fichas = cargar_fichas(mtime)

if debug_fichas:
    st.subheader("Debug: fichas cargadas (muestra)")
    st.dataframe(fichas[["Codigo del Producto", "Linea", "Corrida", "__Linea_clean", "__Corrida_clean"]].drop_duplicates().head(100))

# estado inicial
if "pedido_total" not in st.session_state:
    st.session_state["pedido_total"] = []
if "corrida_seleccionada" not in st.session_state:
    st.session_state["corrida_seleccionada"] = None

st.markdown("---")
st.title("Generador de Pedido SUOLMEX (B1–B4 consolidado)")

# uploader
uploaded_prog = st.file_uploader("Sube tu Excel de programación", type=["xlsx"], key="prog_excel")
codigo_producto_override = st.text_input("Código del Producto (override si no lo detecta)", value="")

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
                    "Código": codigo,
                    "Modelo": modelo,
                    "Talla": corrida,
                    "Cantidad pares": cantidad,
                    "Peso Total (g)": peso_total,
                    "Poliol (g)": cantidad_poliol,
                    "ISO (g)": cantidad_iso,
                    "Hoja": ficha['Hoja']
                }
                if nuevo not in st.session_state["pedido_total"]:
                    st.session_state["pedido_total"].append(nuevo)
            else:
                st.warning(f"No se encontró ficha para código='{codigo}', modelo='{modelo}', talla='{corrida}'.")

        st.success("Explosión automática agregada al pedido.")

# ingreso manual
st.markdown("---")
st.subheader("Ingreso manual")
codigo_manual = st.selectbox("Código del Producto:", sorted(fichas["Codigo del Producto"].unique()))
modelos = fichas[fichas["Codigo del Producto"] == codigo_manual]["Linea"].unique()
modelo_manual = st.selectbox("Modelo:", sorted(modelos))
cantidad = st.number_input("Cantidad de pares:", min_value=1, step=1)
corridas = fichas[
    (fichas["Codigo del Producto"] == codigo_manual) &
    (fichas["Linea"] == modelo_manual)
]["Corrida"].unique()
st.markdown("#### Selecciona una talla:")
cols = st.columns(min(5, len(corridas)))
for i, talla in enumerate(sorted(corridas)):
    if cols[i % 5].button(talla):
        st.session_state["corrida_seleccionada"] = talla

if st.session_state["corrida_seleccionada"]:
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
            "Código": codigo_manual,
            "Modelo": modelo_manual,
            "Talla": corrida,
            "Cantidad pares": cantidad,
            "Peso Total (g)": peso_total,
            "Poliol (g)": cantidad_poliol,
            "ISO (g)": cantidad_iso,
            "Hoja": ficha_manual['Hoja']
        }
        if nuevo not in st.session_state["pedido_total"]:
            st.session_state["pedido_total"].append(nuevo)
        st.success(f"Agregado: {modelo_manual} - Talla {corrida} - {cantidad} pares.")
        st.session_state["corrida_seleccionada"] = None

# resumen pedido
if st.session_state["pedido_total"]:
    st.markdown("---")
    st.subheader("Resumen del Pedido")

    resumen_df = pd.DataFrame(st.session_state["pedido_total"])

    for idx, row in resumen_df.iterrows():
        cols = st.columns([5, 1])
        with cols[0]:
            st.markdown(
                f"**Código:** {row['Código']} | **Modelo:** {row['Modelo']} | "
                f"**Talla:** {row['Talla']} | **Cantidad:** {row['Cantidad pares']} pares | "
                f"**Poliol:** {row['Poliol (g)']:.2f} g | **ISO:** {row['ISO (g)']:.2f} g"
            )
        with cols[1]:
            if st.button("Eliminar", key=f"eliminar_{idx}"):
                st.session_state["pedido_total"].pop(idx)
                st.success("Elemento eliminado.")

    resumen_df = pd.DataFrame(st.session_state["pedido_total"])
    total_poliol = resumen_df["Poliol (g)"].sum() / 1000
    total_iso = resumen_df["ISO (g)"].sum() / 1000
    poliol_merma = total_poliol * 0.03
    iso_merma = total_iso * 0.03
    total_poliol_con_merma = total_poliol + poliol_merma
    total_iso_con_merma = total_iso + iso_merma
    mezcla_total_kg = total_poliol + total_iso
    mezcla_con_merma = total_poliol_con_merma + total_iso_con_merma

    st.markdown(f"*Poliol necesario:* {total_poliol:.2f} kg")
    st.markdown(f"*ISO necesario:* {total_iso:.2f} kg")
    st.markdown(f"*Poliol con merma (3%):* {total_poliol_con_merma:.2f} kg")
    st.markdown(f"*ISO con merma (3%):* {total_iso_con_merma:.2f} kg")
    st.markdown(f"*Mezcla sin merma:* {mezcla_total_kg:.2f} kg")
    st.markdown(f"*Mezcla total con merma:* {mezcla_con_merma:.2f} kg")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Generar PDF"):
            Path("historial_pedidos").mkdir(exist_ok=True)
            fecha_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            nombre_usuario = st.session_state.usuario
            nombre_archivo = f"historial_pedidos/pedido_{nombre_usuario}_{fecha_hora}.pdf"

            class PDF(FPDF):
                def header(self):
                    self.image(LOGO_PATH, 10, 8, 33)
                    self.set_font("Arial", 'B', 12)
                    self.cell(0, 10, "Resumen de Pedido SUOLMEX", 0, 1, 'C')
                    self.ln(10)

            pdf = PDF()
            pdf.add_page()
            pdf.set_font("Arial", size=10)

            pdf.cell(0, 10, f"Usuario: {nombre_usuario}", ln=True)
            pdf.cell(0, 10, f"Fecha y hora: {fecha_hora.replace('_', ' ')}", ln=True)
            pdf.ln(5)

            headers = ["Código", "Modelo", "Talla", "Pares", "Poliol (g)", "ISO (g)"]
            for h in headers:
                pdf.cell(32, 10, h, 1, 0, 'C')
            pdf.ln()

            for _, row in resumen_df.iterrows():
                pdf.cell(32, 10, str(row['Código']), 1)
                pdf.cell(32, 10, str(row['Modelo']), 1)
                pdf.cell(32, 10, str(row['Talla']), 1)
                pdf.cell(32, 10, str(row['Cantidad pares']), 1)
                pdf.cell(32, 10, f"{row['Poliol (g)']:.1f}", 1)
                pdf.cell(32, 10, f"{row['ISO (g)']:.1f}", 1)
                pdf.ln()

            pdf.ln(5)
            pdf.set_font("Arial", 'B', 10)
            pdf.cell(0, 10, "Totales:", ln=True)
            pdf.set_font("Arial", size=10)
            pdf.cell(0, 10, f"Poliol necesario: {total_poliol:.2f} kg", ln=True)
            pdf.cell(0, 10, f"ISO necesario: {total_iso:.2f} kg", ln=True)
            pdf.cell(0, 10, f"Poliol con merma (3%): {total_poliol_con_merma:.2f} kg", ln=True)
            pdf.cell(0, 10, f"ISO con merma (3%): {total_iso_con_merma:.2f} kg", ln=True)
            pdf.cell(0, 10, f"Mezcla sin merma: {mezcla_total_kg:.2f} kg", ln=True)
            pdf.cell(0, 10, f"Mezcla total con merma: {mezcla_con_merma:.2f} kg", ln=True)

            pdf.output(nombre_archivo)
            st.success(f"PDF generado: {Path(nombre_archivo).name}")

            with open(nombre_archivo, "rb") as f:
                st.download_button("Descargar PDF", data=f, file_name=Path(nombre_archivo).name)
    with col2:
        if st.button("Reiniciar Pedido"):
            st.session_state["pedido_total"] = []
            st.success("Pedido reiniciado.")
