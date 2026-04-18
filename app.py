"""
Generador de Contratos de Financiación v2
Streamlit app — plantilla Word con marcadores «» + Excel estructurado por producción
Con sistema de login por contraseña via Streamlit Secrets
"""

import streamlit as st
import pandas as pd
import zipfile
import os
import re
import tempfile
import hashlib
from datetime import date, datetime

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Generador de Contratos",
    page_icon="📄",
    layout="centered"
)

# ══════════════════════════════════════════════════════════════════════════════
# LOGIN
# ══════════════════════════════════════════════════════════════════════════════

def check_password(password: str) -> bool:
    """Check password against stored hash in Streamlit secrets."""
    try:
        stored_hash = st.secrets["auth"]["password_hash"]
        input_hash  = hashlib.sha256(password.encode()).hexdigest()
        return input_hash == stored_hash
    except Exception:
        return False

def login_screen():
    st.title("🔐 Acceso Restringido")
    st.markdown("Esta aplicación es de uso privado. Introduce la contraseña para continuar.")
    st.divider()

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        password = st.text_input("Contraseña", type="password", key="pwd_input")
        if st.button("Entrar", type="primary", use_container_width=True):
            if check_password(password):
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("❌ Contraseña incorrecta. Inténtalo de nuevo.")

def logout():
    st.session_state["authenticated"] = False
    st.rerun()

# ── Auth gate ─────────────────────────────────────────────────────────────────
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    login_screen()
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# APP (only shown after login)
# ══════════════════════════════════════════════════════════════════════════════

# Header with logout button
col_title, col_logout = st.columns([5, 1])
with col_title:
    st.title("📄 Generador de Contratos de Financiación")
with col_logout:
    st.write("")
    if st.button("🚪 Salir"):
        logout()

st.markdown("Genera contratos Word rellenos automáticamente a partir de la plantilla y el Excel de la producción.")

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS — XML replacement
# ══════════════════════════════════════════════════════════════════════════════

def replace_in_xml(xml: str, old: str, new: str) -> str:
    """Replace `old` with `new` inside every <w:t>…</w:t> node."""
    new_xml = (str(new)
               .replace("&", "&amp;")
               .replace("<", "&lt;")
               .replace(">", "&gt;"))
    result, i = [], 0
    while i < len(xml):
        wt_start = xml.find("<w:t", i)
        if wt_start == -1:
            result.append(xml[i:])
            break
        tag_end  = xml.find(">", wt_start) + 1
        wt_close = xml.find("</w:t>", tag_end)
        if wt_close == -1:
            result.append(xml[i:])
            break
        result.append(xml[i:wt_start])
        result.append(xml[wt_start:tag_end])
        result.append(xml[tag_end:wt_close].replace(old, new_xml))
        result.append("</w:t>")
        i = wt_close + 6
    return "".join(result)


def apply_replacements(xml: str, reps: dict) -> str:
    for old, new in reps.items():
        if old and new is not None:
            xml = replace_in_xml(xml, old, str(new))
    return xml


def fmt_euros(n) -> str:
    try:
        return f"{int(n):,}€".replace(",", ".")
    except Exception:
        return str(n)


def nombre_mes(n: int) -> str:
    return ["enero","febrero","marzo","abril","mayo","junio",
            "julio","agosto","septiembre","octubre","noviembre","diciembre"][n-1]


def parse_fecha(val) -> date:
    if isinstance(val, (datetime, date)):
        return val if isinstance(val, date) else val.date()
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return date.today()


# ══════════════════════════════════════════════════════════════════════════════
# CONTRACT GENERATION
# ══════════════════════════════════════════════════════════════════════════════

def build_replacements(prod: dict, cliente: dict, tipo: str) -> dict:
    fecha = parse_fecha(cliente.get("fecha", date.today()))
    r = {}

    r["«DIA»"]  = f"{fecha.day:02d}"
    r["«MES»"]  = nombre_mes(fecha.month)
    r["«ANYO»"] = str(fecha.year)

    r["«PROD_REPRESENTANTE»"]     = str(prod.get("representante", "")).upper()
    r["«PROD_DNI»"]               = str(prod.get("dni_representante", ""))
    r["«PROD_DOMICILIO»"]         = str(prod.get("domicilio_representante", ""))
    r["«PROD_CP_LOCALIDAD»"]      = str(prod.get("cp_localidad_representante", ""))
    r["«PROD_NOMBRE_EMPRESA»"]    = str(prod.get("nombre_empresa", "")).upper()
    r["«PROD_CIF»"]               = str(prod.get("cif", ""))
    r["«PROD_DOMICILIO_EMPRESA»"] = str(prod.get("domicilio_empresa", ""))
    r["«PROD_CP_EMPRESA»"]        = str(prod.get("cp_localidad_empresa", ""))
    r["«ENTIDAD_BANCARIA»"]       = str(prod.get("entidad_bancaria", ""))
    r["«IBAN»"]                   = str(prod.get("iban", ""))
    r["«NOMBRE_PRODUCCION»"]      = str(prod.get("nombre_produccion", ""))
    r["«PRODUCTORA»"]             = str(prod.get("nombre_empresa", "")).upper()

    if tipo == "PF":
        nombre = str(cliente.get("nombre", "")).upper()
        r["«NOMBRE_Y_APELLIDOS_PF»"] = nombre
        r["«DNI»"]                   = str(cliente.get("dni", ""))
        r["«DOMICILIO»"]             = str(cliente.get("domicilio", ""))
        r["«CP_Y_LOCALIDAD»"]        = str(cliente.get("cp_localidad", ""))
        r["«CIF»"]                   = ""
        r["«NOMBRE_EMPRESA»"]        = nombre
        r["«IMPORTE_EN_LETRAS»"]     = str(cliente.get("importe_letras", "")).upper()
        r["«IMP_NUMERO»"]            = fmt_euros(cliente.get("importe_num", 0))
        r["«IMPORTE_LETRAS__BASE»"]  = str(cliente.get("deduccion_letras", "")).upper()
        r["«IMP_NUMERO__BASE»"]      = fmt_euros(cliente.get("deduccion_num", 0))
        r["«CLIENTE»"]               = nombre
        r["«FIN_REPRESENTANTE»"]     = nombre

    elif tipo == "PJ":
        rep     = str(cliente.get("representante", "")).upper()
        empresa = str(cliente.get("nombre_empresa", "")).upper()
        r["«NOMBRE_Y_APELLIDOS_PF»"] = rep
        r["«DNI»"]                   = str(cliente.get("dni_representante", ""))
        r["«DOMICILIO»"]             = str(cliente.get("domicilio_empresa", ""))
        r["«CP_Y_LOCALIDAD»"]        = str(cliente.get("cp_localidad_empresa", ""))
        r["«NOMBRE_EMPRESA»"]        = empresa
        r["«CIF»"]                   = str(cliente.get("cif", ""))
        r["«IMPORTE_EN_LETRAS»"]     = str(cliente.get("importe_letras", "")).upper()
        r["«IMP_NUMERO»"]            = fmt_euros(cliente.get("importe_num", 0))
        r["«IMPORTE_LETRAS__BASE»"]  = str(cliente.get("deduccion_letras", "")).upper()
        r["«IMP_NUMERO__BASE»"]      = fmt_euros(cliente.get("deduccion_num", 0))
        r["«CLIENTE»"]               = empresa
        r["«FIN_REPRESENTANTE»"]     = rep

    return r


def generate_contract(template_bytes: bytes, reps: dict) -> bytes:
    with tempfile.TemporaryDirectory() as tmp:
        tpl_path = os.path.join(tmp, "template.docx")
        out_path = os.path.join(tmp, "output.docx")

        with open(tpl_path, "wb") as f:
            f.write(template_bytes)

        with zipfile.ZipFile(tpl_path, "r") as z:
            z.extractall(tmp)

        doc_xml = os.path.join(tmp, "word", "document.xml")
        with open(doc_xml, "r", encoding="utf-8") as f:
            xml = f.read()

        xml = apply_replacements(xml, reps)

        with open(doc_xml, "w", encoding="utf-8") as f:
            f.write(xml)

        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for root, _, files in os.walk(tmp):
                for fname in files:
                    if fname in ("template.docx", "output.docx"):
                        continue
                    fpath = os.path.join(root, fname)
                    arcname = os.path.relpath(fpath, tmp)
                    zout.write(fpath, arcname)

        with open(out_path, "rb") as f:
            return f.read()


# ══════════════════════════════════════════════════════════════════════════════
# EXCEL READING
# ══════════════════════════════════════════════════════════════════════════════

def read_productora(xls: pd.ExcelFile) -> dict | None:
    try:
        df = pd.read_excel(xls, sheet_name="PRODUCTORA", header=1)
        df = df.dropna(how="all").iloc[1:].reset_index(drop=True).dropna(how="all")
        if df.empty:
            return None
        row = df.iloc[0]
        return {
            "representante":              str(row.iloc[0]) if pd.notna(row.iloc[0]) else "",
            "dni_representante":          str(row.iloc[1]) if pd.notna(row.iloc[1]) else "",
            "cargo_representante":        str(row.iloc[2]) if pd.notna(row.iloc[2]) else "Administrador único",
            "domicilio_representante":    str(row.iloc[3]) if pd.notna(row.iloc[3]) else "",
            "cp_localidad_representante": str(row.iloc[4]) if pd.notna(row.iloc[4]) else "",
            "nombre_empresa":             str(row.iloc[5]) if pd.notna(row.iloc[5]) else "",
            "cif":                        str(row.iloc[6]) if pd.notna(row.iloc[6]) else "",
            "domicilio_empresa":          str(row.iloc[7]) if pd.notna(row.iloc[7]) else "",
            "cp_localidad_empresa":       str(row.iloc[8]) if pd.notna(row.iloc[8]) else "",
            "nombre_produccion":          str(row.iloc[9]) if pd.notna(row.iloc[9]) else "",
            "entidad_bancaria":           str(row.iloc[10]) if pd.notna(row.iloc[10]) else "",
            "iban":                       str(row.iloc[11]) if pd.notna(row.iloc[11]) else "",
        }
    except Exception as e:
        st.error(f"Error leyendo hoja PRODUCTORA: {e}")
        return None


def read_clientes_pf(xls: pd.ExcelFile) -> pd.DataFrame:
    try:
        df = pd.read_excel(xls, sheet_name="PERSONAS_FISICAS", header=1)
        df = df.iloc[1:].reset_index(drop=True)
        df = df[df.iloc[:,0].notna() & ~df.iloc[:,0].astype(str).str.startswith("ℹ️")]
        df = df.reset_index(drop=True)
        df.columns = ["nombre","dni","domicilio","cp_localidad",
                      "importe_num","importe_letras",
                      "deduccion_num","deduccion_letras","fecha"]
        return df
    except Exception as e:
        st.error(f"Error leyendo hoja PERSONAS_FISICAS: {e}")
        return pd.DataFrame()


def read_clientes_pj(xls: pd.ExcelFile) -> pd.DataFrame:
    try:
        df = pd.read_excel(xls, sheet_name="PERSONAS_JURIDICAS", header=1)
        df = df.iloc[1:].reset_index(drop=True)
        df = df[df.iloc[:,0].notna() & ~df.iloc[:,0].astype(str).str.startswith("ℹ️")]
        df = df.reset_index(drop=True)
        df.columns = ["representante","dni_representante","cargo_representante",
                      "nombre_empresa","cif","domicilio_empresa","cp_localidad_empresa",
                      "importe_num","importe_letras",
                      "deduccion_num","deduccion_letras","fecha"]
        return df
    except Exception as e:
        st.error(f"Error leyendo hoja PERSONAS_JURIDICAS: {e}")
        return pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════════════════════

col1, col2 = st.columns(2)
with col1:
    st.subheader("1️⃣ Plantilla Word")
    template_file = st.file_uploader("Sube la plantilla del contrato (.docx)",
                                     type=["docx"], key="template")
with col2:
    st.subheader("2️⃣ Excel de la producción")
    excel_file = st.file_uploader("Sube el Excel de la producción (.xlsx)",
                                  type=["xlsx"], key="excel")

st.divider()

prod_data = None
df_pf     = pd.DataFrame()
df_pj     = pd.DataFrame()

if excel_file:
    try:
        xls       = pd.ExcelFile(excel_file)
        prod_data = read_productora(xls)
        df_pf     = read_clientes_pf(xls)
        df_pj     = read_clientes_pj(xls)
    except Exception as e:
        st.error(f"Error al leer el Excel: {e}")

if prod_data:
    st.subheader("🎬 Productora detectada")
    c1, c2, c3 = st.columns(3)
    c1.info(f"**Empresa:** {prod_data.get('nombre_empresa','—')}")
    c2.info(f"**Producción:** {prod_data.get('nombre_produccion','—')}")
    c3.info(f"**Banco:** {prod_data.get('entidad_bancaria','—')}")
    st.divider()

if prod_data and (not df_pf.empty or not df_pj.empty):
    st.subheader("3️⃣ Selecciona el cliente")

    opciones = []
    if not df_pf.empty:
        for i, row in df_pf.iterrows():
            opciones.append({"label": f"👤 {row['nombre']}", "tipo": "PF", "idx": i})
    if not df_pj.empty:
        for i, row in df_pj.iterrows():
            opciones.append({"label": f"🏢 {row['nombre_empresa']} ({row['representante']})",
                             "tipo": "PJ", "idx": i})

    sel_label = st.selectbox("Cliente:", [o["label"] for o in opciones])
    sel = next(o for o in opciones if o["label"] == sel_label)

    st.markdown("**Datos del cliente:**")
    if sel["tipo"] == "PF":
        row = df_pf.iloc[sel["idx"]]
        c1, c2 = st.columns(2)
        with c1:
            st.info(f"**Nombre:** {row['nombre']}")
            st.info(f"**DNI:** {row['dni']}")
            st.info(f"**Domicilio:** {row['domicilio']}")
            st.info(f"**CP / Localidad:** {row['cp_localidad']}")
        with c2:
            st.info(f"**Importe inversión:** {fmt_euros(row['importe_num'])} — {row['importe_letras']}")
            st.info(f"**Importe deducción:** {fmt_euros(row['deduccion_num'])} — {row['deduccion_letras']}")
            st.info(f"**Fecha contrato:** {row['fecha']}")
        cliente_dict = row.to_dict()
        tipo_sel = "PF"
    else:
        row = df_pj.iloc[sel["idx"]]
        c1, c2 = st.columns(2)
        with c1:
            st.info(f"**Empresa:** {row['nombre_empresa']}")
            st.info(f"**CIF:** {row['cif']}")
            st.info(f"**Representante:** {row['representante']}")
            st.info(f"**DNI rep.:** {row['dni_representante']}")
            st.info(f"**Cargo:** {row['cargo_representante']}")
        with c2:
            st.info(f"**Domicilio:** {row['domicilio_empresa']}")
            st.info(f"**CP / Localidad:** {row['cp_localidad_empresa']}")
            st.info(f"**Importe inversión:** {fmt_euros(row['importe_num'])} — {row['importe_letras']}")
            st.info(f"**Importe deducción:** {fmt_euros(row['deduccion_num'])} — {row['deduccion_letras']}")
            st.info(f"**Fecha contrato:** {row['fecha']}")
        cliente_dict = row.to_dict()
        tipo_sel = "PJ"

    st.divider()

    if template_file:
        if st.button("⚡ Generar Contrato", type="primary", use_container_width=True):
            with st.spinner("Generando contrato..."):
                try:
                    reps       = build_replacements(prod_data, cliente_dict, tipo_sel)
                    docx_bytes = generate_contract(template_file.read(), reps)

                    nombre_f = (cliente_dict.get("nombre","cliente") if tipo_sel == "PF"
                                else cliente_dict.get("nombre_empresa","empresa")).replace(" ","_")
                    fecha_f  = parse_fecha(cliente_dict.get("fecha", date.today())).strftime("%d_%m_%Y")
                    prod_f   = prod_data.get("nombre_produccion","").replace(" ","_").replace('"',"").replace("'","")
                    filename = f"Contrato_{prod_f}_{nombre_f}_{fecha_f}.docx"

                    st.success("✅ Contrato generado correctamente.")
                    st.download_button(
                        label="📥 Descargar Contrato (.docx)",
                        data=docx_bytes,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error al generar el contrato: {e}")
                    import traceback
                    st.code(traceback.format_exc())
    else:
        st.info("⬆️ Sube la plantilla Word para poder generar el contrato.")

elif excel_file and not prod_data:
    st.warning("No se encontraron datos en la hoja PRODUCTORA. Revisa que la fila 4 esté rellena.")
elif excel_file and df_pf.empty and df_pj.empty:
    st.warning("No hay clientes en el Excel. Añade filas en PERSONAS_FISICAS o PERSONAS_JURIDICAS.")
else:
    st.info("⬆️ Sube la plantilla Word y el Excel de la producción para comenzar.")

st.divider()
st.caption("💡 Descarga la plantilla Excel vacía en la pestaña INSTRUCCIONES del propio fichero.")
