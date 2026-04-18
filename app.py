"""
Generador de Contratos de Financiación v4
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
import io
from datetime import date, datetime
from pathlib import Path

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="VCapital · Contratos",
    page_icon="assets/favicon.ico",
    layout="centered"
)

# ══════════════════════════════════════════════════════════════════════════════
# LOGIN
# ══════════════════════════════════════════════════════════════════════════════

def check_password(password: str) -> bool:
    try:
        stored_hash = st.secrets["auth"]["password_hash"]
        input_hash  = hashlib.sha256(password.encode()).hexdigest()
        return input_hash == stored_hash
    except Exception:
        return False

def login_screen():
    col_l, col_c, col_r = st.columns([1, 2, 1])
    with col_c:
        logo_path = Path("assets/logo_vcapital.png")
        if logo_path.exists():
            st.image(str(logo_path), use_container_width=True)
        st.markdown("---")
        st.markdown("### 🔐 Acceso Restringido")
        st.markdown("Esta aplicación es de uso privado. Introduce la contraseña para continuar.")
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

# ── Inicializar historial de sesión ───────────────────────────────────────────
if "historial" not in st.session_state:
    st.session_state["historial"] = []

# ══════════════════════════════════════════════════════════════════════════════
# APP HEADER
# ══════════════════════════════════════════════════════════════════════════════

col_logo, col_title, col_logout = st.columns([1, 4, 1])
with col_logo:
    logo_path = Path("assets/logo_vcapital.png")
    if logo_path.exists():
        st.image(str(logo_path), use_container_width=True)
with col_title:
    st.title("Generador de Contratos")
    st.markdown("Genera contratos Word rellenos automáticamente a partir de la plantilla y el Excel de la producción.")
with col_logout:
    st.write("")
    st.write("")
    if st.button("🚪 Salir"):
        logout()

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS — XML replacement
# ══════════════════════════════════════════════════════════════════════════════

def replace_in_xml(xml: str, old: str, new: str) -> str:
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


def generate_contract(template_bytes: bytes, reps: dict) -> tuple[bytes, list[str]]:
    """Genera el contrato y devuelve (docx_bytes, marcadores_sin_sustituir)."""
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

        # ── Mejora 2: detectar marcadores sin sustituir ───────────────────────
        sin_sustituir = sorted(set(re.findall(r"«[^»]+»", xml)))

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
            return f.read(), sin_sustituir


def build_filename(prod_data: dict, cliente_dict: dict, tipo_sel: str) -> str:
    nombre_f = (cliente_dict.get("nombre", "cliente") if tipo_sel == "PF"
                else cliente_dict.get("nombre_empresa", "empresa")).replace(" ", "_")
    fecha_f  = parse_fecha(cliente_dict.get("fecha", date.today())).strftime("%d_%m_%Y")
    prod_f   = prod_data.get("nombre_produccion", "").replace(" ", "_").replace('"', "").replace("'", "")
    return f"Contrato_{prod_f}_{nombre_f}_{fecha_f}.docx"


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
# MEJORA 3: Validación del Excel
# ══════════════════════════════════════════════════════════════════════════════

def validar_clientes_pf(df: pd.DataFrame) -> list[str]:
    campos = {
        "nombre": "Nombre", "dni": "DNI", "domicilio": "Domicilio",
        "cp_localidad": "CP/Localidad", "importe_num": "Importe inversión (€)",
        "importe_letras": "Importe inversión (letras)", "deduccion_num": "Importe deducción (€)",
        "deduccion_letras": "Importe deducción (letras)", "fecha": "Fecha",
    }
    avisos = []
    for i, row in df.iterrows():
        nombre_c = str(row.get("nombre", f"Fila {i+1}"))
        for campo, etiqueta in campos.items():
            val = row.get(campo)
            if pd.isna(val) or str(val).strip() in ("", "nan"):
                avisos.append(f"👤 **{nombre_c}** — falta: {etiqueta}")
    return avisos


def validar_clientes_pj(df: pd.DataFrame) -> list[str]:
    campos = {
        "nombre_empresa": "Nombre empresa", "cif": "CIF",
        "representante": "Representante", "dni_representante": "DNI representante",
        "domicilio_empresa": "Domicilio", "cp_localidad_empresa": "CP/Localidad",
        "importe_num": "Importe inversión (€)", "importe_letras": "Importe inversión (letras)",
        "deduccion_num": "Importe deducción (€)", "deduccion_letras": "Importe deducción (letras)",
        "fecha": "Fecha",
    }
    avisos = []
    for i, row in df.iterrows():
        empresa_c = str(row.get("nombre_empresa", f"Fila {i+1}"))
        for campo, etiqueta in campos.items():
            val = row.get(campo)
            if pd.isna(val) or str(val).strip() in ("", "nan"):
                avisos.append(f"🏢 **{empresa_c}** — falta: {etiqueta}")
    return avisos


# ══════════════════════════════════════════════════════════════════════════════
# HELPER — plantillas del repositorio
# ══════════════════════════════════════════════════════════════════════════════

def get_plantillas_disponibles() -> list[str]:
    plantillas_dir = Path("plantillas")
    if not plantillas_dir.exists():
        return []
    return sorted([f.name for f in plantillas_dir.glob("*.docx")])


def load_plantilla_bytes(nombre: str) -> bytes | None:
    path = Path("plantillas") / nombre
    if path.exists():
        return path.read_bytes()
    return None


# ══════════════════════════════════════════════════════════════════════════════
# UI — Paso 1: Plantilla Word + Paso 2: Excel
# ══════════════════════════════════════════════════════════════════════════════

col1, col2 = st.columns(2)

with col1:
    st.subheader("1️⃣ Plantilla Word")
    plantillas_disponibles = get_plantillas_disponibles()

    if plantillas_disponibles:
        opciones_plantilla = ["📁 Selecciona una plantilla…"] + plantillas_disponibles + ["⬆️ Subir plantilla propia…"]
        seleccion_plantilla = st.selectbox("Elige la plantilla del contrato:", opciones_plantilla, key="sel_plantilla")
        if seleccion_plantilla == "⬆️ Subir plantilla propia…":
            template_file  = st.file_uploader("Sube tu plantilla (.docx)", type=["docx"], key="template_upload")
            template_bytes = template_file.read() if template_file else None
            template_name  = template_file.name if template_file else None
        elif seleccion_plantilla == "📁 Selecciona una plantilla…":
            template_bytes = None
            template_name  = None
        else:
            template_bytes = load_plantilla_bytes(seleccion_plantilla)
            template_name  = seleccion_plantilla
            if template_bytes:
                st.success(f"✅ Plantilla cargada: **{seleccion_plantilla}**")
    else:
        st.caption("No hay plantillas en la carpeta `plantillas/`. Sube una manualmente.")
        template_file  = st.file_uploader("Sube la plantilla del contrato (.docx)", type=["docx"], key="template_upload_fallback")
        template_bytes = template_file.read() if template_file else None
        template_name  = template_file.name if template_file else None

with col2:
    st.subheader("2️⃣ Excel de la producción")
    excel_file = st.file_uploader("Sube el Excel de la producción (.xlsx)", type=["xlsx"], key="excel")

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# UI — Lectura del Excel + Validación
# ══════════════════════════════════════════════════════════════════════════════

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

    # ── Mejora 3: Validación ──────────────────────────────────────────────────
    avisos_pf = validar_clientes_pf(df_pf) if not df_pf.empty else []
    avisos_pj = validar_clientes_pj(df_pj) if not df_pj.empty else []
    todos_avisos = avisos_pf + avisos_pj
    if todos_avisos:
        with st.expander(f"⚠️ {len(todos_avisos)} campo(s) vacío(s) detectado(s) en el Excel — pulsa para ver", expanded=False):
            for aviso in todos_avisos:
                st.warning(aviso)
    else:
        st.success("✅ Excel validado. Todos los campos obligatorios están rellenos.")

    st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# UI — Clientes y generación
# ══════════════════════════════════════════════════════════════════════════════

if prod_data and (not df_pf.empty or not df_pj.empty):

    opciones = []
    if not df_pf.empty:
        for i, row in df_pf.iterrows():
            opciones.append({"label": f"👤 {row['nombre']}", "tipo": "PF", "idx": i})
    if not df_pj.empty:
        for i, row in df_pj.iterrows():
            opciones.append({"label": f"🏢 {row['nombre_empresa']} ({row['representante']})",
                             "tipo": "PJ", "idx": i})

    if template_bytes:
        # ── Mejora 1: Tabs individual / lote ─────────────────────────────────
        tab_individual, tab_lote = st.tabs(["📄 Contrato individual", "📦 Generar todos"])

        # ─────────────────────────────────────────────────────────────────────
        with tab_individual:
            st.subheader("3️⃣ Selecciona el cliente")
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

            # ── Mejora 4: Vista previa de sustituciones ───────────────────────
            reps_preview = build_replacements(prod_data, cliente_dict, tipo_sel)
            with st.expander("🔍 Revisar valores antes de generar", expanded=False):
                items = list(reps_preview.items())
                mitad = len(items) // 2
                col_a, col_b = st.columns(2)
                with col_a:
                    for marcador, valor in items[:mitad]:
                        st.markdown(f"`{marcador}` → **{valor}**")
                with col_b:
                    for marcador, valor in items[mitad:]:
                        st.markdown(f"`{marcador}` → **{valor}**")

            st.write("")
            if st.button("⚡ Generar Contrato", type="primary", use_container_width=True, key="btn_individual"):
                with st.spinner("Generando contrato..."):
                    try:
                        reps = build_replacements(prod_data, cliente_dict, tipo_sel)
                        docx_bytes, sin_sustituir = generate_contract(template_bytes, reps)
                        filename = build_filename(prod_data, cliente_dict, tipo_sel)

                        # ── Mejora 2: Advertencia marcadores sin sustituir ────
                        if sin_sustituir:
                            st.warning(
                                f"⚠️ {len(sin_sustituir)} marcador(es) sin sustituir: "
                                f"{', '.join(sin_sustituir)}"
                            )
                        else:
                            st.success("✅ Contrato generado. Todos los marcadores sustituidos.")

                        # ── Mejora 5: Añadir al historial ─────────────────────
                        nombre_cliente = (cliente_dict.get("nombre", "—") if tipo_sel == "PF"
                                          else cliente_dict.get("nombre_empresa", "—"))
                        st.session_state["historial"].append({
                            "Hora": datetime.now().strftime("%H:%M:%S"),
                            "Producción": prod_data.get("nombre_produccion", "—"),
                            "Cliente": nombre_cliente,
                            "Tipo": "Persona Física" if tipo_sel == "PF" else "Persona Jurídica",
                            "Plantilla": template_name or "—",
                            "Archivo": filename,
                            "Marcadores sin sustituir": len(sin_sustituir),
                        })

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

        # ─────────────────────────────────────────────────────────────────────
        with tab_lote:
            st.subheader("📦 Generar todos los contratos")
            total = len(opciones)
            st.info(
                f"Se generarán **{total} contratos** "
                f"({len(df_pf)} persona(s) física(s) + {len(df_pj)} persona(s) jurídica(s)) "
                f"y se descargarán en un único archivo `.zip`."
            )

            if st.button(f"⚡ Generar los {total} contratos", type="primary", use_container_width=True, key="btn_lote"):
                errores_lote  = []
                archivos_zip  = {}
                progress      = st.progress(0, text="Iniciando generación en lote…")

                for i, opcion in enumerate(opciones):
                    try:
                        if opcion["tipo"] == "PF":
                            cd = df_pf.iloc[opcion["idx"]].to_dict()
                        else:
                            cd = df_pj.iloc[opcion["idx"]].to_dict()

                        reps_l              = build_replacements(prod_data, cd, opcion["tipo"])
                        docx_bytes_l, ss_l  = generate_contract(template_bytes, reps_l)
                        filename_l          = build_filename(prod_data, cd, opcion["tipo"])
                        archivos_zip[filename_l] = docx_bytes_l

                        nombre_c = (cd.get("nombre", "—") if opcion["tipo"] == "PF"
                                    else cd.get("nombre_empresa", "—"))
                        st.session_state["historial"].append({
                            "Hora": datetime.now().strftime("%H:%M:%S"),
                            "Producción": prod_data.get("nombre_produccion", "—"),
                            "Cliente": nombre_c,
                            "Tipo": "Persona Física" if opcion["tipo"] == "PF" else "Persona Jurídica",
                            "Plantilla": template_name or "—",
                            "Archivo": filename_l,
                            "Marcadores sin sustituir": len(ss_l),
                        })
                    except Exception as e:
                        errores_lote.append(f"{opcion['label']}: {e}")

                    progress.progress((i + 1) / total, text=f"Generando {i+1}/{total}: {opcion['label']}")

                progress.empty()

                # Empaquetar en ZIP
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for nombre_doc, contenido in archivos_zip.items():
                        zf.writestr(nombre_doc, contenido)
                zip_buffer.seek(0)

                prod_nombre  = prod_data.get("nombre_produccion", "produccion").replace(" ", "_")
                zip_filename = f"Contratos_{prod_nombre}.zip"

                if errores_lote:
                    st.warning(f"⚠️ {len(errores_lote)} error(es) durante la generación:")
                    for err in errores_lote:
                        st.error(err)

                st.success(f"✅ {len(archivos_zip)} contrato(s) generado(s) correctamente.")
                st.download_button(
                    label=f"📥 Descargar ZIP ({len(archivos_zip)} contratos)",
                    data=zip_buffer,
                    file_name=zip_filename,
                    mime="application/zip",
                    use_container_width=True
                )

    else:
        # Sin plantilla: mostrar datos del cliente igualmente
        st.subheader("3️⃣ Selecciona el cliente")
        sel_label = st.selectbox("Cliente:", [o["label"] for o in opciones])
        sel = next(o for o in opciones if o["label"] == sel_label)

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

        st.info("⬆️ Selecciona o sube la plantilla Word para poder generar el contrato.")

elif excel_file and not prod_data:
    st.warning("No se encontraron datos en la hoja PRODUCTORA. Revisa que la fila 4 esté rellena.")
elif excel_file and df_pf.empty and df_pj.empty:
    st.warning("No hay clientes en el Excel. Añade filas en PERSONAS_FISICAS o PERSONAS_JURIDICAS.")
else:
    st.info("⬆️ Sube la plantilla Word y el Excel de la producción para comenzar.")

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# Mejora 5: Historial de la sesión
# ══════════════════════════════════════════════════════════════════════════════

if st.session_state["historial"]:
    with st.expander(f"📋 Historial de esta sesión — {len(st.session_state['historial'])} contrato(s) generado(s)", expanded=False):
        df_hist = pd.DataFrame(st.session_state["historial"])
        st.dataframe(df_hist, use_container_width=True, hide_index=True)
        csv_hist = df_hist.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Descargar historial CSV",
            data=csv_hist,
            file_name=f"historial_contratos_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv"
        )

st.caption("💡 Descarga la plantilla Excel vacía en la pestaña INSTRUCCIONES del propio fichero.")
