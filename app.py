"""
Generador de Contratos de Financiación v6
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

# ── CSS: badge contador en sidebar, quitar padding superior excesivo ──────────
st.markdown("""
<style>
/* Reduce top padding so the sidebar header sits higher */
[data-testid="stSidebar"] > div:first-child { padding-top: 1.5rem; }

/* Badge pill reutilizable */
.badge {
    display: inline-block;
    background: #1b3a6b;
    color: white;
    border-radius: 999px;
    padding: 2px 10px;
    font-size: 0.78rem;
    font-weight: 600;
    margin-left: 6px;
    vertical-align: middle;
}

/* Botones primarios en azul marino corporativo */
.stButton > button[data-testid="baseButton-primary"] {
    background-color: #1b3a6b !important;
    border-color: #1b3a6b !important;
    color: white !important;
}
.stButton > button[data-testid="baseButton-primary"]:hover {
    background-color: #142f5e !important;
    border-color: #142f5e !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

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
    st.markdown("<div style='max-width: 480px; margin: 0 auto;'>", unsafe_allow_html=True)
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
    st.markdown("</div>", unsafe_allow_html=True)

# def login_screen():
#     col_l, col_c, col_r = st.columns([1, 3, 1])
#     with col_c:
#         logo_path = Path("assets/logo_vcapital.png")
#         if logo_path.exists():
#             st.image(str(logo_path), use_container_width=True)
#         st.markdown("---")
#         st.markdown("### 🔐 Acceso Restringido")
#         st.markdown("Esta aplicación es de uso privado. Introduce la contraseña para continuar.")
#         password = st.text_input("Contraseña", type="password", key="pwd_input")
#         if st.button("Entrar", type="primary", use_container_width=True):
#             if check_password(password):
#                 st.session_state["authenticated"] = True
#                 st.rerun()
#             else:
#                 st.error("❌ Contraseña incorrecta. Inténtalo de nuevo.")

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
# SIDEBAR — logo + botón salir + historial (siempre visible)
# ══════════════════════════════════════════════════════════════════════════════

with st.sidebar:
    logo_path = Path("assets/logo_vcapital.png")
    if logo_path.exists():
        st.image(str(logo_path), use_container_width=True)
    st.markdown("---")

    # ── Historial ─────────────────────────────────────────────────────────────
    n_hist = len(st.session_state["historial"])
    if n_hist == 0:
        st.caption("📋 Sin contratos generados aún.")
    else:
        st.markdown(
            f"**📋 Historial** "
            f"<span class='badge'>{n_hist}</span>",
            unsafe_allow_html=True
        )
        df_hist = pd.DataFrame(st.session_state["historial"])
        st.dataframe(
            df_hist[["Hora", "Cliente", "Producción"]],
            use_container_width=True,
            hide_index=True
        )
        csv_hist = df_hist.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Exportar historial CSV",
            data=csv_hist,
            file_name=f"historial_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True
        )

    # ── Espaciador — empuja footer y botón al fondo ───────────────────────────
    st.markdown("<div style='margin-top: 3rem;'></div>", unsafe_allow_html=True)

    # ── Footer ────────────────────────────────────────────────────────────────
    st.markdown("---")
    st.caption("**v6.0** · Desarrollado por")
    st.markdown(
        "<small>Marc Mestres Mejias<br>"
        "<a href='mailto:m.mestres87@gmail.com'>m.mestres87@gmail.com</a></small>",
        unsafe_allow_html=True
    )
    st.markdown("---")

    # ── Cerrar sesión — siempre al fondo ──────────────────────────────────────
    if st.button("🚪 Cerrar sesión", use_container_width=True):
        logout()

# ══════════════════════════════════════════════════════════════════════════════
# HEADER — compacto (header en lugar de title)
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("## Generador de Contratos de Financiación")
st.caption("Genera contratos Word rellenos automáticamente a partir de la plantilla y el Excel de la producción.")

st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# BARRA DE DESCARGA DE PLANTILLAS
# ══════════════════════════════════════════════════════════════════════════════

def get_archivos_plantillas() -> dict[str, Path]:
    """Devuelve {nombre_archivo: path} para todos los .xlsx y .docx en plantillas/."""
    d = Path("plantillas")
    if not d.exists():
        return {}
    archivos = {}
    for ext in ("*.xlsx", "*.docx"):
        for f in sorted(d.glob(ext)):
            archivos[f.name] = f
    return archivos

_plantillas_repo = get_archivos_plantillas()

if _plantillas_repo:
    with st.expander("📎 Descargar plantillas", expanded=False):
        st.caption("Descarga las plantillas Word y Excel para preparar tus contratos.")
        cols = st.columns(min(len(_plantillas_repo), 3))
        for i, (nombre_f, path_f) in enumerate(_plantillas_repo.items()):
            icono = "📊" if nombre_f.endswith(".xlsx") else "📄"
            mime  = (
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                if nombre_f.endswith(".xlsx") else
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            with cols[i % 3]:
                st.download_button(
                    label=f"{icono} {nombre_f}",
                    data=path_f.read_bytes(),
                    file_name=nombre_f,
                    mime=mime,
                    use_container_width=True,
                    key=f"dl_plantilla_{i}"
                )

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
# VALIDACIÓN DEL EXCEL
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
# HELPER — plantillas .docx disponibles para el selector
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
# HELPER — tarjeta de datos del cliente (sustituye los st.info apilados)
# ══════════════════════════════════════════════════════════════════════════════

def mostrar_datos_cliente(row, tipo: str):
    """Muestra los datos del cliente como tabla compacta en lugar de st.info apilados."""
    if tipo == "PF":
        datos = {
            "Campo": ["Nombre", "DNI", "Domicilio", "CP / Localidad",
                      "Importe inversión", "Importe deducción", "Fecha contrato"],
            "Valor": [
                str(row["nombre"]),
                str(row["dni"]),
                str(row["domicilio"]),
                str(row["cp_localidad"]),
                f"{fmt_euros(row['importe_num'])}  ·  {row['importe_letras']}",
                f"{fmt_euros(row['deduccion_num'])}  ·  {row['deduccion_letras']}",
                parse_fecha(row["fecha"]).strftime("%d/%m/%Y"),
            ]
        }
    else:
        datos = {
            "Campo": ["Empresa", "CIF", "Representante", "DNI representante", "Cargo",
                      "Domicilio", "CP / Localidad",
                      "Importe inversión", "Importe deducción", "Fecha contrato"],
            "Valor": [
                str(row["nombre_empresa"]),
                str(row["cif"]),
                str(row["representante"]),
                str(row["dni_representante"]),
                str(row["cargo_representante"]),
                str(row["domicilio_empresa"]),
                str(row["cp_localidad_empresa"]),
                f"{fmt_euros(row['importe_num'])}  ·  {row['importe_letras']}",
                f"{fmt_euros(row['deduccion_num'])}  ·  {row['deduccion_letras']}",
                parse_fecha(row["fecha"]).strftime("%d/%m/%Y"),
            ]
        }
    st.dataframe(
        pd.DataFrame(datos),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Campo": st.column_config.TextColumn(width="small"),
            "Valor": st.column_config.TextColumn(width="large"),
        }
    )


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

    avisos_pf    = validar_clientes_pf(df_pf) if not df_pf.empty else []
    avisos_pj    = validar_clientes_pj(df_pj) if not df_pj.empty else []
    todos_avisos = avisos_pf + avisos_pj
    if todos_avisos:
        with st.expander(f"⚠️ {len(todos_avisos)} campo(s) vacío(s) en el Excel — pulsa para ver", expanded=False):
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
        tab_individual, tab_lote = st.tabs(["📄 Contrato individual", "📦 Generar todos"])

        # ── TAB INDIVIDUAL ────────────────────────────────────────────────────
        with tab_individual:
            st.subheader("3️⃣ Selecciona el cliente")
            sel_label = st.selectbox("Cliente:", [o["label"] for o in opciones])
            sel = next(o for o in opciones if o["label"] == sel_label)

            if sel["tipo"] == "PF":
                row = df_pf.iloc[sel["idx"]]
                cliente_dict = row.to_dict()
                tipo_sel = "PF"
            else:
                row = df_pj.iloc[sel["idx"]]
                cliente_dict = row.to_dict()
                tipo_sel = "PJ"

            # Tabla compacta en lugar de st.info apilados
            mostrar_datos_cliente(row, tipo_sel)

            # Vista previa de sustituciones
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

                        if sin_sustituir:
                            st.warning(
                                f"⚠️ {len(sin_sustituir)} marcador(es) sin sustituir: "
                                f"{', '.join(sin_sustituir)}"
                            )
                        else:
                            st.success("✅ Contrato generado. Todos los marcadores sustituidos.")

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

        # ── TAB LOTE ──────────────────────────────────────────────────────────
        with tab_lote:
            st.subheader("📦 Generar todos los contratos")
            total = len(opciones)
            st.info(
                f"Se generarán **{total} contratos** "
                f"({len(df_pf)} persona(s) física(s) + {len(df_pj)} persona(s) jurídica(s)) "
                f"y se descargarán en un único archivo `.zip`."
            )

            if st.button(f"⚡ Generar los {total} contratos", type="primary", use_container_width=True, key="btn_lote"):
                errores_lote = []
                archivos_zip = {}
                progress     = st.progress(0, text="Iniciando generación en lote…")

                for i, opcion in enumerate(opciones):
                    try:
                        cd = (df_pf.iloc[opcion["idx"]] if opcion["tipo"] == "PF"
                              else df_pj.iloc[opcion["idx"]]).to_dict()

                        reps_l             = build_replacements(prod_data, cd, opcion["tipo"])
                        docx_bytes_l, ss_l = generate_contract(template_bytes, reps_l)
                        filename_l         = build_filename(prod_data, cd, opcion["tipo"])
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
        # Sin plantilla: mostrar selector y datos de cliente igualmente
        st.subheader("3️⃣ Selecciona el cliente")
        sel_label = st.selectbox("Cliente:", [o["label"] for o in opciones])
        sel = next(o for o in opciones if o["label"] == sel_label)
        row = (df_pf.iloc[sel["idx"]] if sel["tipo"] == "PF" else df_pj.iloc[sel["idx"]])
        mostrar_datos_cliente(row, sel["tipo"])
        st.info("⬆️ Selecciona o sube la plantilla Word para poder generar el contrato.")

elif excel_file and not prod_data:
    st.warning("No se encontraron datos en la hoja PRODUCTORA. Revisa que la fila 4 esté rellena.")
elif excel_file and df_pf.empty and df_pj.empty:
    st.warning("No hay clientes en el Excel. Añade filas en PERSONAS_FISICAS o PERSONAS_JURIDICAS.")
else:
    st.info("⬆️ Sube la plantilla Word y el Excel de la producción para comenzar.")

st.divider()
st.caption("VCapital · Generador de Contratos de Financiación")
