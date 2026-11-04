import streamlit as st
import pandas as pd
import json, io, hashlib, re
from datetime import datetime
from zoneinfo import ZoneInfo
from PIL import Image, ImageOps
try:
    import pdfplumber
except Exception:
    pdfplumber = None
from google import genai
from google.genai import types
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

st.set_page_config(page_title="Lector de Remitos", page_icon="📦", layout="wide")

def secreto(nombre):
    return str(st.secrets.get(nombre, "")).strip().strip('"').strip("'")

# ---------- Contraseña (protege tu cuota de IA) ----------
clave = secreto("APP_PASSWORD")
if clave and st.text_input("Contraseña", type="password") != clave:
    st.stop()

api_key = secreto("GEMINI_API_KEY")
if not api_key:
    st.error("Falta GEMINI_API_KEY en Secrets.")
    st.stop()
client = genai.Client(api_key=api_key)

COLS = ["Localidad", "Cliente", "Domicilio", "Teléfono", "Celular", "Condición",
        "Detalle de Productos"]
ANCHOS = [28, 35, 32, 12, 16, 35, 55]
CENTRADAS = {3, 4}
CON_AJUSTE = {5, 6}

MODELOS = ["gemini-3.6-flash", "gemini-3.7-flash", "gemini-3.5-flash"]
MAX_ARCHIVOS = 10          # archivos por consulta
MAX_MB = 15                # tamaño total por consulta

PROMPT = """Recibís uno o varios archivos (fotos o PDF) de remitos. Cada archivo puede tener uno o varios remitos.
Devolvé SOLO un JSON: una lista con un objeto por remito, en el mismo orden de los archivos. Cada objeto tiene estas claves:
cliente (código y nombre tal como figura, ej: "C305634 - Apellido Nombre"), domicilio, telefono, celular,
localidad (tal como figura, con el código postal entre paréntesis), condicion,
items (lista de objetos con cantidad, codigo, descripcion).
Copiá el texto exactamente como está impreso. Si un dato no figura, poné "-". No inventes nada."""

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

def nombre_archivo():
    h = datetime.now(ZoneInfo("America/Argentina/Buenos_Aires"))
    return f"ESTO SE ENTREGÓ HOY {h.day} de {MESES[h.month - 1]} de {h.year}.xlsx"

def limpiar_localidad(t):
    t = re.sub(r"^\s*\(\d+\)\s*", "", str(t))                      # quita el código postal
    t = re.sub(r"[\s,\-–]*chubut\s*$", "", t, flags=re.IGNORECASE)   # quita "Chubut"
    return t.strip() or "-"

def _c(t, patron):
    m = re.search(patron, t, re.M)
    return (m.group(1).strip() if m else "") or "-"

def leer_pdf_texto(data):
    """Remito Famago preimpreso (PDF con texto): devuelve las filas, o None si no es ese formato."""
    if pdfplumber is None:
        return None
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            paginas = [p.extract_text(x_tolerance=2) or "" for p in pdf.pages]
    except Exception:
        return None
    filas = []
    for t in paginas:
        if "Nro Interno" not in t or "Cliente:" not in t or "Detalle" not in t:
            return None
        items = []
        for linea in t.splitlines():
            m = re.match(r"^\s*(\d+)\s+-\s+(.+?)\s+(\d+)\s*$", linea)
            if m:
                items.append(f"{m.group(3)} x [{m.group(1)}] {re.sub(r'\s+', ' ', m.group(2)).strip()}")
        if not items:
            return None
        filas.append([
            limpiar_localidad(_c(t, r"LOCALIDAD:\s*(.+?)\s*$")),
            _c(t, r"Cliente:\s*(.+?)\s+C\.U\.I\.T"),
            _c(t, r"Domicilio:\s*(.+?)\s+LOCALIDAD:"),
            _c(t, r"Tel[eé]fono:[ \t]*(\d[\d \t\-]*)"),
            _c(t, r"Celular:\s*(\d+)"),
            _c(t, r"Condici[oó]n:\s*(.+?)\s*$"),
            "\n".join(items),
        ])
    return filas

def preparar(f):
    data = f.getvalue()
    if getattr(f, "name", "").lower().endswith(".pdf"):
        return types.Part.from_bytes(data=data, mime_type="application/pdf"), len(data)
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
    img.thumbnail((1800, 1800))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=85)
    jpg = buf.getvalue()
    return types.Part.from_bytes(data=jpg, mime_type="image/jpeg"), len(jpg)

def leer(partes):
    errores = []
    for modelo in MODELOS:
        try:
            r = client.models.generate_content(
                model=modelo,
                contents=partes + [PROMPT],
                config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0),
            )
            break
        except Exception as e:
            errores.append(f"{modelo}: {str(e)[:160]}")
    else:
        raise RuntimeError(" | ".join(errores))
    d = json.loads(r.text)
    if isinstance(d, dict):
        d = [d]
    claves = ["localidad", "cliente", "domicilio", "telefono", "celular", "condicion"]
    filas = []
    for x in d:
        detalle = "\n".join(
            f"{i.get('cantidad', '1')} x [{i.get('codigo', '')}] {i.get('descripcion', '')}"
            for i in x.get("items", [])
        )
        fila = [str(x.get(k) or "-") for k in claves] + [detalle or "-"]
        fila[0] = limpiar_localidad(fila[0])
        filas.append(fila)
    return filas, modelo

def crear_excel(df):
    df = df.fillna("").sort_values(["Localidad", "Cliente"], key=lambda c: c.astype(str).str.lower(),
                                   kind="stable").reset_index(drop=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Remitos Procesados"
    ws.merge_cells("A1:G1")
    ws["A1"] = "Resumen de Remitos Procesados"
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color="1F4E78")
    ws.row_dimensions[1].height = 25
    hoy = datetime.now(ZoneInfo("America/Argentina/Buenos_Aires"))
    ws["A2"] = f"Fecha de procesamiento: {hoy:%d/%m/%Y} | Total Remitos: {len(df)}"
    ws["A2"].font = Font(name="Calibri", size=10, color="595959")
    ws.row_dimensions[4].height = 28
    lado = Side(style="thin", color="D9D9D9")
    borde = Border(left=lado, right=lado, top=lado, bottom=lado)
    for i, titulo in enumerate(COLS):
        c = ws.cell(row=4, column=i + 1, value=titulo)
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[chr(65 + i)].width = ANCHOS[i]
    for n, fila in enumerate(df.itertuples(index=False)):
        r = 5 + n
        for i, v in enumerate(fila):
            c = ws.cell(row=r, column=i + 1, value=v)
            c.font = Font(name="Calibri", size=10, color="000000")
            c.alignment = Alignment(horizontal="center" if i in CENTRADAS else "left",
                                    vertical="top", wrap_text=(i in CON_AJUSTE) or None)
            c.border = borde
            if r % 2 == 0:
                c.fill = PatternFill("solid", fgColor="F2F5F9")
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ---------- Pantalla ----------
st.title("📦 Lector de Remitos a Excel")
if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame(columns=COLS)
    st.session_state.hechos = set()

tab1, tab2 = st.tabs(["📷 Sacar foto", "📁 Subir archivos"])
fotos = []
with tab1:
    c = st.camera_input("Sacá la foto del remito")
    if c:
        fotos.append(c)
with tab2:
    fotos += st.file_uploader("Elegí una o varias imágenes o PDF", type=["jpg", "jpeg", "png", "pdf"],
                              accept_multiple_files=True)

if fotos and st.button("🚀 Procesar", type="primary"):
    lotes, actual, peso, locales = [], [], 0, 0
    for f in fotos:
        data = f.getvalue()
        h = hashlib.md5(data).hexdigest()
        if h in st.session_state.hechos:
            continue
        if getattr(f, "name", "").lower().endswith(".pdf"):
            filas = leer_pdf_texto(data)          # PDF con texto: sin Gemini
            if filas:
                for fila in filas:
                    st.session_state.df.loc[len(st.session_state.df)] = fila
                st.session_state.hechos.add(h)
                locales += len(filas)
                continue
        try:
            parte, n = preparar(f)
        except Exception as e:
            st.error(f"No se pudo abrir {getattr(f, 'name', 'la foto')}: {e}")
            continue
        if actual and (len(actual) >= MAX_ARCHIVOS or peso + n > MAX_MB * 1024 * 1024):
            lotes.append(actual)
            actual, peso = [], 0
        actual.append((h, parte))
        peso += n
    if actual:
        lotes.append(actual)
    if locales:
        st.caption(f"{locales} remito(s) leído(s) directamente del PDF, sin usar Gemini.")
    for lote in lotes:
        with st.spinner(f"Leyendo {len(lote)} archivo(s) en una sola consulta..."):
            try:
                filas, modelo = leer([p for _, p in lote])
                for fila in filas:
                    st.session_state.df.loc[len(st.session_state.df)] = fila
                st.session_state.hechos.update(h for h, _ in lote)
                st.caption(f"Leído con {modelo}: {len(lote)} archivo(s), {len(filas)} remito(s).")
                if len(filas) != len(lote):
                    st.warning("La cantidad de remitos no coincide con la de archivos. Revisá la tabla.")
            except Exception as e:
                st.error(f"No se pudo leer: {e}")
    st.session_state.df = st.session_state.df.sort_values(
        ["Localidad", "Cliente"], key=lambda c: c.astype(str).str.lower(),
        kind="stable").reset_index(drop=True)

if not st.session_state.df.empty:
    st.caption("Revisá los datos: podés corregir cualquier celda antes de descargar.")
    st.session_state.df = st.data_editor(st.session_state.df, use_container_width=True,
                                         hide_index=True, num_rows="dynamic")
    c1, c2 = st.columns(2)
    c1.download_button("📥 Descargar Excel", crear_excel(st.session_state.df),
                       file_name=nombre_archivo(),
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if c2.button("🗑️ Empezar de nuevo"):
        st.session_state.df = pd.DataFrame(columns=COLS)
        st.session_state.hechos = set()
        st.rerun()
