# -*- coding: utf-8 -*-
"""
CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.
Sistema Unificado de Control Operativo, Auditoría Automatizada y Rentabilidad.
Integra Módulo de Operación (Boletos, QR, Impresión) y Módulo de Supervisión (OCR, Finanzas, Nube).
"""

import io
import json
import os
import re
import shutil
import uuid
import base64
import hmac
import urllib.parse
import html
import zipfile
import math
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps, ImageStat

# =====================================================================
# VERIFICACIÓN SEGURA DE LIBRERÍAS EXTERNAS
# =====================================================================
try:
    import qrcode
    QR_OK = True
except ImportError:
    qrcode = None
    QR_OK = False

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

try:
    import xlsxwriter
    XLSXWRITER_OK = True
except ImportError:
    XLSXWRITER_OK = False

try:
    from cryptography.fernet import Fernet, InvalidToken
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    CRYPTOGRAPHY_OK = True
except ImportError:
    CRYPTOGRAPHY_OK = False

try:
    import pytesseract
    PYTESSERACT_OK = True
except ImportError:
    PYTESSERACT_OK = False

try:
    from supabase import create_client, Client
    SUPABASE_OK = True
except ImportError:
    SUPABASE_OK = False

# =====================================================================
# CONFIGURACIÓN GENERAL Y DIRECTORIOS
# =====================================================================
st.set_page_config(
    page_title="Club Park - Sistema Unificado",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("CLUBPARK_DATA_DIR", str(BASE_DIR)))
DATA_DIR.mkdir(parents=True, exist_ok=True)

PHOTO_DIR = DATA_DIR / "fotos_vehiculos"
PHOTO_DIR.mkdir(parents=True, exist_ok=True)
FOTOS_DIR = DATA_DIR / "fotos_reportes"
FOTOS_DIR.mkdir(parents=True, exist_ok=True)

DB_FILE = DATA_DIR / "vehiculos_produccion_db.json"
REPORTES_FILE = DATA_DIR / "reportes_clubpark.json"
CONFIG_FILE = DATA_DIR / "config_operativa.json"
AUDIT_FILE = DATA_DIR / "auditoria_operativa.json"
PAYMENTS_FILE = DATA_DIR / "pagos_registrados.json"
AI_MEMORIA_FILE = DATA_DIR / "ia_memoria_aprendizaje.json"
NOTIF_LOG_FILE = DATA_DIR / "bitacora_notificaciones.json"
COSTOS_FILE = DATA_DIR / "gastos_operativos.json"
CIERRES_DIARIOS_FILE = DATA_DIR / "cierres_diarios_sucursal.json"
ARCHIVO_REGISTRO_IMPRESION = DATA_DIR / "registro_impresion.json"

# =====================================================================
# ESTILOS CSS UNIFICADOS
# =====================================================================
CSS_INSTITUCIONAL = """
<style>
:root {
  --cpv-verde-botella: #0B3B24;
  --cpv-verde-oscuro: #072617;
  --cpv-verde-medio: #145A32;
  --cpv-verde-claro: #EBF3EE;
  --cpv-dorado-oficial: #5B8F6E;
  --cpv-fondo-suave: #F4F7F5;
  --cpv-texto-principal: #1A2B20;
}
h1, h2, h3 { color: #0B3B24 !important; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; font-weight: 700; }
button[kind="primary"], .stButton > button[kind="primary"] { background-color: #0B3B24 !important; color: #FFFFFF !important; border: 1px solid #072617 !important; border-radius: 8px !important; font-weight: 600 !important; }
button[kind="secondary"], .stButton > button[kind="secondary"] { background-color: #FFFFFF !important; color: #0B3B24 !important; border: 1px solid #C4D7CC !important; border-radius: 8px !important; font-weight: 600 !important; }
.clubpark-header-card { background: linear-gradient(135deg, #072617 0%, #0B3B24 65%, #145A32 100%); border: 2px solid #5B8F6E; border-radius: 12px; padding: 20px 24px; color: #FFFFFF; margin-bottom: 20px; box-shadow: 0 4px 14px rgba(11, 59, 36, 0.25); }
.clubpark-badge { background: rgba(46, 107, 71, 0.30); border: 1px solid #5B8F6E; color: #E8F3EC; padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: 700; display: inline-block; }
.card-auditoria { background-color: #ffffff; border-radius: 10px; padding: 15px; border-left: 6px solid #3182ce; box-shadow: 0 2px 4px rgba(0,0,0,0.08); margin-bottom: 12px; }
.badge-verde { background-color: #C6F6D5; color: #22543D; padding: 4px 8px; border-radius: 6px; font-weight: bold; }
.badge-amarillo { background-color: #FEFCBF; color: #744210; padding: 4px 8px; border-radius: 6px; font-weight: bold; }
.badge-rojo { background-color: #FED7D7; color: #742A2A; padding: 4px 8px; border-radius: 6px; font-weight: bold; }
</style>
"""
st.markdown(CSS_INSTITUCIONAL, unsafe_allow_html=True)

def encabezado_institucional(titulo, badge="", subtitulo=""):
    insignia = f'<div class="clubpark-badge" style="margin-top:2px">{badge}</div>' if badge else ""
    sub = f'<div style="font-size:13px;color:#EBF3EE;margin-top:2px">{subtitulo}</div>' if subtitulo else ""
    st.markdown(f"""<div class="clubpark-header-card">{insignia}<h3 style="color:#FFFFFF !important;margin:6px 0 2px 0">{titulo}</h3>{sub}</div>""", unsafe_allow_html=True)

_CONTADOR_LLAVES = 0
def llave_unica(base):
    global _CONTADOR_LLAVES
    _CONTADOR_LLAVES += 1
    return f"{base}__{_CONTADOR_LLAVES}"

# =====================================================================
# CONSTANTES Y CATÁLOGOS UNIFICADOS
# =====================================================================
PRECIOS_SUCURSALES = {
    "PAMPAS INSURGENTES": 50, "PAMPAS COAPA": 50, "RAMMA": 70, "SUSHITTO": 40,
    "CARAJILLO POLANCO": 100, "BENINGAN´S": 40, "SANATORIO DURANGO": 100, "BARRANCA": 0,
    "VETERINARIA": 0, "ROMANCE PALACE": 65, "JUNGLA TLALPAN": 100, "JUNGLA DIVISIÓN": 100,
    "JUNGLA PICACHO": 100, "JUNGLA EMPRESA": 100, "SIPIRILY COAPA": 80, "MAROMETA SATÉLITE": 70,
    "MAROMETA MARINA": 70, "SUMO STA. MONICA": 50, "SUMO ECHEGARAY": 50, "RYOSHI": 80,
    "ENCANTO": 60, "S. AZURA": 0, "E-BNI": 50, "CHURRASQUERIA": 70,
}
SUCURSALES = sorted(PRECIOS_SUCURSALES.keys())
SUCURSALES_LISTA = ["-- Selecciona o busca la sucursal --"] + SUCURSALES

CATALOGO_OFICIAL_23 = [
    {"num": 1, "sitio": "PAMPAS INSURGENTES", "dias": "LUNES A DOMINGO", "horarios": "11:30 AM A 12:00 PM", "empleados": 5, "supervisor": "JESÚS", "tarifa": 50},
    {"num": 2, "sitio": "PAMPAS COAPA", "dias": "LUNES A DOMINGO", "horarios": "-", "empleados": 0, "supervisor": "JESÚS", "tarifa": 50},
    {"num": 3, "sitio": "RAMMA", "dias": "LUNES A DOMINGO", "horarios": "12:00 PM A 02:00 AM", "empleados": 2, "supervisor": "JESÚS", "tarifa": 70},
    {"num": 4, "sitio": "SUSHITTO", "dias": "LUNES A DOMINGO", "horarios": "01:00 PM A 08:00 PM", "empleados": 1, "supervisor": "JESÚS", "tarifa": 40},
    {"num": 5, "sitio": "CARAJILLO POLANCO", "dias": "LUNES A DOMINGO", "horarios": "01:00 PM A 01:00 AM", "empleados": 2, "supervisor": "JESÚS", "tarifa": 100},
    {"num": 6, "sitio": "BENINGAN´S", "dias": "LUNES A DOMINGO", "horarios": "12:00 PM A 09:00 PM", "empleados": 2, "supervisor": "JESÚS", "tarifa": 40},
    {"num": 7, "sitio": "SANATORIO DURANGO", "dias": "LUNES A DOMINGO", "horarios": "12:00 PM A 09:00 PM", "empleados": 4, "supervisor": "JESÚS", "tarifa": 100},
    {"num": 8, "sitio": "BARRANCA", "dias": "LUNES A VIERNES", "horarios": "08:00 AM A 08:00 PM", "empleados": 2, "supervisor": "JESÚS", "tarifa": 0},
    {"num": 9, "sitio": "VETERINARIA", "dias": "LUNES A DOMINGO", "horarios": "09:00 AM A 19:00 PM", "empleados": 1, "supervisor": "JESÚS", "tarifa": 0},
    {"num": 10, "sitio": "ROMANCE PALACE", "dias": "VIERNES A DOMINGO", "horarios": "-", "empleados": 0, "supervisor": "JESÚS", "tarifa": 65},
    {"num": 11, "sitio": "JUNGLA TLALPAN", "dias": "SÁBADO Y DOMINGO", "horarios": "-", "empleados": 0, "supervisor": "JESÚS", "tarifa": 100},
    {"num": 12, "sitio": "JUNGLA DIVISIÓN", "dias": "SÁBADO Y DOMINGO", "horarios": "-", "empleados": 0, "supervisor": "JESÚS", "tarifa": 100},
    {"num": 13, "sitio": "JUNGLA PICACHO", "dias": "SÁBADO Y DOMINGO", "horarios": "-", "empleados": 0, "supervisor": "JESÚS", "tarifa": 100},
    {"num": 14, "sitio": "JUNGLA EMPRESA", "dias": "SÁBADO Y DOMINGO", "horarios": "-", "empleados": 0, "supervisor": "JESÚS", "tarifa": 100},
    {"num": 15, "sitio": "SIPIRILY COAPA", "dias": "SÁBADO Y DOMINGO", "horarios": "-", "empleados": 0, "supervisor": "JESÚS", "tarifa": 80},
    {"num": 16, "sitio": "MAROMETA SATÉLITE", "dias": "SÁBADO Y DOMINGO", "horarios": "-", "empleados": 0, "supervisor": "JESÚS", "tarifa": 70},
    {"num": 17, "sitio": "MAROMETA MARINA", "dias": "SÁBADO Y DOMINGO", "horarios": "-", "empleados": 0, "supervisor": "JESÚS", "tarifa": 70},
    {"num": 18, "sitio": "SUMO STA. MONICA", "dias": "LUNES A DOMINGO", "horarios": "12:30 PM A 11:00 PM", "empleados": 1, "supervisor": "VÍCTOR", "tarifa": 50},
    {"num": 19, "sitio": "SUMO ECHEGARAY", "dias": "LUNES A DOMINGO", "horarios": "12:30 PM A 11:00 PM", "empleados": 1, "supervisor": "VÍCTOR", "tarifa": 50},
    {"num": 20, "sitio": "RYOSHI", "dias": "LUNES A DOMINGO", "horarios": "12:00 PM A 1:00 AM", "empleados": 4, "supervisor": "VÍCTOR", "tarifa": 80},
    {"num": 21, "sitio": "ENCANTO", "dias": "LUNES A DOMINGO", "horarios": "12:30 PM A 01 AM", "empleados": 1, "supervisor": "VÍCTOR", "tarifa": 60},
    {"num": 22, "sitio": "E-BNI", "dias": "LUNES A DOMINGO", "horarios": "-", "empleados": 0, "supervisor": "VÍCTOR", "tarifa": 50},
    {"num": 23, "sitio": "CHURRASQUERIA", "dias": "LUNES A DOMINGO", "horarios": "12:30 PM A 11:00 PM", "empleados": 1, "supervisor": "VÍCTOR", "tarifa": 70},
]
UNIDADES_JESUS = [i["sitio"] for i in CATALOGO_OFICIAL_23 if i["supervisor"] == "JESÚS"]
UNIDADES_VICTOR = [i["sitio"] for i in CATALOGO_OFICIAL_23 if i["supervisor"] == "VÍCTOR"]

TOKENS_POR_DEFECTO = {
    "CPV-GLOBAL-AUDITOR-2026": {"nombre": "Víctor (Contador)", "sucursal": "TODAS", "rol": "Contador Auditor"},
    "CPV-GLOBAL-JESUS-2026": {"nombre": "José de Jesús", "sucursal": "TODAS", "rol": "Supervisor Global"},
    "CPV-GLOBAL-VICTOR-2026": {"nombre": "Víctor Manuel", "sucursal": "TODAS", "rol": "Supervisor Global"},
}
for s in SUCURSALES:
    k = f"CPV-QR-{s.replace(' ', '')[:8]}"
    if k not in TOKENS_POR_DEFECTO:
        TOKENS_POR_DEFECTO[k] = {"nombre": f"Encargado {s}", "sucursal": s, "rol": "Supervisor de Sucursal"}

def cargar_accesos_protegidos():
    accesos = None
    fuente = "por_defecto"
    try:
        if "CLUBPARK_TOKENS_JSON" in st.secrets:
            accesos = json.loads(str(st.secrets["CLUBPARK_TOKENS_JSON"]))
            fuente = "secrets"
        elif "CLUBPARK_TOKENS" in st.secrets:
            accesos = {clave: dict(valor) for clave, valor in st.secrets["CLUBPARK_TOKENS"].items()}
            fuente = "secrets"
    except Exception: pass
    if not accesos:
        accesos = dict(TOKENS_POR_DEFECTO)
    return accesos, fuente

ACCESOS, ACCESOS_FUENTE = cargar_accesos_protegidos()
APP_URL = str(st.secrets.get("CLUBPARK_APP_URL", "") or os.environ.get("CLUBPARK_APP_URL", "")).strip()

def obtener_secreto(clave: str, valor_por_defecto: str) -> str:
    try:
        if hasattr(st, "secrets") and clave in st.secrets:
            return st.secrets[clave]
    except Exception: pass
    return valor_por_defecto

# =====================================================================
# FUNCIONES BASE DE DATOS Y ESTADO
# =====================================================================
def cargar_json_seguro(ruta_archivo, default_tipo):
    if Path(ruta_archivo).exists():
        try:
            with open(ruta_archivo, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, type(default_tipo)): return data
        except Exception: pass
    return default_tipo

def guardar_json_atomico(ruta_archivo, datos):
    try:
        tmp_file = str(ruta_archivo) + ".tmp"
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False, indent=2)
        os.replace(tmp_file, ruta_archivo)
        return True
    except Exception as e:
        st.error(f"Error guardando {ruta_archivo}: {e}")
        return False

# Cargas específicas
def cargar_bd(): return cargar_json_seguro(DB_FILE, [])
def guardar_bd(registros): return guardar_json_atomico(DB_FILE, registros)
def cargar_registros_guardados(): return cargar_json_seguro(REPORTES_FILE, [])
def guardar_registros_permanentes(registros): return guardar_json_atomico(REPORTES_FILE, registros)
def cargar_config(): return cargar_json_seguro(CONFIG_FILE, {"siguiente_folio": {}, "historial_folios": []})
def guardar_config(c): return guardar_json_atomico(CONFIG_FILE, c)
def cargar_auditoria(): return cargar_json_seguro(AUDIT_FILE, [])

def registrar_evento_auditoria(tipo_evento, usuario, id_reporte, sucursal, fecha_op, detalles, score=100, fallas=None):
    historial = cargar_auditoria()
    evento = {
        "id_evento": f"AUD_{uuid.uuid4().hex[:8]}",
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "tipo_evento": tipo_evento,
        "usuario": usuario,
        "id_reporte": id_reporte,
        "sucursal": sucursal,
        "fecha_operacion": fecha_op,
        "detalles": detalles,
        "score_auditoria": score,
        "fallas_detectadas": fallas or []
    }
    historial.insert(0, evento)
    guardar_json_atomico(AUDIT_FILE, historial)
    return evento

def guardar_evento_auditoria(tipo, detalle):
    registrar_evento_auditoria(tipo, st.session_state.get("usuario_autenticado", "Sistema"), "-", st.session_state.get("sucursal_asignada", "TODAS"), datetime.now().strftime("%Y-%m-%d"), detalle)

def cargar_memoria_ia(): return cargar_json_seguro(AI_MEMORIA_FILE, {"version": "2.0_adaptive", "total_reportes_procesados": 0, "sucursales": {}})
def guardar_memoria_ia(m): return guardar_json_atomico(AI_MEMORIA_FILE, m)
def cargar_config_operativa(): return cargar_json_seguro(CONFIG_FILE, {"tel_jesus": "525500000001", "tel_victor": "525500000002"})
def guardar_config_operativa(cfg): return guardar_json_atomico(CONFIG_FILE, cfg)

def cargar_pagos_respaldo(): return cargar_json_seguro(PAYMENTS_FILE, [])
def guardar_pago_respaldo(pago, registro):
    pagos = cargar_pagos_respaldo()
    pago_respaldo = dict(pago)
    pago_respaldo.update({"registro_id": str(registro.get("id", "")), "sucursal": registro.get("sucursal", "")})
    if pago_respaldo.get("id_pago") and any(p.get("id_pago") == pago_respaldo.get("id_pago") for p in pagos): return True
    pagos.insert(0, pago_respaldo)
    return guardar_json_atomico(PAYMENTS_FILE, pagos)

def cargar_registro_impresion(): return cargar_json_seguro(ARCHIVO_REGISTRO_IMPRESION, [])
ETAPAS_IMPRESION = {
    "CONEXION": "Fallo de conexión", "ENVIO_ESCPOS": "Fallo de envío ESC/POS",
    "IMPRESORA": "Fallo de la impresora", "IMPRESION_OK": "Impresión correcta",
    "INTENTO": "Intento de impresión", "REINTENTO": "Reintento de impresión"
}
def registrar_evento_impresion(etapa, detalle, resultado="FALLO", datos=None):
    eventos = cargar_registro_impresion()
    evento = {
        "id_impresion": uuid.uuid4().hex[:10],
        "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "usuario": st.session_state.get("usuario_autenticado", "Local"),
        "etapa": etapa, "resultado": resultado, "diagnóstico": ETAPAS_IMPRESION.get(etapa, ""),
        "detalle": str(detalle)[:300]
    }
    if datos: evento.update(datos)
    eventos.insert(0, evento)
    guardar_json_atomico(ARCHIVO_REGISTRO_IMPRESION, eventos[:300])
    return eventos[:300]

def cargar_cierres_diarios(): return cargar_json_seguro(CIERRES_DIARIOS_FILE, [])
def guardar_cierres_diarios(c): return guardar_json_atomico(CIERRES_DIARIOS_FILE, c)

# =====================================================================
# UTILIDADES OPERATIVAS (APP 1)
# =====================================================================
def codigo_sucursal(sucursal):
    s = str(sucursal or "").strip().upper()
    return "".join(p[0] for p in re.findall(r"[A-ZÁÉÍÓÚÜÑ0-9]+", s))[:4] or "CP"

def asignar_folio(sucursal, registros):
    config = cargar_config()
    configurado = int(config["siguiente_folio"].get(sucursal, 1))
    mayor = 0
    for r in registros:
        if r.get("sucursal") == sucursal:
            m = re.search(r"-(\d+)$", str(r.get("id", "")))
            if m: mayor = max(mayor, int(m.group(1)))
    numero = max(configurado, mayor + 1)
    folio = f"CP-{codigo_sucursal(sucursal)}-{numero:06d}"
    config["siguiente_folio"][sucursal] = numero + 1
    guardar_config(config)
    return folio, numero

def qr_payload(registro): return f"CLUBPARK|VEHICULO|{registro.get('id', '')}"

def generar_qr_bytes(registro):
    if not QR_OK: return None
    img = qrcode.make(qr_payload(registro))
    salida = io.BytesIO()
    img.save(salida, format="PNG")
    return salida.getvalue()

def crear_ticket(registro):
    f, s, m, p, c = registro.get("id", ""), registro.get("sucursal", ""), registro.get("marca", ""), registro.get("placas", ""), registro.get("color", "")
    t = float(registro.get("tarifa", 0))
    ent = registro.get("fecha_entrada") or registro.get("fecha", "")
    sal = registro.get("salida_fecha", "PENDIENTE")
    cust = registro.get("custodia", "SIN OBJETOS DECLARADOS")
    linea = "=" * 32
    guion = "-" * 32
    return f"""{linea}
 CLUB PARK VALET Y ESTACIONAMIENTOS
         S.A. DE C.V.
{linea}
SUCURSAL: {s}
FOLIO TICKET: {f}
TARIFA: ${t:,.2f} MXN
{guion}
MARCA: {m}
PLACA: {p}  COLOR: {c}
FECHA/HORA ENTRADA: {ent}
FECHA/HORA SALIDA: {sal}
{guion}
[X] OBSERVACIONES Y CUSTODIA:
{cust}
{guion}
       BOLETO DEL CLIENTE
{guion}
NO NOS HACEMOS RESPONSABLES POR
OBJETOS DE VALOR NO DECLARADOS.
{linea}
           [ CORTE ]
{linea}
      BOLETO SIN VALOR LLAVERO
FOLIO: {f} | MARCA: {m}
PLACA: {p} | UBICACIÓN: _______
{linea}
           [ CORTE ]
{linea}
          CONTRASEÑAS
     BOLETO SIN VALOR PARABRISAS
FOLIO: {f}
{linea}"""

def crear_partes_ticket(registro):
    ticket = registro.get("ticket_texto") or crear_ticket(registro)
    corte = "=" * 32 + "\n           [ CORTE ]\n" + "=" * 32 + "\n"
    partes = ticket.split(corte, 1)
    if len(partes) == 2: return partes[0] + "\n" + corte.rstrip() + "\n", partes[1]
    return ticket, ticket

def dinero(valor):
    try: return float(valor or 0)
    except: return 0.0

def guardar_foto(uploaded_file, folio, vista):
    img = uploaded_file.copy() if isinstance(uploaded_file, Image.Image) else Image.open(uploaded_file)
    img = ImageOps.exif_transpose(img).convert("RGB")
    img.thumbnail((1400, 1400))
    nombre = f"{folio}_{vista}.jpg"
    ruta = PHOTO_DIR / nombre
    img.save(ruta, format="JPEG", quality=78, optimize=True)
    return str(ruta)

def ruta_foto(registro, vista):
    r = (registro.get("fotos") or {}).get(vista, "")
    if not r: return None
    p = Path(r)
    if not p.is_absolute(): p = DATA_DIR / p
    return p if p.exists() else None

# Funciones de Impresión Térmica
def construir_payload_escpos(texto, qr_payload_txt=None):
    partes = bytearray(b"\x1B\x40") # ESC @
    try: partes += str(texto).encode("cp437", errors="replace")
    except: partes += str(texto).encode("utf-8", errors="replace")
    partes += b"\n\n\n"
    if qr_payload_txt:
        d = str(qr_payload_txt).encode("utf-8")
        L, H = (len(d)+3) & 0xFF, ((len(d)+3) >> 8) & 0xFF
        partes += b"\x1B\x61\x01" # Centrado
        partes += bytes((0x1D, 0x28, 0x6B, 0x04, 0x00, 0x31, 0x41, 0x32, 0x00))
        partes += bytes((0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x43, 0x06))
        partes += bytes((0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x45, 0x31))
        partes += bytes((0x1D, 0x28, 0x6B, L, H, 0x31, 0x50, 0x30)) + d
        partes += bytes((0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x51, 0x30))
        partes += b"\x1B\x61\x00" # Izquierda
    partes += b"\x0A\x0A\x0A\x0A"
    return bytes(partes)

def imprimir_bluetooth(texto, key_suffix, qr_payload=None, contexto=None):
    safe_key = re.sub(r"[^a-zA-Z0-9_-]", "_", str(key_suffix))
    binario = construir_payload_escpos(texto, qr_payload)
    b64 = base64.b64encode(binario).decode("ascii")
    if contexto:
        registrar_evento_impresion("INTENTO", "Impresion térmica", "PENDIENTE", datos={"payload_b64": b64, "folio": contexto.get("folio","")})
    st.markdown(f'<a href="rawbt:base64,{b64}" style="padding:10px;background:#0B3B24;color:white;border-radius:5px;text-decoration:none;">🖨️ Imprimir RawBT</a>', unsafe_allow_html=True)

# Lector QR Integrado
def leer_codigo_qr(imagen_pil):
    if not isinstance(imagen_pil, Image.Image):
        try: imagen_pil = Image.open(imagen_pil)
        except: return None
    try: base = ImageOps.exif_transpose(imagen_pil).convert("RGB")
    except: base = imagen_pil.convert("RGB")
    candidatos = [base, ImageOps.autocontrast(base.convert("L"))]
    def _t(res):
        for r in res or []:
            txt = getattr(r, "text", getattr(r, "data", b"").decode("utf-8", "replace")) if hasattr(r, "text") or hasattr(r, "data") else ""
            if txt.strip(): return txt.strip()
        return None
    try:
        import zxingcpp
        for c in candidatos:
            t = _t(zxingcpp.read_barcodes(c, try_harder=True, try_rotate=True))
            if t: return t
    except: pass
    try:
        from pyzbar.pyzbar import decode
        for c in candidatos:
            t = _t(decode(c))
            if t: return t
    except: pass
    return None

def lector_qr_camara_trasera(key, height=390):
    foto = st.camera_input("Capturar QR", key=key)
    if foto: return leer_codigo_qr(foto)
    return None

# =====================================================================
# UTILIDADES FINANCIERAS Y NUBE (APP 2)
# =====================================================================
def calcular_boletos_block100(f_ini, f_fin):
    i, f = int(f_ini), int(f_fin)
    if i < 0 or f < 0 or (i==0 and f==0): return 0
    return ((f % 100) - (i % 100)) % 100 + 1

def obtener_tarifa_sucursal(sucursal, tarifa_ocr=None):
    cat = float(PRECIOS_SUCURSALES.get(sucursal, 0))
    try:
        if tarifa_ocr and float(tarifa_ocr) > 0: return float(tarifa_ocr)
    except: pass
    return cat

def evaluar_rentabilidad(margen):
    if margen >= 40.0: return "🟢 Alta"
    elif margen >= 15.0: return "🟡 Moderada"
    elif margen > 0.0: return "🟠 Baja"
    else: return "🔴 Pérdida"

def evaluar_auditoria_reporte(rep, registros, memoria_ia):
    fallas, obs, score = [], [], 100
    bruto, gastos = float(rep.get("Ingreso Bruto",0)), float(rep.get("Total Gastos",0))
    neto, margen = float(rep.get("Ingreso Neto",0)), float(rep.get("Margen (%)",0))
    neto_calc = bruto - gastos
    if abs(neto - neto_calc) > 1.0:
        score -= 25; fallas.append({"regla": "Cuadre Neto", "detalle": "Neto no cuadra con Bruto - Gastos", "gravedad": "CRÍTICA"})
    if neto < 0:
        score -= 25; fallas.append({"regla": "Pérdida", "detalle": "Operación en números rojos", "gravedad": "CRÍTICA"})
    return max(0, score), "Conforme" if score>=90 else "Falla", fallas, obs

def actualizar_memoria_ia(reporte):
    mem = cargar_memoria_ia()
    suc = reporte.get("Sitio / Unidad")
    if suc in mem["sucursales"]:
        d = mem["sucursales"][suc]
        d["reportes_aprendidos"] = d.get("reportes_aprendidos", 0) + 1
        d["promedio_boletos"] = (d.get("promedio_boletos",0) + float(reporte.get("Boletos Cobrados",0))) / 2.0
    else:
        mem["sucursales"][suc] = {"reportes_aprendidos": 1, "promedio_boletos": float(reporte.get("Boletos Cobrados",0))}
    guardar_memoria_ia(mem)

def sincronizar_reporte_supabase(rep, foto_bytes=None, foto_filename=None):
    sb = create_client(os.environ.get("SUPABASE_URL",""), os.environ.get("SUPABASE_KEY","")) if SUPABASE_OK else None
    if not sb: return False, "Sin Supabase"
    try:
        url = None
        if foto_bytes:
            sb.storage.from_("reportes").upload(foto_filename, foto_bytes, {"upsert":"true"})
            url = sb.storage.from_("reportes").get_public_url(foto_filename)
        sb.table("incomes").upsert({"id": rep.get("ID"), "amount": rep.get("Ingreso Neto"), "site": rep.get("Sitio / Unidad"), "photo_path": url}).execute()
        return True, "Sincronizado"
    except Exception as e: return False, str(e)

# =====================================================================
# INICIALIZACIÓN DE SESIÓN Y AUTENTICACIÓN UNIFICADA
# =====================================================================
if "vehiculos" not in st.session_state: st.session_state.vehiculos = cargar_bd()
if "reportes" not in st.session_state: st.session_state.reportes = cargar_registros_guardados()
if "usuario_autenticado" not in st.session_state: st.session_state.usuario_autenticado = None
if "rol" not in st.session_state: st.session_state.rol = None
if "sucursal_asignada" not in st.session_state: st.session_state.sucursal_asignada = None

if st.session_state.usuario_autenticado is None:
    encabezado_institucional("SISTEMA DE CONTROL UNIFICADO", "ACCESO", "Club Park Valet y Estacionamientos")
    st.info("Ingresa por Token/QR para Operación de Valet, o con Contraseña para Supervisión.")
    t1, t2 = st.tabs(["🎫 Acceso Operativo (QR/Token)", "🔐 Acceso Supervisión (Password)"])
    
    with t1:
        token_input = st.text_input("Ingresa Token Operativo", type="password")
        if st.button("Validar Token"):
            if token_input in ACCESOS:
                info = ACCESOS[token_input]
                st.session_state.usuario_autenticado = info["nombre"]
                st.session_state.rol = info["rol"]
                st.session_state.sucursal_asignada = info["sucursal"]
                st.rerun()
            else: st.error("Token no válido")
    with t2:
        u_sel = st.selectbox("Usuario", ["Víctor Manuel", "José de Jesús", "Víctor (Contador)"])
        p_sel = st.text_input("Contraseña", type="password")
        if st.button("Iniciar Sesión"):
            p_sup = obtener_secreto("PASS_SUPERVISOR", "sup123")
            p_aud = obtener_secreto("PASS_AUDITOR", "aud123")
            if u_sel in ["Víctor Manuel", "José de Jesús"] and p_sel == p_sup:
                st.session_state.usuario_autenticado = u_sel
                st.session_state.rol = "Supervisor"
                st.rerun()
            elif u_sel == "Víctor (Contador)" and p_sel == p_aud:
                st.session_state.usuario_autenticado = u_sel
                st.session_state.rol = "Contador Auditor"
                st.rerun()
            else: st.error("Credenciales incorrectas")
    st.stop()

# =====================================================================
# NAVEGACIÓN PRINCIPAL
# =====================================================================
st.sidebar.markdown("### CLUB PARK VALET")
st.sidebar.write(f"**Usuario:** {st.session_state.usuario_autenticado}")
st.sidebar.write(f"**Rol:** {st.session_state.rol}")
if st.sidebar.button("🚪 Cerrar sesión", use_container_width=True):
    for k in ["usuario_autenticado", "rol", "sucursal_asignada"]: st.session_state[k] = None
    st.rerun()

es_auditor = st.session_state.rol == "Contador Auditor"
es_supervisor = "Supervisor" in str(st.session_state.rol)

modulos = ["🚗 Módulo Valet (Boletos en Sitio)"]
if es_auditor or es_supervisor:
    modulos.append("📊 Módulo Finanzas (Reportes de Turno)")

modulo_activo = st.sidebar.radio("Selecciona Módulo:", modulos)

# =====================================================================
# MÓDULO 1: OPERACIÓN VALET (APP 1)
# =====================================================================
if "Valet" in modulo_activo:
    encabezado_institucional("OPERACIÓN EN SITIO", "VALET", "Ingreso, Cobro e Impresión")
    tb_in, tb_out, tb_list, tb_imp = st.tabs(["📥 Ingreso", "🚘 Salida/Cobro", "📋 Vehículos", "🖨️ Re-Impresión"])
    
    with tb_in:
        suc_in = st.selectbox("Sucursal", SUCURSALES) if es_auditor or es_supervisor else st.session_state.sucursal_asignada
        m_in = st.text_input("Marca")
        p_in = st.text_input("Placas")
        t_in = st.number_input("Tarifa", value=float(PRECIOS_SUCURSALES.get(suc_in, 50)))
        if st.button("Registrar Ingreso y Generar Ticket", type="primary"):
            fol, num = asignar_folio(suc_in, st.session_state.vehiculos)
            reg = {
                "id": fol, "sucursal": suc_in, "marca": m_in, "placas": p_in, "tarifa": t_in,
                "fecha_entrada": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
                "estado": "En Resguardo", "operador": st.session_state.usuario_autenticado
            }
            reg["ticket_texto"] = crear_ticket(reg)
            st.session_state.vehiculos.insert(0, reg)
            guardar_bd(st.session_state.vehiculos)
            st.success(f"Ingresado: {fol}")
            imprimir_bluetooth(reg["ticket_texto"], "imp_in", qr_payload(reg))
            
    with tb_out:
        fol_out = st.text_input("Buscar Folio")
        if fol_out:
            v = next((x for x in st.session_state.vehiculos if x["id"]==fol_out), None)
            if v:
                st.write(f"Vehículo: {v['marca']} ({v['placas']})")
                m_pago = st.radio("Método", ["Efectivo", "Tarjeta", "QR"])
                if st.button("Registrar Salida y Cobro"):
                    v["estado"] = "Entregado"
                    v["salida_fecha"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    v["pago_estado"] = "Pagado"
                    v["metodo_pago"] = m_pago
                    guardar_bd(st.session_state.vehiculos)
                    st.success("Salida Registrada")
            else: st.error("Folio no encontrado")
            
    with tb_list:
        st.dataframe(pd.DataFrame(st.session_state.vehiculos))
        
    with tb_imp:
        if st.session_state.vehiculos:
            f_imp = st.selectbox("Folio", [v["id"] for v in st.session_state.vehiculos])
            v_i = next(x for x in st.session_state.vehiculos if x["id"]==f_imp)
            st.code(v_i.get("ticket_texto", ""))
            imprimir_bluetooth(v_i.get("ticket_texto", ""), "imp_re")

# =====================================================================
# MÓDULO 2: FINANZAS Y REPORTES (APP 2)
# =====================================================================
else:
    encabezado_institucional("SUPERVISIÓN Y FINANZAS", "REPORTES", "OCR, Auditoría y Rentabilidad")
    tr_sub, tr_rent, tr_aud = st.tabs(["📷 Subir Reporte", "📈 Rentabilidad", "🕵️ Auditoría"])
    
    with tr_sub:
        foto_up = st.file_uploader("Foto del Reporte Físico", type=["jpg", "png"])
        if foto_up:
            st.image(foto_up, width=300)
            suc_r = st.selectbox("Sucursal Confirmada", SUCURSALES_LISTA)
            bol_r = st.number_input("Boletos Cobrados", min_value=0, value=50)
            tar_r = st.number_input("Tarifa", value=float(PRECIOS_SUCURSALES.get(suc_r, 50) if suc_r in PRECIOS_SUCURSALES else 50))
            gas_r = st.number_input("Gastos (Nómina/Especiales)", value=0.0)
            
            bruto = bol_r * tar_r
            neto = bruto - gas_r
            st.metric("Neto a Entregar", f"${neto:,.2f}")
            
            if st.button("Guardar Reporte", type="primary"):
                rid = f"REP_{uuid.uuid4().hex[:6]}"
                r = {
                    "ID": rid, "Sitio / Unidad": suc_r, "Fecha": date.today().strftime("%d/%m/%Y"),
                    "Boletos Cobrados": bol_r, "Costo Boleto": tar_r, "Ingreso Bruto": bruto,
                    "Total Gastos": gas_r, "Ingreso Neto": neto, "Margen (%)": round((neto/bruto*100) if bruto>0 else 0, 2),
                    "Elaboró": st.session_state.usuario_autenticado
                }
                st.session_state.reportes.append(r)
                guardar_registros_permanentes(st.session_state.reportes)
                actualizar_memoria_ia(r)
                st.success("Reporte guardado")
                
    with tr_rent:
        if st.session_state.reportes:
            df = pd.DataFrame(st.session_state.reportes)
            if not df.empty:
                res = df.groupby("Sitio / Unidad")[["Ingreso Bruto", "Total Gastos", "Ingreso Neto"]].sum().reset_index()
                st.dataframe(res)
                st.bar_chart(res.set_index("Sitio / Unidad")["Ingreso Neto"])
                
                # Botón de Descarga Excel (si aplica)
                if OPENPYXL_OK or XLSXWRITER_OK:
                    out = io.BytesIO()
                    with pd.ExcelWriter(out, engine="openpyxl" if OPENPYXL_OK else "xlsxwriter") as writer:
                        df.to_excel(writer, index=False)
                    st.download_button("📥 Descargar Excel Oficial", data=out.getvalue(), file_name="Reporte_Oficial.xlsx")
        else: st.info("No hay reportes de turno guardados.")
        
    with tr_aud:
        st.dataframe(pd.DataFrame(cargar_auditoria()))

# Fin del sistema
