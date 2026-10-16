# -*- coding: utf-8 -*-
"""
CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.
Sistema de control operativo con Streamlit.

IMPORTANTE:
- Para conservar archivos en Streamlit Cloud después de reinicios/despliegues,
  configura CLUBPARK_DATA_DIR con una ruta de almacenamiento persistente o usa
  un servicio externo (S3, Azure Blob, Supabase Storage, etc.).
- Cambia los tokens de ejemplo antes de usar la aplicación en producción.
- La impresión Bluetooth Web Bluetooth requiere navegador/dispositivo compatible,
  contexto HTTPS y una impresora que exponga un servicio GATT compatible.
"""

import io
import json
import os
import re
import uuid
import zipfile
import urllib.parse
import html
import base64
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from PIL import Image, ImageOps

try:
    import qrcode
    QR_OK = True
except ImportError:
    qrcode = None
    QR_OK = False

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

try:
    import xlsxwriter  # noqa: F401
    XLSXWRITER_OK = True
except ImportError:
    XLSXWRITER_OK = False


# ---------------------------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------------------------
st.set_page_config(
    page_title="Club Park - Control Operativo",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------
# ESTILO CONSTITUCIONAL E INSTITUCIONAL - VERDE BOTELLA (#0B3B24)
# CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.
# ---------------------------------------------------------------------
CSS_INSTITUCIONAL = """
<style>
/* Paleta Institucional Club Park: Verde Botella Oficial */
:root {
  --cpv-verde-botella: #0B3B24;
  --cpv-verde-oscuro: #072617;
  --cpv-verde-medio: #145A32;
  --cpv-verde-claro: #EBF3EE;
  --cpv-dorado-oficial: #C5A059;
  --cpv-dorado-brillante: #D4AF37;
  --cpv-fondo-suave: #F4F7F5;
  --cpv-texto-principal: #1A2B20;
}

/* Encabezados y títulos */
h1, h2, h3 {
  color: #0B3B24 !important;
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
  font-weight: 700;
}

/* Botones Primarios */
button[kind="primary"], .stButton > button[kind="primary"], div[data-testid="stFormSubmitButton"] > button {
  background-color: #0B3B24 !important;
  color: #FFFFFF !important;
  border: 1px solid #072617 !important;
  border-radius: 8px !important;
  font-weight: 600 !important;
  box-shadow: 0 2px 5px rgba(11, 59, 36, 0.25) !important;
  transition: all 0.2s ease !important;
}
button[kind="primary"]:hover, .stButton > button[kind="primary"]:hover {
  background-color: #145A32 !important;
  border-color: #0B3B24 !important;
  color: #FFFFFF !important;
}

/* Botones Secundarios */
button[kind="secondary"], .stButton > button[kind="secondary"] {
  background-color: #FFFFFF !important;
  color: #0B3B24 !important;
  border: 1px solid #C4D7CC !important;
  border-radius: 8px !important;
  font-weight: 600 !important;
}
button[kind="secondary"]:hover, .stButton > button[kind="secondary"]:hover {
  background-color: #EBF3EE !important;
  border-color: #0B3B24 !important;
  color: #0B3B24 !important;
}

/* Pestañas (Tabs) en Verde Botella */
button[data-baseweb="tab"] {
  color: #1A2B20 !important;
  font-weight: 600 !important;
  font-size: 14px !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
  color: #0B3B24 !important;
  border-bottom-color: #0B3B24 !important;
  font-weight: 700 !important;
}

/* Alertas y métricas */
div[data-testid="stMetricValue"] {
  color: #0B3B24 !important;
}
div[data-testid="stAlert"] {
  border-radius: 8px !important;
}

/* Banners institucionales */
.clubpark-header-card {
  background: linear-gradient(135deg, #072617 0%, #0B3B24 65%, #145A32 100%);
  border: 2px solid #C5A059;
  border-radius: 12px;
  padding: 20px 24px;
  color: #FFFFFF;
  margin-bottom: 20px;
  box-shadow: 0 4px 14px rgba(11, 59, 36, 0.25);
}
.clubpark-gold-badge {
  background: rgba(197, 160, 89, 0.18);
  border: 1px solid #C5A059;
  color: #F7E7CE;
  padding: 4px 10px;
  border-radius: 6px;
  font-size: 12px;
  font-weight: 700;
  display: inline-block;
  letter-spacing: 0.5px;
}

/* Presentación formal y elegante - Tipografía institucional */
h1, h2, h3 {
  font-family: Georgia, "Palatino Linotype", "Times New Roman", serif !important;
  letter-spacing: 0.4px;
}
h1 { font-size: 1.65rem !important; }
h2 { font-size: 1.35rem !important; }
h3 { font-size: 1.12rem !important; }
/* Barra lateral institucional en verde botella */
section[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #072617 0%, #0B3B24 70%, #145A32 100%);
  border-right: 2px solid #C5A059;
}
section[data-testid="stSidebar"] * { color: #F4F7F5 !important; }
section[data-testid="stSidebar"] button[kind="primary"],
section[data-testid="stSidebar"] button[kind="secondary"],
section[data-testid="stSidebar"] .stButton > button {
  background-color: #F4F7F5 !important; color: #0B3B24 !important;
  border: 1px solid #C5A059 !important; font-weight: 700 !important;
}
section[data-testid="stSidebar"] .stButton > button:hover {
  background-color: #EBF3EE !important;
}
/* Tarjetas de métricas formales */
div[data-testid="stMetric"] {
  background: #FFFFFF;
  border: 1px solid #C4D7CC;
  border-left: 3px solid #C5A059;
  border-radius: 10px;
  padding: 10px 14px;
  box-shadow: 0 1px 4px rgba(11, 59, 36, 0.10);
}
/* Expanders y contenedores */
div[data-testid="stExpander"] {
  border: 1px solid #C4D7CC !important;
  border-radius: 10px !important;
  background: #FFFFFF !important;
}
/* Pestaña activa con subrayado dorado */
button[data-baseweb="tab"][aria-selected="true"] {
  box-shadow: inset 0 -3px 0 #C5A059;
}
/* Entradas de texto con borde institucional */
input, textarea, div[data-baseweb="select"] > div {
  border: 1px solid #C4D7CC !important;
  border-radius: 8px !important;
}
input:focus, textarea:focus {
  border-color: #0B3B24 !important;
  box-shadow: 0 0 0 1px #0B3B24 !important;
}
/* Regla dorada institucional */
.clubpark-gold-rule { border: 0; border-top: 2px solid #C5A059; margin: 12px 0; }
</style>
"""
st.markdown(CSS_INSTITUCIONAL, unsafe_allow_html=True)

# ---------------------------------------------------------------------
# LOGOTIPO INSTITUCIONAL DE CLUB PARK (embebido, fondo transparente)
# ---------------------------------------------------------------------
LOGO_CLUBPARK_B64 = "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"

def logo_clubpark_html(ancho=170, margen="0 auto 10px auto", sombra=True, fondo=None):
    sombra_css = "filter:drop-shadow(0 2px 5px rgba(0,0,0,0.35));" if sombra else ""
    if fondo:
        envoltura = (
            'background:' + fondo + ';border:1px solid #C5A059;border-radius:12px;'
            'padding:10px 18px;display:inline-block;'
        )
    else:
        envoltura = ""
    return (
        '<div style="text-align:center;margin:' + margen + '">'
        '<span style="' + envoltura + '">'
        '<img alt="Logotipo oficial CLUB PARK VALET Y ESTACIONAMIENTOS" '
        'src="data:image/png;base64,' + LOGO_CLUBPARK_B64 + '" '
        'style="width:' + str(ancho) + 'px;max-width:92%;height:auto;display:inline-block;' + sombra_css + '">'
        '</span>'
        '</div>'
    )

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("CLUBPARK_DATA_DIR", str(BASE_DIR)))
DATA_DIR.mkdir(parents=True, exist_ok=True)
PHOTO_DIR = DATA_DIR / "fotos_vehiculos"
PHOTO_DIR.mkdir(parents=True, exist_ok=True)
def obtener_dir_componente():
    candidatos = [
        BASE_DIR / "clubpark_component",
        BASE_DIR,
        DATA_DIR / "clubpark_component",
        DATA_DIR,
    ]
    for c in candidatos:
        if (c / "index.html").exists():
            return c
    # Si no existe, crear la carpeta clubpark_component
    destino = BASE_DIR / "clubpark_component"
    destino.mkdir(parents=True, exist_ok=True)
    if (BASE_DIR / "index.html").exists():
        import shutil
        shutil.copy(BASE_DIR / "index.html", destino / "index.html")
    return destino

DIR_COMPONENTE = obtener_dir_componente()
try:
    CAMARA_COMPONENT = st.components.v1.declare_component(
        "clubpark_camara_trasera", path=str(DIR_COMPONENTE)
    )
except Exception:
    CAMARA_COMPONENT = None
DB_FILE = DATA_DIR / "vehiculos_produccion_db.json"
CONFIG_FILE = DATA_DIR / "config_operativa.json"
AUDIT_FILE = DATA_DIR / "auditoria_operativa.json"
PAYMENTS_FILE = DATA_DIR / "pagos_registrados.json"

PRECIOS_SUCURSALES = {
    "PAMPAS INSURGENTES": 50,
    "PAMPAS COAPA": 50,
    "RAMMA": 70,
    "SUSHITTO": 40,
    "CARAJILLO POLANCO": 100,
    "BENINGAN´S": 40,
    "SANATORIO DURANGO": 100,
    "BARRANCA": 0,
    "VETERINARIA": 0,
    "ROMANCE PALACE": 65,
    "JUNGLA TLALPAN": 100,
    "JUNGLA DIVISIÓN": 100,
    "JUNGLA PICACHO": 100,
    "JUNGLA EMPRESA": 100,
    "SIPIRILY COAPA": 80,
    "MAROMETA SATÉLITE": 70,
    "MAROMETA MARINA": 70,
    "SUMO STA. MONICA": 50,
    "SUMO ECHEGARAY": 50,
    "RYOSHI": 80,
    "ENCANTO": 60,
    "E-BNI": 50,
    "CHURRASQUERIA": 70,
}
SUCURSALES = sorted(PRECIOS_SUCURSALES)

# Tokens de ejemplo: reemplázalos por secretos/variables de entorno en producción.
# Si defines CLUBPARK_TOKENS_JSON, debe contener un objeto JSON con el mismo formato.
TOKENS_POR_DEFECTO = {
    "CPV-GLOBAL-AUDITOR-2026": {
        "nombre": "Víctor (Contador)",
        "sucursal": "TODAS",
        "rol": "Contador Auditor (Máxima Jerarquía)",
    },
    "CPV-GLOBAL-JESUS-2026": {
        "nombre": "José de Jesús",
        "sucursal": "TODAS",
        "rol": "Supervisor Global",
    },
    "CPV-GLOBAL-VICTOR-2026": {
        "nombre": "Víctor Manuel",
        "sucursal": "TODAS",
        "rol": "Supervisor Global",
    },
    "CPV-QR-PAMPAS-INS": {"nombre": "Encargado Pampas Ins.", "sucursal": "PAMPAS INSURGENTES", "rol": "Supervisor de Sucursal"},
    "CPV-QR-PAMPAS-COAPA": {"nombre": "Encargado Pampas Coapa", "sucursal": "PAMPAS COAPA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-RAMMA": {"nombre": "Encargado Ramma", "sucursal": "RAMMA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-SUSHITTO": {"nombre": "Encargado Sushitto", "sucursal": "SUSHITTO", "rol": "Supervisor de Sucursal"},
    "CPV-QR-CARAJILLO": {"nombre": "Encargado Carajillo", "sucursal": "CARAJILLO POLANCO", "rol": "Supervisor de Sucursal"},
    "CPV-QR-BENINGANS": {"nombre": "Encargado Beningan's", "sucursal": "BENINGAN´S", "rol": "Supervisor de Sucursal"},
    "CPV-QR-DURANGO": {"nombre": "Encargado Sanatorio Durango", "sucursal": "SANATORIO DURANGO", "rol": "Supervisor de Sucursal"},
    "CPV-QR-BARRANCA": {"nombre": "Encargado Barranca", "sucursal": "BARRANCA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-VETERINARIA": {"nombre": "Encargado Veterinaria", "sucursal": "VETERINARIA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-ROMANCE": {"nombre": "Encargado Romance", "sucursal": "ROMANCE PALACE", "rol": "Supervisor de Sucursal"},
    "CPV-QR-TLALPAN": {"nombre": "Encargado Jungla Tlalpan", "sucursal": "JUNGLA TLALPAN", "rol": "Supervisor de Sucursal"},
    "CPV-QR-DIVISION": {"nombre": "Encargado Jungla División", "sucursal": "JUNGLA DIVISIÓN", "rol": "Supervisor de Sucursal"},
    "CPV-QR-PICACHO": {"nombre": "Encargado Jungla Picacho", "sucursal": "JUNGLA PICACHO", "rol": "Supervisor de Sucursal"},
    "CPV-QR-EMPRESA": {"nombre": "Encargado Jungla Empresa", "sucursal": "JUNGLA EMPRESA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-SIPIRILY": {"nombre": "Encargado Sipirily", "sucursal": "SIPIRILY COAPA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-SATELITE": {"nombre": "Encargado Marometa Satélite", "sucursal": "MAROMETA SATÉLITE", "rol": "Supervisor de Sucursal"},
    "CPV-QR-MARINA": {"nombre": "Encargado Marometa Marina", "sucursal": "MAROMETA MARINA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-SUMOSTA": {"nombre": "Encargado Sumo Santa Mónica", "sucursal": "SUMO STA. MONICA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-CHEGARAY": {"nombre": "Encargado Sumo Echegaray", "sucursal": "SUMO ECHEGARAY", "rol": "Supervisor de Sucursal"},
    "CPV-QR-RYOSHI": {"nombre": "Encargado Ryoshi", "sucursal": "RYOSHI", "rol": "Supervisor de Sucursal"},
    "CPV-QR-ENCANTO": {"nombre": "Encargado Encanto", "sucursal": "ENCANTO", "rol": "Supervisor de Sucursal"},
    "CPV-QR-EBNI": {"nombre": "Encargado E-BNI", "sucursal": "E-BNI", "rol": "Supervisor de Sucursal"},
    "CPV-QR-CHURRAS": {"nombre": "Encargado Churrasquería", "sucursal": "CHURRASQUERIA", "rol": "Supervisor de Sucursal"},
}
# ---------------------------------------------------------------------
# CARGA PROTEGIDA DE CREDENCIALES QR
# Prioridad: 1) st.secrets  2) variable de entorno CLUBPARK_TOKENS_JSON
# 3) tokens por defecto del código (solo contingencia; la app avisa en
# pantalla para moverlos a st.secrets lo antes posible).
# ---------------------------------------------------------------------
def cargar_accesos_protegidos():
    fuente = "por_defecto"
    accesos = None
    try:
        if "CLUBPARK_TOKENS_JSON" in st.secrets:
            accesos = json.loads(str(st.secrets["CLUBPARK_TOKENS_JSON"]))
            fuente = "secrets"
        elif "CLUBPARK_TOKENS" in st.secrets:
            accesos = {clave: dict(valor) for clave, valor in st.secrets["CLUBPARK_TOKENS"].items()}
            fuente = "secrets"
    except Exception:
        accesos = None
    if not accesos:
        tokens_env = os.environ.get("CLUBPARK_TOKENS_JSON", "").strip()
        if tokens_env:
            try:
                accesos = json.loads(tokens_env)
                fuente = "entorno"
            except json.JSONDecodeError:
                st.warning("CLUBPARK_TOKENS_JSON no contiene JSON válido; se ignoró.")
                accesos = None
    if not (isinstance(accesos, dict) and accesos):
        accesos = dict(TOKENS_POR_DEFECTO)
        fuente = "por_defecto"
    return accesos, fuente


ACCESOS, ACCESOS_FUENTE = cargar_accesos_protegidos()

try:
    APP_URL = str(st.secrets.get("CLUBPARK_APP_URL", "") or os.environ.get("CLUBPARK_APP_URL", "")).strip()
except Exception:
    APP_URL = str(os.environ.get("CLUBPARK_APP_URL", "")).strip()


def contenido_qr_para(token):
    """Contenido del QR de la credencial: si se configura CLUBPARK_APP_URL el QR
    abre la app con la sesión lista al escanearlo con la cámara del teléfono."""
    if APP_URL:
        separador = "&" if "?" in APP_URL else "?"
        return f"{APP_URL}{separador}qrscan={token}"
    return token

# Agrega aquí los tokens QR de encargados de sucursal o define todos en
# CLUBPARK_TOKENS_JSON. Formato:
# {"TOKEN":{"nombre":"Encargado","sucursal":"RAMMA","rol":"Supervisor de Sucursal"}}


# ---------------------------------------------------------------------
# DATOS Y UTILIDADES
# ---------------------------------------------------------------------
def cargar_bd():
    if not DB_FILE.exists():
        return []
    try:
        with DB_FILE.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except (OSError, json.JSONDecodeError):
        st.warning("No se pudo leer la base de datos JSON. Se iniciará una lista vacía.")
        return []


def guardar_bd(registros):
    temp_file = DB_FILE.with_suffix(".json.tmp")
    try:
        with temp_file.open("w", encoding="utf-8") as f:
            json.dump(registros, f, ensure_ascii=False, indent=2)
        os.replace(temp_file, DB_FILE)
        return True
    except OSError as exc:
        st.error(f"No se pudo guardar la base de datos: {exc}")
        return False


def cargar_pagos_respaldo():
    """Carga el libro independiente de cobros, además del pago dentro del vehículo."""
    if not PAYMENTS_FILE.exists():
        return []
    try:
        with PAYMENTS_FILE.open("r", encoding="utf-8") as f:
            datos = json.load(f)
        return datos if isinstance(datos, list) else []
    except (OSError, json.JSONDecodeError):
        return []


def guardar_pago_respaldo(pago, registro):
    """Guarda cada cobro en un libro separado para facilitar respaldo y auditoría."""
    pagos = cargar_pagos_respaldo()
    pago_respaldo = dict(pago)
    pago_respaldo["registro_id"] = str(registro.get("id", ""))
    pago_respaldo["sucursal"] = registro.get("sucursal", "")
    pago_respaldo["marca"] = registro.get("marca", "")
    pago_respaldo["placas"] = registro.get("placas", "")
    # Evita duplicar el mismo pago si Streamlit repite el rerun.
    pago_id = pago_respaldo.get("id_pago")
    if pago_id and any(p.get("id_pago") == pago_id for p in pagos):
        return True
    pagos.insert(0, pago_respaldo)
    temporal = PAYMENTS_FILE.with_suffix(".tmp")
    try:
        with temporal.open("w", encoding="utf-8") as f:
            json.dump(pagos, f, ensure_ascii=False, indent=2)
        os.replace(temporal, PAYMENTS_FILE)
        return True
    except OSError as exc:
        st.error(f"No se pudo respaldar el cobro en el libro de pagos: {exc}")
        return False



def cargar_config():
    """Configuración persistente: siguiente folio por sucursal y secuencias."""
    config = {"siguiente_folio": {}, "historial_folios": []}
    if CONFIG_FILE.exists():
        try:
            with CONFIG_FILE.open("r", encoding="utf-8") as f:
                guardada = json.load(f)
            if isinstance(guardada, dict):
                config.update(guardada)
        except (OSError, json.JSONDecodeError):
            pass
    config.setdefault("siguiente_folio", {})
    config.setdefault("historial_folios", [])
    for suc in SUCURSALES:
        config["siguiente_folio"].setdefault(suc, 1)
    return config


def guardar_config(config):
    temporal = CONFIG_FILE.with_suffix(".tmp")
    with temporal.open("w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)
    os.replace(temporal, CONFIG_FILE)


def cargar_auditoria():
    if not AUDIT_FILE.exists():
        return []
    try:
        with AUDIT_FILE.open("r", encoding="utf-8") as f:
            datos = json.load(f)
        return datos if isinstance(datos, list) else []
    except (OSError, json.JSONDecodeError):
        return []


def guardar_evento_auditoria(tipo, detalle):
    eventos = cargar_auditoria()
    eventos.insert(0, {
        "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "usuario": st.session_state.get("usuario_autenticado", "Sistema"),
        "rol": st.session_state.get("rol", ""),
        "tipo": tipo,
        "detalle": detalle,
    })
    temporal = AUDIT_FILE.with_suffix(".tmp")
    with temporal.open("w", encoding="utf-8") as f:
        json.dump(eventos, f, ensure_ascii=False, indent=2)
    os.replace(temporal, AUDIT_FILE)


CODIGOS_SUCURSAL = {
    "PAMPAS INSURGENTES": "PINS",
    "PAMPAS COAPA": "PCOA",
    "RAMMA": "RAM",
    "SUSHITTO": "SUS",
    "CARAJILLO POLANCO": "CARA",
    "BENINGAN´S": "BEN",
    "SANATORIO DURANGO": "SDUR",
    "BARRANCA": "BARR",
    "VETERINARIA": "VET",
    "ROMANCE PALACE": "ROM",
    "JUNGLA TLALPAN": "JTL",
    "JUNGLA DIVISIÓN": "JDV",
    "JUNGLA PICACHO": "JPI",
    "JUNGLA EMPRESA": "JEM",
    "SIPIRILY COAPA": "SIP",
    "MAROMETA SATÉLITE": "MSAT",
    "MAROMETA MARINA": "MMAR",
    "SUMO STA. MONICA": "SSTA",
    "SUMO ECHEGARAY": "SEC",
    "RYOSHI": "RYO",
    "ENCANTO": "ENC",
    "E-BNI": "EBN",
    "CHURRASQUERIA": "CHUR",
}


def codigo_sucursal(sucursal):
    """Código estable y exclusivo para que los folios no choquen entre sucursales."""
    nombre = str(sucursal or "").strip().upper()
    if nombre in CODIGOS_SUCURSAL:
        return CODIGOS_SUCURSAL[nombre]
    palabras = re.findall(r"[A-ZÁÉÍÓÚÜÑ0-9]+", nombre)
    return ("".join(p[0] for p in palabras)[:4] or "CP")


def siguiente_numero_usado(sucursal, registros):
    """Obtiene el mayor consecutivo de esa sucursal, incluidos folios de versiones anteriores."""
    mayor = 0
    for registro in registros:
        if registro.get("sucursal") != sucursal:
            continue
        folio = str(registro.get("id", ""))
        # La sucursal ya está filtrada; así conservamos secuencias anteriores aunque
        # haya cambiado el prefijo histórico del folio.
        coincidencia = re.search(r"-(\d+)$", folio)
        if coincidencia:
            try:
                mayor = max(mayor, int(coincidencia.group(1)))
            except ValueError:
                pass
        try:
            mayor = max(mayor, int(registro.get("folio_numero", 0) or 0))
        except (ValueError, TypeError):
            pass
    return mayor


def asignar_folio(sucursal, registros):
    config = cargar_config()
    configurado = int(config["siguiente_folio"].get(sucursal, 1))
    minimo_seguro = siguiente_numero_usado(sucursal, registros) + 1
    numero = max(configurado, minimo_seguro)
    folio = f"CP-{codigo_sucursal(sucursal)}-{numero:06d}"
    config["siguiente_folio"][sucursal] = numero + 1
    guardar_config(config)
    return folio, numero


def qr_payload(registro):
    return f"CLUBPARK|VEHICULO|{registro.get('id', '')}"


def generar_qr_bytes(registro):
    if not QR_OK:
        return None
    imagen = qrcode.make(qr_payload(registro))
    salida = io.BytesIO()
    imagen.save(salida, format="PNG")
    return salida.getvalue()


def extraer_folio_qr(valor):
    valor = (valor or "").strip()
    if valor.startswith("CLUBPARK|VEHICULO|"):
        return valor.split("|", 2)[-1].strip()
    return valor


def crear_partes_ticket(registro):
    """Devuelve el boleto del cliente y la segunda impresión de contraseñas."""
    ticket = registro.get("ticket_texto") or crear_ticket(registro)
    corte = "=" * 32 + "\n           [ CORTE ]\n" + "=" * 32 + "\n"
    partes = ticket.split(corte, 1)
    if len(partes) == 2:
        primera = partes[0] + "\n" + corte.rstrip() + "\n"
        segunda = partes[1]
    else:
        primera = ticket
        segunda = (
            "=" * 32 + "\n      BOLETO SIN VALOR LLAVERO\n"
            f"CLUB PARK VALET Y ESTACIONAMIENTOS\nFOLIO: {registro.get('id', '')}\n"
            f"MARCA: {registro.get('marca', '')}\n"
            f"PLACA: {registro.get('placas', '')} | COLOR: {registro.get('color', '')}\n"
            "UBICACIÓN: _____________________\n"
            + "=" * 32 + "\n          CONTRASEÑAS\n"
            "     BOLETO SIN VALOR PARABRISAS\n"
            f"CLUB PARK VALET Y ESTACIONAMIENTOS\nFOLIO: {registro.get('id', '')}\n"
            + "=" * 32 + "\n"
        )
    return primera, segunda



def vista_impresion_con_qr(registro, parte="cliente", key="print_qr"):
    """Vista imprimible del boleto con QR y soporte térmico directo (sin bloqueador de ventanas)."""
    primera, segunda = crear_partes_ticket(registro)
    texto = primera if parte == "cliente" else segunda
    qr_data = generar_qr_bytes(registro) if QR_OK else None
    qr_b64 = base64.b64encode(qr_data).decode("ascii") if qr_data else ""
    qr_img = (
        f'<img alt="QR del vehículo" style="width:130px;height:130px;margin:8px auto;display:block;image-rendering:pixelated" src="data:image/png;base64,{qr_b64}">'
        if qr_data and parte == "cliente" else ""
    )
    seguro = html.escape(texto)
    if parte == "cliente" and qr_img:
        marca_corte = "[ CORTE ]"
        if marca_corte in seguro:
            antes, despues = seguro.split(marca_corte, 1)
            seguro = antes + '<br><center>' + qr_img + '</center><center><b>ESCANEE PARA REGISTRAR LA SALIDA</b></center><br>' + marca_corte + despues
        else:
            seguro += '<br><center>' + qr_img + '</center><center><b>ESCANEE PARA REGISTRAR LA SALIDA</b></center>'

    # Base64 para RawBT (impresión Bluetooth instantánea en Android)
    texto_rawbt_b64 = base64.b64encode(texto.encode("utf-8")).decode("ascii")
    rawbt_url = f"rawbt:data:text/plain;base64,{texto_rawbt_b64}"
    safe_key = re.sub(r"[^a-zA-Z0-9_-]", "_", str(key))
    titulo_btn = 'Boleto del cliente + QR de salida' if parte == 'cliente' else 'Llavero y contraseña parabrisas'

    html_doc = f"""
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;color:#1A2B20;margin:6px 0">
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
        <!-- Botón 1: Impresión directa térmica (sin ventanas emergentes bloqueadas) -->
        <button id="btn_print_{safe_key}" type="button" style="padding:10px 16px;font-weight:700;border:1px solid #072617;border-radius:8px;background:#0B3B24;color:#FFFFFF;cursor:pointer;display:inline-flex;align-items:center;gap:6px">
          🖨️ Imprimir {titulo_btn}
        </button>
        <!-- Botón 2: Enlace directo RawBT para impresoras Bluetooth Android -->
        <a href="{rawbt_url}" target="_top" style="text-decoration:none;padding:10px 14px;font-weight:700;border:1px solid #C5A059;border-radius:8px;background:#FAF8F5;color:#0B3B24;display:inline-flex;align-items:center;gap:6px;font-size:13px">
          📱 Enviar a Mini Impresora Bluetooth (Android)
        </a>
      </div>
      <div id="status_{safe_key}" style="margin-top:6px;font-size:12px;color:#4A6B56">
        El botón verde usa el diálogo de impresión del navegador (impresoras de red o servicio Mopria). Para una impresora térmica <b>Bluetooth</b> como la MP210 (58 mm), usa los botones RawBT del panel de impresión Bluetooth de abajo.
      </div>
      <div id="ticket_html_{safe_key}" style="display:none">{seguro}</div>
    </div>
    <script>
      (function() {{
        const btn = document.getElementById("btn_print_{safe_key}");
        const status = document.getElementById("status_{safe_key}");
        btn.addEventListener("click", function() {{
          try {{
            const contenido = document.getElementById("ticket_html_{safe_key}").innerHTML;
            let iframe = document.getElementById("iframe_print_{safe_key}");
            if (!iframe) {{
              iframe = document.createElement("iframe");
              iframe.id = "iframe_print_{safe_key}";
              iframe.style.position = "fixed";
              iframe.style.right = "0";
              iframe.style.bottom = "0";
              iframe.style.width = "0";
              iframe.style.height = "0";
              iframe.style.border = "0";
              document.body.appendChild(iframe);
            }}
            const doc = iframe.contentWindow.document;
            doc.open();
            doc.write('<!DOCTYPE html><html><head><title>Ticket Club Park</title><style>@page{{size:58mm auto;margin:1.5mm}}body{{font-family:monospace;font-size:9.5pt;line-height:1.25;white-space:pre-wrap;width:54mm;margin:0;padding:2px;color:#000}}img{{display:block;margin:6px auto;max-width:130px}}center{{text-align:center}}</style></head><body>' + contenido + '</body></html>');
            doc.close();
            setTimeout(() => {{
              try {{
                iframe.contentWindow.focus();
                iframe.contentWindow.print();
                status.textContent = "Diálogo de impresión activado correctamente.";
              }} catch(err) {{
                // Fallback por si el navegador bloquea print() en iframe
                const w = window.open("", "_blank", "width=420,height=650");
                if (w) {{
                  w.document.write('<!DOCTYPE html><html><head><title>Ticket Club Park</title><style>@page{{size:58mm auto;margin:2mm}}body{{font-family:monospace;font-size:10pt;white-space:pre-wrap;width:54mm}}</style></head><body>' + contenido + '</body></html>');
                  w.document.close();
                  w.focus();
                  setTimeout(() => w.print(), 350);
                }} else {{
                  status.textContent = "El navegador bloqueó la impresión. Permite ventanas emergentes para este sitio.";
                }}
              }}
            }}, 400);
          }} catch(e) {{
            status.textContent = "Error al preparar impresión: " + (e.message || e);
          }}
        }});
      }})();
    </script>
    """
    st.components.v1.html(html_doc, height=105)


def vista_comprobante_cobro(pago, registro, key="imprimir_comprobante_cobro"):
    """Muestra un comprobante imprimible del cobro ya registrado con estilo institucional."""
    datos = {
        "Sucursal": registro.get("sucursal", ""),
        "Folio": registro.get("id", ""),
        "Vehículo": f"{registro.get('marca', '')} · {registro.get('placas', '')}",
        "Fecha y hora": pago.get("fecha_hora", ""),
        "Importe": f"${dinero(pago.get('importe')):,.2f} MXN",
        "Método": pago.get("metodo", ""),
        "Estado": pago.get("estado", ""),
        "Referencia": pago.get("referencia", "") or "No aplica",
        "Atendió": pago.get("usuario", ""),
    }
    filas = "".join(
        f"<tr><td style='padding:3px 0;font-weight:bold;width:40%'>{html.escape(str(etiqueta))}</td><td style='padding:3px 0'>{html.escape(str(valor))}</td></tr>"
        for etiqueta, valor in datos.items()
    )
    safe_key = re.sub(r"[^a-zA-Z0-9_-]", "_", str(key))
    html_doc = f"""
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;color:#1A2B20;max-width:440px">
      <button id="{safe_key}" style="padding:10px 16px;font-weight:700;border:1px solid #072617;border-radius:8px;background:#0B3B24;color:#FFFFFF;cursor:pointer">
        🧾 Imprimir comprobante de cobro oficial
      </button>
      <div id="contenido_{safe_key}" style="display:none">
        <div style="width:72mm;padding:3mm;font-family:sans-serif;color:#000">
          <center>
            <b style="font-size:12pt">CLUB PARK</b><br>
            <span style="font-size:9pt">VALET Y ESTACIONAMIENTOS S.A. DE C.V.</span><br>
            <b style="font-size:10pt">COMPROBANTE DE COBRO OFICIAL</b>
          </center>
          <hr style="border:0;border-top:1px dashed #000;margin:6px 0">
          <table style="width:100%;font-size:9.5pt;border-collapse:collapse">{filas}</table>
          <hr style="border:0;border-top:1px dashed #000;margin:6px 0">
          <center style="font-size:8.5pt">Conserve este comprobante para cualquier aclaración.<br>Gracias por su preferencia.</center>
        </div>
      </div>
      <script>
      (function() {{
        const btn = document.getElementById("{safe_key}");
        btn.addEventListener("click", function() {{
          const contenido = document.getElementById("contenido_{safe_key}").innerHTML;
          let iframe = document.getElementById("iframe_{safe_key}");
          if (!iframe) {{
            iframe = document.createElement("iframe");
            iframe.id = "iframe_{safe_key}";
            iframe.style.position = "fixed";
            iframe.style.right = "0";
            iframe.style.bottom = "0";
            iframe.style.width = "0";
            iframe.style.height = "0";
            iframe.style.border = "0";
            document.body.appendChild(iframe);
          }}
          const doc = iframe.contentWindow.document;
          doc.open();
          doc.write('<!DOCTYPE html><html><head><title>Comprobante Club Park</title><style>@page{{size:80mm auto;margin:3mm}}body{{font-family:sans-serif;margin:0;padding:2px}}</style></head><body>' + contenido + '</body></html>');
          doc.close();
          setTimeout(() => {{
            try {{ iframe.contentWindow.focus(); iframe.contentWindow.print(); }}
            catch(e) {{ window.print(); }}
          }}, 350);
        }});
      }})();
      </script>
    </div>
    """
    st.components.v1.html(html_doc, height=75)


def imprimir_partes_ticket(registro, sufijo):
    primera, segunda = crear_partes_ticket(registro)
    st.markdown("**Impresión 1: Boleto del cliente**")
    st.code(primera, language="text")
    vista_impresion_con_qr(registro, parte="cliente", key=f"{sufijo}_cliente_directo")
    imprimir_bluetooth(primera, f"{sufijo}_cliente", qr_payload=qr_payload(registro))

    st.markdown("**Impresión 2: Llavero y contraseña de parabrisas**")
    st.code(segunda, language="text")
    vista_impresion_con_qr(registro, parte="llavero", key=f"{sufijo}_contrasena_directo")
    imprimir_bluetooth(segunda, f"{sufijo}_contrasena")


def construir_payload_escpos(texto, qr_payload_txt=None):
    """Construye el binario ESC/POS completo: inicialización, texto, código QR y avance.
    El mismo binario se envía por RawBT (Bluetooth clásico SPP en Android) o por Web Bluetooth (BLE)."""
    partes = bytearray()
    partes += bytes((0x1B, 0x40))  # ESC @ : inicializar impresora
    try:
        partes += str(texto).encode("cp437", errors="replace")
    except Exception:
        partes += str(texto).encode("utf-8", errors="replace")
    partes += b"\n\n\n"
    if qr_payload_txt:
        datos_qr = str(qr_payload_txt).encode("utf-8")
        longitud = len(datos_qr) + 3
        pL, pH = longitud & 0xFF, (longitud >> 8) & 0xFF
        partes += bytes((0x1B, 0x61, 0x01))  # ESC a 1 : centrado
        partes += bytes((0x1D, 0x28, 0x6B, 0x04, 0x00, 0x31, 0x41, 0x32, 0x00))  # modelo QR
        partes += bytes((0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x43, 0x06))  # tamaño módulo
        partes += bytes((0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x45, 0x31))  # nivel corrección M
        partes += bytes((0x1D, 0x28, 0x6B, pL, pH, 0x31, 0x50, 0x30)) + datos_qr  # datos
        partes += bytes((0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x51, 0x30))  # imprimir QR
        partes += bytes((0x1B, 0x61, 0x00))  # ESC a 0 : alineación izquierda
    partes += bytes((0x0A, 0x0A, 0x0A, 0x0A))  # avance de papel
    return bytes(partes)


def imprimir_bluetooth(texto, key_suffix, qr_payload=None):
    """Impresión térmica por Bluetooth: binario ESC/POS completo (texto + QR) para
    impresoras Bluetooth clásico SPP vía RawBT, botón de compartir, y Web Bluetooth BLE."""
    texto_json = json.dumps(str(texto), ensure_ascii=False)
    qr_json = json.dumps(str(qr_payload or ""), ensure_ascii=False)
    safe_key = re.sub(r"[^a-zA-Z0-9_-]", "_", str(key_suffix))

    payload_binario = construir_payload_escpos(texto, qr_payload)
    # Formato oficial RawBT: "rawbt:base64,<datos>" (sin MIME). El binario ESC/POS
    # llega intacto a la impresora Bluetooth clasica (SPP), incluido el QR.
    payload_b64 = base64.b64encode(payload_binario).decode("ascii")
    texto_b64 = base64.b64encode(str(texto).encode("utf-8")).decode("ascii")

    html_doc = f"""
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;border:1px solid #C4D7CC;background:#F9FBF9;border-radius:10px;padding:12px;margin:8px 0">
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <a href="rawbt:base64,{payload_b64}" target="_top"
           style="flex:1;min-width:200px;text-decoration:none;background:#0B3B24;color:#FFFFFF;padding:12px;border:1px solid #072617;border-radius:8px;font-weight:700;display:inline-flex;align-items:center;justify-content:center">
          🖨️ Imprimir por Bluetooth (RawBT - Android)
        </a>
        <a href="rawbt:data:text/plain;base64,{texto_b64}" target="_top"
           style="text-decoration:none;background:#FAF8F5;color:#0B3B24;padding:12px;border:1px solid #C5A059;border-radius:8px;font-weight:700;display:inline-flex;align-items:center;justify-content:center">
          📄 RawBT solo texto
        </a>
        <button id="cp_share_{safe_key}" type="button"
           style="background:#FAF8F5;color:#0B3B24;padding:12px;border:1px solid #C5A059;border-radius:8px;font-weight:700;cursor:pointer">
          📤 Compartir / Copiar
        </button>
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:8px">
        <button id="cp_bt_{safe_key}" type="button"
           style="background:#145A32;color:#FFFFFF;padding:11px;border:1px solid #072617;border-radius:8px;font-weight:700;cursor:pointer">
          📶 Conectar e imprimir por Bluetooth BLE (Chrome)
        </button>
        <button id="cp_dl_{safe_key}" type="button"
           style="background:#FAF8F5;color:#0B3B24;padding:11px;border:1px solid #C5A059;border-radius:8px;font-weight:700;cursor:pointer">
          ⬇️ Descargar boleto .bin (abrir con RawBT)
        </button>
      </div>
      <div id="cp_status_{safe_key}" role="status" aria-live="polite" style="font-size:12px;line-height:1.45;margin-top:10px;color:#2B4A37">
        <b>Impresora térmica Bluetooth (p. ej. MP210, 58 mm):</b> 1) Empareja la impresora en <b>Ajustes › Bluetooth</b> de Android. 2) Instala la app <b>RawBT</b>: <a href="https://play.google.com/store/apps/details?id=ru.a402d.rawbtprinter" target="_top" style="color:#0B3B24;font-weight:700">📲 Instalar RawBT</a>. 3) En RawBT, Ajustes ※ <b>Default printer</b>, elige la MP210. 4) Regresa aquí y pulsa <b>🖨🸏 Imprimir por Bluetooth (RawBT)</b>: se abre RawBT y envía el boleto completo con su QR. Si el botón no abre nada, RawBT no está instalada o no tiene impresora configurada: usa <b>⬇🸏 Descargar boleto .bin</b> y ábrelo o compártelo a RawBT. Impresoras solo BLE: botón BLE en Chrome.
      </div>
    </div>
    <script>
    (() => {{
      const ticket = {texto_json};
      const qr = {qr_json};
      const btnBle = document.getElementById('cp_bt_{safe_key}');
      const btnShare = document.getElementById('cp_share_{safe_key}');
      const status = document.getElementById('cp_status_{safe_key}');
      const btnDl = document.getElementById('cp_dl_{safe_key}');
      const say = t => status.innerHTML = t;
      const enc = new TextEncoder();
      const concat = (...parts) => {{ let n=parts.reduce((a,p)=>a+p.length,0), o=new Uint8Array(n), k=0; for(const p of parts){{o.set(p,k);k+=p.length;}} return o; }};
      const cmd = (...a) => new Uint8Array(a);
      function qrCommands(data) {{
        if (!data) return new Uint8Array();
        const b=enc.encode(data), len=b.length+3, pL=len&255, pH=(len>>8)&255;
        return concat(
          cmd(0x1B,0x61,0x01),
          cmd(0x1D,0x28,0x6B,0x04,0x00,0x31,0x41,0x32,0x00),
          cmd(0x1D,0x28,0x6B,0x03,0x00,0x31,0x43,0x06),
          cmd(0x1D,0x28,0x6B,0x03,0x00,0x31,0x45,0x31),
          cmd(0x1D,0x28,0x6B,pL,pH,0x31,0x50,0x30), b,
          cmd(0x1D,0x28,0x6B,0x03,0x00,0x31,0x51,0x30),
          cmd(0x1B,0x61,0x00)
        );
      }}
      async function sendChunks(ch, data) {{
        const max = 160;
        for(let i=0;i<data.length;i+=max) {{
          const chunk=data.slice(i,Math.min(i+max,data.length));
          if(ch.writeValueWithoutResponse) await ch.writeValueWithoutResponse(chunk);
          else if(ch.writeValueWithResponse) await ch.writeValueWithResponse(chunk);
          else await ch.writeValue(chunk);
          await new Promise(r=>setTimeout(r,15));
        }}
      }}
      btnShare.addEventListener('click', async () => {{
        try {{
          if (navigator.share) {{
            await navigator.share({{ title: 'Boleto Club Park', text: ticket }});
            say('<b>Boleto enviado al menú de compartir.</b> Elige RawBT (o tu app de impresión) para mandarlo a la impresora.');
          }} else {{
            await navigator.clipboard.writeText(ticket);
            say('<b>Boleto copiado al portapapeles.</b> Pégalo en RawBT o en la app de tu impresora.');
          }}
        }} catch(e) {{ say('No se pudo compartir: ' + (e.message || e)); }}
      }});
      btnDl.addEventListener('click', () => {{
        try {{
          const bin = Uint8Array.from(atob('{payload_b64}'), c => c.charCodeAt(0));
          const blob = new Blob([bin], {{type: 'application/octet-stream'}});
          const a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          a.download = 'boleto_clubpark_escpos.bin';
          document.body.appendChild(a); a.click(); a.remove();
          say('<b>Boleto descargado.</b> Ábrelo desde tus descargas con RawBT (o compártelo a RawBT) para imprimirlo por Bluetooth.');
        }} catch(e) {{ say('No se pudo descargar el binario: ' + (e.message || e)); }}
      }});
      btnBle.addEventListener('click', async () => {{
        if (!navigator.bluetooth) {{
          say('<span style="color:#B3261E">Este navegador no soporta Web Bluetooth (requiere Chrome en Android o desktop). Usa el botón verde <b>RawBT</b>, que funciona con TODAS las impresoras Bluetooth clásicas de Android.</span>');
          return;
        }}
        btnBle.disabled=true;
        try {{
          say('Selecciona tu impresora térmica en el cuadro de diálogo…');
          const device=await navigator.bluetooth.requestDevice({{
            acceptAllDevices: true,
            optionalServices: [
              '000018f0-0000-1000-8000-00805f9b34fb',
              '0000ffe0-0000-1000-8000-00805f9b34fb',
              '0000ff00-0000-1000-8000-00805f9b34fb',
              '0000ae00-0000-1000-8000-00805f9b34fb',
              '49535343-fe7d-4ae5-8fa9-9fafd205e455',
              'e7810a71-73ae-499d-8c15-faa9aef0c3f2',
              '0000af30-0000-1000-8000-00805f9b34fb'
            ]
          }});
          const nombre = device.name || 'impresora térmica';
          say('Conectando con <b>' + nombre + '</b>…');
          if (!device.gatt) throw new Error('El dispositivo no expone conexión GATT (es Bluetooth clásico). Usa el botón verde RawBT.');
          const server=await device.gatt.connect();
          const services=await server.getPrimaryServices();
          let characteristic=null;
          for(const service of services) {{
            try {{
              const chars=await service.getCharacteristics();
              characteristic=chars.find(c=>c.properties.writeWithoutResponse||c.properties.write||c.properties.writeWithResponse);
              if(characteristic) break;
            }} catch(err){{}}
          }}
          if(!characteristic) throw new Error('No se encontró canal de escritura ESC/POS en ' + nombre + '. Usa RawBT.');
          say('Enviando boleto y QR a <b>' + nombre + '</b>…');
          const init=cmd(0x1B,0x40);
          const body=enc.encode(ticket+'\n');
          const feed=cmd(0x0A,0x0A,0x0A);
          await sendChunks(characteristic,init);
          await sendChunks(characteristic,body);
          if(qr) await sendChunks(characteristic,qrCommands(qr));
          await sendChunks(characteristic,feed);
          say('<span style="color:#1E4620"><b>✅ Impresión enviada a ' + nombre + ' con éxito.</b></span>');
          try {{ device.gatt.disconnect(); }} catch(e){{}}
        }} catch(e) {{
          say('<span style="color:#B3261E">Aviso: ' + (e.message || e) + '. Si la impresora se empareja por los Ajustes de Android, es Bluetooth clásico y DEBES usar el botón verde <b>RawBT</b>.</span>');
        }} finally {{ btnBle.disabled=false; }}
      }});
    }})();
    </script>
    """
    st.components.v1.html(html_doc, height=310)


def dinero(valor):
    try:
        return float(valor or 0)
    except (TypeError, ValueError):
        return 0.0


def normalizar_texto(valor):
    return str(valor or "").strip()


def usuario_es_contador():
    return "Contador" in str(st.session_state.get("rol", ""))


def usuario_es_global():
    return usuario_es_contador() or st.session_state.get("rol") == "Supervisor Global"


def vehiculos_visibles(registros):
    if usuario_es_global():
        return registros
    sucursal = st.session_state.get("sucursal_asignada")
    return [r for r in registros if r.get("sucursal") == sucursal]


def guardar_foto(uploaded_file, folio, vista):
    """Guarda una foto optimizada y devuelve su ruta relativa al directorio de datos."""
    if isinstance(uploaded_file, Image.Image):
        imagen = uploaded_file.copy()
    else:
        imagen = Image.open(uploaded_file)
    imagen = ImageOps.exif_transpose(imagen).convert("RGB")
    imagen.thumbnail((1400, 1400))
    nombre = f"{folio}_{vista}.jpg"
    ruta = PHOTO_DIR / nombre
    imagen.save(ruta, format="JPEG", quality=78, optimize=True)
    return str(ruta)


def ruta_foto(registro, vista):
    ruta = (registro.get("fotos") or {}).get(vista, "")
    if not ruta:
        return None
    p = Path(ruta)
    if not p.is_absolute():
        p = DATA_DIR / p
    return p if p.exists() else None


def crear_ticket(registro):
    """Crea ticket térmico con marcas temporales explícitas de entrada y salida."""
    folio = registro.get("id", "")
    sucursal = registro.get("sucursal", "")
    marca = registro.get("marca", "")
    placas = registro.get("placas", "")
    color = registro.get("color", "")
    tarifa = float(registro.get("tarifa", 0) or 0)
    entrada_dt = registro.get("fecha_entrada") or registro.get("fecha", "")
    salida_dt = registro.get("salida_fecha") or registro.get("salida_fecha_hora") or "PENDIENTE"
    custodia = registro.get("custodia", "") or "SIN OBJETOS DECLARADOS"
    linea = "=" * 32
    guion = "-" * 32
    return f"""{linea}
 CLUB PARK VALET Y ESTACIONAMIENTOS
         S.A. DE C.V.
{linea}
SUCURSAL: {sucursal}
FOLIO TICKET: {folio}
TARIFA: ${tarifa:,.2f} MXN
{guion}
MARCA: {marca}
PLACA: {placas}  COLOR: {color}
FECHA/HORA ENTRADA: {entrada_dt}
FECHA/HORA SALIDA: {salida_dt}
{guion}
[X] OBSERVACIONES Y CUSTODIA:
{custodia}
{guion}
       BOLETO DEL CLIENTE
{guion}
NO NOS HACEMOS RESPONSABLES POR
OBJETOS COMO SON: EFECTIVO, CHEQUES,
ARMAS, BOLSAS, TARJETAS, CELULARES,
APARATOS ELECTRONICOS.

SIN LA PRESENTACION DE ESTE TALON,
NO SE ENTREGARA EL VEHICULO.
LEER CLAUSULAS AL REVERSO.
{linea}
           [ CORTE ]
{linea}
      BOLETO SIN VALOR LLAVERO
CLUB PARK VALET Y ESTACIONAMIENTOS
FOLIO: {folio}
MARCA: {marca}
PLACA: {placas} | COLOR: {color}
UBICACIÓN: _____________________
{linea}
           [ CORTE ]
{linea}
          CONTRASEÑAS
     BOLETO SIN VALOR PARABRISAS
CLUB PARK VALET Y ESTACIONAMIENTOS
FOLIO: {folio}
{linea}
"""


def generar_excel_oficial(df):
    """Genera un archivo XLSX usando el motor disponible."""
    output = io.BytesIO()
    columnas = ["id", "fecha", "sucursal", "marca", "placas", "tarifa", "operador", "estado"]
    encabezados = ["FOLIO ID", "FECHA Y HORA", "SUCURSAL", "MARCA/MODELO", "PLACAS", "TARIFA ($)", "OPERADOR", "ESTADO"]
    datos = df.reindex(columns=columnas).copy()
    datos.columns = encabezados

    if OPENPYXL_OK:
        from openpyxl.styles import Alignment, Font, PatternFill

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Reporte Operativo"
        ws.merge_cells("A1:H1")
        ws["A1"] = "CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V."
        ws["A1"].font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
        ws["A1"].fill = PatternFill("solid", fgColor="1F4E78")
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

        ws.append([])
        ws.append(encabezados)
        header_fill = PatternFill("solid", fgColor="D9EAF7")
        for cell in ws[3]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")

        for row in datos.itertuples(index=False, name=None):
            ws.append(["" if pd.isna(v) else v for v in row])

        for row_num in range(4, ws.max_row + 1):
            ws.cell(row_num, 6).number_format = '"$"#,##0.00'
        for idx, width in enumerate([17, 22, 26, 25, 16, 16, 25, 18], start=1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(idx)].width = width
        ws.freeze_panes = "A4"
        ws.auto_filter.ref = f"A3:H{ws.max_row}"
        wb.save(output)
        output.seek(0)
        return output

    if XLSXWRITER_OK:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            datos.to_excel(writer, sheet_name="Reporte Operativo", index=False, startrow=2)
            workbook = writer.book
            worksheet = writer.sheets["Reporte Operativo"]
            titulo = workbook.add_format({
                "bold": True, "font_size": 14, "font_color": "white",
                "bg_color": "#1F4E78", "align": "center", "valign": "vcenter"
            })
            header = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
            money = workbook.add_format({"num_format": '$#,##0.00'})
            worksheet.merge_range("A1:H1", "CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.", titulo)
            for col, name in enumerate(encabezados):
                worksheet.write(2, col, name, header)
            worksheet.set_column("A:A", 17)
            worksheet.set_column("B:B", 22)
            worksheet.set_column("C:D", 26)
            worksheet.set_column("E:E", 16)
            worksheet.set_column("F:F", 16, money)
            worksheet.set_column("G:G", 25)
            worksheet.set_column("H:H", 18)
            worksheet.freeze_panes(3, 0)
            if len(datos):
                worksheet.autofilter(2, 0, len(datos) + 2, len(encabezados) - 1)
        output.seek(0)
        return output

    raise RuntimeError(
        "No hay motor Excel instalado. Agrega openpyxl y xlsxwriter a requirements.txt."
    )


def leer_codigo_qr(imagen_pil):
    """Lector de QR robusto: prepara la imagen (contraste, ampliacion, binarizacion,
    inversiones y rotaciones) y prueba tres decodificadores antes de rendirse."""
    if imagen_pil is None:
        return None
    if not isinstance(imagen_pil, Image.Image):
        try:
            imagen_pil = Image.open(imagen_pil)
        except Exception:
            return None
    try:
        base = ImageOps.exif_transpose(imagen_pil).convert("RGB")
    except Exception:
        base = imagen_pil.convert("RGB")

    candidatos = [base]
    try:
        gris = ImageOps.autocontrast(base.convert("L"))
        ancho, alto = gris.size
        if max(ancho, alto) < 1100:
            factor = max(2, int(1100 / max(ancho, alto)))
            candidatos.append(gris.resize((ancho * factor, alto * factor), Image.LANCZOS))
        candidatos.append(gris)
        candidatos.append(ImageOps.invert(gris))
        candidatos.append(ImageOps.autocontrast(base))
        try:
            import cv2
            import numpy as np
            arr = np.array(gris)
            _, otsu = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            candidatos.append(Image.fromarray(otsu))
            candidatos.append(Image.fromarray(cv2.adaptiveThreshold(
                arr, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 31, 2)))
        except Exception:
            pass
    except Exception:
        pass

    def _texto_de(resultados):
        for r in resultados or []:
            texto = getattr(r, "text", None)
            if texto is None:
                datos = getattr(r, "data", None)
                if isinstance(datos, bytes):
                    texto = datos.decode("utf-8", "replace")
                else:
                    texto = str(datos or "")
            if str(texto).strip():
                return str(texto).strip()
        return None

    # 1) zxing-cpp con TRY_HARDER y rotaciones
    try:
        import zxingcpp
        for imagen in candidatos:
            try:
                hallado = _texto_de(zxingcpp.read_barcodes(imagen, try_harder=True, try_rotate=True))
                if hallado:
                    return hallado
            except Exception:
                pass
    except Exception:
        pass

    # 2) pyzbar con rotaciones manuales
    try:
        from pyzbar.pyzbar import decode
        for imagen in candidatos:
            for angulo in (0, 90, 180, 270):
                prueba = imagen if angulo == 0 else imagen.rotate(angulo, expand=True)
                try:
                    hallado = _texto_de(decode(prueba))
                    if hallado:
                        return hallado
                except Exception:
                    pass
    except Exception:
        pass

    # 3) detector de OpenCV
    try:
        import cv2
        import numpy as np
        detector = cv2.QRCodeDetector()
        for imagen in candidatos:
            arr = cv2.cvtColor(np.array(imagen), cv2.COLOR_RGB2BGR)
            for angulo in (0, 90, 180, 270):
                prueba = arr if angulo == 0 else np.rot90(arr, angulo // 90)
                try:
                    valor, _, _ = detector.detectAndDecode(prueba)
                    if valor and valor.strip():
                        return valor.strip()
                except Exception:
                    pass
    except Exception:
        pass
    return None

def lector_qr_camara_trasera(key="lector_qr_trasero", height=390):
    """Componente real de cámara trasera; devuelve el texto QR a Streamlit."""
    return CAMARA_COMPONENT(mode="qr", key=key, height=height, default=None)


def camara_trasera_foto_pertenencias(key="foto_pertenencias_trasera", height=430):
    """Captura evidencia usando explícitamente la cámara posterior del dispositivo."""
    resultado = CAMARA_COMPONENT(mode="photo", key=key, height=height, default=None)
    if not isinstance(resultado, str) or not resultado.startswith("data:image/"):
        return None
    try:
        encabezado, encoded = resultado.split(",", 1)
        imagen_bytes = base64.b64decode(encoded)
        return Image.open(io.BytesIO(imagen_bytes)).copy()
    except Exception:
        st.warning("No se pudo procesar la foto. Vuelve a tomarla.")
        return None


# ---------------------------------------------------------------------
# PRUEBAS OPERATIVAS (autodiagnóstico institucional)
# ---------------------------------------------------------------------
def ejecutar_pruebas_operativas():
    """Ejecuta una batería de pruebas sobre los componentes críticos de la app."""
    resultados = []

    def registrar(nombre, funcion):
        try:
            detalle = funcion()
            resultados.append({"prueba": nombre, "estado": "✅", "detalle": detalle or "OK"})
        except Exception as exc:
            resultados.append({"prueba": nombre, "estado": "❌", "detalle": str(exc)[:220]})

    def prueba_base_datos():
        archivo = DATA_DIR / "_prueba_operativa.json"
        datos = [{"id": "CP-TEST-000001", "sucursal": "RAMMA"}]
        with archivo.open("w", encoding="utf-8") as f:
            json.dump(datos, f, ensure_ascii=False)
        with archivo.open("r", encoding="utf-8") as f:
            leido = json.load(f)
        archivo.unlink()
        assert leido == datos
        return "Escritura y lectura JSON correctas"

    def prueba_folios():
        registros = [
            {"id": "CP-RAM-000005", "sucursal": "RAMMA", "folio_numero": 5},
            {"id": "CP-SUS-000002", "sucursal": "SUSHITTO", "folio_numero": 2},
        ]
        folio1, num1 = asignar_folio("RAMMA", registros)
        assert folio1.startswith("CP-RAM-") and num1 >= 6, f"folio inesperado {folio1}"
        folio2, num2 = asignar_folio("SUSHITTO", registros)
        assert folio2.startswith("CP-SUS-") and num2 >= 3
        config = cargar_config()
        config["siguiente_folio"]["RAMMA"] = 6
        config["siguiente_folio"]["SUSHITTO"] = 3
        guardar_config(config)
        return f"Secuencias independientes por sucursal ({folio1} / {folio2})"

    def prueba_registro():
        """Prueba de registro: simula el alta completa de un vehículo (folio, ticket y foto)."""
        config_previa = cargar_config()
        folio_previo = int(config_previa["siguiente_folio"].get("RAMMA", 1))
        registros_sim = [{"id": "CP-RAM-000005", "sucursal": "RAMMA", "folio_numero": 5}]
        folio, numero = asignar_folio("RAMMA", registros_sim)
        # Restaura el contador real para no consumir folios con la prueba
        config_post = cargar_config()
        config_post["siguiente_folio"]["RAMMA"] = folio_previo
        guardar_config(config_post)
        assert re.match(r"^CP-[A-Z0-9]+-\d{6}$", folio), f"Formato de folio inválido: {folio}"
        ahora = datetime.now()
        registro = {
            "id": folio,
            "folio_numero": numero,
            "qr_payload": f"CLUBPARK|VEHICULO|{folio}",
            "fecha": ahora.strftime("%d/%m/%Y %H:%M:%S"),
            "sucursal": "RAMMA",
            "marca": "Vehículo de prueba", "color": "Verde Botella", "placas": "PRB-000-0",
            "tarifa": 70, "operador": "Prueba operativa",
            "custodia": "SIN OBJETOS DE VALOR DECLARADOS",
            "fotos": {},
        }
        registro["ticket_texto"] = crear_ticket(registro)
        assert folio in registro["ticket_texto"], "El ticket no contiene el folio asignado"
        primera, segunda = crear_partes_ticket(registro)
        assert folio in primera and folio in segunda, "El folio falta en el boleto o en el llavero"
        imagen_prueba = Image.new("RGB", (320, 240), (11, 59, 36))
        ruta_foto_prueba = guardar_foto(imagen_prueba, folio, "prueba_registro")
        assert Path(ruta_foto_prueba).exists(), "La foto del registro no se guardó"
        Path(ruta_foto_prueba).unlink()
        return f"Registro simulado válido ({folio}): folio, boleto, llavero y foto correctos"

    def prueba_qr():
        if not QR_OK:
            raise RuntimeError("Falta el paquete qrcode[pil] en el servidor: actualiza requirements.txt en el repositorio y reinicia la app (Manage app → Reboot)")
        registro = {"id": "CP-RAM-000099"}
        png = generar_qr_bytes(registro)
        assert png and png[:8] == b"\x89PNG\r\n\x1a\n", "El QR no es un PNG válido"
        texto = None
        try:
            import zxingcpp
            decodificados = zxingcpp.read_barcodes(Image.open(io.BytesIO(png)))
            texto = decodificados[0].text if decodificados else None
        except Exception:
            try:
                from pyzbar.pyzbar import decode as decode_zbar
                detectados = decode_zbar(Image.open(io.BytesIO(png)))
                texto = detectados[0].data.decode() if detectados else None
            except Exception:
                raise RuntimeError("Falta el paquete zxing-cpp en el servidor: actualiza requirements.txt en el repositorio y reinicia la app")
        assert texto == qr_payload(registro), f"QR decodificado {texto} != {qr_payload(registro)}"
        return f"QR generado y releído correctamente ({texto})"

    def prueba_partes_ticket():
        registro = {
            "id": "CP-RAM-000099", "sucursal": "RAMMA", "marca": "Aveo", "color": "Azúl",
            "placas": "ABC-123-D", "tarifa": 70, "operador": "Prueba",
            "fecha": "01/01/2026 10:00:00", "custodia": "Prueba",
        }
        registro["ticket_texto"] = crear_ticket(registro)
        primera, segunda = crear_partes_ticket(registro)
        assert registro["id"] in primera, "El folio no aparece en el boleto del cliente"
        assert "LLAVERO" in segunda.upper() or "CONTRASE" in segunda.upper(), "Falta el llavero"
        return "Boleto del cliente y llavero separados con folio"

    def prueba_escpos():
        payload = construir_payload_escpos("PRUEBA BOLETO", "CLUBPARK|VEHICULO|CP-RAM-000099")
        assert payload[:2] == b"\x1b@", "Falta inicialización ESC @"
        assert b"\x1d(\x6b" in payload, "Faltan comandos del código QR"
        assert b"PRUEBA BOLETO" in payload
        return f"Binario ESC/POS válido de {len(payload)} bytes con QR"

    def prueba_credenciales():
        assert isinstance(ACCESOS, dict) and ACCESOS, "No hay credenciales cargadas"
        for token, datos in ACCESOS.items():
            assert datos.get("nombre") and datos.get("sucursal") and datos.get("rol"), f"Credencial incompleta: {token}"
        faltantes = [suc for suc in PRECIOS_SUCURSALES if not any(d.get("sucursal") == suc for d in ACCESOS.values())]
        assert not faltantes, f"Sucursales sin credencial: {faltantes}"
        return f"{len(ACCESOS)} credenciales completas (fuente: {ACCESOS_FUENTE})"

    def prueba_auditoria():
        guardar_evento_auditoria("PRUEBA_OPERATIVA", {"mensaje": "Prueba automática de bitácora"})
        eventos = cargar_auditoria()
        assert any(e.get("tipo") == "PRUEBA_OPERATIVA" for e in eventos), "El evento no se registró"
        return f"Bitácora con {len(eventos)} eventos"

    def prueba_fotos():
        assert PHOTO_DIR.exists() and os.access(PHOTO_DIR, os.W_OK), "La carpeta de fotos no es escribible"
        return f"Carpeta de fotografías escribible ({PHOTO_DIR.name})"

    def prueba_config():
        config = cargar_config()
        assert isinstance(config.get("siguiente_folio"), dict)
        return f"Configuración con {len(config['siguiente_folio'])} sucursales"

    registrar("Base de datos (JSON) escritura y lectura", prueba_base_datos)
    registrar("Folios independientes por sucursal", prueba_folios)
    registrar("Registro de ingreso de vehículo (simulado)", prueba_registro)
    registrar("Generación y lectura de código QR", prueba_qr)
    registrar("Partes del ticket (boleto + llavero)", prueba_partes_ticket)
    registrar("Binario de impresión térmica ESC/POS", prueba_escpos)
    registrar("Integridad de credenciales QR", prueba_credenciales)
    registrar("Bitácora de auditoría", prueba_auditoria)
    registrar("Carpeta de fotografías del vehículo", prueba_fotos)
    registrar("Configuración de folios", prueba_config)
    return resultados


# ---------------------------------------------------------------------
# ESTADO DE SESIÓN
# ---------------------------------------------------------------------
if "usuario_autenticado" not in st.session_state:
    st.session_state.usuario_autenticado = None
if "sucursal_asignada" not in st.session_state:
    st.session_state.sucursal_asignada = None
if "rol" not in st.session_state:
    st.session_state.rol = None
if "vehiculos" not in st.session_state:
    st.session_state.vehiculos = cargar_bd()


# ---------------------------------------------------------------------
# LOGIN
# ---------------------------------------------------------------------
if st.session_state.usuario_autenticado is None:
    left, center, right = st.columns([1, 2.2, 1])
    with center:
        st.markdown(f"""
        <div class="clubpark-header-card" style="text-align:center;padding:18px 20px;margin-bottom:16px">
          {logo_clubpark_html(150, "0 0 10px 0")}
          <div class="clubpark-gold-badge" style="margin-bottom:6px">Acceso Operativo Oficial</div>
          <h2 style="color:#FFFFFF !important;margin:4px 0 2px 0;font-size:22px">CLUB PARK VALET</h2>
          <div style="font-size:12px;color:#D4AF37;letter-spacing:1px">Y ESTACIONAMIENTOS S.A. DE C.V.</div>
        </div>
        """, unsafe_allow_html=True)

        st.info("Identifícate mediante credencial QR o ingresa tu token operativo.")

        if CAMARA_COMPONENT is not None:
            opciones_login = [
                "📷 Cámara trasera en vivo (Recomendado): lee el QR solo",
                "📸 Foto de QR con la cámara del dispositivo",
                "🖼️ Subir imagen QR",
                "🔑 Token manual",
            ]
        else:
            opciones_login = [
                "📸 Foto de QR con la cámara del dispositivo (Recomendado)",
                "🖼️ Subir imagen QR",
                "🔑 Token manual",
            ]
        metodo = st.radio(
            "Método de autenticación:",
            opciones_login,
            horizontal=False,
            key="metodo_login_operativo"
        )

        token = str(st.query_params.get("qrscan", "") or "").strip() or None

        if metodo.startswith("📷"):
            if not token:
                st.caption("Permite el acceso a la cámara cuando el navegador lo pida. El visor usa la cámara TRASERA y detecta el QR automáticamente: acerca la credencial a 10-15 cm con buena luz.")
                token_leido = lector_qr_camara_trasera("lector_login_trasero", height=390)
                if token_leido:
                    token = str(token_leido).strip()

        elif metodo.startswith("📸"):
            if not token:
                st.caption("Pulsa el botón: se abrirá la cámara de tu teléfono. En Chrome Android usa el botón Cambiar cámara para pasar a la trasera. Enfoca la credencial QR a 10-15 cm.")
                foto_qr = st.camera_input("Capturar QR para iniciar sesión", key="camara_nativa_login")
                if foto_qr:
                    token_detectado = leer_codigo_qr(Image.open(foto_qr))
                    if token_detectado:
                        token = token_detectado.strip()
                    else:
                        st.warning("No se detectó un QR legible. Acércate, mejora la luz y vuelve a tomar la foto.")

        elif metodo.startswith("🖼️"):
            archivo_qr = st.file_uploader("Selecciona el archivo de tu credencial QR", type=["png", "jpg", "jpeg"])
            if archivo_qr:
                token = leer_codigo_qr(Image.open(archivo_qr))

        else:
            token_input = st.text_input("Ingresa tu token de acceso asignado", type="password")
            if st.button("Validar acceso", type="primary", use_container_width=True):
                token = token_input.strip()

        if token:
            info = ACCESOS.get(token)
            if info:
                st.session_state.usuario_autenticado = info.get("nombre", "Usuario")
                st.session_state.sucursal_asignada = info.get("sucursal", "")
                st.session_state.rol = info.get("rol", "Supervisor de Sucursal")
                if "qrscan" in st.query_params:
                    del st.query_params["qrscan"]
                st.success(f"Acceso concedido: **{info.get('nombre')}** ({info.get('rol')})")
                st.rerun()
            else:
                st.error(f"Token o código QR no reconocido: `{token}`. Verifica con el Contador Auditor.")

    st.stop()


# ---------------------------------------------------------------------
# ENCABEZADO Y NAVEGACIÓN
# ---------------------------------------------------------------------
es_contador = usuario_es_contador()
es_global = usuario_es_global()

sucursal_texto = "Todas (Global)" if es_global else st.session_state.sucursal_asignada

st.markdown(f"""
<div class="clubpark-header-card">
  <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px">
    <div>
      {logo_clubpark_html(140, "0 0 10px 0")}
      <div class="clubpark-gold-badge">SISTEMA INSTITUCIONAL DE CONTROL OPERATIVO</div>
      <h2 style="color:#FFFFFF !important;margin:6px 0 2px 0;font-size:24px">CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.</h2>
      <div style="font-size:13px;color:#EBF3EE">
        <b>Usuario:</b> {st.session_state.usuario_autenticado} &nbsp;|&nbsp;
        <b>Rol:</b> {st.session_state.rol} &nbsp;|&nbsp;
        <b>Sucursal:</b> <span style="color:#D4AF37;font-weight:700">{sucursal_texto}</span>
      </div>
    </div>
    <div style="text-align:right">
      <div style="font-size:11px;color:#D4AF37;text-transform:uppercase;font-weight:700">Identidad Corporativa</div>
      <div style="font-size:13px;font-weight:700;color:#FFFFFF">COLOR CONSTITUCIONAL: VERDE BOTELLA (#0B3B24)</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown(logo_clubpark_html(150, "0 0 8px 0", fondo="#FFFFFF") + '<div style="font-weight:800;color:#F4F7F5;letter-spacing:2px;font-size:15px;margin:2px 0 6px 0;text-align:center">CLUB PARK VALET</div>', unsafe_allow_html=True)
    st.subheader("Panel de sesión")
    st.write(f"**Usuario:** {st.session_state.usuario_autenticado}")
    st.write(f"**Rol:** {st.session_state.rol}")
    st.write(f"**Vehículos registrados:** {len(st.session_state.vehiculos)}")
    if st.button("🚪 Cerrar sesión", use_container_width=True):
        st.session_state.usuario_autenticado = None
        st.session_state.sucursal_asignada = None
        st.session_state.rol = None
        st.rerun()

def analisis_inteligente(registros):
    """Análisis automático de la operación: KPIs, alertas tempranas y resumen
    ejecutivo. Se recalcula con la información real cada vez que se consulta."""
    ahora = datetime.now()
    activos = [r for r in registros if r.get("estado") != "Entregado"]
    entregados = [r for r in registros if r.get("estado") == "Entregado"]

    def horas_estancia(r):
        try:
            iso = r.get("fecha_entrada_iso") or ""
            if not iso:
                return None
            return (ahora - datetime.fromisoformat(iso)).total_seconds() / 3600
        except Exception:
            return None

    estancias = [(r, h) for r, h in ((r, horas_estancia(r)) for r in activos) if h is not None]

    alertas = []
    for r, h in sorted(estancias, key=lambda par: -par[1]):
        ident = r.get("id", "")
        placas = r.get("placas", "")
        suc = r.get("sucursal", "")
        if h >= 12:
            alertas.append({"tipo": "🚨 Estancia crítica", "sucursal": suc, "detalle": f"{ident} ({placas}) en {suc}: {h:.1f} h sin salida", "accion": "Verificar con el encargado y registrar salida o incidencia hoy."})
        elif h >= 6:
            alertas.append({"tipo": "⚠️ Estancia prolongada", "sucursal": suc, "detalle": f"{ident} ({placas}) en {suc}: {h:.1f} h", "accion": "Dar seguimiento en la próxima revisión."})
    placas_activas = {}
    for r in activos:
        clave = normalizar_texto(r.get("placas")).upper()
        if clave:
            placas_activas.setdefault(clave, []).append(r)
    for clave, lista in placas_activas.items():
        if len(lista) > 1:
            alertas.append({"tipo": "⚠️ Placa repetida en resguardo", "sucursal": ", ".join(sorted({str(r.get("sucursal", "")) for r in lista})), "detalle": f"Placa {clave} en {len(lista)} vehículos activos: " + ", ".join(str(r.get('id', '')) for r in lista), "accion": "Revisar duplicados: posible error de captura."})
    for r in entregados:
        if normalizar_texto(r.get("pago_estado", "")).lower() not in ("pagado", "cortado", ""):
            alertas.append({"tipo": "⚠️ Entrega sin cobro confirmado", "sucursal": str(r.get("sucursal", "")), "detalle": f"{r.get('id', '')} ({r.get('placas', '')}) en {r.get('sucursal', '')}", "accion": "Conciliar el cobro en el corte de caja."})

    entregados_hoy = [r for r in entregados if str(r.get("fecha_salida", "")).startswith(ahora.strftime("%d/%m/%Y"))]
    ingresos_hoy = sum(dinero(r.get("tarifa")) for r in entregados_hoy)
    estancia_prom = (sum(h for _, h in estancias) / len(estancias)) if estancias else 0.0
    sucursales_activas = sorted({r.get("sucursal", "Sin sucursal") for r in activos})
    por_sucursal = max(sucursales_activas, key=lambda s: sum(1 for r in activos if r.get("sucursal") == s)) if sucursales_activas else "—"
    conteo_top = sum(1 for r in activos if r.get("sucursal") == por_sucursal) if por_sucursal != "—" else 0
    ticket_prom = (sum(dinero(r.get("tarifa")) for r in entregados_hoy) / len(entregados_hoy)) if entregados_hoy else 0.0

    kpis = [
        ("Vehículos en resguardo", f"{len(activos)}"),
        ("Entregados hoy", f"{len(entregados_hoy)}"),
        ("Ingresos de hoy", f"${ingresos_hoy:,.0f}"),
        ("Estancia promedio", f"{estancia_prom:.1f} h" if activos else "—"),
    ]

    resumen = (
        f"Hay {len(activos)} vehículos en resguardo en {len(sucursales_activas)} sucursal(es). "
        f"Hoy se han entregado {len(entregados_hoy)} vehículos con ingresos de ${ingresos_hoy:,.0f} "
        f"(ticket promedio ${ticket_prom:,.0f}). "
        + (f"La sucursal con más vehículos en sitio es {por_sucursal} ({conteo_top}). " if por_sucursal != "—" else "")
        + (f"La estancia promedio en sitio es de {estancia_prom:.1f} horas. " if activos else "")
        + f"Análisis generado automáticamente el {ahora.strftime('%d/%m/%Y %H:%M')}."
    )

    recomendaciones = []
    if len(activos) >= 20:
        recomendaciones.append("Volumen alto en sitio: refuerza personal en la zona de resguardo durante horas pico.")
    if estancia_prom >= 5 and activos:
        recomendaciones.append("La estancia promedio supera 5 horas: revisa tarifas por tiempo o promueve pagos anticipados.")
    sucursales_alertadas = sorted({a.get("sucursal", "") for a in alertas if a.get("sucursal")})
    if sucursales_alertadas:
        recomendaciones.append("Concentra la revisión en: " + ", ".join(sucursales_alertadas[:5]) + ".")
    if not entregados_hoy and activos:
        recomendaciones.append("Aún sin salidas registradas hoy: verifica que los encargados estén usando el módulo de Salida y Cobro.")
    if not recomendaciones:
        recomendaciones.append("Operación dentro de parámetros normales. Continúa con el protocolo fotográfico establecido.")

    return kpis, alertas, resumen, recomendaciones


tab_ing, tab_list, tab_salida, tab_balance, tab_imp, tab_rep, tab_corte, tab_cierre_diario, tab_aud, tab_qrs, tab_pruebas, tab_intel = st.tabs([
    "📥 Ingreso y Fotografías",
    "📋 Vehículos Registrados",
    "🚘 Salida y Cobro",
    "📈 Balances y Rentabilidad",
    "🖨️ Re-Impresión Térmica",
    "📊 Reportes y Excel",
    "💵 Cortes de Caja",
    "📒 Cierre diario por sucursal",
    "🕵️ Auditoría y WhatsApp",
    "🏷️ Credenciales QR Institucionales",
    "🧪 Pruebas Operativas",
    "🧠 Inteligencia Operativa",
])


# ---------------------------------------------------------------------
# TAB 1: INGRESO
# ---------------------------------------------------------------------
with tab_ing:
    st.subheader("1. Datos del vehículo y recepción")
    c1, c2 = st.columns(2)
    with c1:
        if es_global:
            sucursal = st.selectbox("Sucursal operativa", ["-- Selecciona la Sucursal --"] + SUCURSALES)
        else:
            sucursal = st.session_state.sucursal_asignada
            st.info(f"Sucursal asignada: **{sucursal}**")
        marca = st.text_input("Marca y modelo", key="ing_marca")
        color = st.text_input("Color", key="ing_color")
    with c2:
        placas = st.text_input("Placas / folio de placa", key="ing_placas").strip().upper()
        tarifa_default = float(PRECIOS_SUCURSALES.get(sucursal, 50))
        tarifa = tarifa_default
        st.metric("Tarifa fija de la sucursal", f"${tarifa:,.2f} MXN")
        operador = st.text_input("Valet / operador", value=st.session_state.usuario_autenticado)
    st.subheader("Observaciones y custodia")
    sin_valores = st.checkbox("⚡ Sin objetos de valor declarados", key="ing_sin_valores")
    if sin_valores:
        custodia = "SIN OBJETOS DE VALOR DECLARADOS"
        st.success("Se registrará: SIN OBJETOS DE VALOR DECLARADOS")
    else:
        custodia = st.text_area(
            "Objetos declarados y observaciones de custodia",
            placeholder="Ej. computadora portátil, mochila negra, cargador...",
            key="ing_custodia_texto",
        )
    st.caption("Marca la opción rápida o describe los objetos que el cliente deja bajo resguardo.")
    st.markdown("**📷 Fotografías de pertenencias bajo resguardo (opcional)**")
    st.caption("Puedes capturar con la cámara nativa del celular, con el visor en vivo o adjuntando fotos.")
    
    opc_cam_p = st.radio(
        "Herramienta de captura:",
        ["📸 Cámara nativa del dispositivo (Recomendado)", "📷 Visor en vivo cámara trasera", "📁 Subir fotos desde galería"],
        horizontal=True,
        key="selector_camara_pertenencias"
    )
    foto_pertenencias_camara = None
    fotos_pertenencias_archivos = []

    if opc_cam_p.startswith("📷"):
        foto_pertenencias_camara = camara_trasera_foto_pertenencias("foto_pertenencias_trasera", height=430)
        if foto_pertenencias_camara:
            st.image(foto_pertenencias_camara, caption="Fotografía de pertenencias capturada", width=260)
    elif opc_cam_p.startswith("📸"):
        foto_pertenencias_camara = camara_trasera_foto_pertenencias("foto_pertenencias_trasera", height=430)
        if foto_pertenencias_camara:
            st.image(foto_pertenencias_camara, caption="Fotografía de pertenencias capturada", width=260)
    else:
        fotos_pertenencias_archivos = st.file_uploader(
            "Selecciona fotos de pertenencias", type=["jpg", "jpeg", "png"],
            accept_multiple_files=True, key="fotos_pertenencias_archivos"
        )
        if fotos_pertenencias_archivos:
            st.caption(f"Fotos seleccionadas: {len(fotos_pertenencias_archivos)}")

    st.subheader("2. Fotografías del vehículo (protocolo de 4 o 6 fotos)")
    st.caption("Las vistas fotográficas son OBLIGATORIAS antes de registrar el ingreso: protegen al operador y a la empresa frente a reclamaciones de daños o pérdida de pertenencias. Elige el protocolo completo (6) o el rápido (4) según la operación.")
    if CAMARA_COMPONENT is not None:
        opciones_cam_v = ["📷 Visor en vivo cámara trasera (Recomendado)", "📸 Cámara nativa del dispositivo", "📁 Subir fotos desde galería"]
    else:
        opciones_cam_v = ["📸 Cámara nativa del dispositivo (Recomendado)", "📁 Subir fotos desde galería"]
    opc_cam_v = st.radio(
        "Herramienta de captura del vehículo:",
        opciones_cam_v,
        horizontal=True,
        key="selector_camara_vehiculo",
    )

    VISTAS_COMPLETAS = [
        ("frente", "1. Vista frontal"),
        ("trasera", "2. Vista trasera"),
        ("izquierda", "3. Costado izquierdo"),
        ("derecha", "4. Costado derecho"),
        ("tablero", "5. Tablero / kilometraje"),
        ("interior", "6. Interior / asientos"),
    ]
    VISTAS_RAPIDAS = VISTAS_COMPLETAS[:4]

    modo_fotos = st.radio(
        "Protocolo de evidencia fotográfica:",
        [
            "📋 Protocolo completo (6 fotos): frente, trasera, ambos costados, tablero/kilometraje e interior",
            "⚡ Protocolo rápido (4 fotos): frente, trasera y ambos costados",
        ],
        index=0,
        key="modo_fotos_vehiculo",
        help="El protocolo completo es el oficial: documenta el kilometraje y el interior. El rápido agiliza el ingreso en horas de alta demanda.",
    )
    VISTAS_VEHICULO = VISTAS_COMPLETAS if modo_fotos.startswith("📋") else VISTAS_RAPIDAS
    total_fotos = len(VISTAS_VEHICULO)

    fotos_vehiculo = {}
    for clave_vista, titulo_vista in VISTAS_VEHICULO:
        if opc_cam_v.startswith("📷"):
            captura = camara_trasera_foto_pertenencias(key=f"foto_veh_trasera_{clave_vista}", height=380)
            fotos_vehiculo[clave_vista] = captura
            if captura is not None:
                st.image(captura, caption=f"{titulo_vista} capturada con cámara trasera", width=240)
        elif opc_cam_v.startswith("📸"):
            fotos_vehiculo[clave_vista] = st.camera_input(
                f"{titulo_vista} (obligatoria)", key=f"foto_veh_{clave_vista}"
            )
        else:
            fotos_vehiculo[clave_vista] = st.file_uploader(
                f"{titulo_vista} (obligatoria)", type=["jpg", "jpeg", "png"], key=f"foto_veh_{clave_vista}"
            )
    faltan_fotos = [titulo_vista for clave_vista, titulo_vista in VISTAS_VEHICULO if not fotos_vehiculo.get(clave_vista)]
    capturadas = total_fotos - len(faltan_fotos)
    st.progress(capturadas / total_fotos, text=f"Fotografías del vehículo capturadas: {capturadas} de {total_fotos}")
    if faltan_fotos:
        st.info("Faltan las siguientes vistas: " + ", ".join(faltan_fotos))

    if st.button("💾 Registrar ingreso y generar ticket", type="primary", use_container_width=True):
        faltantes = []
        if not es_global and not sucursal:
            faltantes.append("sucursal asignada")
        if es_global and sucursal == "-- Selecciona la Sucursal --":
            faltantes.append("sucursal")
        if not marca.strip():
            faltantes.append("marca y modelo")
        if not placas:
            faltantes.append("placas")
        if not operador.strip():
            faltantes.append("operador")
        faltantes.extend(faltan_fotos)

        if faltantes:
            st.error("Completa los siguientes campos: " + ", ".join(faltantes) + ".")
        else:
            folio, folio_numero = asignar_folio(sucursal, st.session_state.vehiculos)
            ahora_entrada = datetime.now()
            fecha = ahora_entrada.strftime("%d/%m/%Y %H:%M:%S")
            fecha_entrada_iso = ahora_entrada.isoformat(timespec="seconds")
            rutas = {}
            archivos = [
                (clave_vista, fotos_vehiculo[clave_vista])
                for clave_vista, _ in VISTAS_VEHICULO
                if fotos_vehiculo.get(clave_vista)
            ]
            try:
                for vista, archivo in archivos:
                    rutas[vista] = guardar_foto(archivo, folio, vista)
   
