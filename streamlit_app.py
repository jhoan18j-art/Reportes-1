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
# ESTILO CONSTITUCIONAL E INSTITUCIONAL - BLANCO (#0B3B24)
# CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.
# ---------------------------------------------------------------------
CSS_INSTITUCIONAL = """
<style>
/* Paleta Institucional Club Park: BLANCO Oficial */
:root {
  --cpv-verde-botella: #0B3B24;
  --cpv-blanco-oscuro: #072617;
  --cpv-verde-medio: #145A32;
  --cpv-blanco-claro: #EBF3EE;
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
    payload_b64 = base64.b64encode(payload_binario).decode("ascii")
    texto_b64 = base64.b64encode(str(texto).encode("utf-8")).decode("ascii")

    html_doc = f"""
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;border:1px solid #C4D7CC;background:#F9FBF9;border-radius:10px;padding:12px;margin:8px 0">
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <a href="rawbt:data:application/octet-stream;base64,{payload_b64}" target="_top"
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
        <b>Impresora Bluetooth clásica (p. ej. MP210, 58 mm, USB+Bluetooth):</b> 1) Empareja la impresora en <b>Ajustes › Bluetooth</b> de Android. 2) Instala la app gratuita <b>RawBT</b> desde Play Store. 3) Pulsa el botón verde <b>RawBT</b> y elige la impresora emparejada: el boleto sale con su código QR. Si el botón no hace nada, RawBT no está instalada en este teléfono: usa <b>⬇️ Descargar boleto .bin</b> y abre el archivo con RawBT (o compártelo a RawBT). Las impresoras solo BLE pueden usar el botón BLE en Chrome.
      </div>
    </div>
    <script>
    (() => {{
      const ticket = {texto_json};
      const qr = {qr_json};
      const btnBle = document.getElementById('cp_bt_{safe_key}');
      const btnShare = document.getElementById('cp_share_{safe_key}');
      const status = document.getElementById('cp_status_{safe_key}');
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
    # 1) zxing-cpp: decodificador robusto, no requiere librerías del sistema
    try:
        import zxingcpp
        resultados = zxingcpp.read_barcodes(imagen_pil)
        for r in resultados:
            texto = getattr(r, "text", "") or ""
            if texto.strip():
                return texto.strip()
    except Exception:
        pass
    try:
        from pyzbar.pyzbar import decode
        resultado = decode(imagen_pil)
        if resultado:
            return resultado[0].data.decode("utf-8").strip()
    except Exception:
        pass
    try:
        import cv2
        import numpy as np
        img = np.array(imagen_pil.convert("RGB"))
        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
        valor, _, _ = cv2.QRCodeDetector().detectAndDecode(img)
        return valor.strip() if valor else None
    except Exception:
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

        metodo = st.radio(
            "Método de autenticación:",
            [
                "📸 Foto de QR con la cámara del dispositivo (Recomendado)",
                "🖼️ Subir imagen QR",
                "🔑 Token manual",
                "📷 Visor en vivo de QR (experimental)",
            ],
            horizontal=False,
            key="metodo_login_operativo"
        )

        token = str(st.query_params.get("qrscan", "") or "").strip() or None

        if metodo.startswith("📸"):
            if not token:
                st.caption("Pulsa el botón: se abrirá la cámara de tu teléfono (usa la trasera). Enfoca la credencial QR a 10-15 cm.")
                foto_qr = st.camera_input("Capturar QR para iniciar sesión", key="camara_nativa_login")
                if foto_qr:
                    token_detectado = leer_codigo_qr(Image.open(foto_qr))
                    if token_detectado:
                        token = token_detectado.strip()
                    else:
                        st.warning("No se detectó un QR legible. Acércate, mejora la luz y vuelve a tomar la foto.")

        elif metodo.startswith("📷"):
            if not token:
                st.caption("Si el navegador bloquea la cámara del visor (común en móviles por los permisos de ventanas integradas), usa la opción recomendada de arriba o el token manual.")
                token_leido = lector_qr_camara_trasera("lector_login_trasero", height=390)
                if token_leido:
                    token = str(token_leido).strip()

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

tab_ing, tab_list, tab_salida, tab_balance, tab_imp, tab_rep, tab_corte, tab_cierre_diario, tab_aud, tab_qrs, tab_pruebas = st.tabs([
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

    if opc_cam_p == "📸 Cámara nativa del dispositivo (Recomendado)":
        foto_nativa_subida = st.camera_input("Capturar foto de pertenencias", key="camara_nativa_pert")
        if foto_nativa_subida:
            foto_pertenencias_camara = Image.open(foto_nativa_subida)
            st.image(foto_pertenencias_camara, caption="Fotografía de pertenencias capturada", width=260)
    elif opc_cam_p == "📷 Visor en vivo cámara trasera":
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
    opc_cam_v = st.radio(
        "Herramienta de captura del vehículo:",
        ["📸 Cámara nativa del dispositivo (Recomendado)", "📁 Subir fotos desde galería"],
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
        if opc_cam_v.startswith("📸"):
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
                fotos_custodia = []
                if foto_pertenencias_camara:
                    fotos_custodia.append(guardar_foto(foto_pertenencias_camara, folio, "pertenencia_camara"))
                for indice, archivo_extra in enumerate(fotos_pertenencias_archivos or [], start=1):
                    fotos_custodia.append(guardar_foto(archivo_extra, folio, f"pertenencia_{indice}"))
                registro = {
                    "id": folio,
                    "folio_numero": folio_numero,
                    "qr_payload": f"CLUBPARK|VEHICULO|{folio}",
                    "fotos_requeridas": total_fotos,
                    "fecha": fecha,
                    "fecha_entrada": fecha,
                    "fecha_entrada_iso": fecha_entrada_iso,
                    "hora_entrada": ahora_entrada.strftime("%H:%M:%S"),
                    "fecha_salida": "",
                    "hora_salida": "",
                    "salida_fecha": "",
                    "sucursal": sucursal,
                    "marca": marca.strip(),
                    "color": color.strip(),
                    "placas": placas,
                    "tarifa": float(tarifa),
                    "operador": operador.strip(),
                    "custodia": custodia.strip() or "Sin objetos declarados",
                    "fotos": rutas,
                    "fotos_custodia": fotos_custodia,
                    "estado": "En Resguardo",
                    "registrado_por": st.session_state.usuario_autenticado,
                }
                registro["ticket_texto"] = crear_ticket(registro)
                st.session_state.vehiculos.insert(0, registro)
                if guardar_bd(st.session_state.vehiculos):
                    st.session_state.ultimo_ticket = registro["ticket_texto"]
                    st.session_state.ultimo_folio = folio
                    st.success(f"Ingreso registrado. Folio: {folio}")
                    st.rerun()
            except Exception as exc:
                st.error(f"No se pudieron guardar las fotografías o el registro: {exc}")

    if st.session_state.get("ultimo_ticket"):
        st.divider()
        st.subheader(f"Ticket generado: {st.session_state.get('ultimo_folio', '')}")
        ultimo = next((r for r in st.session_state.vehiculos if r.get("id") == st.session_state.get("ultimo_folio")), None)
        if ultimo:
            if QR_OK:
                qr_data = generar_qr_bytes(ultimo)
                st.image(qr_data, caption="QR único del vehículo (identifica el folio; no es QR de pago)", width=180)
                st.download_button(
                    "⬇️ Descargar QR del vehículo",
                    data=qr_data,
                    file_name=f"QR_{ultimo.get('id')}.png",
                    mime="image/png",
                    key="descargar_qr_ingreso",
                )
            else:
                st.warning("Instala qrcode[pil] para generar imágenes QR.")
            vista_impresion_con_qr(ultimo, "cliente", "imprimir_cliente_ingreso")
            vista_impresion_con_qr(ultimo, "llavero", "imprimir_contrasena_ingreso")
            st.warning("IMPORTANTE: para imprimir el QR gráfico usa el botón 'Imprimir boleto del cliente + QR de salida' de arriba. El botón Bluetooth de texto solo imprime texto y no puede dibujar el QR. El QR se coloca antes del primer corte.")
            imprimir_partes_ticket(ultimo, "ingreso")
        else:
            st.code(st.session_state.ultimo_ticket, language="text")


# ---------------------------------------------------------------------
# TAB 2: VEHÍCULOS REGISTRADOS
# ---------------------------------------------------------------------
with tab_list:
    st.subheader("Vehículos registrados")
    visibles = vehiculos_visibles(st.session_state.vehiculos)
    if not visibles:
        st.info("No hay vehículos registrados para tu alcance.")
    else:
        df_v = pd.DataFrame(visibles)
        columnas_mostrar = ["id", "fecha", "sucursal", "marca", "placas", "tarifa", "operador", "estado", "pago_estado", "metodo_pago"]
        for col in columnas_mostrar:
            if col not in df_v.columns:
                df_v[col] = ""
        st.dataframe(df_v[columnas_mostrar], use_container_width=True, hide_index=True)

        folios = [v.get("id", "") for v in visibles]
        folio_sel = st.selectbox("Selecciona un folio para ver sus datos y fotografías", folios, key="lista_folio")
        detalle = next((v for v in visibles if v.get("id") == folio_sel), None)
        if detalle:
            st.write(
                f"**Sucursal:** {detalle.get('sucursal', '')} · "
                f"**Placas:** {detalle.get('placas', '')} · "
                f"**Color:** {detalle.get('color', '')} · "
                f"**Custodia:** {detalle.get('custodia', '')}"
            )
            cols_foto = st.columns(4)
            for col, (vista, etiqueta) in zip(
                cols_foto,
                [("frente", "Frente"), ("izquierda", "Izquierdo"),
                 ("derecha", "Derecho"), ("trasera", "Trasera")]
            ):
                with col:
                    p = ruta_foto(detalle, vista)
                    st.caption(etiqueta)
                    if p:
                        st.image(str(p), use_container_width=True)
                    else:
                        st.warning("Foto no disponible")
            fotos_custodia_detalle = detalle.get("fotos_custodia", []) or []
            if fotos_custodia_detalle:
                st.markdown("**📷 Evidencia fotográfica de pertenencias**")
                cols_custodia = st.columns(min(3, len(fotos_custodia_detalle)))
                for i, ruta_extra in enumerate(fotos_custodia_detalle):
                    p_extra = Path(ruta_extra)
                    if not p_extra.is_absolute():
                        p_extra = DATA_DIR / p_extra
                    with cols_custodia[i % len(cols_custodia)]:
                        if p_extra.exists():
                            st.image(str(p_extra), caption=f"Pertenencia {i + 1}", use_container_width=True)
                        else:
                            st.warning(f"Foto de pertenencia {i + 1} no disponible")

            if es_global:
                estados = ["En Resguardo", "Entregado", "Incidencia", "Cancelado"]
                estado_actual = detalle.get("estado", "En Resguardo")
                estado_nuevo = st.selectbox(
                    "Cambiar estado",
                    estados,
                    index=estados.index(estado_actual) if estado_actual in estados else 0,
                    key=f"estado_{folio_sel}",
                )
                if st.button("Actualizar estado", key=f"actualizar_{folio_sel}"):
                    for r in st.session_state.vehiculos:
                        if r.get("id") == folio_sel:
                            r["estado"] = estado_nuevo
                            r["actualizado_por"] = st.session_state.usuario_autenticado
                            r["actualizado_en"] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                            break
                    if guardar_bd(st.session_state.vehiculos):
                        st.success("Estado actualizado.")
                        st.rerun()



# ---------------------------------------------------------------------
# TAB: SALIDA Y COBRO
# ---------------------------------------------------------------------
with tab_salida:
    st.subheader("Salida de vehículo y registro de cobro")
    st.caption("Escanea el QR único del boleto o captura el folio. La salida solo se registra una vez.")
    metodo_busqueda = st.radio(
        "Cómo localizar el vehículo",
        ["Capturar folio", "Escanear QR con cámara", "Subir imagen QR"],
        horizontal=True,
        key="metodo_busqueda_salida",
    )
    valor_busqueda = ""
    if metodo_busqueda == "Capturar folio":
        valor_busqueda = st.text_input("Folio del boleto", key="folio_buscar_salida").strip()
    elif metodo_busqueda == "Escanear QR con cámara":
        valor_query = str(st.query_params.get("qrscan", "") or "").strip()
        st.info("Se abrirá la cámara posterior. Mantén el QR dentro del recuadro; al detectarlo, el lector lo enviará directamente al formulario.")
        nonce_lector = int(st.session_state.get("qr_scan_nonce_salida", 0))
        valor_leido = lector_qr_camara_trasera(f"lector_salida_trasero_{nonce_lector}", height=430)
        valor_busqueda = str(valor_leido or valor_query or "").strip()
        if valor_busqueda:
            st.success(f"QR leído: {valor_busqueda}")
            if st.button("🧹 Limpiar QR leído", key="limpiar_qr_salida"):
                st.session_state.qr_scan_nonce_salida = int(st.session_state.get("qr_scan_nonce_salida", 0)) + 1
                st.query_params.clear()
                st.rerun()
    else:
        archivo_qr_salida = st.file_uploader("Sube la imagen QR del ticket", type=["png", "jpg", "jpeg"], key="archivo_qr_salida")
        if archivo_qr_salida:
            valor_busqueda = leer_codigo_qr(Image.open(archivo_qr_salida)) or ""

    folio_busqueda = extraer_folio_qr(valor_busqueda)
    registro_salida = next(
        (r for r in st.session_state.vehiculos if str(r.get("id", "")).upper() == str(folio_busqueda).upper()),
        None,
    ) if folio_busqueda else None

    if folio_busqueda and not registro_salida:
        st.error("No se encontró un vehículo con ese folio/QR.")
    if registro_salida and not usuario_es_global() and registro_salida.get("sucursal") != st.session_state.sucursal_asignada:
        st.error("Este vehículo pertenece a otra sucursal. No tienes permiso para cobrarlo ni entregarlo.")
        registro_salida = None

    if registro_salida:
        st.divider()
        a, b = st.columns([2, 1])
        with a:
            st.markdown(f"### Folio {registro_salida.get('id')}")
            st.write(f"**Sucursal:** {registro_salida.get('sucursal')}")
            st.write(f"**Marca/modelo:** {registro_salida.get('marca', '')}")
            st.write(f"**Placas:** {registro_salida.get('placas', '')} · **Color:** {registro_salida.get('color', '')}")
            st.write(f"**Fecha/hora de entrada:** {registro_salida.get('fecha_entrada', registro_salida.get('fecha', ''))}")
            st.write(f"**Fecha/hora de salida:** {registro_salida.get('salida_fecha', 'Pendiente de salida') or 'Pendiente de salida'}")
            st.write(f"**Estado:** {registro_salida.get('estado', 'En Resguardo')}")
        with b:
            if QR_OK:
                st.image(generar_qr_bytes(registro_salida), caption="QR del vehículo", width=150)
        if registro_salida.get("estado") == "Entregado":
            st.success(f"Este vehículo ya fue entregado el {registro_salida.get('salida_fecha', 'fecha no disponible')}.")
            st.info("Para evitar cobros duplicados, no se permite registrar otra salida.")
        elif registro_salida.get("estado") != "En Resguardo":
            st.warning("El vehículo no está en estado 'En Resguardo'. Revisa el registro antes de continuar.")
        else:
            tarifa_salida = dinero(registro_salida.get("tarifa", PRECIOS_SUCURSALES.get(registro_salida.get("sucursal"), 0)))
            st.metric("Tarifa fija a cobrar", f"${tarifa_salida:,.2f} MXN")
            st.caption("La tarifa se toma del registro de ingreso. Verifica cualquier cambio autorizado antes de cobrar.")
            forma_pago = st.radio("Método de pago", ["Efectivo", "Tarjeta", "QR"], horizontal=True, key=f"pago_metodo_{registro_salida['id']}")
            referencia_pago = st.text_input("Referencia / autorización (si aplica)", key=f"pago_ref_{registro_salida['id']}")
            nota_pago = st.text_input("Observaciones del cobro", key=f"pago_nota_{registro_salida['id']}")
            if forma_pago == "Efectivo":
                st.info("El efectivo se registrará como recibido al confirmar la entrega.")
            elif forma_pago == "Tarjeta":
                st.warning("Registra el pago con la terminal bancaria. Esta versión guarda el comprobante y la referencia capturada, pero no consulta automáticamente al banco.")
            else:
                st.warning("Cobra con la app/proveedor QR. Esta versión guarda la referencia y la confirmación del operador; no valida automáticamente la aprobación del proveedor.")
            referencia_obligatoria = forma_pago != "Efectivo"
            pago_aprobado_externamente = False
            if referencia_obligatoria:
                pago_aprobado_externamente = st.checkbox(
                    f"Confirmo que revisé la aprobación en la terminal/app de {forma_pago} y que el importe es correcto.",
                    key=f"pago_externo_aprobado_{registro_salida['id']}",
                )
            confirmar_pago = st.checkbox(
                "Confirmo que verifiqué el importe y que el vehículo puede ser entregado.",
                key=f"confirmar_salida_{registro_salida['id']}",
            )
            if st.button("💳 Registrar cobro y salida", type="primary", use_container_width=True, disabled=not confirmar_pago, key=f"btn_salida_{registro_salida['id']}"):
                if referencia_obligatoria and not referencia_pago.strip():
                    st.error("Captura la referencia de autorización del proveedor.")
                elif referencia_obligatoria and not pago_aprobado_externamente:
                    st.error("No se registró la salida: confirma primero que el proveedor aprobó el pago.")
                else:
                    momento_operacion = datetime.now()
                    ahora = momento_operacion.strftime("%Y-%m-%d %H:%M:%S")
                    pago = {
                        "id_pago": str(uuid.uuid4()),
                        "fecha_hora": ahora,
                        "fecha_hora_iso": momento_operacion.isoformat(timespec="seconds"),
                        "sucursal": registro_salida.get("sucursal"),
                        "folio": registro_salida.get("id"),
                        "importe": tarifa_salida,
                        "metodo": forma_pago,
                        "estado": "Pagado",
                        "referencia": referencia_pago.strip(),
                        "observaciones": nota_pago.strip(),
                        "usuario": st.session_state.usuario_autenticado,
                        "confirmacion_manual": forma_pago != "Efectivo",
                    }
                    registro_salida.setdefault("pagos", []).append(pago)
                    registro_salida["pago_estado"] = "Pagado"
                    registro_salida["metodo_pago"] = forma_pago
                    registro_salida["importe_cobrado"] = tarifa_salida
                    registro_salida["salida_fecha"] = ahora
                    registro_salida["fecha_salida"] = momento_operacion.strftime("%d/%m/%Y")
                    registro_salida["hora_salida"] = momento_operacion.strftime("%H:%M:%S")
                    registro_salida["salida_fecha_hora_iso"] = momento_operacion.isoformat(timespec="seconds")
                    registro_salida["entregado_por"] = st.session_state.usuario_autenticado
                    registro_salida["estado"] = "Entregado"
                    # Actualiza el ticket para que las reimpresiones posteriores reflejen la salida real.
                    registro_salida["ticket_texto"] = crear_ticket(registro_salida)
                    respaldo_ok = guardar_pago_respaldo(pago, registro_salida)
                    if respaldo_ok and guardar_bd(st.session_state.vehiculos):
                        st.session_state.ultimo_comprobante_cobro = {
                            "pago": pago.copy(),
                            "registro": {
                                "id": registro_salida.get("id", ""),
                                "sucursal": registro_salida.get("sucursal", ""),
                                "marca": registro_salida.get("marca", ""),
                                "placas": registro_salida.get("placas", ""),
                            },
                        }
                        guardar_evento_auditoria("SALIDA Y COBRO", f"Folio {registro_salida.get('id')} · ${tarifa_salida:.2f} · {forma_pago} · {registro_salida.get('sucursal')}")
                        st.success(f"Salida registrada. Folio {registro_salida.get('id')}; cobro de ${tarifa_salida:,.2f} MXN por {forma_pago}.")
                        st.rerun()


    comprobante_reciente = st.session_state.get("ultimo_comprobante_cobro")
    if comprobante_reciente:
        st.divider()
        st.subheader("🧾 Último comprobante de cobro")
        st.caption("El comprobante se genera después del cobro y la salida. Los cobros también se guardan en un libro independiente de pagos para respaldo. En tarjeta y QR, la aprobación se confirma manualmente en la terminal/app.")
        vista_comprobante_cobro(
            comprobante_reciente.get("pago", {}),
            comprobante_reciente.get("registro", {}),
            "imprimir_comprobante_cobro_reciente",
        )
        pago_reciente = comprobante_reciente.get("pago", {})
        registro_reciente = comprobante_reciente.get("registro", {})
        texto_comprobante = "\n".join([
            "CLUB PARK",
            "VALET Y ESTACIONAMIENTOS S.A. DE C.V.",
            "COMPROBANTE DE COBRO",
            "--------------------------------",
            f"Sucursal: {registro_reciente.get('sucursal', '')}",
            f"Folio: {registro_reciente.get('id', '')}",
            f"Vehiculo: {registro_reciente.get('marca', '')} {registro_reciente.get('placas', '')}",
            f"Fecha: {pago_reciente.get('fecha_hora', '')}",
            f"Importe: ${dinero(pago_reciente.get('importe')):,.2f} MXN",
            f"Metodo: {pago_reciente.get('metodo', '')}",
            f"Estado: {pago_reciente.get('estado', '')}",
            f"Referencia: {pago_reciente.get('referencia', '') or 'No aplica'}",
            f"Atendio: {pago_reciente.get('usuario', '')}",
            "--------------------------------",
            "Conserve este comprobante.",
            "Gracias por su preferencia.",
        ])
        st.markdown("**Impresión térmica Bluetooth del comprobante**")
        imprimir_bluetooth(texto_comprobante, f"comprobante_{registro_reciente.get('id', 'reciente')}")
        st.download_button(
            "⬇️ Descargar comprobante en texto",
            data="CLUB PARK\nCOMPROBANTE DE COBRO\n" + "\n".join(
                f"{k}: {v}" for k, v in {
                    "Sucursal": comprobante_reciente.get("registro", {}).get("sucursal", ""),
                    "Folio": comprobante_reciente.get("registro", {}).get("id", ""),
                    "Fecha y hora": comprobante_reciente.get("pago", {}).get("fecha_hora", ""),
                    "Importe": f"${dinero(comprobante_reciente.get('pago', {}).get('importe')):,.2f} MXN",
                    "Método": comprobante_reciente.get("pago", {}).get("metodo", ""),
                    "Estado": comprobante_reciente.get("pago", {}).get("estado", ""),
                    "Referencia": comprobante_reciente.get("pago", {}).get("referencia", "") or "No aplica",
                }.items()
            ),
            file_name=f"Comprobante_{comprobante_reciente.get('registro', {}).get('id', 'CLUBPARK')}.txt",
            mime="text/plain",
            key="descargar_comprobante_cobro_reciente",
        )


# ---------------------------------------------------------------------
# TAB: BALANCES Y RENTABILIDAD POR SUCURSAL
# ---------------------------------------------------------------------
with tab_balance:
    st.subheader("Balances operativos y rentabilidad por sucursal")
    st.caption(
        "El conteo de vehículos y los ingresos se calculan desde los registros. "
        "La rentabilidad usa los gastos operativos capturados por el supervisor o auditor."
    )

    COSTOS_FILE = DATA_DIR / "gastos_operativos.json"

    def cargar_gastos_operativos():
        if not COSTOS_FILE.exists():
            return []
        try:
            datos = json.loads(COSTOS_FILE.read_text(encoding="utf-8"))
            return datos if isinstance(datos, list) else []
        except (OSError, json.JSONDecodeError):
            return []

    def guardar_gastos_operativos(datos):
        temporal = COSTOS_FILE.with_suffix(".tmp")
        temporal.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temporal, COSTOS_FILE)

    if not (es_contador or "Supervisor" in str(st.session_state.get("rol", ""))):
        st.error("No tienes permisos para consultar balances de rentabilidad.")
    else:
        sucursales_balance = SUCURSALES if es_global else [st.session_state.sucursal_asignada]
        colf1, colf2 = st.columns(2)
        with colf1:
            fecha_desde = st.date_input("Desde", value=datetime.now().date(), key="balance_desde")
        with colf2:
            fecha_hasta = st.date_input("Hasta", value=datetime.now().date(), key="balance_hasta")

        if fecha_desde > fecha_hasta:
            st.error("La fecha inicial no puede ser posterior a la fecha final.")
        else:
            registros_balance = vehiculos_visibles(st.session_state.vehiculos)

            def fecha_registro(registro):
                for campo in ("salida_fecha", "fecha"):
                    valor = registro.get(campo, "")
                    for formato in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y"):
                        try:
                            return datetime.strptime(valor, formato).date()
                        except (ValueError, TypeError):
                            continue
                return None

            registros_periodo = []
            for registro in registros_balance:
                f_ingreso = fecha_registro(registro)
                try:
                    f_ingreso = datetime.strptime(registro.get("fecha", ""), "%d/%m/%Y %H:%M:%S").date()
                except (ValueError, TypeError):
                    try:
                        f_ingreso = datetime.strptime(registro.get("fecha", ""), "%Y-%m-%d %H:%M:%S").date()
                    except (ValueError, TypeError):
                        f_ingreso = None
                if f_ingreso and fecha_desde <= f_ingreso <= fecha_hasta:
                    registros_periodo.append(registro)

            gastos = cargar_gastos_operativos()
            gastos_periodo = []
            for gasto in gastos:
                if not es_global and gasto.get("sucursal") != st.session_state.sucursal_asignada:
                    continue
                try:
                    f_gasto = datetime.strptime(gasto.get("fecha", ""), "%Y-%m-%d").date()
                except (ValueError, TypeError):
                    continue
                if fecha_desde <= f_gasto <= fecha_hasta:
                    gastos_periodo.append(gasto)

            filas = []
            for suc in sucursales_balance:
                todos_suc = [r for r in registros_balance if r.get("sucursal") == suc]
                entradas = 0
                salidas_periodo = []
                ingresos = 0.0
                folios_cobrados = set()

                for reg in todos_suc:
                    try:
                        ingreso_dt = datetime.strptime(reg.get("fecha", ""), "%d/%m/%Y %H:%M:%S").date()
                    except (ValueError, TypeError):
                        try:
                            ingreso_dt = datetime.strptime(reg.get("fecha", ""), "%Y-%m-%d %H:%M:%S").date()
                        except (ValueError, TypeError):
                            ingreso_dt = None
                    if ingreso_dt and fecha_desde <= ingreso_dt <= fecha_hasta:
                        entradas += 1

                    salida_dt = None
                    try:
                        salida_dt = datetime.strptime(reg.get("salida_fecha", ""), "%Y-%m-%d %H:%M:%S").date()
                    except (ValueError, TypeError):
                        pass
                    if reg.get("estado") == "Entregado" and salida_dt and fecha_desde <= salida_dt <= fecha_hasta:
                        salidas_periodo.append(reg)
                        pagos = [p for p in reg.get("pagos", []) if p.get("estado") == "Pagado"]
                        if pagos:
                            for pago in pagos:
                                pago_dt = None
                                try:
                                    pago_dt = datetime.strptime(pago.get("fecha_hora", ""), "%Y-%m-%d %H:%M:%S").date()
                                except (ValueError, TypeError):
                                    pass
                                if pago_dt is None or fecha_desde <= pago_dt <= fecha_hasta:
                                    ingresos += dinero(pago.get("importe", 0))
                                    folios_cobrados.add(reg.get("id"))
                        elif reg.get("pago_estado") == "Pagado":
                            ingresos += dinero(reg.get("importe_cobrado", reg.get("tarifa", 0)))
                            folios_cobrados.add(reg.get("id"))

                gastos_suc = sum(dinero(g.get("importe", 0)) for g in gastos_periodo if g.get("sucursal") == suc)
                utilidad = ingresos - gastos_suc
                margen = (utilidad / ingresos * 100) if ingresos > 0 else 0.0
                filas.append({
                    "Sucursal": suc,
                    "Vehículos ingresados": entradas,
                    "Vehículos que salieron": len(salidas_periodo),
                    "Vehículos cobrados": len(folios_cobrados),
                    "Pendientes de cobro": max(0, len(salidas_periodo) - len(folios_cobrados)),
                    "Ingresos ($)": ingresos,
                    "Gastos operativos ($)": gastos_suc,
                    "Utilidad estimada ($)": utilidad,
                    "Margen (%)": margen,
                })

            df_balance = pd.DataFrame(filas)
            if not es_global:
                suc_sel = st.session_state.sucursal_asignada
                df_balance = df_balance[df_balance["Sucursal"] == suc_sel]

            total_ingresos = float(df_balance["Ingresos ($)"].sum()) if not df_balance.empty else 0
            total_gastos = float(df_balance["Gastos operativos ($)"].sum()) if not df_balance.empty else 0
            total_utilidad = total_ingresos - total_gastos
            vehiculos_ingresados = int(df_balance["Vehículos ingresados"].sum()) if not df_balance.empty else 0
            vehiculos_salieron = int(df_balance["Vehículos que salieron"].sum()) if not df_balance.empty else 0
            vehiculos_cobrados = int(df_balance["Vehículos cobrados"].sum()) if not df_balance.empty else 0
            margen_total = total_utilidad / total_ingresos * 100 if total_ingresos else 0

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Vehículos ingresados", vehiculos_ingresados)
            k2.metric("Vehículos que salieron", vehiculos_salieron)
            k3.metric("Vehículos cobrados", vehiculos_cobrados)
            k4.metric("Pendientes de cobro", max(0, vehiculos_salieron - vehiculos_cobrados))
            k5, k6, k7 = st.columns(3)
            k5.metric("Ingresos confirmados", f"${total_ingresos:,.2f} MXN")
            k6.metric("Gastos operativos", f"${total_gastos:,.2f} MXN")
            k7.metric("Utilidad estimada", f"${total_utilidad:,.2f} MXN", delta=f"{margen_total:.1f}% margen")

            st.markdown("#### Balance por sucursal")
            st.dataframe(
                df_balance.style.format({
                    "Ingresos ($)": "${:,.2f}",
                    "Gastos operativos ($)": "${:,.2f}",
                    "Utilidad estimada ($)": "${:,.2f}",
                    "Margen (%)": "{:.1f}%",
                }),
                use_container_width=True,
                hide_index=True,
            )

            if es_contador or "Supervisor" in str(st.session_state.get("rol", "")):
                st.markdown("#### Registrar gasto operativo")
                st.caption("El Contador Auditor puede registrar gastos de cualquier sucursal; cada supervisor solo puede registrar gastos de su propia sucursal.")
                with st.form("form_gasto_operativo", clear_on_submit=True):
                    opciones_gasto = SUCURSALES if es_contador else [st.session_state.sucursal_asignada]
                    suc_gasto = st.selectbox("Sucursal del gasto", opciones_gasto, key="gasto_sucursal")
                    fecha_gasto = st.date_input("Fecha del gasto", value=datetime.now().date(), key="gasto_fecha")
                    categoria_gasto = st.selectbox(
                        "Categoría",
                        ["Nómina", "Renta", "Insumos", "Mantenimiento", "Comisiones", "Servicios", "Otros"],
                        key="gasto_categoria",
                    )
                    concepto_gasto = st.text_input("Concepto / descripción", key="gasto_concepto")
                    importe_gasto = st.number_input("Importe ($)", min_value=0.01, value=1.0, step=50.0, key="gasto_importe")
                    if st.form_submit_button("Guardar gasto"):
                        if not concepto_gasto.strip():
                            st.error("Escribe una descripción para el gasto.")
                        else:
                            nuevos_gastos = cargar_gastos_operativos()
                            nuevo_gasto = {
                                "id": uuid.uuid4().hex[:12].upper(),
                                "fecha": str(fecha_gasto),
                                "sucursal": suc_gasto,
                                "categoria": categoria_gasto,
                                "concepto": concepto_gasto.strip(),
                                "importe": float(importe_gasto),
                                "registrado_por": st.session_state.usuario_autenticado,
                                "registrado_en": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            }
                            nuevos_gastos.insert(0, nuevo_gasto)
                            try:
                                guardar_gastos_operativos(nuevos_gastos)
                                guardar_evento_auditoria("GASTO OPERATIVO", f"{suc_gasto} · {categoria_gasto} · ${importe_gasto:.2f} · {concepto_gasto.strip()}")
                                st.success("Gasto registrado y añadido al balance.")
                                st.rerun()
                            except OSError as exc:
                                st.error(f"No se pudo guardar el gasto: {exc}")

            st.markdown("#### Rentabilidad por sucursal")
            st.caption("Margen = (ingresos confirmados − gastos registrados) ÷ ingresos confirmados. No incluye costos que todavía no se hayan capturado.")
            if not df_balance.empty:
                st.bar_chart(df_balance.set_index("Sucursal")[["Ingresos ($)", "Gastos operativos ($)", "Utilidad estimada ($)"]])
                st.download_button(
                    "⬇️ Descargar balance CSV",
                    data=df_balance.to_csv(index=False).encode("utf-8-sig"),
                    file_name=f"balance_clubpark_{fecha_desde}_{fecha_hasta}.csv",
                    mime="text/csv",
                    key="descargar_balance_csv",
                )


# ---------------------------------------------------------------------
# TAB: CORTES DE CAJA
# ---------------------------------------------------------------------
with tab_corte:
    st.subheader("Corte de caja")
    st.caption("Los totales se calculan con salidas registradas y pagos confirmados. Pagos QR/tarjeta no confirmados no se suman a lo cobrado.")
    registros_entregados = [
        r for r in vehiculos_visibles(st.session_state.vehiculos)
        if r.get("estado") == "Entregado" and r.get("salida_fecha")
    ]
    sucursales_corte = SUCURSALES if es_global else [st.session_state.sucursal_asignada]
    suc_corte = st.selectbox("Sucursal", sucursales_corte, key="corte_sucursal")
    fecha_corte = st.date_input("Fecha del corte", value=datetime.now().date(), key="fecha_corte")
    inicio_corte = st.text_input("Hora inicial (HH:MM)", value="00:00", key="hora_inicio_corte")
    fin_corte = st.text_input("Hora final (HH:MM)", value="23:59", key="hora_fin_corte")
    try:
        desde_dt = datetime.strptime(f"{fecha_corte:%Y-%m-%d} {inicio_corte}", "%Y-%m-%d %H:%M")
        hasta_dt = datetime.strptime(f"{fecha_corte:%Y-%m-%d} {fin_corte}", "%Y-%m-%d %H:%M")
    except ValueError:
        desde_dt = hasta_dt = None
        st.error("Usa formato de hora HH:MM, por ejemplo 08:00.")
    movimientos_corte = []
    if desde_dt and hasta_dt and desde_dt <= hasta_dt:
        for reg in registros_entregados:
            if reg.get("sucursal") != suc_corte:
                continue
            try:
                salida_dt = datetime.strptime(reg.get("salida_fecha", ""), "%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue
            if not (desde_dt <= salida_dt <= hasta_dt):
                continue
            for pago in reg.get("pagos", []):
                if pago.get("estado") == "Pagado":
                    movimientos_corte.append(pago)
    elif desde_dt and hasta_dt:
        st.error("La hora inicial no puede ser posterior a la hora final.")

    efectivo_total = sum(dinero(p.get("importe")) for p in movimientos_corte if p.get("metodo") == "Efectivo")
    qr_total = sum(dinero(p.get("importe")) for p in movimientos_corte if p.get("metodo") == "QR")
    tarjeta_total = sum(dinero(p.get("importe")) for p in movimientos_corte if p.get("metodo") == "Tarjeta")
    total_cobrado = efectivo_total + qr_total + tarjeta_total
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Efectivo", f"${efectivo_total:,.2f}")
    m2.metric("QR confirmado", f"${qr_total:,.2f}")
    m3.metric("Tarjeta confirmada", f"${tarjeta_total:,.2f}")
    m4.metric("Total cobrado", f"${total_cobrado:,.2f}")
    st.write(f"**Movimientos confirmados:** {len(movimientos_corte)}")
    if movimientos_corte:
        st.dataframe(pd.DataFrame(movimientos_corte), use_container_width=True, hide_index=True)
    else:
        st.info("No hay pagos confirmados en ese periodo.")

    st.markdown("#### Arqueo de efectivo")
    apertura = st.number_input("Fondo inicial de caja ($)", min_value=0.0, value=0.0, step=50.0, key="fondo_apertura")
    efectivo_contado = st.number_input("Efectivo contado físicamente al cierre ($)", min_value=0.0, value=0.0, step=50.0, key="efectivo_contado")
    efectivo_esperado = apertura + efectivo_total
    diferencia = efectivo_contado - efectivo_esperado
    c1, c2, c3 = st.columns(3)
    c1.metric("Efectivo esperado", f"${efectivo_esperado:,.2f}")
    c2.metric("Efectivo contado", f"${efectivo_contado:,.2f}")
    c3.metric("Diferencia", f"${diferencia:,.2f}")
    if st.button("🔒 Guardar corte de caja", type="primary", key="guardar_corte"):
        corte = {
            "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "fecha_operativa": str(fecha_corte),
            "desde": inicio_corte,
            "hasta": fin_corte,
            "sucursal": suc_corte,
            "usuario": st.session_state.usuario_autenticado,
            "efectivo": efectivo_total,
            "qr_confirmado": qr_total,
            "tarjeta_confirmada": tarjeta_total,
            "total_cobrado": total_cobrado,
            "fondo_inicial": apertura,
            "efectivo_contado": efectivo_contado,
            "diferencia": diferencia,
            "folios": [p.get("folio") for p in movimientos_corte],
        }
        cortes_file = DATA_DIR / "cortes_caja.json"
        try:
            cortes = []
            if cortes_file.exists():
                cortes = json.loads(cortes_file.read_text(encoding="utf-8"))
            cortes.insert(0, corte)
            cortes_file.write_text(json.dumps(cortes, ensure_ascii=False, indent=2), encoding="utf-8")
            guardar_evento_auditoria("CORTE DE CAJA", f"{suc_corte} · {fecha_corte} · total ${total_cobrado:.2f} · diferencia ${diferencia:.2f}")
            st.success("Corte guardado con usuario, sucursal, periodo, métodos de pago y diferencia.")
        except (OSError, json.JSONDecodeError) as exc:
            st.error(f"No se pudo guardar el corte: {exc}")


# ---------------------------------------------------------------------
# CIERRE DIARIO POR SUCURSAL
# ---------------------------------------------------------------------
with tab_cierre_diario:
    st.subheader("Reporte diario de cierre por sucursal")
    st.caption(
        "Captura y entrega el balance diario de boletos, cobros y gastos. "
        "Cada sucursal puede guardar un cierre por fecha; el contador puede consultar todas."
    )

    CIERRES_DIARIOS_FILE = DATA_DIR / "cierres_diarios_sucursal.json"
    GASTOS_CIERRE_FILE = DATA_DIR / "gastos_operativos.json"

    def cargar_cierres_diarios():
        if not CIERRES_DIARIOS_FILE.exists():
            return []
        try:
            contenido = json.loads(CIERRES_DIARIOS_FILE.read_text(encoding="utf-8"))
            return contenido if isinstance(contenido, list) else []
        except (OSError, json.JSONDecodeError):
            return []

    def guardar_cierres_diarios(contenido):
        temporal = CIERRES_DIARIOS_FILE.with_suffix(".tmp")
        temporal.write_text(json.dumps(contenido, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temporal, CIERRES_DIARIOS_FILE)

    def fecha_a_date(valor):
        for formato in ("%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(str(valor or ""), formato).date()
            except (ValueError, TypeError):
                continue
        return None

    def numero_folio(valor):
        try:
            return int(str(valor or "").strip().rsplit("-", 1)[-1])
        except (ValueError, TypeError):
            return None

    sucursales_cierre = SUCURSALES if es_global else [st.session_state.sucursal_asignada]
    fecha_cierre = st.date_input("Fecha operativa del cierre", value=datetime.now().date(), key="fecha_reporte_cierre_diario")
    sucursal_cierre = st.selectbox("Sucursal del reporte", sucursales_cierre, key="sucursal_reporte_cierre_diario")

    registros_cierre = [
        r for r in st.session_state.vehiculos
        if r.get("sucursal") == sucursal_cierre and fecha_a_date(r.get("fecha")) == fecha_cierre
    ]
    gastos_cierre = []
    if GASTOS_CIERRE_FILE.exists():
        try:
            datos_gastos = json.loads(GASTOS_CIERRE_FILE.read_text(encoding="utf-8"))
            if isinstance(datos_gastos, list):
                gastos_cierre = [
                    g for g in datos_gastos
                    if g.get("sucursal") == sucursal_cierre and fecha_a_date(g.get("fecha")) == fecha_cierre
                ]
        except (OSError, json.JSONDecodeError):
            gastos_cierre = []

    total_gastos_registrados = sum(dinero(g.get("importe")) for g in gastos_cierre)
    pagos_confirmados_cierre = []
    folios_pagados = set()
    ingresos_confirmados_cierre = 0.0
    for reg in registros_cierre:
        pagos_reg = [pago for pago in reg.get("pagos", []) if pago.get("estado") == "Pagado"]
        if pagos_reg:
            for pago in pagos_reg:
                pagos_confirmados_cierre.append(pago)
                ingresos_confirmados_cierre += dinero(pago.get("importe"))
            folios_pagados.add(reg.get("id"))
        elif reg.get("pago_estado") == "Pagado":
            folios_pagados.add(reg.get("id"))
            ingresos_confirmados_cierre += dinero(reg.get("importe_cobrado", reg.get("tarifa", 0)))

    cancelados_auto = sum(1 for r in registros_cierre if str(r.get("estado", "")).lower() in ("cancelado", "cancelada"))
    perdidos_auto = sum(1 for r in registros_cierre if str(r.get("estado", "")).lower() in ("perdido", "perdida", "extraviado", "extraviada"))
    pares_folios = [(numero_folio(r.get("id")), str(r.get("id", ""))) for r in registros_cierre]
    pares_folios = [(n, folio) for n, folio in pares_folios if n is not None]
    folio_inicial_default = min(pares_folios, key=lambda x: x[0])[1] if pares_folios else ""
    folio_final_default = max(pares_folios, key=lambda x: x[0])[1] if pares_folios else ""
    clave_widget_cierre = f"{sucursal_cierre}_{fecha_cierre:%Y%m%d}".replace(" ", "_")

    st.markdown("#### Datos del turno y control de boletos")
    st.caption("Los conteos se precargan desde los registros del sistema. Puedes corregirlos si hubo boletos físicos no capturados digitalmente.")
    c1, c2, c3 = st.columns(3)
    with c1:
        fecha_precio = st.number_input(
            "Costo unitario del boleto ($)", min_value=0.0,
            value=float(PRECIOS_SUCURSALES.get(sucursal_cierre, 0)), step=5.0,
            key=f"cierre_precio_{clave_widget_cierre}"
        )
        folio_inicial = st.text_input(
            "Folio inicial de la sesión", value=folio_inicial_default,
            key=f"cierre_folio_ini_{clave_widget_cierre}"
        )
    with c2:
        folio_final = st.text_input(
            "Folio final de la sesión", value=folio_final_default,
            key=f"cierre_folio_fin_{clave_widget_cierre}"
        )
        total_boletos_auto = len(registros_cierre)
        ni = numero_folio(folio_inicial)
        nf = numero_folio(folio_final)
        if ni is not None and nf is not None and nf >= ni:
            total_boletos_auto = nf - ni + 1
        total_boletos = st.number_input(
            "Total de boletos utilizados", min_value=0, value=int(total_boletos_auto), step=1,
            key=f"cierre_total_{clave_widget_cierre}"
        )
    with c3:
        boletos_pagados = st.number_input(
            "Boletos pagados", min_value=0, value=int(len(folios_pagados)), step=1,
            key=f"cierre_pagados_{clave_widget_cierre}"
        )
        boletos_cancelados = st.number_input(
            "Boletos cancelados", min_value=0, value=int(cancelados_auto), step=1,
            key=f"cierre_cancelados_{clave_widget_cierre}"
        )
        boletos_perdidos = st.number_input(
            "Boletos perdidos", min_value=0, value=int(perdidos_auto), step=1,
            key=f"cierre_perdidos_{clave_widget_cierre}"
        )

    st.markdown("#### Gastos generales y balance")
    st.write(f"**Gastos operativos ya registrados en el sistema:** ${total_gastos_registrados:,.2f} MXN")
    st.caption("Si agregas un gasto adicional aquí, quedará en este cierre diario y no se duplicará en el catálogo de gastos operativos.")
    colg1, colg2 = st.columns(2)
    with colg1:
        gasto_adicional = st.number_input(
            "Gastos generales adicionales de este cierre ($)", min_value=0.0, value=0.0, step=50.0,
            key=f"cierre_gasto_extra_{clave_widget_cierre}"
        )
    with colg2:
        detalle_gasto_adicional = st.text_input(
            "Concepto / observaciones de gastos", key=f"cierre_gasto_concepto_{clave_widget_cierre}"
        )

    total_gastos_cierre = total_gastos_registrados + float(gasto_adicional)
    balance_cierre = ingresos_confirmados_cierre - total_gastos_cierre
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Vehículos registrados ese día", len(registros_cierre))
    m2.metric("Ingresos confirmados", f"${ingresos_confirmados_cierre:,.2f}")
    m3.metric("Gastos totales", f"${total_gastos_cierre:,.2f}")
    m4.metric("Balance del día", f"${balance_cierre:,.2f}")
    st.caption("El balance se calcula con cobros confirmados menos gastos registrados y gastos adicionales del cierre; los boletos cancelados/perdidos no se cuentan como cobros salvo que exista un pago confirmado.")

    cierres_guardados = cargar_cierres_diarios()
    cierre_existente = next((c for c in cierres_guardados if c.get("fecha") == str(fecha_cierre) and c.get("sucursal") == sucursal_cierre), None)
    if cierre_existente:
        st.warning(f"Ya existe un cierre para {sucursal_cierre} el {fecha_cierre}. Marca la casilla para reemplazarlo con los datos actuales.")
    reemplazar_cierre = st.checkbox(
        "Confirmo que deseo reemplazar el cierre existente de esta sucursal y fecha",
        value=False, key=f"cierre_reemplazar_{clave_widget_cierre}", disabled=not bool(cierre_existente)
    )
    if st.button("💾 Guardar cierre diario", type="primary", key=f"guardar_cierre_diario_{clave_widget_cierre}"):
        if not folio_inicial.strip() or not folio_final.strip():
            st.error("Captura el folio inicial y el folio final de la sesión.")
        elif int(boletos_pagados) + int(boletos_cancelados) + int(boletos_perdidos) > int(total_boletos):
            st.error("La suma de boletos pagados, cancelados y perdidos no puede superar el total de boletos utilizados.")
        elif cierre_existente and not reemplazar_cierre:
            st.error("El cierre ya existe. Marca la casilla de reemplazo si realmente necesitas corregirlo.")
        else:
            cierre = {
                "id": cierre_existente.get("id") if cierre_existente else uuid.uuid4().hex[:12].upper(),
                "fecha": str(fecha_cierre),
                "sucursal": sucursal_cierre,
                "costo_unitario_boleto": float(fecha_precio),
                "folio_inicial": folio_inicial.strip(),
                "folio_final": folio_final.strip(),
                "total_boletos": int(total_boletos),
                "boletos_pagados": int(boletos_pagados),
                "boletos_cancelados": int(boletos_cancelados),
                "boletos_perdidos": int(boletos_perdidos),
                "ingresos_confirmados": round(ingresos_confirmados_cierre, 2),
                "gastos_operativos_registrados": round(total_gastos_registrados, 2),
                "gastos_generales_adicionales": round(float(gasto_adicional), 2),
                "concepto_gastos_adicionales": detalle_gasto_adicional.strip(),
                "gastos_totales": round(total_gastos_cierre, 2),
                "balance_neto": round(balance_cierre, 2),
                "usuario": st.session_state.usuario_autenticado,
                "rol": st.session_state.rol,
                "guardado_en": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            if cierre_existente:
                cierres_guardados = [cierre if c.get("id") == cierre_existente.get("id") else c for c in cierres_guardados]
            else:
                cierres_guardados.insert(0, cierre)
            try:
                guardar_cierres_diarios(cierres_guardados)
                guardar_evento_auditoria("CIERRE DIARIO POR SUCURSAL", f"{sucursal_cierre} · {fecha_cierre} · ingresos ${ingresos_confirmados_cierre:.2f} · gastos ${total_gastos_cierre:.2f} · balance ${balance_cierre:.2f}")
                st.success("Cierre diario guardado. Ya puedes descargar la hoja de cálculo para entregar el balance.")
                st.rerun()
            except OSError as exc:
                st.error(f"No se pudo guardar el cierre diario: {exc}")

    st.divider()
    st.markdown("#### Historial y entrega de hoja de cálculo")
    cierres_guardados = cargar_cierres_diarios()
    cierres_visibles = cierres_guardados if es_global else [c for c in cierres_guardados if c.get("sucursal") == st.session_state.sucursal_asignada]
    if cierres_visibles:
        df_cierres = pd.DataFrame(cierres_visibles)
        columnas_mostrar_cierres = [
            "fecha", "sucursal", "costo_unitario_boleto", "folio_inicial", "folio_final",
            "total_boletos", "boletos_pagados", "boletos_cancelados", "boletos_perdidos",
            "ingresos_confirmados", "gastos_totales", "balance_neto", "usuario", "guardado_en"
        ]
        st.dataframe(df_cierres.reindex(columns=columnas_mostrar_cierres), use_container_width=True, hide_index=True)
        opciones_cierre = {
            f"{c.get('fecha')} · {c.get('sucursal')} · Balance ${dinero(c.get('balance_neto')):,.2f}": c
            for c in cierres_visibles
        }
        etiqueta_cierre = st.selectbox("Selecciona el cierre que vas a entregar", list(opciones_cierre.keys()), key="seleccionar_cierre_descarga")
        cierre_descarga = opciones_cierre[etiqueta_cierre]
        fecha_exportacion = fecha_a_date(cierre_descarga.get("fecha"))
        registros_excel_cierre = [
            r for r in st.session_state.vehiculos
            if r.get("sucursal") == cierre_descarga.get("sucursal")
            and fecha_a_date(r.get("fecha")) == fecha_exportacion
        ]
        gastos_excel_cierre = []
        if GASTOS_CIERRE_FILE.exists():
            try:
                gastos_archivo_excel = json.loads(GASTOS_CIERRE_FILE.read_text(encoding="utf-8"))
                if isinstance(gastos_archivo_excel, list):
                    gastos_excel_cierre = [
                        g for g in gastos_archivo_excel
                        if g.get("sucursal") == cierre_descarga.get("sucursal")
                        and fecha_a_date(g.get("fecha")) == fecha_exportacion
                    ]
            except (OSError, json.JSONDecodeError):
                gastos_excel_cierre = []
        salida_excel = io.BytesIO()
        df_resumen_cierre = pd.DataFrame([{
            "Fecha": cierre_descarga.get("fecha"),
            "Sucursal": cierre_descarga.get("sucursal"),
            "Costo unitario boleto ($)": cierre_descarga.get("costo_unitario_boleto"),
            "Folio inicial": cierre_descarga.get("folio_inicial"),
            "Folio final": cierre_descarga.get("folio_final"),
            "Total boletos": cierre_descarga.get("total_boletos"),
            "Boletos pagados": cierre_descarga.get("boletos_pagados"),
            "Boletos cancelados": cierre_descarga.get("boletos_cancelados"),
            "Boletos perdidos": cierre_descarga.get("boletos_perdidos"),
            "Ingresos confirmados ($)": cierre_descarga.get("ingresos_confirmados"),
            "Gastos operativos registrados ($)": cierre_descarga.get("gastos_operativos_registrados"),
            "Gastos generales adicionales ($)": cierre_descarga.get("gastos_generales_adicionales"),
            "Concepto gastos adicionales": cierre_descarga.get("concepto_gastos_adicionales"),
            "Gastos totales ($)": cierre_descarga.get("gastos_totales"),
            "Balance neto ($)": cierre_descarga.get("balance_neto"),
            "Elaboró": cierre_descarga.get("usuario"),
            "Rol": cierre_descarga.get("rol"),
            "Fecha/hora de elaboración": cierre_descarga.get("guardado_en"),
        }])
        try:
            with pd.ExcelWriter(salida_excel, engine="openpyxl") as writer:
                df_resumen_cierre.to_excel(writer, sheet_name="Cierre diario", index=False, startrow=2)
                ws = writer.sheets["Cierre diario"]
                ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_resumen_cierre.columns))
                ws.cell(row=1, column=1, value="CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V. — CIERRE DIARIO")
                ws.cell(row=1, column=1).font = openpyxl.styles.Font(bold=True, size=14, color="FFFFFF")
                ws.cell(row=1, column=1).fill = openpyxl.styles.PatternFill("solid", fgColor="1F4E78")
                ws.freeze_panes = "A4"
                ws.auto_filter.ref = f"A3:{openpyxl.utils.get_column_letter(len(df_resumen_cierre.columns))}4"
                for col_cells in ws.columns:
                    max_len = max(len(str(cell.value or "")) for cell in col_cells)
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col_cells[0].column)].width = min(max(max_len + 2, 14), 36)
                ws.append([])
                ws.append(["Detalle de gastos operativos registrados en la sucursal y fecha"])
                ws.append(["Categoría", "Concepto", "Importe ($)", "Registrado por"])
                for gasto in gastos_excel_cierre:
                    ws.append([gasto.get("categoria", ""), gasto.get("concepto", ""), dinero(gasto.get("importe")), gasto.get("registrado_por", "")])
                ws2 = writer.book.create_sheet("Boletos del día")
                ws2.append(["Folio", "Fecha de entrada", "Sucursal", "Tarifa ($)", "Estado", "Pagado ($)"])
                for reg in registros_excel_cierre:
                    pagado_reg = sum(dinero(p.get("importe")) for p in reg.get("pagos", []) if p.get("estado") == "Pagado")
                    if not reg.get("pagos") and reg.get("pago_estado") == "Pagado":
                        pagado_reg = dinero(reg.get("importe_cobrado", reg.get("tarifa", 0)))
                    ws2.append([reg.get("id", ""), reg.get("fecha", ""), reg.get("sucursal", ""), dinero(reg.get("tarifa")), reg.get("estado", ""), pagado_reg])
                ws2.freeze_panes = "A2"
                ws2.auto_filter.ref = f"A1:F{max(ws2.max_row, 1)}"
                for col in ws2.columns:
                    max_len = max(len(str(cell.value or "")) for cell in col)
                    ws2.column_dimensions[openpyxl.utils.get_column_letter(col[0].column)].width = min(max(max_len + 2, 14), 34)
            st.download_button(
                "📥 Descargar cierre diario en Excel (.xlsx)",
                data=salida_excel.getvalue(),
                file_name=f"ClubPark_Cierre_{cierre_descarga.get('sucursal','Sucursal').replace(' ','_')}_{cierre_descarga.get('fecha','')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="descargar_excel_cierre_diario",
            )
            st.download_button(
                "📄 Descargar cierre diario en CSV",
                data=df_resumen_cierre.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"ClubPark_Cierre_{cierre_descarga.get('fecha','')}.csv",
                mime="text/csv",
                key="descargar_csv_cierre_diario",
            )
        except Exception as exc:
            st.error(f"No se pudo preparar el Excel del cierre diario: {exc}")
            st.info("Confirma que openpyxl esté instalado en requirements.txt.")
    else:
        st.info("Todavía no hay cierres diarios guardados para esta sucursal.")


# ---------------------------------------------------------------------
# TAB 3: REIMPRESIÓN
# ---------------------------------------------------------------------
with tab_imp:
    st.subheader("Re-impresión de tickets")
    visibles = vehiculos_visibles(st.session_state.vehiculos)
    if not visibles:
        st.info("No hay tickets disponibles.")
    else:
        folio_imp = st.selectbox(
            "Selecciona folio a imprimir",
            [r.get("id", "") for r in visibles],
            key="folio_reimpresion",
        )
        registro_imp = next((r for r in visibles if r.get("id") == folio_imp), None)
        if registro_imp:
            ticket = registro_imp.get("ticket_texto") or crear_ticket(registro_imp)
            st.markdown("**Las dos impresiones son independientes.**")
            vista_impresion_con_qr(registro_imp, "cliente", "imprimir_cliente_reimp")
            vista_impresion_con_qr(registro_imp, "llavero", "imprimir_contrasena_reimp")
            imprimir_partes_ticket(registro_imp, "reimpresion")
            st.download_button(
                "📄 Descargar ticket completo TXT",
                data=ticket.encode("utf-8"),
                file_name=f"ticket_{folio_imp}.txt",
                mime="text/plain",
                use_container_width=True,
            )
            if QR_OK:
                qr_data = generar_qr_bytes(registro_imp)
                st.image(qr_data, caption="QR único para localizar el vehículo", width=160)
                st.download_button(
                    "⬇️ Descargar QR del vehículo",
                    data=qr_data,
                    file_name=f"QR_{folio_imp}.png",
                    mime="image/png",
                    key=f"qr_reimp_{folio_imp}",
                )


# ---------------------------------------------------------------------
# TAB 4: REPORTES Y EXCEL
# ---------------------------------------------------------------------
with tab_rep:
    st.subheader("Reporte operativo y descarga Excel")
    visibles = vehiculos_visibles(st.session_state.vehiculos)
    if not visibles:
        st.info("No hay registros para generar un reporte.")
    else:
        df_rep = pd.DataFrame(visibles)
        if "fecha" in df_rep.columns:
            fechas = df_rep["fecha"].fillna("").astype(str)
            opciones_sucursal = ["Todas"] + sorted(df_rep["sucursal"].dropna().astype(str).unique().tolist())
            filtro_sucursal = st.selectbox("Filtrar por sucursal", opciones_sucursal, key="rep_sucursal")
            if filtro_sucursal != "Todas":
                df_rep = df_rep[df_rep["sucursal"] == filtro_sucursal]
            st.metric("Vehículos en el reporte", len(df_rep))
            ingresos = pd.to_numeric(df_rep.get("tarifa", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
            st.metric("Suma de tarifas registradas", f"${ingresos:,.2f} MXN")
            st.dataframe(
                df_rep.reindex(columns=["id", "fecha", "sucursal", "marca", "placas", "tarifa", "operador", "estado"]),
                use_container_width=True,
                hide_index=True,
            )
            try:
                excel_data = generar_excel_oficial(df_rep)
                st.download_button(
                    "📥 Descargar Excel Oficial",
                    data=excel_data.getvalue(),
                    file_name=f"Club_Park_Reporte_{datetime.now():%Y%m%d_%H%M}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as exc:
                st.error(f"No se pudo generar el Excel: {exc}")
                st.info("Verifica requirements.txt y vuelve a desplegar la aplicación.")
            csv_data = df_rep.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "📥 Descargar respaldo CSV",
                data=csv_data,
                file_name=f"Club_Park_Reporte_{datetime.now():%Y%m%d_%H%M}.csv",
                mime="text/csv",
                use_container_width=True,
            )


# ---------------------------------------------------------------------
# TAB 5: AUDITORÍA Y WHATSAPP
# ---------------------------------------------------------------------
with tab_aud:
    st.subheader("Auditoría y comunicación")
    visibles = vehiculos_visibles(st.session_state.vehiculos)
    total = len(visibles)
    entregados = sum(1 for r in visibles if r.get("estado") == "Entregado")
    resguardo = sum(1 for r in visibles if r.get("estado", "En Resguardo") == "En Resguardo")
    c1, c2, c3 = st.columns(3)
    c1.metric("Registros visibles", total)
    c2.metric("En resguardo", resguardo)
    c3.metric("Entregados", entregados)

    if visibles:
        st.markdown("#### Revisión de registros")
        df_aud = pd.DataFrame(visibles)
        sucursales_aud = ["Todas"] + sorted(df_aud["sucursal"].dropna().astype(str).unique().tolist())
        suc_aud = st.selectbox("Sucursal para auditar", sucursales_aud, key="aud_sucursal")
        if suc_aud != "Todas":
            df_aud = df_aud[df_aud["sucursal"] == suc_aud]
        st.dataframe(
            df_aud.reindex(columns=["id", "fecha", "sucursal", "marca", "placas", "tarifa", "operador", "estado"]),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("#### Validación de las cuatro fotografías")
        for registro in visibles[:100]:
            faltan = []
            for vista in ("frente", "izquierda", "derecha", "trasera"):
                if not ruta_foto(registro, vista):
                    faltan.append(vista)
            if faltan:
                st.warning(f"{registro.get('id')} — faltan/no están disponibles: {', '.join(faltan)}")

    st.divider()
    st.markdown("#### Enviar recordatorio por WhatsApp")
    st.caption("El enlace abre WhatsApp con el texto preparado; debes revisar y enviarlo manualmente.")
    telefono = st.text_input("Teléfono con clave de país o celular mexicano de 10 dígitos", key="wa_telefono")
    mensaje = st.text_area(
        "Mensaje",
        value="Buen día. Favor de revisar y completar los registros operativos pendientes de CLUB PARK.",
        key="wa_mensaje",
    )
    if telefono.strip():
        st.link_button(
            "💬 Abrir WhatsApp con mensaje",
            enlace_whatsapp(telefono, mensaje),
            use_container_width=True,
        )

    if es_contador:
        st.divider()
        st.markdown("#### Respaldo independiente de cobros")
        st.caption("Cada cobro guardado queda también en pagos_registrados.json, separado de la ficha del vehículo. Descarga una copia periódicamente y guárdala fuera del servidor.")
        pagos_respaldo = cargar_pagos_respaldo()
        st.metric("Cobros en libro de respaldo", len(pagos_respaldo))
        if pagos_respaldo:
            df_pagos_respaldo = pd.DataFrame(pagos_respaldo)
            st.dataframe(df_pagos_respaldo, use_container_width=True, hide_index=True)
            st.download_button(
                "⬇️ Descargar respaldo de cobros (JSON)",
                data=json.dumps(pagos_respaldo, ensure_ascii=False, indent=2).encode("utf-8"),
                file_name=f"ClubPark_Respaldo_Cobros_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                mime="application/json",
                key="descargar_respaldo_cobros_json",
            )
            st.download_button(
                "📊 Descargar libro de cobros (CSV)",
                data=df_pagos_respaldo.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"ClubPark_Cobros_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                key="descargar_respaldo_cobros_csv",
            )
        else:
            st.info("Todavía no hay cobros en el libro de respaldo. Los siguientes cobros se registrarán aquí.")

        st.divider()
        st.markdown("#### Conteo y consecutivo de folios por sucursal")
        st.caption("Cada sucursal tiene un prefijo exclusivo. El consecutivo continúa al día siguiente y no se reinicia al cerrar caja.")
        config_folios = cargar_config()
        resumen_folios = []
        for sucursal_folio_resumen in SUCURSALES:
            maximo_folio_resumen = siguiente_numero_usado(sucursal_folio_resumen, st.session_state.vehiculos)
            siguiente_configurado_resumen = int(config_folios["siguiente_folio"].get(sucursal_folio_resumen, 1))
            resumen_folios.append({
                "Sucursal": sucursal_folio_resumen,
                "Prefijo exclusivo": f"CP-{codigo_sucursal(sucursal_folio_resumen)}-",
                "Último consecutivo utilizado": maximo_folio_resumen,
                "Siguiente consecutivo": max(siguiente_configurado_resumen, maximo_folio_resumen + 1),
                "Vehículos registrados": sum(1 for reg in st.session_state.vehiculos if reg.get("sucursal") == sucursal_folio_resumen),
            })
        st.dataframe(pd.DataFrame(resumen_folios), use_container_width=True, hide_index=True)
        st.markdown("#### Configuración de folios por sucursal")
        st.caption("Define el siguiente consecutivo. El sistema nunca bajará del máximo ya utilizado en esa sucursal para evitar reutilizar folios.")
        suc_folio_admin = st.selectbox("Sucursal para configurar folio", SUCURSALES, key="admin_sucursal_folio")
        valor_actual = int(config_folios["siguiente_folio"].get(suc_folio_admin, 1))
        st.write(f"**Siguiente folio configurado:** {valor_actual}")
        nuevo_inicio = st.number_input(
            "Número del siguiente folio",
            min_value=1,
            value=max(1, valor_actual),
            step=1,
            key="admin_nuevo_inicio_folio",
        )
        if st.button("💾 Guardar siguiente folio", type="primary", key="guardar_folio_sucursal"):
            maximo_usado = siguiente_numero_usado(suc_folio_admin, st.session_state.vehiculos)
            if int(nuevo_inicio) <= maximo_usado:
                st.error(f"No se puede iniciar en {int(nuevo_inicio)}: ya existe un folio de esta sucursal hasta {maximo_usado}. Usa al menos {maximo_usado + 1}.")
            else:
                anterior = int(config_folios["siguiente_folio"].get(suc_folio_admin, 1))
                config_folios["siguiente_folio"][suc_folio_admin] = int(nuevo_inicio)
                config_folios.setdefault("historial_folios", []).insert(0, {
                    "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "usuario": st.session_state.usuario_autenticado,
                    "sucursal": suc_folio_admin,
                    "anterior": anterior,
                    "nuevo": int(nuevo_inicio),
                })
                try:
                    guardar_config(config_folios)
                    guardar_evento_auditoria("CAMBIO DE FOLIO", f"{suc_folio_admin}: {anterior} -> {int(nuevo_inicio)}")
                    st.success("Siguiente folio actualizado y cambio registrado en auditoría.")
                    st.rerun()
                except OSError as exc:
                    st.error(f"No se pudo guardar la configuración: {exc}")

        st.markdown("#### Historial de cambios de folios")
        historial_folios = cargar_config().get("historial_folios", [])
        if historial_folios:
            st.dataframe(pd.DataFrame(historial_folios), use_container_width=True, hide_index=True)
        else:
            st.info("Todavía no hay cambios manuales de folio.")

        st.markdown("#### Reinicio administrativo de datos operativos")
        st.warning(
            "Esta herramienta es exclusiva del Contador Auditor. Antes de reiniciar, descarga el respaldo. "
            "El historial de auditoría se conserva y los consecutivos no se reinician salvo que marques la opción correspondiente."
        )
        cierres_file_admin = DATA_DIR / "cierres_diarios_sucursal.json"
        gastos_file_admin = DATA_DIR / "gastos_operativos.json"
        cortes_file_admin = DATA_DIR / "cortes_caja.json"
        config_admin_backup = cargar_config()
        try:
            cierres_admin_backup = json.loads(cierres_file_admin.read_text(encoding="utf-8")) if cierres_file_admin.exists() else []
        except (OSError, json.JSONDecodeError):
            cierres_admin_backup = []
        try:
            gastos_admin_backup = json.loads(gastos_file_admin.read_text(encoding="utf-8")) if gastos_file_admin.exists() else []
        except (OSError, json.JSONDecodeError):
            gastos_admin_backup = []
        try:
            cortes_admin_backup = json.loads(cortes_file_admin.read_text(encoding="utf-8")) if cortes_file_admin.exists() else []
        except (OSError, json.JSONDecodeError):
            cortes_admin_backup = []
        respaldo_admin = {
            "generado_en": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "vehiculos": st.session_state.vehiculos,
            "cierres_diarios": cierres_admin_backup,
            "gastos_operativos": gastos_admin_backup,
            "cortes_caja": cortes_admin_backup,
            "configuracion_folios": config_admin_backup,
        }
        st.download_button(
            "📦 Descargar respaldo antes de reiniciar",
            data=json.dumps(respaldo_admin, ensure_ascii=False, indent=2).encode("utf-8"),
            file_name=f"respaldo_clubpark_{datetime.now():%Y%m%d_%H%M%S}.json",
            mime="application/json",
            key="descargar_respaldo_antes_reinicio",
        )
        alcance_reinicio = st.radio(
            "Alcance del reinicio",
            ["Solo una sucursal", "Toda la aplicación"],
            horizontal=True,
            key="alcance_reinicio_admin",
        )
        sucursal_reinicio = st.selectbox(
            "Sucursal que se reiniciará",
            SUCURSALES,
            key="sucursal_reinicio_admin",
            disabled=alcance_reinicio == "Toda la aplicación",
        )
        reiniciar_folios = st.checkbox(
            "También reiniciar el consecutivo de folios a 1 (puede reutilizar números históricos)",
            value=False,
            key="reiniciar_folios_admin",
        )
        texto_confirmacion = "REINICIAR TODO" if alcance_reinicio == "Toda la aplicación" else f"REINICIAR {sucursal_reinicio}"
        
        def normalizar_confirmacion(txt):
            import unicodedata
            t = unicodedata.normalize('NFKD', str(txt or '')).encode('ASCII', 'ignore').decode('ASCII')
            return re.sub(r'[^A-Z0-9]', '', t.upper())

        col_conf1, col_conf2 = st.columns([2, 1])
        with col_conf1:
            confirmacion_reinicio = st.text_input(
                f'Para confirmar, escribe: {texto_confirmacion}',
                key="texto_confirmacion_reinicio_admin",
                placeholder=texto_confirmacion
            )
        with col_conf2:
            st.write("")
            st.write("")
            check_autorizar_directo = st.checkbox(
                "✅ Autorizar reinicio",
                key="check_autorizar_reinicio_directo",
                help="Activa esta casilla para habilitar el botón de reinicio sin preocuparte por acentos o mayúsculas."
            )

        coincide_texto = normalizar_confirmacion(confirmacion_reinicio) == normalizar_confirmacion(texto_confirmacion)
        puede_reiniciar = coincide_texto or check_autorizar_directo

        if not puede_reiniciar:
            st.info(f"💡 Escribe exactamente `{texto_confirmacion}` o marca la casilla 'Autorizar reinicio' para activar el botón.")

        if st.button(
            "🧹 Reiniciar datos operativos seleccionados",
            type="primary" if puede_reiniciar else "secondary",
            disabled=not puede_reiniciar,
            key="ejecutar_reinicio_admin",
        ):
            reiniciar_todo = alcance_reinicio == "Toda la aplicación"
            sucursales_objetivo = set(SUCURSALES if reiniciar_todo else [sucursal_reinicio])
            registros_borrados = [r for r in st.session_state.vehiculos if r.get("sucursal") in sucursales_objetivo]
            registros_conservar = [r for r in st.session_state.vehiculos if r.get("sucursal") not in sucursales_objetivo]
            fotos_borradas = 0
            for registro_borrado in registros_borrados:
                rutas = []
                for vista in ("frente", "izquierda", "derecha", "trasera"):
                    ruta = ruta_foto(registro_borrado, vista)
                    if ruta:
                        rutas.append(ruta)
                for ruta_extra in registro_borrado.get("fotos_custodia", []) or []:
                    try:
                        ruta = Path(ruta_extra)
                        if not ruta.is_absolute():
                            ruta = DATA_DIR / ruta
                        rutas.append(ruta)
                    except (TypeError, ValueError):
                        pass
                for ruta in rutas:
                    try:
                        if ruta.exists() and ruta.is_file():
                            ruta.unlink()
                            fotos_borradas += 1
                    except OSError:
                        pass
            if guardar_bd(registros_conservar):
                st.session_state.vehiculos = registros_conservar
                for archivo_datos, lista_datos in (
                    (cierres_file_admin, cierres_admin_backup),
                    (gastos_file_admin, gastos_admin_backup),
                    (cortes_file_admin, cortes_admin_backup),
                ):
                    if reiniciar_todo:
                        filtrados = []
                    else:
                        filtrados = [d for d in lista_datos if d.get("sucursal") not in sucursales_objetivo]
                    try:
                        archivo_datos.write_text(json.dumps(filtrados, ensure_ascii=False, indent=2), encoding="utf-8")
                    except OSError as exc:
                        st.error(f"No se pudo actualizar {archivo_datos.name}: {exc}")
                if reiniciar_folios:
                    config_nueva = cargar_config()
                    for suc in sucursales_objetivo:
                        anterior = int(config_nueva["siguiente_folio"].get(suc, 1))
                        config_nueva["siguiente_folio"][suc] = 1
                        config_nueva.setdefault("historial_folios", []).insert(0, {
                            "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "usuario": st.session_state.usuario_autenticado,
                            "sucursal": suc,
                            "anterior": anterior,
                            "nuevo": 1,
                            "motivo": "Reinicio administrativo de datos",
                        })
                    guardar_config(config_nueva)
                guardar_evento_auditoria(
                    "REINICIO ADMINISTRATIVO",
                    f"Alcance: {'todas las sucursales' if reiniciar_todo else sucursal_reinicio}; "
                    f"registros eliminados: {len(registros_borrados)}; fotos eliminadas: {fotos_borradas}; "
                    f"reinicio de folios: {'sí' if reiniciar_folios else 'no'}",
                )
                ultimo_comprobante = st.session_state.get("ultimo_comprobante_cobro", {})
                sucursal_comprobante = (ultimo_comprobante.get("registro") or {}).get("sucursal")
                if reiniciar_todo or sucursal_comprobante in sucursales_objetivo:
                    st.session_state.pop("ultimo_comprobante_cobro", None)
                ultimo_folio_sesion = st.session_state.get("ultimo_folio")
                if any(r.get("id") == ultimo_folio_sesion for r in registros_borrados):
                    st.session_state.pop("ultimo_folio", None)
                st.success(
                    f"Reinicio terminado. Registros eliminados: {len(registros_borrados)}; "
                    f"fotografías eliminadas: {fotos_borradas}. El historial de auditoría se conservó."
                )
                st.rerun()

        st.markdown("#### Administración de registros")
        st.warning("Eliminar un registro es una acción permanente. Descarga un reporte antes de continuar.")
        ids = [r.get("id", "") for r in st.session_state.vehiculos]
        if ids:
            id_borrar = st.selectbox("Folio para eliminar", ids, key="aud_borrar_id")
            confirmar = st.checkbox("Confirmo que deseo eliminar este registro y sus fotografías.")
            if st.button("🗑️ Eliminar registro seleccionado", type="secondary", disabled=not confirmar):
                registro = next((r for r in st.session_state.vehiculos if r.get("id") == id_borrar), None)
                if registro:
                    for vista in ("frente", "izquierda", "derecha", "trasera"):
                        p = ruta_foto(registro, vista)
                        if p:
                            try:
                                p.unlink()
                            except OSError:
                                pass
                    st.session_state.vehiculos = [
                        r for r in st.session_state.vehiculos if r.get("id") != id_borrar
                    ]
                    if guardar_bd(st.session_state.vehiculos):
                        st.success(f"Registro {id_borrar} eliminado.")
                        st.rerun()

st.sidebar.caption(
    "Datos guardados en: " + str(DATA_DIR) +
    "\n\nEn Streamlit Cloud, configura almacenamiento persistente externo para evitar pérdidas."
)


# ---------------------------------------------------------------------
# REGISTRO DE AUDITORÍA (eventos recientes)
# ---------------------------------------------------------------------
if es_contador:
    st.divider()
    st.subheader("Historial de auditoría del sistema")
    eventos = cargar_auditoria()
    if eventos:
        st.dataframe(pd.DataFrame(eventos), use_container_width=True, hide_index=True)
    else:
        st.info("Los cambios de folios, salidas y cortes aparecerán aquí.")


# ---------------------------------------------------------------------
# TAB 10: CREDENCIALES Y CÓDIGOS QR INSTITUCIONALES
# ---------------------------------------------------------------------
with tab_qrs:
    if es_contador:
        st.markdown(f"""
        <div class="clubpark-header-card">
          {logo_clubpark_html(170, "0 0 12px 0")}
          <div class="clubpark-gold-badge" style="margin-top:4px">CATÁLOGO OFICIAL DE ACCESOS Y CREDENCIALES</div>
          <h3 style="color:#FFFFFF !important;margin:6px 0 2px 0">CÓDIGOS QR OFICIALES DE SUCURSALES Y SUPERVISIÓN</h3>
          <div style="font-size:13px;color:#EBF3EE">
            CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V. &nbsp;|&nbsp; Color Institucional: Verde Botella (#0B3B24)
          </div>
        </div>
        """, unsafe_allow_html=True)

        st.write(
            "A continuación se presentan todas las credenciales QR para el Contador Auditor, "
            "los Supervisores Globales y los Encargados de las 23 sucursales operativas. "
            "Puedes descargar individualmente cada código QR o descargarlos todos juntos en un archivo ZIP."
        )

        if ACCESOS_FUENTE == "por_defecto":
            st.warning(
                "⚠️ **Credenciales en modo contingencia:** los tokens por defecto todavía viven en el código fuente. "
                "Muévelos a `.streamlit/secrets.toml` (clave `CLUBPARK_TOKENS`) y gírelos de vez en cuando. "
                "Consulta la guía de implementación adjunta."
            )
        else:
            st.success(
                "✅ Credenciales protegidas: se cargan desde "
                + ("Streamlit Secrets" if ACCESOS_FUENTE == "secrets" else "variables de entorno")
                + " y no están expuestas en el código fuente."
            )

        # Paquete ZIP completo generado en memoria (sin rutas de máquina de desarrollo)
        if QR_OK and qrcode:
            buffer_zip = io.BytesIO()
            with zipfile.ZipFile(buffer_zip, "w", zipfile.ZIP_DEFLATED) as archivo_zip:
                for token_zip in ACCESOS:
                    qr_zip = qrcode.QRCode(version=1, box_size=8, border=2)
                    qr_zip.add_data(contenido_qr_para(token_zip))
                    qr_zip.make(fit=True)
                    imagen_zip = qr_zip.make_image(fill_color="#0B3B24", back_color="#FFFFFF")
                    buffer_imagen = io.BytesIO()
                    imagen_zip.save(buffer_imagen, format="PNG")
                    archivo_zip.writestr(f"QR_{token_zip}.png", buffer_imagen.getvalue())
            st.download_button(
                f"📦 Descargar paquete completo de {len(ACCESOS)} Códigos QR (ZIP)",
                data=buffer_zip.getvalue(),
                file_name="Credenciales_QR_ClubPark_Completas.zip",
                mime="application/zip",
                use_container_width=True,
                type="primary",
                key="descargar_todos_qrs_zip",
            )
        if APP_URL:
            st.info(f"Los QR incluyen el enlace de acceso directo: {APP_URL} — al escanearlos con la cámara del teléfono la sesión se abre sola.")

        filtro_rol = st.radio(
            "Filtrar catálogo:",
            ["Todos (26 credenciales)", "Auditoría y Supervisión Global", "Encargados de Sucursal"],
            horizontal=True,
            key="filtro_credenciales_qr"
        )

        # Filtrar tokens
        tokens_mostrar = {}
        for tok, data in ACCESOS.items():
            es_sup = "Global" in data.get("rol", "") or "Auditor" in data.get("rol", "")
            if filtro_rol == "Auditoría y Supervisión Global" and not es_sup:
                continue
            if filtro_rol == "Encargados de Sucursal" and es_sup:
                continue
            tokens_mostrar[tok] = data

        cols = st.columns(3)
        for idx, (tok, data) in enumerate(tokens_mostrar.items()):
            col = cols[idx % 3]
            with col:
                st.markdown(f"""
                <div style="border:2px solid #0B3B24;border-radius:10px;padding:12px;background:#FFFFFF;margin-bottom:14px;box-shadow:0 3px 8px rgba(11,59,36,0.12)">
                  <div style="background:#0B3B24;color:#FFFFFF;padding:6px 10px;border-radius:6px;font-size:12px;font-weight:700;margin-bottom:8px;text-align:center">
                    {data.get('rol', 'Supervisor')}
                  </div>
                  <div style="font-size:14px;font-weight:800;color:#0B3B24;text-align:center">{data.get('nombre', '')}</div>
                  <div style="font-size:12px;color:#C5A059;font-weight:700;text-align:center;margin-bottom:6px">SUCURSAL: {data.get('sucursal', '')}</div>
                  <div style="font-family:monospace;font-size:11px;background:#EBF3EE;padding:4px;border-radius:4px;text-align:center;word-break:break-all;color:#072617;margin-bottom:8px">
                    {tok}
                  </div>
                </div>
                """, unsafe_allow_html=True)

                if QR_OK and qrcode:
                    qr_obj = qrcode.QRCode(version=1, box_size=8, border=2)
                    qr_obj.add_data(contenido_qr_para(tok))
                    qr_obj.make(fit=True)
                    img_qr = qr_obj.make_image(fill_color="#0B3B24", back_color="#FFFFFF")
                    buf_qr = io.BytesIO()
                    img_qr.save(buf_qr, format="PNG")
                    qr_bytes = buf_qr.getvalue()
                    st.image(qr_bytes, width=160)
                    st.download_button(
                        label=f"⬇️ Descargar QR",
                        data=qr_bytes,
                        file_name=f"QR_{tok}.png",
                        mime="image/png",
                        key=f"dl_qr_tab_{tok}",
                        use_container_width=True
                    )
    else:
        st.error("🔒 Catálogo restringido: solo el Contador Auditor (máxima jerarquía) puede consultar y descargar las credenciales QR oficiales. Pide la rotación y descarga de credenciales al auditor.")


# ---------------------------------------------------------------------
# TAB: PRUEBAS OPERATIVAS
# ---------------------------------------------------------------------
with tab_pruebas:
    st.subheader("Pruebas operativas del sistema")
    st.caption(
        "Verifica en segundos que los componentes críticos funcionan antes de abrir la jornada: "
        "base de datos, folios, códigos QR, boletos térmicos, credenciales, bitácora y fotografías."
    )
    if st.button("▶️ Ejecutar pruebas operativas", type="primary", use_container_width=True, key="btn_ejecutar_pruebas"):
        resultados_pruebas = ejecutar_pruebas_operativas()
        st.session_state["resultados_pruebas"] = resultados_pruebas
    if st.session_state.get("resultados_pruebas"):
        resultados_pruebas = st.session_state["resultados_pruebas"]
        exitos = sum(1 for r in resultados_pruebas if r["estado"] == "✅")
        col_a, col_b = st.columns(2)
        col_a.metric("Pruebas aprobadas", f"{exitos}/{len(resultados_pruebas)}")
        col_b.metric("Estado general", "OPERATIVO ✅" if exitos == len(resultados_pruebas) else "REVISAR ❌")
        st.dataframe(pd.DataFrame(resultados_pruebas), use_container_width=True, hide_index=True)
        if exitos == len(resultados_pruebas):
            st.success("Todas las pruebas operativas pasaron. El sistema está listo para la jornada.")
        else:
            st.error("Hay pruebas fallidas. Atiende los detalles antes de operar.")
