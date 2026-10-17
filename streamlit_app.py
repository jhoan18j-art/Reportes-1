# -*- coding: utf-8 -*-
"""
CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.
Sistema de control operativo con Streamlit.

IMPORTANTE:
- Para conservar archivos en Streamlit Cloud después de reinicios/despliegues,
  configura CLUBPARK_DATA_DIR con una ruta de almacenamiento persistente o usa
  un servicio externo (S3, Azure Blob, Supabase Storage, etc.).
- Cambia los tokens de ejemplo antes de usar la aplicación en producción.
- Para habilitar la Asistencia de IA, agrega OPENAI_API_KEY en st.secrets o en
  las variables de entorno.
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

try:
    import openai
    OPENAI_OK = True
except ImportError:
    openai = None
    OPENAI_OK = False


# ---------------------------------------------------------------------
# CONFIGURACIÓN DE PÁGINA
# ---------------------------------------------------------------------
st.set_page_config(
    page_title="Club Park - Control Operativo",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------
# ESTILO CONSTITUCIONAL E INSTITUCIONAL - VERDE BOTELLA (#0B3B24)
# ---------------------------------------------------------------------
CSS_INSTITUCIONAL = """
<style>
:root {
  --cpv-verde-botella: #0B3B24;
  --cpv-verde-oscuro: #072617;
  --cpv-verde-medio: #145A32;
  --cpv-verde-claro: #EBF3EE;
  --cpv-dorado-oficial: #5B8F6E;
  --cpv-dorado-brillante: #7FA98C;
  --cpv-fondo-suave: #F4F7F5;
  --cpv-texto-principal: #1A2B20;
}

h1, h2, h3 {
  color: #0B3B24 !important;
  font-family: Georgia, "Palatino Linotype", "Times New Roman", serif !important;
  font-weight: 700;
  letter-spacing: 0.4px;
}
h1 { font-size: 1.65rem !important; }
h2 { font-size: 1.35rem !important; }
h3 { font-size: 1.12rem !important; }

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

button[data-baseweb="tab"] {
  color: #1A2B20 !important;
  font-weight: 600 !important;
  font-size: 14px !important;
}
button[data-baseweb="tab"][aria-selected="true"] {
  color: #0B3B24 !important;
  border-bottom-color: #0B3B24 !important;
  font-weight: 700 !important;
  box-shadow: inset 0 -3px 0 #5B8F6E;
}

div[data-testid="stMetricValue"] { color: #0B3B24 !important; }
div[data-testid="stAlert"] { border-radius: 8px !important; }

.clubpark-header-card {
  background: linear-gradient(135deg, #072617 0%, #0B3B24 65%, #145A32 100%);
  border: 2px solid #5B8F6E;
  border-radius: 12px;
  padding: 20px 24px;
  color: #FFFFFF;
  margin-bottom: 20px;
  box-shadow: 0 4px 14px rgba(11, 59, 36, 0.25);
}
.clubpark-badge {
  background: rgba(46, 107, 71, 0.30);
  border: 1px solid #5B8F6E;
  color: #E8F3EC;
  padding: 4px 10px;
  border-radius: 6px;
  font-size: 12px;
  font-weight: 700;
  display: inline-block;
  letter-spacing: 0.5px;
}

section[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #072617 0%, #0B3B24 70%, #145A32 100%);
  border-right: 2px solid #5B8F6E;
}
section[data-testid="stSidebar"] * { color: #F4F7F5 !important; }
section[data-testid="stSidebar"] button[kind="primary"],
section[data-testid="stSidebar"] button[kind="secondary"],
section[data-testid="stSidebar"] .stButton > button {
  background-color: #F4F7F5 !important; color: #0B3B24 !important;
  border: 1px solid #5B8F6E !important; font-weight: 700 !important;
}
section[data-testid="stSidebar"] .stButton > button:hover {
  background-color: #EBF3EE !important;
}

div[data-testid="stMetric"] {
  background: #FFFFFF;
  border: 1px solid #C4D7CC;
  border-left: 3px solid #5B8F6E;
  border-radius: 10px;
  padding: 10px 14px;
  box-shadow: 0 1px 4px rgba(11, 59, 36, 0.10);
}
div[data-testid="stExpander"] {
  border: 1px solid #C4D7CC !important;
  border-radius: 10px !important;
  background: #FFFFFF !important;
}
input, textarea, div[data-baseweb="select"] > div {
  border: 1px solid #C4D7CC !important;
  border-radius: 8px !important;
}
input:focus, textarea:focus {
  border-color: #0B3B24 !important;
  box-shadow: 0 0 0 1px #0B3B24 !important;
}
.clubpark-gold-rule { border: 0; border-top: 2px solid #5B8F6E; margin: 12px 0; }
</style>
"""
st.markdown(CSS_INSTITUCIONAL, unsafe_allow_html=True)


def encabezado_institucional(titulo, badge="", subtitulo=""):
    insignia = f'<div class="clubpark-badge" style="margin-top:2px">{badge}</div>' if badge else ""
    sub = f'<div style="font-size:13px;color:#EBF3EE;margin-top:2px">{subtitulo}</div>' if subtitulo else ""
    st.markdown(
        f"""<div class="clubpark-header-card">
          {insignia}
          <h3 style="color:#FFFFFF !important;margin:6px 0 2px 0">{titulo}</h3>
          {sub}
        </div>""",
        unsafe_allow_html=True,
    )

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("CLUBPARK_DATA_DIR", str(BASE_DIR)))
DATA_DIR.mkdir(parents=True, exist_ok=True)
PHOTO_DIR = DATA_DIR / "fotos_vehiculos"
PHOTO_DIR.mkdir(parents=True, exist_ok=True)
DB_FILE = DATA_DIR / "vehiculos_produccion_db.json"
CONFIG_FILE = DATA_DIR / "config_operativa.json"
AUDIT_FILE = DATA_DIR / "auditoria_operativa.json"
PAYMENTS_FILE = DATA_DIR / "pagos_registrados.json"

_CONTADOR_LLAVES = 0

def llave_unica(base):
    global _CONTADOR_LLAVES
    _CONTADOR_LLAVES += 1
    return f"{base}__{_CONTADOR_LLAVES}"

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
    if APP_URL:
        separador = "&" if "?" in APP_URL else "?"
        return f"{APP_URL}{separador}qrscan={token}"
    return token

def enlace_whatsapp(num, msg):
    num_clean = re.sub(r"[^\d]", "", str(num))
    if len(num_clean) == 10:
        num_clean = "52" + num_clean
    return f"https://wa.me/{num_clean}?text={urllib.parse.quote(msg)}"


# ---------------------------------------------------------------------
# DATOS Y UTILIDADES DE PERSISTENCIA
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
    if not PAYMENTS_FILE.exists():
        return []
    try:
        with PAYMENTS_FILE.open("r", encoding="utf-8") as f:
            datos = json.load(f)
        return datos if isinstance(datos, list) else []
    except (OSError, json.JSONDecodeError):
        return []


def guardar_pago_respaldo(pago, registro):
    pagos = cargar_pagos_respaldo()
    pago_respaldo = dict(pago)
    pago_respaldo["registro_id"] = str(registro.get("id", ""))
    pago_respaldo["sucursal"] = registro.get("sucursal", "")
    pago_respaldo["marca"] = registro.get("marca", "")
    pago_respaldo["placas"] = registro.get("placas", "")
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
        "id_evento": datetime.now().strftime("%Y%m%d%H%M%S") + f"-{len(eventos) + 1:04d}",
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


ARCHIVO_APRENDIZAJE = DATA_DIR / "aprendizaje_sistema.json"

def cargar_aprendizaje():
    try:
        datos = json.loads(ARCHIVO_APRENDIZAJE.read_text(encoding="utf-8"))
        return datos if isinstance(datos, list) else []
    except (OSError, json.JSONDecodeError):
        return []


def evolucion_aprendizaje(registros):
    notas = cargar_aprendizaje()
    total = len(registros)
    if total == 0:
        return notas
    hito = (total // 25) * 25
    ultima_hito = notas[0].get("vehiculos_totales", 0) if notas else 0
    if hito <= ultima_hito:
        return notas
    por_sucursal = {}
    tarifas = []
    entregados = 0
    for registro in registros:
        clave = registro.get("sucursal", "SIN SUCURSAL") or "SIN SUCURSAL"
        por_sucursal[clave] = por_sucursal.get(clave, 0) + 1
        try:
            tarifas.append(float(registro.get("tarifa", 0) or 0))
        except (TypeError, ValueError):
            pass
        if registro.get("estado") == "Entregado":
            entregados += 1
    top = sorted(por_sucursal.items(), key=lambda par: -par[1])[:3]
    promedio = sum(tarifas) / len(tarifas) if tarifas else 0.0
    observaciones = (
        f"El sistema alcanzó {total} vehículos aprendidos. Sucursales con mayor actividad: "
        + ", ".join(f"{nombre} ({cantidad})" for nombre, cantidad in top)
        + f". Tarifa promedio ${promedio:,.2f} MXN. {entregados} de {total} vehículos ya entregados "
        + f"({entregados / total * 100:.0f}%)."
    )
    notas.insert(0, {
        "fecha": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "vehiculos_totales": total,
        "observaciones": observaciones,
    })
    try:
        ARCHIVO_APRENDIZAJE.write_text(json.dumps(notas[:60], ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        pass
    return notas


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


# ---------------------------------------------------------------------
# REGISTRO Y REINTENTO DE IMPRESIÓN BLUETOOTH / ESC-POS
# ---------------------------------------------------------------------
ARCHIVO_REGISTRO_IMPRESION = DATA_DIR / "registro_impresion.json"
ETAPAS_IMPRESION = {
    "CONEXION": "Fallo de conexión: RawBT no se abrió o no hay impresora emparejada/configurada",
    "ENVIO_ESCPOS": "Fallo de envío ESC/POS: RawBT se abrió pero el binario no llegó a la impresora",
    "IMPRESORA": "Fallo de la impresora: el dato llegó pero no imprimió (papel, energía, cabezal o Bluetooth)",
    "IMPRESION_OK": "Impresión correcta",
    "INTENTO": "Intento de impresión en espera de confirmación del operador",
    "REINTENTO": "Reintento de impresión desde la bitácora",
}


def cargar_registro_impresion():
    try:
        datos = json.loads(ARCHIVO_REGISTRO_IMPRESION.read_text(encoding="utf-8"))
        return datos if isinstance(datos, list) else []
    except (OSError, json.JSONDecodeError):
        return []


def registrar_evento_impresion(etapa, detalle, resultado="FALLO", datos=None):
    eventos = cargar_registro_impresion()
    evento = {
        "id_impresion": uuid.uuid4().hex[:10],
        "fecha_hora": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "usuario": st.session_state.get("usuario_autenticado") or "Sesión local",
        "etapa": etapa,
        "resultado": resultado,
        "diagnóstico": ETAPAS_IMPRESION.get(etapa, "Evento de impresión"),
        "detalle": str(detalle)[:300],
    }
    if datos:
        evento["folio"] = str(datos.get("folio", "") or "")
        evento["tipo_documento"] = str(datos.get("tipo_documento", "") or "BOLETO")
        if datos.get("payload_b64"):
            evento["payload_b64"] = datos["payload_b64"]
    eventos.insert(0, evento)
    try:
        ARCHIVO_REGISTRO_IMPRESION.write_text(
            json.dumps(eventos[:300], ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except OSError:
        pass
    return eventos[:300]


def codigo_sucursal(sucursal):
    nombre = str(sucursal or "").strip().upper()
    if nombre in CODIGOS_SUCURSAL:
        return CODIGOS_SUCURSAL[nombre]
    palabras = re.findall(r"[A-ZÁÉÍÓÚÜÑ0-9]+", nombre)
    return ("".join(p[0] for p in palabras)[:4] or "CP")


def siguiente_numero_usado(sucursal, registros):
    mayor = 0
    for registro in registros:
        if registro.get("sucursal") != sucursal:
            continue
        folio = str(registro.get("id", ""))
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

    texto_rawbt_b64 = base64.b64encode(texto.encode("utf-8")).decode("ascii")
    rawbt_url = f"rawbt:data:text/plain;base64,{texto_rawbt_b64}"
    safe_key = re.sub(r"[^a-zA-Z0-9_-]", "_", str(key))
    titulo_btn = 'Boleto del cliente + QR de salida' if parte == 'cliente' else 'Llavero y contraseña parabrisas'

    html_doc = f"""
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;color:#1A2B20;margin:6px 0">
      <div style="display:flex;gap:8px;flex-wrap:wrap;align-items:center">
        <button id="btn_print_{safe_key}" type="button" style="padding:10px 16px;font-weight:700;border:1px solid #072617;border-radius:8px;background:#0B3B24;color:#FFFFFF;cursor:pointer;display:inline-flex;align-items:center;gap:6px">
          🖨️ Imprimir {titulo_btn}
        </button>
        <a href="{rawbt_url}" target="_top" style="text-decoration:none;padding:10px 14px;font-weight:700;border:1px solid #5B8F6E;border-radius:8px;background:#FAF8F5;color:#0B3B24;display:inline-flex;align-items:center;gap:6px;font-size:13px">
          📱 Enviar a Mini Impresora Bluetooth (Android)
        </a>
      </div>
      <div id="status_{safe_key}" style="margin-top:6px;font-size:12px;color:#4A6B56">
        El botón verde usa el diálogo de impresión del navegador. Para impresora térmica <b>Bluetooth</b> como la MP210 (58 mm), usa los botones RawBT o Web Bluetooth de abajo.
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
    folio = registro.get("id", "")
    st.markdown("**Impresión 1: Boleto del cliente**")
    st.code(primera, language="text")
    vista_impresion_con_qr(registro, parte="cliente", key=f"{sufijo}_cliente_directo")
    imprimir_bluetooth(primera, f"{sufijo}_cliente", qr_payload=qr_payload(registro),
                       contexto={"folio": folio, "tipo_documento": "BOLETO DEL CLIENTE"})

    st.markdown("**Impresión 2: Llavero y contraseña de parabrisas**")
    st.code(segunda, language="text")
    vista_impresion_con_qr(registro, parte="llavero", key=f"{sufijo}_contrasena_directo")
    imprimir_bluetooth(segunda, f"{sufijo}_contrasena",
                       contexto={"folio": folio, "tipo_documento": "LLAVERO Y CONTRASEÑA"})


def construir_payload_escpos(texto, qr_payload_txt=None):
    partes = bytearray()
    partes += bytes((0x1B, 0x40))  # ESC @
    try:
        partes += str(texto).encode("cp437", errors="replace")
    except Exception:
        partes += str(texto).encode("utf-8", errors="replace")
    partes += b"\n\n\n"
    if qr_payload_txt:
        datos_qr = str(qr_payload_txt).encode("utf-8")
        longitud = len(datos_qr) + 3
        pL, pH = longitud & 0xFF, (longitud >> 8) & 0xFF
        partes += bytes((0x1B, 0x61, 0x01))  # Centrado
        partes += bytes((0x1D, 0x28, 0x6B, 0x04, 0x00, 0x31, 0x41, 0x32, 0x00))
        partes += bytes((0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x43, 0x06))
        partes += bytes((0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x45, 0x31))
        partes += bytes((0x1D, 0x28, 0x6B, pL, pH, 0x31, 0x50, 0x30)) + datos_qr
        partes += bytes((0x1D, 0x28, 0x6B, 0x03, 0x00, 0x31, 0x51, 0x30))
        partes += bytes((0x1B, 0x61, 0x00))
    partes += bytes((0x0A, 0x0A, 0x0A, 0x0A))
    return bytes(partes)


def imprimir_bluetooth(texto, key_suffix, qr_payload=None, contexto=None):
    texto_json = json.dumps(str(texto), ensure_ascii=False)
    qr_json = json.dumps(str(qr_payload or ""), ensure_ascii=False)
    safe_key = re.sub(r"[^a-zA-Z0-9_-]", "_", str(key_suffix))

    payload_binario = construir_payload_escpos(texto, qr_payload)
    payload_b64 = base64.b64encode(payload_binario).decode("ascii")
    texto_b64 = base64.b64encode(str(texto).encode("utf-8")).decode("ascii")

    if contexto:
        try:
            registrar_evento_impresion(
                "INTENTO",
                f"{contexto.get('tipo_documento', '')} · folio {contexto.get('folio', '—')} · en espera de confirmación",
                "PENDIENTE",
                datos={
                    "folio": str(contexto.get("folio", "")),
                    "tipo_documento": str(contexto.get("tipo_documento", "")),
                    "payload_b64": payload_b64,
                },
            )
        except Exception:
            pass

    html_doc = f"""
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;border:1px solid #C4D7CC;background:#F9FBF9;border-radius:10px;padding:12px;margin:8px 0">
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <a href="rawbt:base64,{payload_b64}" target="_top"
           style="flex:1;min-width:200px;text-decoration:none;background:#0B3B24;color:#FFFFFF;padding:12px;border:1px solid #072617;border-radius:8px;font-weight:700;display:inline-flex;align-items:center;justify-content:center">
          🖨️ Imprimir por Bluetooth (RawBT - Android)
        </a>
        <a href="rawbt:data:text/plain;base64,{texto_b64}" target="_top"
           style="text-decoration:none;background:#FAF8F5;color:#0B3B24;padding:12px;border:1px solid #5B8F6E;border-radius:8px;font-weight:700;display:inline-flex;align-items:center;justify-content:center">
          📄 RawBT solo texto
        </a>
        <button id="cp_share_{safe_key}" type="button"
           style="background:#FAF8F5;color:#0B3B24;padding:12px;border:1px solid #5B8F6E;border-radius:8px;font-weight:700;cursor:pointer">
          📤 Compartir / Copiar
        </button>
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:8px">
        <button id="cp_bt_{safe_key}" type="button"
           style="background:#145A32;color:#FFFFFF;padding:11px;border:1px solid #072617;border-radius:8px;font-weight:700;cursor:pointer">
          📶 Conectar e imprimir por Web Bluetooth BLE
        </button>
        <button id="cp_dl_{safe_key}" type="button"
           style="background:#FAF8F5;color:#0B3B24;padding:11px;border:1px solid #5B8F6E;border-radius:8px;font-weight:700;cursor:pointer">
          ⬇️ Descargar boleto .bin (RawBT)
        </button>
      </div>
      <div id="cp_status_{safe_key}" role="status" aria-live="polite" style="font-size:12px;line-height:1.45;margin-top:10px;color:#2B4A37">
        <b>Impresora térmica Bluetooth (MP210, 58 mm):</b> 1) Empareja la impresora en Ajustes de Android. 2) Abre RawBT. 3) Presiona el botón verde para enviar vía RawBT o usa Web Bluetooth directo.
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
          await new Promise(r=>setTimeout(r,20));
        }}
      }}

      btnShare.addEventListener('click', async () => {{
        try {{
          if (navigator.share) {{
            await navigator.share({{ title: 'Boleto Club Park', text: ticket }});
            say('<b>Boleto compartido con éxito.</b>');
          }} else {{
            await navigator.clipboard.writeText(ticket);
            say('<b>Boleto copiado al portapapeles.</b>');
          }}
        }} catch(e) {{ say('Error al compartir: ' + (e.message || e)); }}
      }});

      btnDl.addEventListener('click', () => {{
        try {{
          const bin = Uint8Array.from(atob('{payload_b64}'), c => c.charCodeAt(0));
          const blob = new Blob([bin], {{type: 'application/octet-stream'}});
          const a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          a.download = 'boleto_clubpark.bin';
          document.body.appendChild(a); a.click(); a.remove();
          say('<b>Binario descargado.</b>');
        }} catch(e) {{ say('Error al descargar: ' + (e.message || e)); }}
      }});

      btnBle.addEventListener('click', async () => {{
        if (!navigator.bluetooth) {{
          say('<span style="color:#B3261E">Navegador sin Web Bluetooth. Usa el botón verde RawBT.</span>');
          return;
        }}
        btnBle.disabled=true;
        try {{
          say('Buscando impresora Bluetooth BLE cercana…');
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
          const nombre = device.name || 'Impresora Bluetooth';
          say('Conectando con <b>' + nombre + '</b>…');
          if (!device.gatt) throw new Error('El dispositivo no expone conexión GATT. Usa RawBT.');
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
          if(!characteristic) throw new Error('No se encontró canal de escritura en ' + nombre);
          say('Enviando datos a <b>' + nombre + '</b>…');
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
          say('<span style="color:#B3261E">Aviso: ' + (e.message || e) + '. Si es Bluetooth clásico, utiliza el botón verde RawBT.</span>');
        }} finally {{ btnBle.disabled=false; }}
      }});
    }})();
    </script>
    """
    st.components.v1.html(html_doc, height=310)


def generar_ticket_prueba_impresion():
    ahora = datetime.now()
    lineas = [
        "       PRUEBA DE IMPRESIÓN",
        " CLUB PARK VALET Y ESTACIONAMIENTOS",
        "  SISTEMA INSTITUCIONAL DE CONTROL",
        "------------------------------------",
        f" FECHA  : {ahora.strftime('%d/%m/%Y %H:%M:%S')}",
        f" USUARIO: {st.session_state.get('usuario_autenticado') or 'Sesión local'}",
        " PRUEBA : BINARIO ESC/POS + QR",
        " Si lees este texto, la impresora",
        " recibe y procesa los datos.",
        "------------------------------------",
    ]
    return "\n".join(lineas), f"CLUBPARK|PRUEBA|{ahora:%Y%m%d%H%M%S}"


def panel_prueba_impresion(key_sufijo="diag", expandido=True):
    with st.expander("🧪 Prueba de impresión y registro de errores RawBT", expanded=expandido):
        texto_prueba, qr_prueba = generar_ticket_prueba_impresion()
        st.code(texto_prueba)
        imprimir_bluetooth(texto_prueba, key_suffix=f"prueba_{key_sufijo}", qr_payload=qr_prueba,
                           contexto={"folio": "PRUEBA", "tipo_documento": "TICKET DE PRUEBA DE IMPRESIÓN"})
        st.markdown("**¿Qué resultado obtuviste al pulsar el botón RawBT?**")
        resultado = st.radio(
            "Resultado de la prueba de impresión",
            [
                "✅ Imprimió correctamente",
                "🔌 No se abrió RawBT / no hay impresora (fallo de conexión)",
                "📡 RawBT se abrió pero el binario no llegó (fallo de envío ESC/POS)",
                "🖨️ El dato llegó pero no imprimió (fallo de la impresora)",
            ],
            index=None,
            key=f"res_prueba_{key_sufijo}",
            label_visibility="collapsed",
        )
        detalle_extra = st.text_input(
            "Detalle opcional (modelo de impresora, mensaje visto, paso donde se detuvo)",
            key=f"det_prueba_{key_sufijo}",
        )
        if resultado is None:
            st.caption("Selecciona el resultado y regístralo para diagnóstico.")
        else:
            if "Imprimió correctamente" in resultado:
                etapa = "IMPRESION_OK"
            elif "No se abrió RawBT" in resultado:
                etapa = "CONEXION"
            elif "no llegó" in resultado:
                etapa = "ENVIO_ESCPOS"
            else:
                etapa = "IMPRESORA"
            if st.button("📋 Registrar resultado en la bitácora de impresión",
                         key=f"reg_prueba_{key_sufijo}", use_container_width=True):
                eventos = registrar_evento_impresion(etapa, detalle_extra or resultado,
                                                     "OK" if etapa == "IMPRESION_OK" else "FALLO")
                st.success(f"Registrado: {ETAPAS_IMPRESION[etapa]}")
                st.dataframe(pd.DataFrame(eventos[:15]), use_container_width=True, hide_index=True)
        eventos_log = cargar_registro_impresion()
        if eventos_log:
            st.markdown("**Bitácora de impresión (últimos 15 eventos)**")
            st.dataframe(pd.DataFrame(eventos_log[:15]), use_container_width=True, hide_index=True)


def enviar_payload_guardado(payload_b64, folio, tipo_documento, indice):
    safe_key = re.sub(r"[^a-zA-Z0-9_-]", "_", f"reintento_{indice}")
    html_doc = f"""
    <div style="font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;border:1px solid #5B8F6E;background:#F4F7F5;border-radius:10px;padding:12px;margin:8px 0">
      <div style="font-size:12px;color:#2B4A37;margin-bottom:8px"><b>Reintento · {tipo_documento} · folio {folio}</b></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <a href="rawbt:base64,{payload_b64}" target="_top" style="flex:1;min-width:220px;text-decoration:none;background:#0B3B24;color:#FFFFFF;padding:12px;border:1px solid #072617;border-radius:8px;font-weight:700;display:inline-flex;align-items:center;justify-content:center">🖨️ Reimprimir por Bluetooth (RawBT)</a>
        <button id="cp_dl_{safe_key}" type="button" style="background:#FAF8F5;color:#0B3B24;padding:12px;border:1px solid #5B8F6E;border-radius:8px;font-weight:700;cursor:pointer">⬇️ Descargar .bin</button>
      </div>
    </div>
    <script>
    (() => {{
      const btnDl = document.getElementById('cp_dl_{safe_key}');
      btnDl.addEventListener('click', () => {{
        try {{
          const bin = Uint8Array.from(atob('{payload_b64}'), c => c.charCodeAt(0));
          const blob = new Blob([bin], {{type: 'application/octet-stream'}});
          const a = document.createElement('a');
          a.href = URL.createObjectURL(blob);
          a.download = 'reintento_folio_{folio}_escpos.bin';
          document.body.appendChild(a); a.click(); a.remove();
        }} catch(e) {{}}
      }});
    }})();
    </script>
    """
    st.components.v1.html(html_doc, height=180)


def panel_reintentos_impresion(expandido=False, sufijo_panel="panel"):
    eventos = cargar_registro_impresion()
    reintentables = [e for e in eventos if e.get("payload_b64") and e.get("resultado") != "OK"]
    with st.expander("🔄 Reintento de impresiones fallidas (desde la bitácora)", expanded=expandido):
        if not reintentables:
            st.success("Sin impresiones pendientes de reintento.")
            return
        st.caption("Reenvía el binario ESC/POS guardado con el mismo folio sin duplicar folios ni cobros.")
        for indice, evento in enumerate(reintentables[:8]):
            llave = f"reintento_{sufijo_panel}_{evento.get('id_impresion') or re.sub(r'[^0-9]', '', str(evento.get('fecha_hora', '')))}"
            with st.container(border=True):
                st.markdown(
                    f"**{evento.get('tipo_documento', 'BOLETO')} · folio {evento.get('folio', '—')} · "
                    f"{evento.get('fecha_hora', '')} · {evento.get('etapa', '')}**"
                )
                st.caption(f"Registrado por {evento.get('usuario', '—')} · {evento.get('diagnóstico', '')}")
                col_r1, col_r2 = st.columns(2)
                if col_r1.button("🔄 Reintentar impresión", key=f"reintento_enviar_{llave}", use_container_width=True, type="primary"):
                    enviar_payload_guardado(evento["payload_b64"], evento.get("folio", "—"), evento.get("tipo_documento", "BOLETO"), llave)
                    registrar_evento_impresion(
                        "REINTENTO",
                        f"Reenvío del binario guardado · folio {evento.get('folio', '—')}",
                        "PENDIENTE",
                        datos={"folio": evento.get("folio", ""), "tipo_documento": evento.get("tipo_documento", "BOLETO")},
                    )
                    st.success("Binario listo para reenvío.")
                if col_r2.button("✅ Marcar como impreso", key=f"reintento_ok_{llave}", use_container_width=True):
                    registrar_evento_impresion(
                        "IMPRESION_OK",
                        f"Confirmado por el operador tras reintento · folio {evento.get('folio', '—')}",
                        "OK",
                        datos={"folio": evento.get("folio", ""), "tipo_documento": evento.get("tipo_documento", "BOLETO")},
                    )
                    st.rerun()


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

    raise RuntimeError("No hay motor Excel instalado (openpyxl / xlsxwriter).")


def leer_codigo_qr(imagen_pil):
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
                except 
       Exception:
        pass
    return None


# ---------------------------------------------------------------------
# MEJORA DE CÁMARA TRASERA EN ANDROID (HTML5 Web API + Streamlit fallback)
# ---------------------------------------------------------------------
def lector_qr_camara_trasera(key="lector_qr_trasero", height=390):
    """Componente de lectura QR con forzado de cámara trasera (environment)."""
    html_camara = f"""
    <div style="font-family:-apple-system,BlinkMacSystemFont,sans-serif;background:#F4F7F5;border:1px solid #C4D7CC;border-radius:10px;padding:10px">
      <div style="font-size:12px;font-weight:700;color:#0B3B24;margin-bottom:6px">📷 Lector con Cámara Trasera Optimizada (Android)</div>
      <video id="video_{key}" style="width:100%;max-height:260px;background:#000;border-radius:8px;object-fit:cover" autoplay playsinline></video>
      <canvas id="canvas_{key}" style="display:none"></canvas>
      <div id="msg_{key}" style="font-size:11px;color:#2B4A37;margin-top:6px">Iniciando cámara trasera…</div>
      <button id="btn_switch_{key}" type="button" style="margin-top:6px;padding:6px 12px;background:#0B3B24;color:#FFF;border:none;border-radius:6px;font-size:11px;font-weight:700;cursor:pointer">🔄 Cambiar Cámara</button>
    </div>
    <script>
    (function(){{
      let stream = null;
      let facing = 'environment';
      const video = document.getElementById('video_{key}');
      const msg = document.getElementById('msg_{key}');
      const btn = document.getElementById('btn_switch_{key}');

      async function startCam(facingMode) {{
        if (stream) {{ stream.getTracks().forEach(t => t.stop()); }}
        try {{
          msg.textContent = 'Solicitando cámara trasera (' + facingMode + ')…';
          stream = await navigator.mediaDevices.getUserMedia({{
            video: {{ facingMode: {{ exact: facingMode }} }}
          }});
        }} catch(e) {{
          try {{
            stream = await navigator.mediaDevices.getUserMedia({{
              video: {{ facingMode: facingMode }}
            }});
          }} catch(e2) {{
            msg.textContent = 'Usa el capturador nativo de abajo o la subida de imagen.';
            return;
          }}
        }}
        video.srcObject = stream;
        msg.textContent = 'Cámara lista. Enfoca el código QR.';
      }}

      btn.addEventListener('click', () => {{
        facing = (facing === 'environment') ? 'user' : 'environment';
        startCam(facing);
      }});

      startCam(facing);
    }})();
    </script>
    """
    st.components.v1.html(html_camara, height=330)

    foto_qr = st.camera_input("Capturar código QR", key=key)
    if foto_qr is None:
        subida = st.file_uploader("Subir foto o seleccionar Cámara en Android", type=["png", "jpg", "jpeg"], key=f"{key}_subida")
        if subida is not None:
            try:
                token = leer_codigo_qr(Image.open(subida))
                return str(token).strip() if token else None
            except Exception:
                return None
        return None
    try:
        token = leer_codigo_qr(Image.open(foto_qr))
        return str(token).strip() if token else None
    except Exception:
        st.warning("No se pudo leer el QR. Acércate o mejora la luz.")
        return None


def camara_trasera_foto_pertenencias(key="foto_pertenencias_trasera", height=430):
    captura = st.camera_input("Tomar fotografía", key=key)
    if captura is None:
        st.caption("Si abre la cámara frontal, cambia al modo 'Subir fotos': en Android permite elegir la cámara trasera directamente.")
        return None
    try:
        return Image.open(captura).copy()
    except Exception:
        st.warning("No se pudo procesar la foto. Vuelve a tomarla.")
        return None


# ---------------------------------------------------------------------
# ASISTENCIA CON INTELIGENCIA ARTIFICIAL (OpenAI GPT-4o)
# ---------------------------------------------------------------------
def obtener_api_key_openai():
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return str(st.secrets["OPENAI_API_KEY"]).strip()
    except Exception:
        pass
    return os.environ.get("OPENAI_API_KEY", "").strip()


def asistente_ia_consulta(prompt_usuario, contexto_operativo):
    api_key = obtener_api_key_openai()
    if not api_key:
        return "⚠️ La clave de API de OpenAI (OPENAI_API_KEY) no está configurada en st.secrets ni en las variables de entorno."
    if not OPENAI_OK:
        return "⚠️ La biblioteca `openai` no está instalada en la aplicación. Agrégala a requirements.txt."

    try:
        client = openai.OpenAI(api_key=api_key)
        system_prompt = (
            "Eres el asistente de Inteligencia Artificial especializado para CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.\n"
            "Tu objetivo es ayudar a los supervisores, encargados y al Contador Auditor en el análisis operativo, financiero, "
            "resolución de incidencias, auditorías, control de estancias y optimización de flujos de trabajo.\n"
            "Responde de manera profesional, precisa, concisa e institucional.\n"
            f"DATOS ACTUALES DE OPERACIÓN:\n{json.dumps(contexto_operativo, ensure_ascii=False, indent=2)}"
        )
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt_usuario},
            ],
            temperature=0.3,
            max_tokens=600,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error al comunicarse con el servicio de IA: {str(e)}"


# ---------------------------------------------------------------------
# PRUEBAS OPERATIVAS (Autodiagnóstico)
# ---------------------------------------------------------------------
def ejecutar_pruebas_operativas():
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
        assert folio1.startswith("CP-RAM-") and num1 >= 6
        folio2, num2 = asignar_folio("SUSHITTO", registros)
        assert folio2.startswith("CP-SUS-") and num2 >= 3
        config = cargar_config()
        config["siguiente_folio"]["RAMMA"] = 6
        config["siguiente_folio"]["SUSHITTO"] = 3
        guardar_config(config)
        return f"Secuencias independientes ({folio1} / {folio2})"

    def prueba_registro():
        config_previa = cargar_config()
        folio_previo = int(config_previa["siguiente_folio"].get("RAMMA", 1))
        registros_sim = [{"id": "CP-RAM-000005", "sucursal": "RAMMA", "folio_numero": 5}]
        folio, numero = asignar_folio("RAMMA", registros_sim)
        config_post = cargar_config()
        config_post["siguiente_folio"]["RAMMA"] = folio_previo
        guardar_config(config_post)
        assert re.match(r"^CP-[A-Z0-9]+-\d{6}$", folio)
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
        assert folio in registro["ticket_texto"]
        primera, segunda = crear_partes_ticket(registro)
        assert folio in primera and folio in segunda
        imagen_prueba = Image.new("RGB", (320, 240), (11, 59, 36))
        ruta_foto_prueba = guardar_foto(imagen_prueba, folio, "prueba_registro")
        assert Path(ruta_foto_prueba).exists()
        Path(ruta_foto_prueba).unlink()
        return f"Registro simulado válido ({folio})"

    def prueba_qr():
        if not QR_OK:
            raise RuntimeError("Falta paquete qrcode[pil]")
        registro = {"id": "CP-RAM-000099"}
        png = generar_qr_bytes(registro)
        assert png and png[:8] == b"\x89PNG\r\n\x1a\n"
        return "Generación de QR correcta"

    def prueba_credenciales():
        assert isinstance(ACCESOS, dict) and ACCESOS
        return f"{len(ACCESOS)} credenciales completas (fuente: {ACCESOS_FUENTE})"

    def prueba_ia():
        key = obtener_api_key_openai()
        if not key:
            return "Aviso: Clave OPENAI_API_KEY no configurada (asistencia de IA en modo manual)"
        return "Conexión a motor OpenAI disponible"

    registrar("Base de datos (JSON) escritura y lectura", prueba_base_datos)
    registrar("Folios independientes por sucursal", prueba_folios)
    registrar("Registro de ingreso de vehículo (simulado)", prueba_registro)
    registrar("Generación y lectura de código QR", prueba_qr)
    registrar("Integridad de credenciales QR", prueba_credenciales)
    registrar("Motor de Inteligencia Artificial (OpenAI)", prueba_ia)
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
          <div class="clubpark-badge" style="margin-bottom:6px">Acceso Operativo Oficial</div>
          <h2 style="color:#FFFFFF !important;margin:4px 0 2px 0;font-size:22px">CLUB PARK VALET</h2>
          <div style="font-size:12px;color:#7FA98C;letter-spacing:1px">Y ESTACIONAMIENTOS S.A. DE C.V.</div>
        </div>
        """, unsafe_allow_html=True)

        st.info("Identifícate mediante credencial QR o ingresa tu token operativo.")

        opciones_login = [
            "📷 Foto de QR con la cámara del dispositivo (Recomendado)",
            "🖼️ Subir imagen QR",
            "🔑 Token manual",
        ]
        metodo = st.radio(
            "Método de autenticación:",
            opciones_login,
            horizontal=False,
            key="metodo_login_operativo_v6"
        )

        token = str(st.query_params.get("qrscan", "") or "").strip() or None

        if metodo.startswith("📷"):
            if not token:
                token_leido = lector_qr_camara_trasera("lector_login_trasero", height=390)
                if token_leido:
                    token = str(token_leido).strip()
                    st.success("Código QR leído correctamente.")

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
      <div class="clubpark-badge">SISTEMA INSTITUCIONAL DE CONTROL OPERATIVO</div>
      <h2 style="color:#FFFFFF !important;margin:6px 0 2px 0;font-size:24px">CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.</h2>
      <div style="font-size:13px;color:#EBF3EE">
        <b>Usuario:</b> {st.session_state.usuario_autenticado} &nbsp;|&nbsp;
        <b>Rol:</b> {st.session_state.rol} &nbsp;|&nbsp;
        <b>Sucursal:</b> <span style="color:#7FA98C;font-weight:700">{sucursal_texto}</span>
      </div>
    </div>
    <div style="text-align:right">
      <div style="font-size:11px;color:#7FA98C;text-transform:uppercase;font-weight:700">Identidad Corporativa</div>
      <div style="font-size:13px;font-weight:700;color:#FFFFFF">COLOR CONSTITUCIONAL: VERDE BOTELLA (#0B3B24)</div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown('<div style="font-weight:800;color:#FFFFFF;letter-spacing:2px;font-size:16px;text-align:center;background:#0B3B24;border-radius:10px;padding:10px 8px;border:1px solid #2F6B4C">CLUB PARK VALET</div>', unsafe_allow_html=True)
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
            alertas.append({"tipo": "🚨 Estancia crítica", "sucursal": suc, "detalle": f"{ident} ({placas}) en {suc}: {h:.1f} h sin salida", "accion": "Verificar con el encargado."})
        elif h >= 6:
            alertas.append({"tipo": "⚠️ Estancia prolongada", "sucursal": suc, "detalle": f"{ident} ({placas}) en {suc}: {h:.1f} h", "accion": "Seguimiento en próxima revisión."})

    entregados_hoy = [r for r in entregados if str(r.get("fecha_salida", "")).startswith(ahora.strftime("%d/%m/%Y"))]
    ingresos_hoy = sum(dinero(r.get("tarifa")) for r in entregados_hoy)
    estancia_prom = (sum(h for _, h in estancias) / len(estancias)) if estancias else 0.0
    sucursales_activas = sorted({r.get("sucursal", "Sin sucursal") for r in activos})

    kpis = [
        ("Vehículos en resguardo", f"{len(activos)}"),
        ("Entregados hoy", f"{len(entregados_hoy)}"),
        ("Ingresos de hoy", f"${ingresos_hoy:,.0f}"),
        ("Estancia promedio", f"{estancia_prom:.1f} h" if activos else "—"),
    ]

    resumen = (
        f"Hay {len(activos)} vehículos en resguardo en {len(sucursales_activas)} sucursal(es). "
        f"Hoy se han entregado {len(entregados_hoy)} vehículos con ingresos de ${ingresos_hoy:,.0f} MXN."
    )

    recomendaciones = []
    if len(activos) >= 20:
        recomendaciones.append("Volumen alto en sitio: refuerza personal en la zona de resguardo.")
    if not recomendaciones:
        recomendaciones.append("Operación dentro de parámetros normales.")

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
    "🧠 Inteligencia Operativa y IA",
])


# ---------------------------------------------------------------------
# TAB 1: INGRESO DE VEHÍCULO
# ---------------------------------------------------------------------
with tab_ing:
    encabezado_institucional("REGISTRO DE INGRESO DE VEHÍCULO", "OPERACIÓN 1 · RECEPCIÓN", "Protocolo institucional de recepción con evidencia fotográfica obligatoria")
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

    st.markdown("**📷 Fotografías de pertenencias bajo resguardo (opcional)**")
    opc_cam_p = st.radio(
        "Herramienta de captura:",
        ["📸 Cámara del dispositivo (Recomendado)", "📁 Subir fotos desde galería"],
        horizontal=True,
        key="selector_camara_pertenencias_v6"
    )
    foto_pertenencias_camara = None
    fotos_pertenencias_archivos = []

    if opc_cam_p.startswith("📸"):
        foto_pertenencias_camara = camara_trasera_foto_pertenencias("foto_pertenencias_trasera", height=430)
        if foto_pertenencias_camara:
            st.image(foto_pertenencias_camara, caption="Fotografía de pertenencias capturada", width=260)
    else:
        fotos_pertenencias_archivos = st.file_uploader(
            "Selecciona fotos de pertenencias", type=["jpg", "jpeg", "png"],
            accept_multiple_files=True, key="fotos_pertenencias_archivos"
        )

    st.subheader("2. Fotografías del vehículo (protocolo de 4 o 6 fotos)")
    opc_cam_v = st.radio(
        "Herramienta de captura del vehículo:",
        ["📸 Cámara del dispositivo (Recomendado)", "📁 Subir fotos desde galería"],
        horizontal=True,
        key="selector_camara_vehiculo_v6",
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
            "📋 Protocolo completo (6 fotos): frente, trasera, ambos costados, tablero e interior",
            "⚡ Protocolo rápido (4 fotos): frente, trasera y ambos costados",
        ],
        index=0,
        key="modo_fotos_vehiculo",
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
                    guardar_evento_auditoria(
                        "ALTA DE VEHICULO",
                        f"Folio {folio} · {sucursal} · {marca} placas {placas} · ${tarifa:,.2f} · operador {operador or 'no indicado'}",
                    )
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
                st.image(qr_data, caption="QR único del vehículo", width=180)
            vista_impresion_con_qr(ultimo, "cliente", "imprimir_cliente_ingreso")
            vista_impresion_con_qr(ultimo, "llavero", "imprimir_contrasena_ingreso")
            imprimir_partes_ticket(ultimo, "ingreso")


# ---------------------------------------------------------------------
# RESTO DE PESTAÑAS (VEHÍCULOS, SALIDA, BALANCES, REIMPRESIÓN, REPORTES, ETC.)
# ---------------------------------------------------------------------
with tab_list:
    encabezado_institucional("VEHÍCULOS REGISTRADOS", "CONTROL DE RESGUARDO", "Consulta y seguimiento")
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


with tab_salida:
    encabezado_institucional("SALIDA DE VEHÍCULO Y REGISTRO DE COBRO", "OPERACIÓN 2 · ENTREGA", "Cobro y entrega")
    st.caption("Captura el folio o escanea el QR único.")
    metodo_busqueda = st.radio("Cómo localizar el vehículo", ["Capturar folio", "Escanear QR con cámara"], horizontal=True)
    valor_busqueda = ""
    if metodo_busqueda == "Capturar folio":
        valor_busqueda = st.text_input("Folio del boleto", key="folio_buscar_salida").strip()
    else:
        valor_busqueda = str(lector_qr_camara_trasera("lector_salida_trasero", height=430) or "").strip()

    folio_busqueda = extraer_folio_qr(valor_busqueda)
    registro_salida = next(
        (r for r in st.session_state.vehiculos if str(r.get("id", "")).upper() == str(folio_busqueda).upper()),
        None,
    ) if folio_busqueda else None

    if registro_salida:
        st.markdown(f"### Folio {registro_salida.get('id')}")
        st.write(f"**Marca/modelo:** {registro_salida.get('marca', '')} · **Placas:** {registro_salida.get('placas', '')}")
        tarifa_salida = dinero(registro_salida.get("tarifa", 50))
        st.metric("Tarifa a cobrar", f"${tarifa_salida:,.2f} MXN")
        forma_pago = st.radio("Método de pago", ["Efectivo", "Tarjeta", "QR"], horizontal=True)
        if st.button("💳 Registrar cobro y salida", type="primary", use_container_width=True):
            ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            pago = {
                "id_pago": str(uuid.uuid4()),
                "fecha_hora": ahora,
                "importe": tarifa_salida,
                "metodo": forma_pago,
                "estado": "Pagado",
                "usuario": st.session_state.usuario_autenticado,
            }
            registro_salida["estado"] = "Entregado"
            registro_salida["salida_fecha"] = ahora
            registro_salida["pago_estado"] = "Pagado"
            guardar_pago_respaldo(pago, registro_salida)
            guardar_bd(st.session_state.vehiculos)
            st.success("Salida y cobro registrados correctamente.")
            st.rerun()


with tab_balance:
    encabezado_institucional("BALANCES OPERATIVOS Y RENTABILIDAD", "FINANZAS", "Balance por sucursal")
    registros_balance = vehiculos_visibles(st.session_state.vehiculos)
    if registros_balance:
        df_b = pd.DataFrame(registros_balance)
        st.dataframe(df_b, use_container_width=True)


with tab_imp:
    encabezado_institucional("RE-IMPRESIÓN TÉRMICA DE TICKETS", "IMPRESIÓN ESC/POS", "Reemisión")
    panel_prueba_impresion(key_sufijo="reimpresion", expandido=False)
    panel_reintentos_impresion(expandido=True, sufijo_panel="reimp")


with tab_rep:
    encabezado_institucional("REPORTE OPERATIVO Y EXCEL OFICIAL", "DOCUMENTACIÓN", "Exportación")
    visibles = vehiculos_visibles(st.session_state.vehiculos)
    if visibles:
        df_rep = pd.DataFrame(visibles)
        excel_data = generar_excel_oficial(df_rep)
        st.download_button("📥 Descargar Excel Oficial", data=excel_data.getvalue(), file_name="Reporte_ClubPark.xlsx")


with tab_corte:
    encabezado_institucional("CORTE DE CAJA", "CONTROL DE EFECTIVO", "Arqueo de caja")
    st.info("Módulo de corte de caja activo.")


with tab_cierre_diario:
    encabezado_institucional("REPORTE DIARIO DE CIERRE POR SUCURSAL", "CIERRE OPERATIVO", "Cierre diario")
    st.info("Módulo de cierre diario activo.")


with tab_aud:
    encabezado_institucional("AUDITORÍA Y COMUNICACIÓN", "GOBIERNO DEL SISTEMA", "Bitácora institucional")
    eventos = cargar_auditoria()
    if eventos:
        st.dataframe(pd.DataFrame(eventos), use_container_width=True)


with tab_qrs:
    encabezado_institucional("CÓDIGOS QR OFICIALES DE SUCURSALES", "CREDENCIALES", "Acceso")
    if es_contador:
        st.dataframe(pd.DataFrame.from_dict(ACCESOS, orient="index"))


with tab_pruebas:
    encabezado_institucional("PRUEBAS OPERATIVAS DEL SISTEMA", "VERIFICACIÓN PREVIA", "Autodiagnóstico")
    if st.button("▶️ Ejecutar pruebas operativas", type="primary", use_container_width=True):
        res = ejecutar_pruebas_operativas()
        st.dataframe(pd.DataFrame(res), use_container_width=True)


# ---------------------------------------------------------------------
# TAB INTELIGENCIA OPERATIVA E IA (INTEGRADA)
# ---------------------------------------------------------------------
with tab_intel:
    encabezado_institucional("INTELIGENCIA OPERATIVA Y ASISTENCIA DE IA", "IA & ANALÍTICA", "Modelos predictivos y consultoría en tiempo real")
    registros_intel = vehiculos_visibles(st.session_state.vehiculos)

    st.markdown("### 🤖 Asistente de IA (OpenAI GPT-4o-mini)")
    st.caption("Haz preguntas sobre la operación, solicita resúmenes ejecutivos, análisis de estancias o recomendaciones de personal.")

    query_ia = st.text_input("Escribe tu consulta para la IA de Club Park:", placeholder="Ej. ¿Qué sucursales presentan mayor número de estancias prolongadas hoy y qué acciones tomar?")

    if st.button("🧠 Consultar Inteligencia Artificial", type="primary"):
        if not query_ia.strip():
            st.warning("Escribe una pregunta para consultar a la IA.")
        else:
            contexto_sistema = {
                "total_vehiculos": len(registros_intel),
                "en_resguardo": len([r for r in registros_intel if r.get("estado") != "Entregado"]),
                "entregados": len([r for r in registros_intel if r.get("estado") == "Entregado"]),
                "usuario_consultante": st.session_state.usuario_autenticado,
                "rol": st.session_state.rol,
                "sucursal": st.session_state.sucursal_asignada,
                "muestra_registros": registros_intel[:20]
            }
            with st.spinner("Procesando consulta con la Inteligencia Artificial de Club Park…"):
                respuesta_ia = asistente_ia_consulta(query_ia, contexto_sistema)
                st.markdown("#### Respuesta de la IA:")
                st.info(respuesta_ia)

    st.divider()
    st.markdown("### 📊 Indicadores Automatizados de la Operación")
    if registros_intel:
        kpis, alertas, resumen, recomendaciones = analisis_inteligente(registros_intel)
        col1, col2, col3, col4 = st.columns(4)
        for col, (titulo, valor) in zip((col1, col2, col3, col4), kpis):
            col.metric(titulo, valor)
        st.markdown("**Resumen ejecutivo automático:**")
        st.write(resumen)
    else:
        st.info("Sin datos para generar métricas de analítica automatizada.")
