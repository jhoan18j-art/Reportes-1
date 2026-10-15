# -*- coding: utf-8 -*-
"""
CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.
Sistema de control operativo con Streamlit.
"""

import io
import json
import os
import re
import uuid
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

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("CLUBPARK_DATA_DIR", str(BASE_DIR)))
DATA_DIR.mkdir(parents=True, exist_ok=True)
PHOTO_DIR = DATA_DIR / "fotos_vehiculos"
PHOTO_DIR.mkdir(parents=True, exist_ok=True)
CAMARA_COMPONENT = st.components.v1.declare_component(
    "clubpark_camara_trasera", path=str(BASE_DIR / "clubpark_component")
)
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
    "BARRANCA": 25,
    "CLÍNICA ASURA": 100,
    "ASURA": 100,
    "S. AZURA": 100,
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
    "CPV-QR-BARRANCA": {"nombre": "Encargado Barranca", "sucursal": "BARRANCA", "rol": "Supervisor de Sucursal"},
}
TOKENS_JSON = os.environ.get("CLUBPARK_TOKENS_JSON", "").strip()
try:
    ACCESOS = json.loads(TOKENS_JSON) if TOKENS_JSON else TOKENS_POR_DEFECTO
except json.JSONDecodeError:
    ACCESOS = TOKENS_POR_DEFECTO


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

def codigo_sucursal(sucursal):
    nombre = str(sucursal or "").strip().upper()
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

def crear_ticket(registro):
    folio = registro.get("id", "")
    sucursal = registro.get("sucursal", "")
    marca = registro.get("marca", "")
    placas = registro.get("placas", "")
    color = registro.get("color", "")
    tarifa = float(registro.get("tarifa", 0) or 0)
    entrada_dt = registro.get("fecha_entrada") or registro.get("fecha", "")
    salida_dt = registro.get("salida_fecha") or "PENDIENTE"
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

def vista_impresion_con_qr(registro, parte="cliente", key="print_qr"):
    primera, segunda = crear_partes_ticket = (lambda r: (r.get("ticket_texto"), ""))(registro)
    texto = registro.get("ticket_texto", "")
    qr_data = generar_qr_bytes(registro) if QR_OK else None
    qr_img = (
        f'<img alt="QR del vehículo" style="width:128px;height:128px;image-rendering:pixelated" src="data:image/png;base64,{base64.b64encode(qr_data).decode("ascii")}">'
        if qr_data and parte == "cliente" else ""
    )
    seguro = html.escape(texto)
    if parte == "cliente" and qr_img:
        seguro += '<br><center>' + qr_img + '</center><center><b>ESCANEE PARA REGISTRAR LA SALIDA</b></center>'
    html_doc = f"""
    <div style="font-family:monospace;color:#111">
      <button id="{key}" style="padding:10px 14px;font-weight:bold;border:0;border-radius:6px;background:#174a7e;color:white">
        🖨️ Imprimir boleto del cliente + QR de salida
      </button>
      <div id="ticket_{key}" style="display:none;width:58mm;white-space:pre-wrap;font-size:10pt">{seguro}</div>
    </div>
    <script>
      document.getElementById("{key}").addEventListener("click", function() {{
        const body = document.getElementById("ticket_{key}").innerHTML;
        const w = window.open("", "_blank", "width=420,height=650");
        if (!w) {{ alert("Permite ventanas emergentes para imprimir."); return; }}
        w.document.write('<html><head><title>Ticket CLUB PARK</title><style>@page{{size:58mm auto;margin:2mm}}body{{font-family:monospace;font-size:10pt;white-space:pre-wrap;width:54mm}}img{{display:block;margin:4px auto}}</style></head><body>' + body + '</body></html>');
        w.document.close(); w.focus();
        setTimeout(() => {{ w.focus(); w.print(); }}, 700);
      }});
    </script>
    """
    st.components.v1.html(html_doc, height=95)

def imprimir_bluetooth(texto, key_suffix, qr_payload=None):
    texto_json = json.dumps(str(texto), ensure_ascii=False)
    qr_json = json.dumps(str(qr_payload or ""), ensure_ascii=False)
    safe_key = re.sub(r"[^a-zA-Z0-9_-]", "_", str(key_suffix))
    html_doc = f"""
    <div style="font-family:Arial,sans-serif;border:1px solid #ddd;border-radius:10px;padding:12px">
      <button id="bt_{safe_key}" style="background:#174a7e;color:white;padding:12px;border:0;border-radius:8px;font-weight:bold;width:100%">
        🖨️ Conectar e imprimir por Bluetooth
      </button>
      <div id="bt_status_{safe_key}" style="font-size:13px;margin-top:8px">Listo para conectar impresora térmica Bluetooth BLE.</div>
    </div>
    <script>
    (() => {{
      const ticket = {texto_json};
      const qr = {qr_json};
      const btn = document.getElementById('bt_{safe_key}');
      const status = document.getElementById('bt_status_{safe_key}');
      const say = t => status.textContent = t;
      const enc = new TextEncoder();
      const concat = (...parts) => {{ let n=parts.reduce((a,p)=>a+p.length,0), o=new Uint8Array(n), k=0; for(const p of parts){{o.set(p,k);k+=p.length;}} return o; }};
      const cmd = (...a) => new Uint8Array(a);
      async function sendChunks(ch, data) {{
        const max = 180;
        for(let i=0;i<data.length;i+=max) {{
          const chunk=data.slice(i,Math.min(i+max,data.length));
          if(ch.writeValueWithoutResponse) await ch.writeValueWithoutResponse(chunk);
          else await ch.writeValue(chunk);
          await new Promise(r=>setTimeout(r,12));
        }}
      }}
      btn.addEventListener('click', async () => {{
        if (!navigator.bluetooth) {{ say('Web Bluetooth no soportado en este navegador.'); return; }}
        btn.disabled=true;
        try {{
          say('Selecciona la impresora térmica…');
          const device=await navigator.bluetooth.requestDevice({{
            acceptAllDevices:true,
            optionalServices:['0000ffe0-0000-1000-8000-00805f9b34fb','0000ff00-0000-1000-8000-00805f9b34fb','6e400001-b5a3-f393-e0a9-e50e24dcca9e']
          }});
          const server=await device.gatt.connect();
          let characteristic=null;
          const services=await server.getPrimaryServices();
          for(const service of services) {{
            const chars=await service.getCharacteristics();
            characteristic=chars.find(c=>c.properties.writeWithoutResponse||c.properties.write);
            if(characteristic) break;
          }}
          if(!characteristic) throw new Error('No se encontró característica de escritura BLE.');
          await sendChunks(characteristic, cmd(0x1B,0x40));
          await sendChunks(characteristic, enc.encode(ticket+'\\n\\n\\n'));
          say('Impresión completada.');
          try {{ device.gatt.disconnect(); }} catch(e) {{}}
        }} catch(e) {{
          say('Error al imprimir: ' + (e.message || e));
        }} finally {{ btn.disabled=false; }}
      }});
    }})();
    </script>
    """
    st.components.v1.html(html_doc, height=125)

def lector_qr_camara_trasera(key="lector_qr_trasero", height=390):
    return CAMARA_COMPONENT(mode="qr", key=key, height=height, default=None)

# ---------------------------------------------------------------------
# SESIÓN Y VISTA PRINCIPAL
# ---------------------------------------------------------------------
if "usuario_autenticado" not in st.session_state:
    st.session_state.usuario_autenticado = None
if "sucursal_asignada" not in st.session_state:
    st.session_state.sucursal_asignada = None
if "rol" not in st.session_state:
    st.session_state.rol = None
if "vehiculos" not in st.session_state:
    st.session_state.vehiculos = cargar_bd()

if st.session_state.usuario_autenticado is None:
    st.title("🚗 CLUB PARK - Control Operativo")
    token_input = st.text_input("Token de acceso", type="password")
    if st.button("Validar acceso", type="primary"):
        info = ACCESOS.get(token_input.strip())
        if info:
            st.session_state.usuario_autenticado = info.get("nombre")
            st.session_state.sucursal_asignada = info.get("sucursal")
            st.session_state.rol = info.get("rol")
            st.rerun()
        else:
            st.error("Token no reconocido.")
    st.stop()

st.title("CLUB PARK VALET Y ESTACIONAMIENTOS")
tab_ing, tab_salida, tab_imp = st.tabs(["📥 Ingreso", "🚘 Salida", "🖨️ Reimpresión"])

with tab_ing:
    st.subheader("Ingreso de Vehículo")
    sucursal = st.session_state.sucursal_asignada
    marca = st.text_input("Marca y modelo")
    placas = st.text_input("Placas").strip().upper()
    color = st.text_input("Color")
    if st.button("Registrar ingreso", type="primary"):
        if marca and placas:
            folio, folio_numero = asignar_folio(sucursal, st.session_state.vehiculos)
            ahora = datetime.now()
            registro = {
                "id": folio,
                "folio_numero": folio_numero,
                "fecha": ahora.strftime("%d/%m/%Y %H:%M:%S"),
                "sucursal": sucursal,
                "marca": marca,
                "placas": placas,
                "color": color,
                "tarifa": PRECIOS_SUCURSALES.get(sucursal, 50),
                "estado": "En Resguardo",
            }
            registro["ticket_texto"] = crear_ticket(registro)
            st.session_state.vehiculos.insert(0, registro)
            guardar_bd(st.session_state.vehiculos)
            st.success(f"Registrado con éxito. Folio: {folio}")
            vista_impresion_con_qr(registro, "cliente")
        else:
            st.error("Completa marca y placas.")

with tab_salida:
    st.subheader("Salida y escaneo de QR")
    valor_leido = lector_qr_camara_trasera("lector_salida", height=390)
    folio_busqueda = extraer_folio_qr(valor_leido)
    if folio_busqueda:
        st.write(f"Folio detectado: {folio_busqueda}")
        reg = next((r for r in st.session_state.vehiculos if r.get("id") == folio_busqueda), None)
        if reg:
            st.success(f"Vehículo localizado: {reg.get('marca')} - {reg.get('placas')}")
            if st.button("Confirmar Salida"):
                reg["estado"] = "Entregado"
                guardar_bd(st.session_state.vehiculos)
                st.rerun()

with tab_imp:
    st.subheader("Reimpresión")
    folios = [r.get("id") for r in st.session_state.vehiculos]
    if folios:
        f_sel = st.selectbox("Seleccionar folio", folios)
        r_sel = next((r for r in st.session_state.vehiculos if r.get("id") == f_sel), None)
        if r_sel:
            vista_impresion_con_qr(r_sel, "cliente")
            imprimir_bluetooth(r_sel.get("ticket_texto"), f"reimp_{f_sel}")
  
