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
DB_FILE = DATA_DIR / "vehiculos_produccion_db.json"
CONFIG_FILE = DATA_DIR / "config_operativa.json"
AUDIT_FILE = DATA_DIR / "auditoria_operativa.json"

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
TOKENS_JSON = os.environ.get("CLUBPARK_TOKENS_JSON", "").strip()
try:
    ACCESOS = json.loads(TOKENS_JSON) if TOKENS_JSON else TOKENS_POR_DEFECTO
except json.JSONDecodeError:
    ACCESOS = TOKENS_POR_DEFECTO

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


def codigo_sucursal(sucursal):
    palabras = re.findall(r"[A-ZÁÉÍÓÚÜÑ0-9]+", str(sucursal).upper())
    codigo = "".join(p[0] for p in palabras)[:4]
    return codigo or "CP"


def siguiente_numero_usado(sucursal, registros):
    """Obtiene el mayor consecutivo encontrado para esa sucursal, sin reutilizar folios."""
    mayor = 0
    prefijo = f"CP-{codigo_sucursal(sucursal)}-"
    for registro in registros:
        if registro.get("sucursal") != sucursal:
            continue
        folio = str(registro.get("id", ""))
        if folio.startswith(prefijo):
            try:
                mayor = max(mayor, int(folio.rsplit("-", 1)[1]))
            except (ValueError, IndexError):
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
    """Vista imprimible del boleto con QR gráfico; usa el diálogo de impresión del navegador."""
    primera, segunda = crear_partes_ticket(registro)
    texto = primera if parte == "cliente" else segunda
    qr_data = generar_qr_bytes(registro) if QR_OK else None
    qr_img = (
        f'<img alt="QR del vehículo" style="width:128px;height:128px;image-rendering:pixelated" src="data:image/png;base64,{base64.b64encode(qr_data).decode("ascii")}">'
        if qr_data and parte == "cliente" else ""
    )
    seguro = html.escape(texto)
    if parte == "cliente" and qr_img:
        # El QR queda en la parte inferior del boleto del cliente, antes de la línea de corte.
        marca_corte = "[ CORTE ]"
        if marca_corte in seguro:
            antes, despues = seguro.split(marca_corte, 1)
            seguro = antes + '<br><center>' + qr_img + '</center><center><b>ESCANEE PARA REGISTRAR LA SALIDA</b></center><br>' + marca_corte + despues
        else:
            seguro += '<br><center>' + qr_img + '</center><center><b>ESCANEE PARA REGISTRAR LA SALIDA</b></center>'
    html_doc = f"""
    <div style="font-family:monospace;color:#111">
      <button id="{key}" style="padding:10px 14px;font-weight:bold;border:0;border-radius:6px;background:#174a7e;color:white">
        🖨️ Imprimir {('boleto del cliente + QR de salida' if parte == 'cliente' else 'llavero y contraseña')}
      </button>
      <div style="margin-top:8px;font-size:12px">Para que salga el QR gráfico, usa esta impresión del navegador y selecciona la impresora térmica. La impresión Bluetooth de texto no imprime imágenes QR.</div>
      <div id="ticket_{key}" style="display:none;width:58mm;white-space:pre-wrap;font-size:10pt">{seguro}</div>
    </div>
    <script>
      document.getElementById("{key}").addEventListener("click", function() {{
        const body = document.getElementById("ticket_{key}").innerHTML;
        const w = window.open("", "_blank", "width=420,height=650");
        if (!w) {{ alert("El navegador bloqueó la ventana. Permite ventanas emergentes para imprimir."); return; }}
        w.document.write('<html><head><title>Ticket CLUB PARK</title><style>@page{{size:58mm auto;margin:2mm}}body{{font-family:monospace;font-size:10pt;white-space:pre-wrap;width:54mm}}img{{display:block;margin:4px auto}}</style></head><body>' + body + '</body></html>');
        w.document.close(); w.focus();
        setTimeout(() => {{ w.focus(); w.print(); }}, 700);
      }});
    </script>
    """
    st.components.v1.html(html_doc, height=95)


def vista_comprobante_cobro(pago, registro, key="imprimir_comprobante_cobro"):
    """Muestra un comprobante imprimible del cobro ya registrado."""
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
        f"<tr><td>{html.escape(str(etiqueta))}</td><td>{html.escape(str(valor))}</td></tr>"
        for etiqueta, valor in datos.items()
    )
    html_doc = f"""
    <div style="font-family:Arial,sans-serif;color:#111;max-width:440px">
      <button id="{key}" style="padding:10px 14px;font-weight:bold;border:0;border-radius:6px;background:#174a7e;color:white">🧾 Imprimir comprobante de cobro</button>
      <div id="contenido_{key}" style="display:none">
        <div style="width:72mm;padding:3mm;font-family:Arial,sans-serif;color:#111">
          <center><b>CLUB PARK</b><br>VALET Y ESTACIONAMIENTOS S.A. DE C.V.<br><b>COMPROBANTE DE COBRO</b></center>
          <hr><table style="width:100%;font-size:10pt;border-collapse:collapse">{filas}</table>
          <hr><center>Conserve este comprobante.<br>Gracias por su preferencia.</center>
        </div>
      </div>
      <script>
      document.getElementById("{key}").addEventListener("click", function() {{
        const contenido = document.getElementById("contenido_{key}").innerHTML;
        const w = window.open("", "_blank", "width=420,height=650");
        if (!w) {{ alert("Permite ventanas emergentes para imprimir el comprobante."); return; }}
        w.document.write('<html><head><title>Comprobante CLUB PARK</title><style>@page{{size:80mm auto;margin:3mm}}body{{font-family:Arial,sans-serif}}table td{{padding:4px 0;vertical-align:top}}table td:first-child{{font-weight:bold;width:38%}}</style></head><body>' + contenido + '</body></html>');
        w.document.close(); w.focus(); setTimeout(() => w.print(), 300);
      }});
      </script>
    </div>
    """
    st.components.v1.html(html_doc, height=70)


def imprimir_partes_ticket(registro, sufijo):
    primera, segunda = crear_partes_ticket(registro)
    st.markdown("**Impresión 1: boleto del cliente**")
    st.code(primera, language="text")
    imprimir_bluetooth(primera, f"{sufijo}_cliente")
    st.markdown("**Impresión 2: llavero y contraseña de parabrisas**")
    st.code(segunda, language="text")
    imprimir_bluetooth(segunda, f"{sufijo}_contrasena")


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
    fecha = registro.get("fecha", "")
    custodia = registro.get("custodia", "") or "Sin objetos declarados"
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
PLACA: {placas}      COLOR: {color}
ENTRADA: {fecha}    SALIDA: _______
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


def enlace_whatsapp(telefono, mensaje):
    telefono = "".join(ch for ch in str(telefono) if ch.isdigit())
    if len(telefono) == 10:
        telefono = "52" + telefono
    return f"https://wa.me/{telefono}?text={urllib.parse.quote(mensaje)}"


def imprimir_bluetooth(texto, key_suffix):
    """Impresión Bluetooth de texto; recuerda el dispositivo cuando Chrome lo permite.

    Nota: Web Bluetooth no es compatible con muchas impresoras térmicas Bluetooth Classic/SPP.
    Para tickets con QR gráfico debe usarse la impresión de navegador de arriba.
    """
    texto_json = json.dumps(texto, ensure_ascii=False)
    html_doc = f"""
    <div style="font-family:Arial,sans-serif">
      <button id="print_{key_suffix}" style="background:#22543D;color:white;padding:12px;
      border:0;border-radius:8px;font-weight:bold;width:100%">
        🖨️ Imprimir con impresora guardada
      </button>
      <button id="forget_{key_suffix}" style="margin-top:6px;padding:7px;border:1px solid #999;border-radius:6px;width:100%">
        Cambiar impresora guardada
      </button>
      <p id="status_{key_suffix}" style="font-size:12px"></p>
    </div>
    <script>
    const ticket = {texto_json};
    const btn = document.getElementById("print_{key_suffix}");
    const forgetBtn = document.getElementById("forget_{key_suffix}");
    const status = document.getElementById("status_{key_suffix}");
    const storageKey = "clubpark_bt_printer_id";
    forgetBtn.addEventListener("click", () => {{
      try {{ localStorage.removeItem(storageKey); }} catch(e) {{}}
      status.textContent = "Preferencia borrada. Pulsa imprimir y selecciona la impresora que quedará guardada si el navegador lo permite.";
    }});
    btn.addEventListener("click", async () => {{
      if (!navigator.bluetooth) {{
        status.textContent = "Este navegador no admite Web Bluetooth. En Android prueba Chrome; si la impresora es Bluetooth Classic, usa una app/driver de impresión compatible.";
        return;
      }}
      btn.disabled = true;
      try {{
        let device = null;
        let savedId = "";
        try {{ savedId = localStorage.getItem(storageKey) || ""; }} catch(e) {{}}
        if (savedId && navigator.bluetooth.getDevices) {{
          const permitted = await navigator.bluetooth.getDevices();
          device = permitted.find(d => d.id === savedId) || null;
        }}
        if (!device) {{
          status.textContent = "Primera conexión: selecciona tu impresora. Intentaremos recordarla para las siguientes impresiones.";
          device = await navigator.bluetooth.requestDevice({{
            acceptAllDevices: true,
            optionalServices: [
              "00001101-0000-1000-8000-00805f9b34fb",
              "000018f0-0000-1000-8000-00805f9b34fb"
            ]
          }});
          try {{ localStorage.setItem(storageKey, device.id); }} catch(e) {{}}
        }}
        if (!device.gatt) throw new Error("La impresora no ofrece conexión GATT compatible con Web Bluetooth.");
        device.addEventListener("gattserverdisconnected", () => {{ status.textContent = "Impresora desconectada. Pulsa imprimir para reconectar."; }});
        const server = device.gatt.connected ? device.gatt : await device.gatt.connect();
        const services = await server.getPrimaryServices();
        let characteristic = null;
        for (const service of services) {{
          const chars = await service.getCharacteristics();
          characteristic = chars.find(c => c.properties.writeWithoutResponse || c.properties.write);
          if (characteristic) break;
        }}
        if (!characteristic) throw new Error("La impresora no expone una característica GATT de escritura compatible. Muchas térmicas usan Bluetooth Classic y requieren su app/driver.");
        const encoder = new TextEncoder();
        const bytes = encoder.encode(ticket + "\\n\\n\\n");
        const chunkSize = 180;
        for (let i = 0; i < bytes.length; i += chunkSize) {{
          const chunk = bytes.slice(i, i + chunkSize);
          if (characteristic.writeValueWithoutResponse && characteristic.properties.writeWithoutResponse) {{
            await characteristic.writeValueWithoutResponse(chunk);
          }} else {{
            await characteristic.writeValue(chunk);
          }}
        }}
        status.textContent = "Datos enviados a " + (device.name || "la impresora guardada") + ". Verifica el ticket físico.";
      }} catch (e) {{
        status.textContent = "No se pudo imprimir: " + e.message;
      }} finally {{
        btn.disabled = false;
      }}
    }});
    </script>
    """
    st.components.v1.html(html_doc, height=145)


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
    left, center, right = st.columns([1, 2, 1])
    with center:
        st.markdown("<h1 style='text-align:center'>🚗 CLUB PARK</h1>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align:center'>Sistema de Control Operativo</h3>", unsafe_allow_html=True)
        st.info("Identifícate mediante token o escanea una imagen QR con tu cámara.")
        metodo = st.radio(
            "Método de acceso",
            ["Token manual", "Escanear QR con cámara", "Subir imagen QR"],
            horizontal=True,
        )
        token = None
        if metodo == "Token manual":
            token_input = st.text_input("Token de acceso", type="password")
            if st.button("Validar acceso", type="primary", use_container_width=True):
                token = token_input.strip()
        elif metodo == "Escanear QR con cámara":
            foto_qr = st.camera_input("Apunta la cámara al QR")
            if foto_qr:
                token = leer_codigo_qr(Image.open(foto_qr))
        else:
            archivo_qr = st.file_uploader("Sube el QR", type=["png", "jpg", "jpeg"])
            if archivo_qr:
                token = leer_codigo_qr(Image.open(archivo_qr))

        if token:
            info = ACCESOS.get(token)
            if info:
                st.session_state.usuario_autenticado = info.get("nombre", "Usuario")
                st.session_state.sucursal_asignada = info.get("sucursal", "")
                st.session_state.rol = info.get("rol", "Supervisor de Sucursal")
                st.rerun()
            else:
                st.error("Token o QR no reconocido.")
    st.stop()


# ---------------------------------------------------------------------
# ENCABEZADO Y NAVEGACIÓN
# ---------------------------------------------------------------------
es_contador = usuario_es_contador()
es_global = usuario_es_global()
st.title("CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.")
sucursal_texto = "Todas (Global)" if es_global else st.session_state.sucursal_asignada
st.caption(
    f"Usuario: {st.session_state.usuario_autenticado} | "
    f"Rol: {st.session_state.rol} | Sucursal: {sucursal_texto}"
)

with st.sidebar:
    st.subheader("Panel de sesión")
    st.write(f"**Usuario:** {st.session_state.usuario_autenticado}")
    st.write(f"**Rol:** {st.session_state.rol}")
    st.write(f"**Vehículos registrados:** {len(st.session_state.vehiculos)}")
    if st.button("🚪 Cerrar sesión", use_container_width=True):
        st.session_state.usuario_autenticado = None
        st.session_state.sucursal_asignada = None
        st.session_state.rol = None
        st.rerun()

tab_ing, tab_list, tab_salida, tab_balance, tab_imp, tab_rep, tab_corte, tab_cierre_diario, tab_aud = st.tabs([
    "📥 Ingreso y 4 Fotografías",
    "📋 Vehículos Registrados",
    "🚘 Salida y Cobro",
    "📈 Balances y Rentabilidad",
    "🖨️ Re-Impresión Térmica",
    "📊 Reportes y Excel",
    "💵 Cortes de Caja",
    "📒 Cierre diario por sucursal",
    "🕵️ Auditoría y WhatsApp",
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
    st.markdown("**📷 Fotografías de pertenencias adicionales (opcional)**")
    foto_pertenencias_camara = st.camera_input("Fotografiar pertenencias del cliente", key="foto_pertenencias_camara")
    fotos_pertenencias_archivos = st.file_uploader(
        "O selecciona fotos de pertenencias", type=["jpg", "jpeg", "png"],
        accept_multiple_files=True, key="fotos_pertenencias_archivos"
    )
    if foto_pertenencias_camara:
        st.image(foto_pertenencias_camara, caption="Vista previa: pertenencias", width=220)
    if fotos_pertenencias_archivos:
        st.caption(f"Fotos adicionales seleccionadas: {len(fotos_pertenencias_archivos)}")

    st.subheader("2. Fotografías obligatorias del vehículo")
    st.caption("Debes adjuntar las cuatro vistas antes de registrar el ingreso.")
    a, b = st.columns(2)
    with a:
        foto_frente = st.file_uploader("1. Vista frontal", type=["jpg", "jpeg", "png"], key="foto_frente")
        foto_izq = st.file_uploader("2. Costado izquierdo", type=["jpg", "jpeg", "png"], key="foto_izq")
    with b:
        foto_der = st.file_uploader("3. Costado derecho", type=["jpg", "jpeg", "png"], key="foto_der")
        foto_tras = st.file_uploader("4. Vista trasera", type=["jpg", "jpeg", "png"], key="foto_tras")

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
        if not foto_frente:
            faltantes.append("foto frontal")
        if not foto_izq:
            faltantes.append("foto izquierda")
        if not foto_der:
            faltantes.append("foto derecha")
        if not foto_tras:
            faltantes.append("foto trasera")

        if faltantes:
            st.error("Completa los siguientes campos: " + ", ".join(faltantes) + ".")
        else:
            folio, folio_numero = asignar_folio(sucursal, st.session_state.vehiculos)
            fecha = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            rutas = {}
            archivos = [
                ("frente", foto_frente),
                ("izquierda", foto_izq),
                ("derecha", foto_der),
                ("trasera", foto_tras),
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
                    "fecha": fecha,
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
        imagen_qr_salida = st.camera_input("Fotografía el QR del boleto (centrado, enfocado y con buena luz)", key="camara_qr_salida")
        if imagen_qr_salida:
            valor_qr_leido = leer_codigo_qr(Image.open(imagen_qr_salida))
            valor_busqueda = valor_qr_leido or ""
            if not valor_qr_leido:
                st.warning("No se pudo leer el QR de esta foto. Intenta acercar la cámara, enfocar el código o usa 'Subir imagen QR'.")
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
            st.write(f"**Entrada:** {registro_salida.get('fecha', '')}")
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
            forma_pago = st.radio("Método de pago", ["Efectivo", "QR"], horizontal=True, key=f"pago_metodo_{registro_salida['id']}")
            referencia_pago = st.text_input("Referencia / autorización (si aplica)", key=f"pago_ref_{registro_salida['id']}")
            nota_pago = st.text_input("Observaciones del cobro", key=f"pago_nota_{registro_salida['id']}")
            if forma_pago == "Efectivo":
                st.info("El efectivo se registrará como recibido al confirmar la entrega.")
            else:
                st.warning(
                    "Esta versión no conecta aún con una pasarela. Cobra en la terminal/app del proveedor y "
                    "confirma manualmente que la operación esté aprobada antes de entregar el vehículo."
                )
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
                    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    pago = {
                        "fecha_hora": ahora,
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
                    registro_salida["entregado_por"] = st.session_state.usuario_autenticado
                    registro_salida["estado"] = "Entregado"
                    if guardar_bd(st.session_state.vehiculos):
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
        st.caption("El comprobante se genera después de registrar el cobro y la salida. Si el pago fue externo, la referencia es la que capturó el operador.")
        vista_comprobante_cobro(
            comprobante_reciente.get("pago", {}),
            comprobante_reciente.get("registro", {}),
            "imprimir_comprobante_cobro_reciente",
        )
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
        st.markdown("#### Configuración de folios por sucursal")
        st.caption("Define el siguiente consecutivo. El sistema nunca bajará del máximo ya utilizado en esa sucursal para evitar reutilizar folios.")
        config_folios = cargar_config()
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
  
