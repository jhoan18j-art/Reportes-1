# -*- coding: utf-8 -*-
"""
CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.
Sistema Inteligente de Control Operativo - Producción Total
- Autenticación Híbrida (QR Sucursales + Supervisores Globales + Contador Auditor)
- Captura Multivista (4 Fotografías obligatorias por vehículo)
- Impresión Térmica Directa Bluetooth ESC/POS al registrar
- Módulo Ejecutivo de Rentabilidad, Excel Oficial, WhatsApp y Auditoría
"""

import io
import json
import os
import uuid
from datetime import datetime
import pandas as pd
import streamlit as st
from PIL import Image

# Verificación de librerías para Excel corporativo
OPENPYXL_OK = False
try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    OPENPYXL_OK = True
except ImportError:
    pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in locals() else os.getcwd()
DATA_DIR = os.environ.get("CLUBPARK_DATA_DIR", BASE_DIR)
os.makedirs(DATA_DIR, exist_ok=True)

DB_VEHICULOS_FILE = os.path.join(DATA_DIR, "vehiculos_produccion_db.json")
FOTOS_VEH_DIR = os.path.join(DATA_DIR, "fotos_multivista_prod")
os.makedirs(FOTOS_VEH_DIR, exist_ok=True)

def cargar_bd():
    if os.path.exists(DB_VEHICULOS_FILE):
        try:
            with open(DB_VEHICULOS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def guardar_bd(registros):
    try:
        tmp_file = DB_VEHICULOS_FILE + ".tmp"
        with open(tmp_file, "w", encoding="utf-8") as f:
            json.dump(registros, f, ensure_ascii=False, indent=2)
        os.replace(tmp_file, DB_VEHICULOS_FILE)
        return True
    except Exception as e:
        st.error(f"Error guardando base de datos: {e}")
        return False

# Catálogo Oficial de Sucursales y Tarifas
PRECIOS_SUCURSALES = {
    "PAMPAS INSURGENTES": 50, "PAMPAS COAPA": 50, "RAMMA": 70, "SUSHITTO": 40,
    "CARAJILLO POLANCO": 100, "BENINGAN´S": 40, "SANATORIO DURANGO": 100,
    "BARRANCA": 0, "VETERINARIA": 0, "ROMANCE PALACE": 65, "JUNGLA TLALPAN": 100,
    "JUNGLA DIVISIÓN": 100, "JUNGLA PICACHO": 100, "JUNGLA EMPRESA": 100,
    "SIPIRILY COAPA": 80, "MAROMETA SATÉLITE": 70, "MAROMETA MARINA": 70,
    "SUMO STA. MONICA": 50, "SUMO ECHEGARAY": 50, "RYOSHI": 80,
    "ENCANTO": 60, "E-BNI": 50, "CHURRASQUERIA": 70
}
SUCURSALES_LISTA = ["-- Selecciona la Sucursal --"] + sorted(list(PRECIOS_SUCURSALES.keys()))

ACCESOS_QR_Y_ROLES = {
    "CPV-GLOBAL-AUDITOR-2026": {"nombre": "Víctor (Contador)", "sucursal": "TODAS", "rol": "Contador Auditor (Máxima Jerarquía)"},
    "CPV-GLOBAL-JESUS-2026": {"nombre": "José de Jesús", "sucursal": "TODAS", "rol": "Supervisor Global"},
    "CPV-GLOBAL-VICTOR-2026": {"nombre": "Víctor Manuel", "sucursal": "TODAS", "rol": "Supervisor Global"},
    "CPV-QR-PAMPAS-INS": {"nombre": "Encargado Pampas Ins.", "sucursal": "PAMPAS INSURGENTES", "rol": "Supervisor de Sucursal"},
    "CPV-QR-RAMMA": {"nombre": "Encargado Ramma", "sucursal": "RAMMA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-SUSHITTO": {"nombre": "Encargado Sushitto", "sucursal": "SUSHITTO", "rol": "Supervisor de Sucursal"},
    "CPV-QR-CARAJILLO": {"nombre": "Encargado Carajillo", "sucursal": "CARAJILLO POLANCO", "rol": "Supervisor de Sucursal"},
    "CPV-QR-BENINGANS": {"nombre": "Encargado Beningan´s", "sucursal": "BENINGAN´S", "rol": "Supervisor de Sucursal"},
    "CPV-QR-DURANGO": {"nombre": "Encargado S. Durango", "sucursal": "SANATORIO DURANGO", "rol": "Supervisor de Sucursal"},
    "CPV-QR-BARRANCA": {"nombre": "Encargado Barranca", "sucursal": "BARRANCA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-VETERINARIA": {"nombre": "Encargado Veterinaria", "sucursal": "VETERINARIA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-ROMANCE": {"nombre": "Encargado Romance", "sucursal": "ROMANCE PALACE", "rol": "Supervisor de Sucursal"},
    "CPV-QR-TLALPAN": {"nombre": "Encargado J. Tlalpan", "sucursal": "JUNGLA TLALPAN", "rol": "Supervisor de Sucursal"},
    "CPV-QR-DIVISION": {"nombre": "Encargado J. División", "sucursal": "JUNGLA DIVISIÓN", "rol": "Supervisor de Sucursal"},
    "CPV-QR-PICACHO": {"nombre": "Encargado J. Picacho", "sucursal": "JUNGLA PICACHO", "rol": "Supervisor de Sucursal"},
    "CPV-QR-EMPRESA": {"nombre": "Encargado J. Empresa", "sucursal": "JUNGLA EMPRESA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-SIPIRILY": {"nombre": "Encargado Sipirily", "sucursal": "SIPIRILY COAPA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-SATELITE": {"nombre": "Encargado Satélite", "sucursal": "MAROMETA SATÉLITE", "rol": "Supervisor de Sucursal"},
    "CPV-QR-MARINA": {"nombre": "Encargado Marina", "sucursal": "MAROMETA MARINA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-SUMOSTA": {"nombre": "Encargado Sumo Sta.", "sucursal": "SUMO STA. MONICA", "rol": "Supervisor de Sucursal"},
    "CPV-QR-CHEGARAY": {"nombre": "Encargado Chegaray", "sucursal": "SUMO ECHEGARAY", "rol": "Supervisor de Sucursal"},
    "CPV-QR-RYOSHI": {"nombre": "Encargado Ryoshi", "sucursal": "RYOSHI", "rol": "Supervisor de Sucursal"},
    "CPV-QR-ENCANTO": {"nombre": "Encargado Encanto", "sucursal": "ENCANTO", "rol": "Supervisor de Sucursal"},
    "CPV-QR-EBNI": {"nombre": "Encargado E-BNI", "sucursal": "E-BNI", "rol": "Supervisor de Sucursal"},
    "CPV-QR-CHURRAS": {"nombre": "Encargado Churrasqueria", "sucursal": "CHURRASQUERIA", "rol": "Supervisor de Sucursal"}
}

def leer_codigo_qr(imagen_pil):
    try:
        from pyzbar.pyzbar import decode as pyzbar_decode
        decoded = pyzbar_decode(imagen_pil)
        if decoded:
            return decoded[0].data.decode("utf-8").strip()
    except Exception:
        pass
    try:
        import cv2
        import numpy as np
        img_np = np.array(imagen_pil.convert("RGB"))
        img_cv = cv2.cvtColor(img_np, cv2.COLOR_RGB2BGR)
        detector = cv2.QRCodeDetector()
        val, _, _ = detector.detectAndDecode(img_cv)
        if val:
            return val.strip()
    except Exception:
        pass
    return None

def generar_enlace_whatsapp(telefono: str, mensaje: str) -> str:
    tel_clean = "".join(c for c in str(telefono) if c.isdigit())
    if not tel_clean.startswith("52") and len(tel_clean) == 10:
        tel_clean = "52" + tel_clean
    import urllib.parse
    return f"https://wa.me/{tel_clean}?text={urllib.parse.quote(mensaje)}"

def generar_excel_oficial(df):
    output = io.BytesIO()
    if OPENPYXL_OK:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Reporte Operativo"
        
        ws.merge_cells("A1:G1")
        title_cell = ws["A1"]
        title_cell.value = "CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V."
        title_cell.font = Font(name="Arial", size=14, bold=True, color="FFFFFF")
        title_cell.fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        title_cell.alignment = Alignment(horizontal="center", vertical="center")

        headers = ["FOLIO ID", "FECHA Y HORA", "SUCURSAL", "MARCA/MODELO", "PLACAS", "TARIFA ($)", "OPERADOR"]
        ws.append([])
        ws.append(headers)
        
        row_idx = 5
        for _, r in df.iterrows():
            ws.append([
                r.get("id", ""), r.get("fecha", ""), r.get("sucursal", ""),
                r.get("marca", ""), r.get("placas", ""), r.get("tarifa", 0.0),
                r.get("operador", "")
            ])
            ws.cell(row=row_idx, column=6).number_format = '"$"#,##0.00'
            row_idx += 1

        wb.save(output)
    else:
        df.to_excel(output, index=False, engine="openpyxl")
    output.seek(0)
    return output

def componente_impresion_bluetooth(texto_ticket, key_suffix=""):
    """Componente JavaScript autocontenido para imprimir directamente en impresoras térmicas ESC/POS."""
    # Sanitizamos el texto para inyectarlo de forma segura en JS
    texto_js = json.dumps(texto_ticket)
    
    js_code = f"""
    <script>
    async function conectarEImprimir() {{
        const btn = document.getElementById("btn_print_{key_suffix}");
        try {{
            btn.innerText = "⏳ Conectando a impresora...";
            btn.disabled = true;

            const thermalPrinterServices = [
                '00001101-0000-1000-8000-00805f9b34fb',
                '000018f0-0000-1000-8000-00805f9b34fb',
                'e7810a71-73ae-499d-8c15-faa9aef0c3f2'
            ];

            const device = await navigator.bluetooth.requestDevice({{
                acceptAllDevices: true,
                optionalServices: thermalPrinterServices
            }});

            const server = await device.gatt.connect();
            const services = await server.getPrimaryServices();
            
            if (services.length === 0) {{
                throw new Error("No se encontraron servicios Bluetooth disponibles.");
            }}

            const service = services[0];
            const characteristics = await service.getCharacteristics();
            const characteristic = characteristics.find(c => c.properties.write || c.properties.writeWithoutResponse) || characteristics[0];

            const texto = {texto_js};
            const encoder = new TextEncoder();
            
            // Comando ESC/POS GS V 65 16 para corte completo de papel
            const comandoCorte = new Uint8Array([0x1D, 0x56, 0x41, 0x10]);

            await characteristic.writeValue(encoder.encode(texto + "\\n\\n\\n"));
            await characteristic.writeValue(comandoCorte);

            btn.innerText = "✅ ¡Ticket Impreso!";
            btn.style.backgroundColor = "#276749";
            setTimeout(() => {{
                btn.innerText = "🖨️ Re-imprimir Ticket vía Bluetooth";
                btn.disabled = false;
            }}, 3000);

        }} catch (error) {{
            alert("Error al imprimir Bluetooth: " + error.message);
            btn.innerText = "🖨️ Reintentar Impresión Bluetooth";
            btn.style.backgroundColor = "#C53030";
            btn.disabled = false;
        }}
    }}
    </script>
    <button id="btn_print_{key_suffix}" onclick="conectarEImprimir()" 
            style="background-color:#22543D; color:white; padding:14px 20px; border:none; 
                   border-radius:8px; font-weight:bold; font-size:16px; width:100%; cursor:pointer;">
        🖨️ CONECTAR E IMPRIMIR TICKET EN IMPRESORA BLUETOOTH
    </button>
    """
    st.components.v1.html(js_code, height=70)

st.set_page_config(
    page_title="Club Park - Sistema de Producción Total",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded"
)

if "usuario_autenticado" not in st.session_state:
    st.session_state.usuario_autenticado = None
    st.session_state.sucursal_asignada = None
    st.session_state.rol = None

if "vehiculos" not in st.session_state:
    st.session_state.vehiculos = cargar_bd()

# =========================================================================
# ACCESO (LOGIN)
# =========================================================================
if st.session_state.usuario_autenticado is None:
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        st.markdown("<h1 style='text-align: center;'>🚗 CLUB PARK</h1>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align: center; color: #4A5568;'>Acceso Híbrido</h3>", unsafe_allow_html=True)
        st.markdown("---")

        metodo = st.radio("Método de Identificación:", ["📷 Escanear con Cámara", "📁 Subir Imagen QR", "🔑 Token Manual"], horizontal=True)
        token_ing = None

        if metodo == "📷 Escanear con Cámara":
            cam_img = st.camera_input("Apunta tu cámara al código QR:", key="login_cam")
            if cam_img:
                token_ing = leer_codigo_qr(Image.open(cam_img))
        elif metodo == "📁 Subir Imagen QR":
            file_img = st.file_uploader("Sube la imagen del QR:", type=["png", "jpg", "jpeg"], key="login_file")
            if file_img:
                token_ing = leer_codigo_qr(Image.open(file_img))
        else:
            t_man = st.text_input("Ingresa tu token de acceso:", type="password", key="login_token_man")
            if st.button("Validar Token", use_container_width=True, type="primary"):
                token_ing = t_man.strip()

        if token_ing:
            if token_ing in ACCESOS_QR_Y_ROLES:
                info_u = ACCESOS_QR_Y_ROLES[token_ing]
                st.session_state.usuario_autenticado = info_u["nombre"]
                st.session_state.sucursal_asignada = info_u["sucursal"]
                st.session_state.rol = info_u["rol"]
                st.success(f"✅ Acceso concedido: **{info_u['nombre']}** ({info_u['rol']})")
                st.rerun()
            else:
                st.error("❌ Código QR o token no reconocido.")
        st.stop()

# =========================================================================
# PANEL PRINCIPAL
# =========================================================================
es_contador = ("Contador" in st.session_state.rol)
es_global = (es_contador or st.session_state.rol in ["Supervisor Global"])

st.markdown(f"### CLUB PARK VALET Y ESTACIONAMIENTOS S.A. DE C.V.")
st.markdown(f"**Usuario:** `{st.session_state.usuario_autenticado}` | **Rol:** `{st.session_state.rol}` | **Sucursal:** `{'Todas (Global)' if es_global else st.session_state.sucursal_asignada}`")

if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
    st.session_state.usuario_autenticado = None
    st.session_state.sucursal_asignada = None
    st.session_state.rol = None
    st.rerun()

tab_ing, tab_list, tab_imp, tab_rep, tab_aud = st.tabs([
    "📥 Ingreso y 4 Fotografías",
    "📋 Vehículos Registrados",
    "🖨️ Re-Impresión Térmica",
    "📊 Reportes y Excel",
    "🕵️ Auditoría y WhatsApp"
])

# -------------------------------------------------------------------------
# TAB 1: INGRESO DE VEHÍCULOS E IMPRESIÓN DIRECTA
# -------------------------------------------------------------------------
with tab_ing:
    st.markdown("#### 1. Datos del Vehículo y Recepción")
    c1, c2 = st.columns(2)
    with c1:
        if es_global:
            sucursal_op = st.selectbox("Selecciona Sucursal Operativa:", SUCURSALES_LISTA, key="op_suc_g")
        else:
            sucursal_op = st.session_state.sucursal_asignada
            st.info(f"🏢 Sucursal fija asignada: **{sucursal_op}**")
            
        marca_v = st.text_input("Marca y Modelo del Vehículo:", key="op_marca")
        color_v = st.text_input("Color del Vehículo:", key="op_color")
    with c2:
        placas_v = st.text_input("Placas / Folio:", key="op_placas").upper()
        t_base = PRECIOS_SUCURSALES.get(sucursal_op, 50) if sucursal_op in PRECIOS_SUCURSALES else 50
        tarifa_v = st.number_input("Tarifa ($ MXN):", min_value=0.0, value=float(t_base), step=10.0, key="op_tarifa")
        operador_v = st.text_input("Valet / Operador:", value=st.session_state.usuario_autenticado, key="op_valet")

    st.markdown("#### 2. Custodia de Objetos y Observaciones")
    custodia_v = st.text_area("Objetos en custodia (Celulares, armas, efectivo declarado):", key="op_custodia")

    st.markdown("#### 3. Captura Multivista Obligatoria (4 Fotografías)")
    cf1, cf2 = st.columns(2)
    with cf1:
        f_frente = st.file_uploader("📷 1. Vista Frontal", type=["jpg", "jpeg", "png"], key="img_f")
        f_izq = st.file_uploader("📷 2. Costado Izquierdo", type=["jpg", "jpeg", "png"], key="img_i")
    with cf2:
        f_der = st.file_uploader("📷 3. Costado Derecho", type=["jpg", "jpeg", "png"], key="img_d")
        f_tras = st.file_uploader("📷 4. Vista Trasera", type=["jpg", "jpeg", "png"], key="img_t")

    if st.button("💾 Registrar Ingreso, Guardar 4 Fotos y Generar Ticket", type="primary", use_container_width=True):
        if sucursal_op == "-- Selecciona la Sucursal --":
            st.error("❌ Por favor selecciona una sucursal válida.")
        elif not placas_v and not marca_v:
            st.error("❌ Ingresa al menos las placas o la marca del vehículo.")
        else:
            folio_id = f"CP-{uuid.uuid4().hex[:5].upper()}"
            ahora_s = datetime.now().strftime("%d/%m/%Y %H:%M")
            
            rutas_fotos = {}
            for k_f, f_obj in [("frente", f_frente), ("izq", f_izq), ("der", f_der), ("tras", f_tras)]:
                if f_obj:
                    try:
                        img = Image.open(f_obj)
                        img.thumbnail((700, 700))
                        nombre_img = f"{folio_id}_{k_f}.jpg"
                        ruta_img = os.path.join(FOTOS_VEH_DIR, nombre_img)
                        img.save(ruta_img, "JPEG", quality=60, optimize=True)
                        rutas_fotos[k_f] = ruta_img
                    except Exception:
                        pass

            ticket_escpos = f"""
================================
 CLUB PARK VALET Y ESTACIONAMIENTOS
         S.A. DE C.V.
================================
SUCURSAL: {sucursal_op}
FOLIO TICKET: {folio_id}
TARIFA: ${tarifa_v:,.2f} MXN
--------------------------------
MARCA: {marca_v}
PLACA: {placas_v}      COLOR: {color_v}
ENTRADA: {ahora_s}    SALIDA: _______
--------------------------------
[X] OBSERVACIONES Y CUSTODIA:
{custodia_v if custodia_v else 'Sin objetos declarados'}
--------------------------------
       BOLETO DEL CLIENTE
--------------------------------
NO NOS HACEMOS RESPONSABLES POR
OBJETOS COMO SON: EFECTIVO, CHEQUES,
ARMAS, BOLSAS, TARJETAS, CELULARES.
================================
           [ CORTE ]
================================
      BOLETO SIN VALOR LLAVERO
FOLIO: {folio_id} | PLACA: {placas_v}
================================
           [ CORTE ]
================================
     BOLETO SIN VALOR PARABRISAS
FOLIO: {folio_id}
================================
"""

            nuevo_registro = {
                "id": folio_id,
                "fecha": ahora_s,
                "sucursal": sucursal_op,
                "marca": marca_v,
                "color": color_v,
                "placas": placas_v,
                "tarifa": float(tarifa_v),
                "operador": operador_v,
                "custodia": custodia_v,
                "fotos": rutas_fotos,
                "ticket_texto": ticket_escpos,
                "estado": "En Resguardo"
            }

            st.session_state.vehiculos.insert(0, nuevo_registro)
            guardar_bd(st.session_state.vehiculos)
            
            st.session_state["ultimo_ticket"] = ticket_escpos
            st.session_state["ultimo_folio"] = folio_id
            st.success(f"✅ ¡Ingreso registrado exitosamente! Folio: `{folio_id}`")

    # Muestra el botón de impresión inmediata al registrar el ticket
    if "ultimo_ticket" in st.session_state:
        st.markdown("---")
        st.markdown(f"### 🖨️ Impresión Inmediata - Ticket Folio `{st.session_state.get('ultimo_folio', '')}`")
        componente_impresion_bluetooth(st.session_state["ultimo_ticket"], key_suffix="ingreso")

# -------------------------------------------------------------------------
# TAB 2: LISTADO DE VEHÍCULOS
# -------------------------------------------------------------------------
with tab_list:
    st.markdown("#### 📋 Listado General de Vehículos")
    if st.session_state.vehiculos:
        df_v = pd.DataFrame(st.session_state.vehiculos)
        if not es_global:
            df_v = df_v[df_v["sucursal"] == st.session_state.sucursal_asignada]
            
        st.dataframe(df_v[["id", "fecha", "sucursal", "marca", "placas", "estado"]], use_container_width=True, hide_index=True)
        
        sel_folio = st.selectbox("Selecciona Folio para auditar fotos:", [v["id"] for v in st.session_state.vehiculos], key="sel_f_aud")
        v_det = next((v for v in st.session_state.vehiculos if v["id"] == sel_folio), None)
        if v_det:
            st.write(f"**Sucursal:** {v_det['sucursal']} | **Placas:** {v_det['placas']} | **Custodia:** {v_det['custodia']}")
            f_dict = v_det.get("fotos", {})
            if f_dict:
                cols_f = st.columns(4)
                for idx, (k_angle, r_path) in enumerate(f_dict.items()):
                    if os.path.exists(r_path):
                        with cols_f[idx % 4]:
                            st.image(r_path, caption=f"Vista {k_angle.upper()}", use_container_width=True)
            else:
                st.caption("Sin registro fotográfico asociado.")
    else:
        st.info("No hay vehículos registrados.")

# -------------------------------------------------------------------------
# TAB 3: RE-IMPRESIÓN TÉRMICA
# -------------------------------------------------------------------------
with tab_imp:
    st.markdown("#### 🖨️ Re-Impresión de Tickets")
    if st.session_state.vehiculos:
        lista_folios = [v["id"] for v in st.session_state.vehiculos]
        folio_sel = st.selectbox("Selecciona Folio a imprimir:", lista_folios, key="reprint_sel")
        reg_sel = next((v for v in st.session_state.vehiculos if v["id"] == folio_sel), None)
        
        if reg_sel:
            st.text_area("Formato ESC/POS:", value=reg_sel["ticket_texto"], height=280, key="txt_reprint")
            componente_impresion_bluetooth(reg_sel["ticket_texto"], key_suffix="reimpresion")
    else:
        st.info("No hay registros disponibles para re-imprimir.")

# -------------------------------------------------------------------------
# TAB 4: REPORTES Y RENTABILIDAD EXCEL
# -------------------------------------------------------------------------
with tab_rep:
    st.markdown("### 📊 Reportes de Operación y Rentabilidad")
    if st.session_state.vehiculos:
        df_rep = pd.DataFrame(st.session_state.vehiculos)
        if not es_global:
            df_rep = df_rep[df_rep["sucursal"] == st.session_state.sucursal_asignada]

        m1, m2, m3 = st.columns(3)
        m1.metric("Vehículos Registrados", len(df_rep))
        m2.metric("Ingreso Total Acumulado", f"${df_rep['tarifa'].sum():,.2f} MXN")
        m3.metric("Tarifa Promedio", f"${df_rep['tarifa'].mean():,.2f} MXN" if len(df_rep)>0 else "$0.00")

        st.markdown("---")
        excel_data = generar_excel_oficial(df_rep)
        
        st.download_button(
            label="📥 Descargar Reporte en Excel (.xlsx)",
            data=excel_data,
            file_name=f"Reporte_ClubPark_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )
    else:
        st.info("No hay registros para generar reportes.")

# -------------------------------------------------------------------------
# TAB 5: AUDITORÍA Y ALERTAS WHATSAPP
# -------------------------------------------------------------------------
with tab_aud:
    st.markdown("### 🕵️ Auditoría, Eventos Extraordinarios y Alertas WhatsApp")
    
    if not es_contador and not es_global:
        st.warning("🔒 Esta sección está restringida para Supervisores Globales y el Contador Auditor.")
    else:
        st.success(f"🔓 Sesión de Auditoría Activa: **{st.session_state.usuario_autenticado}**")
        
        st.markdown("#### Notificaciones Instantáneas por WhatsApp")
        c_wa1, c_wa2 = st.columns(2)
        with c_wa1:
            num_destino = st.text_input("Número WhatsApp Destino (10 dígitos):", value="5512345678")
            suc_alerta = st.selectbox("Sucursal del Evento:", SUCURSALES_LISTA[1:])
        with c_wa2:
            tipo_evento = st.selectbox("Tipo de Evento:", ["Siniestro / Golpe", "Objeto Extraviado / Reclamación", "Sin Boleto / Pérdida", "Corte de Caja Irregular"])
            detalles_ev = st.text_area("Detalles del Incidente:")

        if st.button("📲 Generar y Enviar Alerta por WhatsApp", type="primary"):
            msg_wa = f"🚨 *ALERTA OPERATIVA CLUB PARK*\n" \
                     f"🏢 *Sucursal:* {suc_alerta}\n" \
                     f"⚠️ *Evento:* {tipo_evento}\n" \
                     f"📝 *Detalles:* {detalles_ev}\n" \
                     f"👤 *Reporta:* {st.session_state.usuario_autenticado} ({st.session_state.rol})"
            
            link_wa = generar_enlace_whatsapp(num_destino, msg_wa)
            st.markdown(f"[👉 Haz clic aquí para abrir WhatsApp y enviar alerta]({link_wa})")
