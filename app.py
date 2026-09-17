import io
import os
import re
import json
import time
import uuid
import threading
import importlib
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd
import requests
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import PieChart, Reference

# Importación del motor de reglas independiente
import reglas_sat as sat
import bd_supabase as bd
import auth

# ---------------------------------------------------------
# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS AUTOCOM
# ---------------------------------------------------------
st.set_page_config(
    page_title="Auditor CFDI 4.0 - AUTOCOM",
    page_icon="🚗",
    layout="wide"
)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    :root {
        --autocom-red: #C8102E;
        --autocom-red-dark: #990B22;
        --autocom-red-light: #E01235;
        --autocom-ink: #1F2937;
        --autocom-slate: #4B5563;
        --autocom-mist: #F3F4F6;
        --autocom-line: #E5E7EB;
        --autocom-green: #15803D;
        --autocom-green-bg: #ECFDF5;
        --autocom-amber: #B45309;
        --autocom-amber-bg: #FFFBEB;
        --autocom-red-bg: #FEF2F2;
        --radius-card: 14px;
        --shadow-soft: 0 2px 10px rgba(17, 24, 39, 0.06);
        --shadow-hover: 0 10px 24px rgba(17, 24, 39, 0.10);
    }

    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
    }

    /* Ocultar el widget / animación nativa de carga de Streamlit (monito nadando) */
    [data-testid="stStatusWidget"] {
        display: none !important;
        visibility: hidden !important;
    }

    /* Tipografía de encabezados más consistente */
    h1, h2, h3 { color: var(--autocom-ink); font-weight: 800 !important; letter-spacing: -0.01em; }
    h4, h5, h6 { color: var(--autocom-ink); font-weight: 700 !important; }
    p, span, label, div { letter-spacing: 0.001em; }

    /* Botones principales con estilo corporativo Autocom */
    .stButton>button, .stDownloadButton>button {
        background: linear-gradient(135deg, #C8102E 0%, #990B22 100%) !important;
        color: white !important;
        border-radius: 9px !important;
        border: none !important;
        font-weight: 700 !important;
        padding: 0.6rem 1.25rem !important;
        box-shadow: 0 3px 8px rgba(200, 16, 46, 0.22) !important;
        transition: all 0.18s ease !important;
        letter-spacing: 0.01em;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 18px rgba(200, 16, 46, 0.32) !important;
        background: linear-gradient(135deg, #E01235 0%, #C8102E 100%) !important;
    }
    .stButton>button:active, .stDownloadButton>button:active {
        transform: translateY(0) !important;
    }

    /* Encabezado principal */
    .header-card {
        background-color: #FFFFFF;
        padding: 1.6rem 1.8rem;
        border-radius: var(--radius-card);
        border-left: 6px solid var(--autocom-red);
        box-shadow: var(--shadow-soft);
        margin-bottom: 1.5rem;
    }
    .author-badge {
        color: var(--autocom-slate);
        font-size: 0.85rem;
        font-weight: 600;
    }

    /* Indicador Visual Personalizado para la IA */
    .ai-loading-box {
        background: var(--autocom-red-bg);
        border: 1px solid #F5B7B1;
        border-radius: 10px;
        padding: 12px 18px;
        display: flex;
        align-items: center;
        gap: 12px;
        margin: 10px 0;
    }
    .ai-spinner {
        width: 20px;
        height: 20px;
        border: 3px solid #F5B7B1;
        border-top: 3px solid var(--autocom-red);
        border-radius: 50%;
        animation: spin 0.8s linear infinite;
    }
    @keyframes spin {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }

    /* Pantalla de login */
    .login-card {
        background-color: #FFFFFF;
        padding: 2.2rem 2.4rem;
        border-radius: 16px;
        border-top: 6px solid var(--autocom-red);
        box-shadow: 0 12px 32px rgba(0,0,0,0.10);
        margin-top: 2.5rem;
    }
    .login-title {
        color: var(--autocom-ink);
        font-size: 1.6rem;
        font-weight: 800;
        margin-bottom: 0.1rem;
    }
    .login-subtitle {
        color: var(--autocom-slate);
        font-size: 0.92rem;
        margin-bottom: 1.2rem;
    }
    .sidebar-user-box {
        background: linear-gradient(135deg, #1F2937 0%, #111827 100%);
        color: #FFFFFF;
        border-radius: 12px;
        padding: 12px 16px;
        margin-bottom: 10px;
        box-shadow: var(--shadow-soft);
    }
    .sidebar-user-box b { color: #FFFFFF; }
    .sidebar-user-role {
        display: inline-block;
        background-color: var(--autocom-red);
        color: #FFFFFF;
        border-radius: 6px;
        padding: 1px 9px;
        font-size: 0.72rem;
        font-weight: 700;
        margin-top: 5px;
        letter-spacing: 0.03em;
    }

    /* Badges de estatus / severidad reutilizables (ver _badge() en Python) */
    .autocom-badge {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        padding: 3px 11px;
        border-radius: 999px;
        font-size: 0.78rem;
        font-weight: 700;
        letter-spacing: 0.02em;
        line-height: 1.6;
    }
    .badge-ok { background: var(--autocom-green-bg); color: var(--autocom-green); border: 1px solid #A7F3D0; }
    .badge-warn { background: var(--autocom-amber-bg); color: var(--autocom-amber); border: 1px solid #FDE68A; }
    .badge-danger { background: var(--autocom-red-bg); color: var(--autocom-red-dark); border: 1px solid #FECACA; }
    .badge-neutral { background: var(--autocom-mist); color: var(--autocom-slate); border: 1px solid var(--autocom-line); }

    /* Tabs principales: más aire y una línea roja activa clara */
    [data-testid="stTabs"] button[role="tab"] {
        font-weight: 700 !important;
        font-size: 0.95rem !important;
        color: var(--autocom-slate) !important;
        padding: 10px 18px !important;
    }
    [data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
        color: var(--autocom-red) !important;
        border-bottom: 3px solid var(--autocom-red) !important;
    }
    [data-testid="stTabs"] { margin-top: 0.4rem; }

    /* Expanders con aspecto de tarjeta: bordes suaves y sombra al pasar el mouse */
    [data-testid="stExpander"] {
        border: 1px solid var(--autocom-line) !important;
        border-radius: var(--radius-card) !important;
        box-shadow: var(--shadow-soft);
        margin-bottom: 0.85rem;
        overflow: hidden;
    }
    [data-testid="stExpander"] summary {
        font-weight: 700 !important;
        padding: 0.9rem 1.1rem !important;
    }
    [data-testid="stExpander"]:hover {
        box-shadow: var(--shadow-hover);
    }

    /* Contenedores con borde (st.container(border=True)) como tarjetas limpias */
    [data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 12px !important;
        border-color: var(--autocom-line) !important;
    }

    /* Métricas como tarjetas KPI */
    [data-testid="stMetric"] {
        background: #FFFFFF;
        border: 1px solid var(--autocom-line);
        border-radius: 12px;
        padding: 0.9rem 1.1rem 0.7rem 1.1rem;
        box-shadow: var(--shadow-soft);
    }
    [data-testid="stMetricLabel"] { font-weight: 600 !important; color: var(--autocom-slate) !important; }
    [data-testid="stMetricValue"] { font-weight: 800 !important; color: var(--autocom-ink) !important; }

    /* Alertas nativas (info/success/warning/error) más redondeadas */
    .stAlert {
        border-radius: 12px !important;
        box-shadow: var(--shadow-soft);
    }

    /* Divisores más sutiles que el <hr> gris plano por defecto */
    hr {
        border: none !important;
        height: 1px !important;
        background: linear-gradient(90deg, var(--autocom-line) 0%, transparent 100%) !important;
        margin: 1.1rem 0 !important;
    }

    /* Sidebar con un poco más de padding y separación visual del contenido */
    section[data-testid="stSidebar"] {
        border-right: 1px solid var(--autocom-line);
    }
    section[data-testid="stSidebar"] .block-container {
        padding-top: 1.2rem;
    }

    /* Inputs de texto y áreas de texto con bordes más suaves */
    .stTextInput input, .stTextArea textarea, .stNumberInput input {
        border-radius: 9px !important;
        border-color: var(--autocom-line) !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: var(--autocom-red) !important;
        box-shadow: 0 0 0 1px var(--autocom-red) !important;
    }

    /* Tablas / dataframes con esquinas redondeadas y borde sutil */
    [data-testid="stDataFrame"] {
        border-radius: 12px !important;
        overflow: hidden;
        border: 1px solid var(--autocom-line) !important;
    }
    </style>
""", unsafe_allow_html=True)


def _badge(texto, tipo="neutral", icono=""):
    """
    HTML de un badge/pill de color consistente con la paleta AUTOCOM, para
    usar dentro de cualquier st.markdown(..., unsafe_allow_html=True).
    `tipo`: "ok" (verde) | "warn" (ámbar) | "danger" (rojo) | "neutral" (gris).
    """
    clase = {"ok": "badge-ok", "warn": "badge-warn", "danger": "badge-danger"}.get(tipo, "badge-neutral")
    prefijo = f"{icono} " if icono else ""
    return f'<span class="autocom-badge {clase}">{prefijo}{texto}</span>'

# ---------------------------------------------------------
# 0. SISTEMA DE AUTENTICACIÓN (LOGIN) Y CONTROL DE ACCESO
# ---------------------------------------------------------
# Los usuarios ya NO viven en un diccionario en el código (con contraseñas
# en texto plano) — viven en la tabla `usuarios` de Supabase, con la
# contraseña guardada como hash (ver auth.py). Cada persona crea SU PROPIA
# contraseña la primera vez que entra; ni el Admin ni nadie con acceso al
# código o al repositorio puede verla.
#
# Alta de gente nueva: el Admin la da de alta desde la pestaña "Ajustes"
# (usuario + nombre + rol, sin contraseña) y esa persona la crea sola en su
# primer ingreso. Si la tabla `usuarios` está completamente vacía (primer
# arranque de la app), la pantalla de login pide crear la cuenta de Admin
# inicial.


@st.cache_data(show_spinner=False, ttl=30)
def _cargar_usuarios_cacheado():
    try:
        return bd.cargar_usuarios()
    except Exception:
        return {}


def obtener_usuarios():
    return _cargar_usuarios_cacheado()


if "autenticado" not in st.session_state:
    st.session_state.autenticado = False


# ---------------------------------------------------------
# UNA SOLA SESIÓN ACTIVA POR USUARIO
# ---------------------------------------------------------
# `st.cache_resource` devuelve el MISMO objeto para todas las sesiones que
# corren dentro del mismo proceso de Streamlit — a diferencia de
# `st.session_state`, que es privado de cada pestaña/navegador. Por eso sirve
# para llevar un registro compartido de "quién tiene abierta cuál cuenta".
#
# LIMITACIÓN a tener presente: si la app se despliega detrás de varios
# procesos/réplicas (balanceo de carga), cada proceso tendría su propio
# registro y esto dejaría de ser 100% confiable. Para el uso normal en un
# solo servidor de Streamlit (el caso de AUTOCOM) funciona bien.
TIMEOUT_SESION_SEGUNDOS = 5 * 60  # sin actividad en 5 min, la sesión se da por cerrada


@st.cache_resource
def _sesiones_activas():
    return {}  # usuario_key -> {"session_id": str, "ultimo_latido": float, "nombre": str}


def _sesion_esta_activa(usuario_key):
    """True si YA hay una sesión viva con este usuario (de otra pestaña/
    navegador), sin contar la sesión actual."""
    with _LOCK_SESIONES:
        datos = _sesiones_activas().get(usuario_key)
        if not datos:
            return False
        return (time.time() - datos["ultimo_latido"]) <= TIMEOUT_SESION_SEGUNDOS


def _registrar_sesion(usuario_key, session_id, nombre):
    with _LOCK_SESIONES:
        _sesiones_activas()[usuario_key] = {
            "session_id": session_id, "ultimo_latido": time.time(), "nombre": nombre,
        }


def _refrescar_latido_sesion(usuario_key, session_id):
    """Se llama en cada ejecución del script mientras el usuario sigue
    autenticado, para que su sesión no expire mientras la sigue usando."""
    with _LOCK_SESIONES:
        datos = _sesiones_activas().get(usuario_key)
        if datos and datos["session_id"] == session_id:
            datos["ultimo_latido"] = time.time()


def _liberar_sesion(usuario_key, session_id):
    with _LOCK_SESIONES:
        datos = _sesiones_activas().get(usuario_key)
        if datos and datos["session_id"] == session_id:
            _sesiones_activas().pop(usuario_key, None)


_LOCK_SESIONES = threading.Lock()


def _cerrar_sesion():
    usuario_key = st.session_state.get("usuario_actual")
    session_id = st.session_state.get("_session_id")
    if usuario_key and session_id:
        _liberar_sesion(usuario_key, session_id)
    for clave in ("autenticado", "usuario_actual", "rol_actual", "nombre_actual", "_session_id"):
        st.session_state.pop(clave, None)


def _render_pantalla_login():
    """
    Devuelve True si el usuario acaba de autenticarse en ESTA ejecución.

    Antes se llamaba st.rerun() al validar las credenciales: eso obligaba a
    Streamlit a reejecutar todo el script desde cero (catálogo del SAT,
    categorías dinámicas, barra lateral de IA...) y es lo que se sentía como
    un login "pegado" de varios segundos. Ahora la pantalla de login se dibuja
    dentro de un contenedor que se vacía en cuanto las credenciales son
    correctas, y la ejecución continúa hacia la app en la MISMA pasada: se
    ahorra un ciclo completo de render.
    """
    contenedor_login = st.empty()
    usuarios = obtener_usuarios()

    with contenedor_login.container():
        col_izq, col_centro, col_der = st.columns([1, 1.3, 1])
        with col_centro:
            st.markdown(
                "<h2 style='color:#C8102E; font-weight:bold; text-align:center;'>🚗 AUTOCOM</h2>",
                unsafe_allow_html=True,
            )
            st.markdown("""
                <div class="login-card">
                    <div class="login-title">Auditor CFDI 4.0</div>
                    <div class="login-subtitle">Acceso restringido — Ingresa tus credenciales del equipo AUTOCOM.</div>
                </div>
            """, unsafe_allow_html=True)

            # --- CASO 1: primer arranque, todavía no existe ningún usuario ---
            if not usuarios:
                st.info("👋 Primer arranque: crea la cuenta de Administrador.")
                with st.form("form_primer_admin", clear_on_submit=False):
                    admin_usuario = st.text_input("Usuario (ej. admin)")
                    admin_nombre = st.text_input("Tu nombre")
                    admin_pass1 = st.text_input("Contraseña (mínimo 8 caracteres)", type="password")
                    admin_pass2 = st.text_input("Repite la contraseña", type="password")
                    crear = st.form_submit_button("Crear cuenta de Administrador", use_container_width=True)
                if crear:
                    admin_usuario_key = admin_usuario.strip().lower()
                    if not admin_usuario_key or not admin_nombre.strip():
                        st.error("🚫 Completa usuario y nombre.")
                    elif not auth.contrasena_es_valida(admin_pass1):
                        st.error("🚫 La contraseña debe tener al menos 8 caracteres.")
                    elif admin_pass1 != admin_pass2:
                        st.error("🚫 Las contraseñas no coinciden.")
                    else:
                        salt = auth.generar_salt()
                        bd.crear_primer_admin(
                            admin_usuario_key, admin_nombre.strip(),
                            auth.hash_password(admin_pass1, salt), salt,
                        )
                        _cargar_usuarios_cacheado.clear()
                        st.success("✅ Cuenta creada. Ingresa con tu usuario y contraseña.")
                        st.rerun()
                return False

            with st.form("form_login", clear_on_submit=False):
                usuario_input = st.text_input("Usuario")
                password_input = st.text_input("Contraseña", type="password")
                enviar = st.form_submit_button("Ingresar", use_container_width=True)

            if enviar:
                usuario_key = usuario_input.strip().lower()
                datos_usuario = usuarios.get(usuario_key)
                if not datos_usuario:
                    st.error("🚫 Usuario o contraseña incorrectos.")
                elif not datos_usuario.get("password_hash"):
                    st.error(
                        "🔑 Todavía no has creado tu contraseña. Abre \"¿Primera vez? "
                        "Crea tu contraseña\" aquí abajo."
                    )
                elif not auth.verificar_password(
                    password_input, datos_usuario.get("password_salt"), datos_usuario.get("password_hash")
                ):
                    st.error("🚫 Usuario o contraseña incorrectos.")
                elif _sesion_esta_activa(usuario_key):
                    st.error(
                        f"🚫 La cuenta **{usuario_key}** ya tiene una sesión abierta en otro "
                        "navegador o pestaña. Ciérrala primero, o espera unos minutos si se "
                        "quedó abierta sin usar el botón \"Cerrar sesión\"."
                    )
                else:
                    session_id = str(uuid.uuid4())
                    _registrar_sesion(usuario_key, session_id, datos_usuario["nombre"])
                    st.session_state.autenticado = True
                    st.session_state.usuario_actual = usuario_key
                    st.session_state.rol_actual = datos_usuario["rol"]
                    st.session_state.nombre_actual = datos_usuario["nombre"]
                    st.session_state._session_id = session_id
                    contenedor_login.empty()
                    return True

            # --- CASO 2: primera vez de un usuario ya dado de alta por el Admin ---
            with st.expander("¿Primera vez? Crea tu contraseña"):
                st.caption(
                    "Solo funciona si el Admin ya te dio de alta como usuario. "
                    "Nadie más ve la contraseña que elijas aquí, ni el Admin."
                )
                with st.form("form_crear_password", clear_on_submit=False):
                    nuevo_usuario = st.text_input("Tu usuario (el que te dio el Admin)")
                    nueva_pass1 = st.text_input("Elige tu contraseña (mínimo 8 caracteres)", type="password")
                    nueva_pass2 = st.text_input("Repítela", type="password")
                    crear_pass = st.form_submit_button("Crear mi contraseña", use_container_width=True)
                if crear_pass:
                    nuevo_usuario_key = nuevo_usuario.strip().lower()
                    datos_pendiente = usuarios.get(nuevo_usuario_key)
                    if not datos_pendiente:
                        st.error("🚫 Ese usuario no está dado de alta. Pídele al Admin que te registre primero.")
                    elif datos_pendiente.get("password_hash"):
                        st.error("🚫 Esa cuenta ya tiene contraseña. Si la olvidaste, pide al Admin que te la restablezca.")
                    elif not auth.contrasena_es_valida(nueva_pass1):
                        st.error("🚫 La contraseña debe tener al menos 8 caracteres.")
                    elif nueva_pass1 != nueva_pass2:
                        st.error("🚫 Las contraseñas no coinciden.")
                    else:
                        salt = auth.generar_salt()
                        bd.establecer_password_usuario(
                            nuevo_usuario_key, auth.hash_password(nueva_pass1, salt), salt
                        )
                        _cargar_usuarios_cacheado.clear()
                        st.success("✅ Contraseña creada. Ciérra este cuadro e ingresa arriba.")

    return False


if not st.session_state.autenticado:
    if not _render_pantalla_login():
        st.stop()

# Mientras la sesión siga autenticada, cada ejecución del script (cualquier
# clic o interacción) renueva el "latido" para que no expire por inactividad
# mientras la persona la sigue usando activamente. Si el registro ya no
# corresponde a esta pestaña (alguien más volvió a entrar con este usuario,
# p. ej. tras un timeout), se cierra esta sesión de forma automática.
_usuario_sesion_actual = st.session_state.get("usuario_actual")
_session_id_actual = st.session_state.get("_session_id")
_registro_vigente = _sesiones_activas().get(_usuario_sesion_actual)
if _registro_vigente and _registro_vigente["session_id"] != _session_id_actual:
    for _clave in ("autenticado", "usuario_actual", "rol_actual", "nombre_actual", "_session_id"):
        st.session_state.pop(_clave, None)
    st.error(
        "🚫 Tu sesión se cerró porque este usuario inició sesión en otro navegador o pestaña."
    )
    st.stop()
_refrescar_latido_sesion(_usuario_sesion_actual, _session_id_actual)

# ---------------------------------------------------------
# AJUSTES GLOBALES CONTROLADOS SOLO POR EL ADMIN
# ---------------------------------------------------------
# "Modo ahorro de tokens" y "Validar vigencia en tiempo real" dejan de ser
# un checkbox por sesión (cada Analista los veía siempre en su valor por
# defecto) y pasan a ser un ajuste COMPARTIDO: solo el Admin lo cambia, y
# todos los Analistas heredan ese mismo valor al auditar o consultar la IA,
# aunque en su barra lateral no vean el control.
@st.cache_resource
def _ajustes_globales():
    return {"modo_ahorro_tokens": True, "validar_estatus_sat_tiempo_real": True}


with st.sidebar:
    st.markdown(f"""
        <div class="sidebar-user-box">
            👤 <b>{st.session_state.get('nombre_actual', 'Usuario')}</b><br>
            <span class="sidebar-user-role">{st.session_state.get('rol_actual', 'N/D')}</span>
        </div>
    """, unsafe_allow_html=True)
    if st.button("🔒 Cerrar sesión", use_container_width=True):
        _cerrar_sesion()
        st.rerun()
    st.markdown("---")

ES_ADMIN = st.session_state.get("rol_actual") == "Admin"

col_logo, col_title = st.columns([1, 4])

with col_logo:
    logo_path = "logo_autocom.png"
    if os.path.exists(logo_path):
        st.image(logo_path, use_container_width=True)
    else:
        st.markdown("<h2 style='color:#C8102E; font-weight:bold;'>AUTOCOM</h2>", unsafe_allow_html=True)

with col_title:
    st.markdown("""
        <div class="header-card">
            <h1 style="color: #1F2937; margin:0; font-size: 2rem;">🚗 Sistema Auditor CFDI 4.0</h1>
            <p style="color: #4B5563; margin: 4px 0 8px 0;">Plataforma de Validación Fiscal de Comprobantes - Sector Automotriz</p>
            <span class="author-badge">⚙️ Desarrollado por: <b>Juan Rogelio Cruz García</b></span>
        </div>
    """, unsafe_allow_html=True)

# ---------------------------------------------------------
# 1.5 CAPA DE IA — GRATIS, BAJO DEMANDA
# ---------------------------------------------------------

MODELS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "models")
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = os.environ.get("GROQ_MODEL", "openai/gpt-oss-120b")


def _obtener_groq_api_key():
    """Prioriza st.secrets (obligatorio en Streamlit Community Cloud, donde
    no se puede "subir" groq_api.txt). Si corres en tu PC y no has
    configurado .streamlit/secrets.toml todavía, cae de regreso al archivo
    viejo para no romper el flujo local."""
    try:
        clave = st.secrets.get("GROQ_API_KEY", "")
        if clave:
            return str(clave).strip()
    except Exception:
        pass
    ruta = os.path.join(os.path.dirname(os.path.abspath(__file__)), "groq_api.txt")
    try:
        with open(ruta, "r", encoding="utf-8") as f:
            return f.read().strip()
    except (OSError, UnicodeError):
        return ""


_ia_lock = threading.Lock()

# ---------------------------------------------------------
# SANITIZACIÓN DEL XML ANTES DE ENVIARLO A LA IA
# ---------------------------------------------------------
# Cuando la IA corre en la nube (Groq), el XML sale de la red de AUTOCOM.
# Antes de enviarlo se le quitan los datos que identifican a la empresa y
# los bloques que no aportan nada al dictamen fiscal pero sí ocupan miles
# de tokens (sellos y certificados en base64).
#
# Qué se ELIMINA / ENMASCARA:
#   - Receptor: Rfc, Nombre y DomicilioFiscalReceptor de AUTOCOM.
#   - Sello, Certificado y NoCertificado del Comprobante.
#   - SelloCFD, SelloSAT y NoCertificadoSAT del Timbre Fiscal Digital.
#   - UUID (se deja solo el prefijo, suficiente para referenciar).
#   - Nodo Addenda completo (suele traer datos internos del proveedor
#     y del cliente que no son fiscales).
#
# Qué se CONSERVA (es justo lo que hay que auditar):
#   - Emisor completo (es el proveedor, no AUTOCOM): RFC, nombre, régimen.
#   - Todos los atributos fiscales: Version, Fecha, MetodoPago, FormaPago,
#     Moneda, TipoDeComprobante, SubTotal, Descuento, Total, UsoCFDI.
#   - Todos los Conceptos con clave, descripción, cantidad, unidad,
#     importe, ObjetoImp y sus impuestos.
#   - Todos los nodos de Impuestos (traslados y retenciones).
ATRIBUTOS_A_ELIMINAR = (
    "Sello", "Certificado", "NoCertificado",
    "SelloCFD", "SelloSAT", "NoCertificadoSAT", "RfcProvCertif",
)

NODOS_A_ELIMINAR = ("Addenda",)


def sanitizar_xml_para_ia(xml_content, enmascarar_receptor=True):
    """
    Devuelve (xml_limpio: str, nota: str) — el XML listo para enviarse a la
    IA, sin datos de la empresa ni bloques criptográficos.

    `enmascarar_receptor=False` solo tiene sentido con IA local (el XML
    nunca sale del equipo); con IA en la nube siempre debe ir en True.
    """
    try:
        if isinstance(xml_content, bytes):
            xml_content = xml_content.decode('utf-8-sig', errors='ignore')

        root = ET.fromstring(xml_content)

        for elem in root.iter():
            if '}' in elem.tag:
                elem.tag = elem.tag.split('}', 1)[1]

        # Quitar bloques completos que no son fiscales.
        for nombre_nodo in NODOS_A_ELIMINAR:
            for padre in root.iter():
                for hijo in list(padre):
                    if hijo.tag == nombre_nodo:
                        padre.remove(hijo)

        # Quitar sellos y certificados (base64 enorme, sin valor fiscal aquí).
        for elem in root.iter():
            for atributo in ATRIBUTOS_A_ELIMINAR:
                elem.attrib.pop(atributo, None)

        # Acortar el UUID: basta el prefijo para poder referenciarlo.
        tfd = root.find('.//TimbreFiscalDigital')
        if tfd is not None and 'UUID' in tfd.attrib:
            tfd.attrib['UUID'] = f"{tfd.attrib['UUID'][:8]}-[...]"

        notas = []
        if enmascarar_receptor:
            receptor = root.find('Receptor')
            if receptor is not None:
                if 'Rfc' in receptor.attrib:
                    receptor.attrib['Rfc'] = "XAXX010101000"
                if 'Nombre' in receptor.attrib:
                    receptor.attrib['Nombre'] = "RECEPTOR (DATOS PROTEGIDOS)"
                receptor.attrib.pop('DomicilioFiscalReceptor', None)
                notas.append("datos del receptor enmascarados")

        xml_limpio = ET.tostring(root, encoding='unicode')
        notas.append("sellos y certificados removidos")
        return xml_limpio, "XML saneado: " + ", ".join(notas) + "."

    except ET.ParseError as e:
        return "", f"No se pudo leer el XML para sanearlo: {e}"
    except Exception as e:
        return "", f"Error al sanear el XML: {type(e).__name__}: {e}"


SYSTEM_PROMPT_AUDITORIA_XML = (
    "Eres un auditor fiscal senior especializado en CFDI 4.0 mexicano, revisando facturas de "
    "proveedores para AUTOCOM (grupo automotriz: agencias, talleres, refacciones, activos fijos).\n"
    "Recibirás el XML COMPLETO de un CFDI (con el receptor enmascarado y los sellos removidos por "
    "política de datos: NO señales eso como error, es intencional) y un bloque de "
    "'HECHOS YA VERIFICADOS por el motor de reglas'.\n\n"
    "REGLA DE ORO: los HECHOS YA VERIFICADOS son de un motor determinista en Python (cuenta "
    "caracteres, aplica catálogos oficiales del SAT, calcula tasas) — son más confiables que tu "
    "propio cálculo. NUNCA generes un hallazgo que contradiga ese bloque. Si algo ahí dice "
    "'correcto' o 'no aplica', no lo reportes como problema aunque tu intuición diga lo contrario. "
    "Enfoca tu análisis en lo que ESE bloque NO cubre: coherencia semántica entre clave ProdServ y "
    "descripción, unidad de medida razonable para el producto, plausibilidad de cantidades y "
    "precios, redacción ambigua, y cualquier otra irregularidad que un motor de reglas no detecta "
    "por ser criterio de especialista y no una fórmula.\n\n"
    "Errores comunes que DEBES evitar (ya se te dan resueltos en el bloque de hechos, no los repitas "
    "por tu cuenta):\n"
    "- NO asumas el tipo de persona del emisor (Física/Moral) por el régimen fiscal: regímenes como "
    "601, 603, 620, 623, 624 y 626 (RESICO) aplican TANTO a Persona Física COMO a Persona Moral. El "
    "tipo de persona se determina ÚNICAMENTE por la longitud del RFC (13 caracteres = Física, 12 = "
    "Moral), y eso ya viene resuelto en el bloque de hechos — nunca lo recalcules tú mismo contando "
    "caracteres del XML.\n"
    "- NO afirmes que falta una retención de ISR/IVA sin antes revisar si el bloque de hechos ya "
    "dice que la retención no aplica o que sí está correcta.\n\n"
    "Revisa al menos, siempre respetando el bloque de hechos:\n"
    "1. Estructura y versión: Version debe ser 4.0; TipoDeComprobante coherente con el contenido.\n"
    "2. UsoCFDI: coherente con lo que se está comprando. Bienes de inversión/activo fijo exigen I01-I08; "
    "gastos generales G03; adquisición de mercancías G01.\n"
    "3. ClaveProdServ de cada concepto: ¿la clave declarada corresponde REALMENTE a lo que dice la "
    "descripción? Señala cada concepto donde la clave no empate con el producto o servicio descrito.\n"
    "4. ClaveUnidad y Unidad: coherentes con el concepto (ej. un producto que normalmente se vende "
    "por pieza pero se factura por litro, o viceversa).\n"
    "5. Cualquier otra irregularidad fiscal relevante que detectes aunque no esté en esta lista, "
    "SIEMPRE que no contradiga el bloque de hechos ya verificados.\n\n"
    "Criterio de severidad:\n"
    "- CRITICO: impide deducir o acreditar, o expone a AUTOCOM ante el SAT.\n"
    "- MEDIO: debe corregirse con el proveedor pero no bloquea el pago por sí solo (clave ProdServ "
    "imprecisa, unidad de medida inadecuada).\n"
    "- MENOR: observación de forma.\n\n"
    "Reglas de respuesta:\n"
    "- Responde ÚNICAMENTE con el objeto JSON, sin texto antes ni después, sin razonamiento visible.\n"
    "- Sé concreto: cita el atributo o el concepto exacto donde está el problema. Nada de generalidades.\n"
    "- Si algo te parece correcto, NO lo reportes como hallazgo. La lista de hallazgos es solo de "
    "problemas reales.\n"
    "- Si no encuentras ningún problema, devuelve \"hallazgos\": [] y veredicto ACEPTADO.\n"
    "- No inventes cifras ni atributos que no estén en el XML. Si un dato no viene, dilo como hallazgo "
    "de dato faltante en vez de suponerlo.\n\n"
    "Esquema JSON obligatorio:\n"
    "{\n"
    '  "veredicto": "ACEPTADO" | "REVISION" | "RECHAZADO",\n'
    '  "resumen": "Dictamen global en máximo 25 palabras",\n'
    '  "hallazgos": [\n'
    "    {\n"
    '      "severidad": "CRITICO" | "MEDIO" | "MENOR",\n'
    '      "rubro": "Retenciones" | "UsoCFDI" | "FormaPago" | "ClaveProdServ" | "Impuestos" | '
    '"Aritmetica" | "Estructura" | "Otro",\n'
    '      "detalle": "Qué está mal y en qué atributo o concepto exacto (máx. 30 palabras)",\n'
    '      "accion": "Qué debe hacer CxP o qué pedirle al proveedor (máx. 15 palabras)"\n'
    "    }\n"
    "  ]\n"
    "}"
)

def _construir_system_prompt_ia(categorias_dinamicas=None):
    """
    Arma el prompt de clasificación por concepto incluyendo, además de las
    categorías fijas, las categorías DINÁMICAS ya aprobadas (para que la IA
    las prefiera si encajan) y le da permiso explícito de proponer una
    categoría de negocio NUEVA cuando ninguna de las existentes describe
    bien el concepto — esa propuesta queda sujeta a aprobación del Admin
    antes de volverse permanente (ver pestaña "🔔 Notificaciones"), nunca se
    aplica sola.
    """
    categorias_dinamicas = categorias_dinamicas or {}
    nombres_dinamicas = list(categorias_dinamicas.keys())
    bloque_dinamicas = (
        ("\nCategorías dinámicas ya aprobadas por el Admin (trátalas igual que las fijas de arriba, "
         "úsalas si el concepto encaja): " + ", ".join(nombres_dinamicas) + ".\n")
        if nombres_dinamicas else ""
    )

    return (
        "Eres un auditor fiscal auxiliar experto en CFDI 4.0 y el catálogo de claves del SAT (México), "
        "trabajando para AUTOCOM (grupo del sector automotriz: agencias, talleres, refacciones y activos fijos).\n"
        "Debes responder EXCLUSIVAMENTE en formato JSON válido, sin texto explicativo adicional fuera del objeto.\n"
        "Categorías fijas del motor de reglas: "
        "ACTIVO_FIJO, REFACCION, HONORARIO_PROFESIONAL, FLETE, COMBUSTIBLE, SEGUROS, SOFTWARE_TI, "
        "PUBLICIDAD_MERCADOTECNIA, ARRENDAMIENTO, CAPACITACION o GENERAL. Respeta estrictamente la misma "
        "lógica que usaría el motor de reglas fiscales de AUTOCOM (reglas_sat.py): un concepto solo puede ser "
        "ACTIVO_FIJO si es un bien de uso duradero (no refacción, no mantenimiento) y de importe relevante; "
        "FLETE es transporte/autotransporte de carga; HONORARIO_PROFESIONAL son servicios profesionales, "
        "legales, contables o de consultoría (NO tecnología ni publicidad, esas tienen su propia categoría); "
        "COMBUSTIBLE es gasolina/diésel/lubricantes de flotilla; SEGUROS son pólizas/primas/deducibles; "
        "SOFTWARE_TI es licencias, suscripciones, hosting o soporte técnico (incluye el DMS Quiter); "
        "PUBLICIDAD_MERCADOTECNIA es campañas, redes sociales, rotulación o material promocional; "
        "ARRENDAMIENTO es renta de local, vehículos de flotilla/cortesía o equipo; CAPACITACION son cursos "
        "o certificaciones de marca/fabricante para personal.\n"
        f"{bloque_dinamicas}\n"
        "Si el concepto NO encaja razonablemente en NINGUNA categoría existente (ni fija ni dinámica) y "
        "tampoco es un gasto genérico (GENERAL), puedes PROPONER una categoría de negocio nueva: usa un "
        "nombre corto en MAYÚSCULAS, una sola palabra o dos unidas con guion bajo (ej. AGUA, "
        "PAPELERIA, LIMPIEZA) — nunca una frase larga. No abuses de esto: solo propón una categoría "
        "nueva cuando el concepto sea claramente un tipo de gasto recurrente y distinto a los ya "
        "cubiertos, no para casos aislados que caben perfectamente en GENERAL. Esta propuesta NUNCA se "
        "aplica sola: siempre queda pendiente de que un Admin la apruebe.\n\n"
        "También recibirás el MetodoPago (PUE/PPD) y el FormaPago declarados del CFDI: úsalos solo como "
        "contexto adicional para tu dictamen (por ejemplo, un pago en efectivo de un importe muy alto puede "
        "ser una señal a mencionar en la justificación) — la validación formal de que FormaPago sea "
        "coherente con MetodoPago ya la hace el motor de reglas por separado, no la repitas como motivo "
        "de la clasificación del concepto.\n\n"
        "Además del veredicto de clasificación, entrega un dictamen MUY breve y directo para el área de "
        "Contabilidad / Cuentas por Pagar de AUTOCOM: nada de rodeos, cortesías ni repetir datos que ya "
        "se te dieron (descripción, clave, importe).\n\n"
        "Reglas de respuesta:\n"
        "- Responde únicamente con la estructura JSON solicitada, en una sola línea, sin razonamiento "
        "ni texto adicional antes o después del objeto JSON.\n"
        "- Si no tienes certeza suficiente, usa \"confianza\": \"baja\" y \"veredicto\": \"REVISION\".\n"
        "- \"justificacion\": UNA sola oración corta (máx. 12 palabras) con el motivo puntual: si el "
        "veredicto es ACEPTADO, di por qué está bien; si es RECHAZADO o REVISION, di qué está mal o "
        "qué falta. Nada de explicaciones largas ni contexto adicional.\n"
        "- \"recomendacion\": UNA sola acción concreta en máx. 10 palabras, en modo imperativo "
        "(ej. \"Solicitar Uso CFDI de Inversión al proveedor.\", \"Aplicar y continuar con CxP.\").\n\n"
        "Esquema JSON obligatorio:\n"
        "{\n"
        '  "categoria": "ACTIVO_FIJO" | "REFACCION" | "HONORARIO_PROFESIONAL" | "FLETE" | "COMBUSTIBLE" | '
        '"SEGUROS" | "SOFTWARE_TI" | "PUBLICIDAD_MERCADOTECNIA" | "ARRENDAMIENTO" | "CAPACITACION" | '
        '"GENERAL" | "<CATEGORIA_NUEVA_PROPUESTA>",\n'
        '  "confianza": "alta" | "media" | "baja",\n'
        '  "motivo": "Explicación breve de la clasificación en menos de 15 palabras",\n'
        '  "veredicto": "ACEPTADO" | "REVISION" | "RECHAZADO",\n'
        '  "justificacion": "Motivo puntual en máx. 12 palabras",\n'
        '  "recomendacion": "Acción concreta en máx. 10 palabras"\n'
        "}"
    )


def _listar_modelos_gguf():
    if not os.path.isdir(MODELS_DIR):
        return []
    return sorted(f for f in os.listdir(MODELS_DIR) if f.lower().endswith(".gguf"))


@st.cache_resource(show_spinner="Cargando modelo de IA local...")
def _cargar_modelo_ia_local(nombre_archivo):
    if not nombre_archivo:
        return None
    ruta_modelo = os.path.join(MODELS_DIR, nombre_archivo)
    if not os.path.exists(ruta_modelo):
        return None
    try:
        from llama_cpp import Llama
        return Llama(
            model_path=ruta_modelo,
            n_ctx=2048,
            n_threads=os.cpu_count() or 4,
            verbose=False,
        )
    except Exception:
        return None


def _construir_mensaje_ia(descripcion, clave_prodserv, uso_cfdi, importe, forma_pago="N/A", metodo_pago="N/A"):
    return (
        f"Descripción del concepto: {descripcion}\n"
        f"Clave ProdServ declarada: {clave_prodserv}\n"
        f"Uso de CFDI declarado: {uso_cfdi}\n"
        f"Método de Pago (MetodoPago) declarado: {metodo_pago}\n"
        f"Forma de Pago (FormaPago) declarada: {forma_pago}\n"
        f"Importe del concepto: ${importe:,.2f} MXN"
    )


def _resultado_error_ia(mensaje, codigo=None, cuota_agotada=False):
    return {
        "_ok": False,
        "_error": mensaje,
        "_codigo": codigo,
        "_cuota_agotada": cuota_agotada,
    }


def _derivar_veredicto_de_categoria(categoria, confianza):
    """Respaldo cuando la IA (sobre todo un modelo local más limitado) no
    devuelve 'veredicto' explícito: lo derivamos de categoría/confianza,
    sin inventar una lógica fiscal distinta a la ya usada en la app."""
    if confianza == "baja":
        return "REVISION"
    if categoria == "GENERAL":
        return "REVISION"
    return "ACEPTADO"


_PATRON_NOMBRE_CATEGORIA_NUEVA = re.compile(r'^[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ_]{1,29}$')


def _validar_respuesta_ia(analisis):
    # Se toma de reglas_sat.obtener_categorias_validas (única fuente de verdad
    # de categorías fiscales, incluyendo las dinámicas dadas de alta en
    # caliente) para que la IA nunca acepte una categoría que el motor de
    # reglas no reconozca, ni al revés. Si la IA propone un nombre NUEVO que
    # no está en esta lista, se acepta igual (ver `_construir_system_prompt_ia`)
    # siempre que tenga forma de nombre de categoría válido — nunca se aplica
    # solo: queda marcado como propuesta pendiente de aprobación del Admin.
    categorias_existentes = set(sat.obtener_categorias_validas(CATEGORIAS_DINAMICAS))
    confianzas_validas = {"alta", "media", "baja"}
    veredictos_validos = {"ACEPTADO", "REVISION", "RECHAZADO"}
    if not isinstance(analisis, dict):
        return _resultado_error_ia("La IA devolvió una respuesta que no es un objeto JSON.")
    categoria = str(analisis.get("categoria", "")).strip().upper().replace(" ", "_")
    confianza = str(analisis.get("confianza", "")).strip().lower()
    motivo = str(analisis.get("motivo", "")).strip()

    categoria_es_nueva = categoria not in categorias_existentes
    if categoria_es_nueva and not _PATRON_NOMBRE_CATEGORIA_NUEVA.match(categoria):
        return _resultado_error_ia(f"La IA devolvió una categoría no válida: {categoria or 'vacía'}.")
    if confianza not in confianzas_validas:
        return _resultado_error_ia(f"La IA devolvió una confianza no válida: {confianza or 'vacía'}.")
    if not motivo:
        return _resultado_error_ia("La IA no devolvió el motivo de la clasificación.")

    veredicto = str(analisis.get("veredicto", "")).strip().upper()
    if veredicto not in veredictos_validos:
        veredicto = _derivar_veredicto_de_categoria(categoria, confianza)
    justificacion = str(analisis.get("justificacion", "")).strip() or motivo
    recomendacion = str(analisis.get("recomendacion", "")).strip() or (
        "Enviar a revisión manual antes de aplicar." if veredicto != "ACEPTADO" else
        "Aplicar y continuar con CxP."
    )

    return {
        "categoria": categoria,
        "categoria_es_nueva": categoria_es_nueva,
        "confianza": confianza,
        "motivo": motivo,
        "veredicto": veredicto,
        "justificacion": justificacion,
        "recomendacion": recomendacion,
        "_ok": True,
    }


# ---------------------------------------------------------
# CAPA DE TRANSPORTE HACIA GROQ (una sola implementación)
# ---------------------------------------------------------
# Antes había dos bloques casi idénticos hablando con Groq. Ahora todo pasa
# por aquí, para que las tres mejoras importantes apliquen siempre:
#
#   1. reasoning_effort="low" en modelos razonadores (gpt-oss, qwen3,
#      deepseek-r1). Esos modelos gastan CIENTOS de tokens "pensando" antes
#      de escribir el JSON, y ese pensamiento se cobra y consume max_tokens.
#      Con esfuerzo bajo el costo cae fuerte y, de paso, se corrige la causa
#      #1 del error HTTP 400 "Failed to validate JSON": el modelo razonaba
#      tanto que el JSON salía cortado a la mitad y el validador de Groq lo
#      rechazaba.
#   2. Reintento automático sin `response_format` cuando Groq devuelve ese
#      400. Sin el validador estricto, Groq sí entrega el texto y nosotros
#      extraemos el JSON a mano.
#   3. Rescate del JSON que viene dentro de `failed_generation` del propio
#      error 400: muchas veces el dictamen ya estaba bien hecho y solo le
#      faltaba una llave de cierre.

_MODELOS_RAZONADORES = ("gpt-oss", "qwen3", "deepseek-r1", "-r1", "reason")


def _modelo_es_razonador(nombre_modelo=None):
    nombre = (nombre_modelo or GROQ_MODEL or "").lower()
    return any(marca in nombre for marca in _MODELOS_RAZONADORES)


def _extraer_json_tolerante(texto):
    """Extrae un objeto JSON de una respuesta de IA aunque venga con texto
    alrededor, dentro de ```json ... ``` o truncado a media llave.
    Devuelve el dict o None."""
    if not texto:
        return None
    texto = str(texto).strip()
    if texto.startswith("```"):
        texto = re.sub(r"^```[a-zA-Z]*\s*", "", texto)
        texto = re.sub(r"```\s*$", "", texto).strip()

    inicio = texto.find("{")
    if inicio == -1:
        return None
    candidato = texto[inicio:]

    try:
        return json.loads(candidato)
    except (TypeError, json.JSONDecodeError):
        pass

    # Reparación de truncado: se cierran comillas/corchetes/llaves abiertos.
    reparado = candidato
    if reparado.count('"') % 2 == 1:
        reparado += '"'
    reparado = re.sub(r",\s*$", "", reparado.rstrip())
    faltan_corchetes = reparado.count("[") - reparado.count("]")
    faltan_llaves = reparado.count("{") - reparado.count("}")
    reparado += "]" * max(0, faltan_corchetes) + "}" * max(0, faltan_llaves)
    try:
        return json.loads(reparado)
    except (TypeError, json.JSONDecodeError):
        return None


def _detectar_cuota_agotada(status_code, error_detalle):
    texto_error = str(error_detalle).lower()
    return status_code == 429 or any(
        palabra in texto_error
        for palabra in (
            "rate limit", "rate_limit", "quota", "tokens per day",
            "tpd", "tokens per minute", "tpm", "insufficient_quota",
        )
    )


def _payload_groq(system_prompt, mensaje_usuario, max_tokens, forzar_json=True):
    payload = {
        "model": GROQ_MODEL,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": mensaje_usuario},
        ],
        "temperature": 0.0,
        "max_tokens": max_tokens,
    }
    if _modelo_es_razonador():
        # Menos tokens de razonamiento = menos gasto y menos JSON truncado.
        payload["reasoning_effort"] = "low"
    if forzar_json:
        payload["response_format"] = {"type": "json_object"}
    return payload


def _peticion_groq(system_prompt, mensaje_usuario, api_key, max_tokens, timeout):
    """
    Hace la petición a Groq y devuelve (analisis_dict, error_dict).
    Exactamente uno de los dos es None.
    """
    headers = {
        "Authorization": f"Bearer {api_key.strip()}",
        "Content-Type": "application/json",
    }

    for intento, forzar_json in enumerate((True, False)):
        try:
            resp = requests.post(
                GROQ_API_URL,
                headers=headers,
                json=_payload_groq(system_prompt, mensaje_usuario, max_tokens, forzar_json),
                timeout=timeout,
            )
        except requests.exceptions.Timeout:
            return None, _resultado_error_ia("Se agotó el tiempo de espera al consultar Groq.")
        except requests.exceptions.ConnectionError:
            return None, _resultado_error_ia("No se pudo conectar con Groq. Revisa tu conexión a internet.")
        except requests.exceptions.RequestException as e:
            return None, _resultado_error_ia(f"Error de comunicación con Groq: {type(e).__name__}.")

        if resp.ok:
            try:
                contenido = resp.json()["choices"][0]["message"]["content"]
            except (KeyError, IndexError, TypeError, ValueError):
                return None, _resultado_error_ia(
                    "Groq respondió, pero la estructura de la respuesta no fue la esperada."
                )
            analisis = _extraer_json_tolerante(contenido)
            if analisis is not None:
                return analisis, None
            if forzar_json:
                continue  # se reintenta en modo libre
            return None, _resultado_error_ia("Groq respondió, pero el contenido no era un JSON válido.")

        # --- respuesta de error ---
        try:
            cuerpo_error = resp.json().get("error", {})
        except Exception:
            cuerpo_error = {}
        error_detalle = cuerpo_error.get("message", resp.text) if isinstance(cuerpo_error, dict) else resp.text

        if _detectar_cuota_agotada(resp.status_code, error_detalle):
            return None, _resultado_error_ia(
                "Se agotaron los tokens/cuota disponibles de la IA en Groq por ahora "
                f"(HTTP {resp.status_code}): {error_detalle}",
                resp.status_code, cuota_agotada=True,
            )

        # Groq devuelve el texto que el modelo alcanzó a generar en
        # `failed_generation`: si ahí ya venía el dictamen, se rescata en vez
        # de perder la llamada (y los tokens que ya se pagaron).
        if isinstance(cuerpo_error, dict) and cuerpo_error.get("failed_generation"):
            rescatado = _extraer_json_tolerante(cuerpo_error["failed_generation"])
            if rescatado is not None:
                return rescatado, None

        es_error_de_validacion_json = resp.status_code == 400 and (
            "failed to validate json" in str(error_detalle).lower()
            or "json_validate_failed" in str(cuerpo_error).lower()
        )
        if es_error_de_validacion_json and forzar_json:
            continue  # segundo intento sin response_format

        return None, _resultado_error_ia(
            f"Groq rechazó la consulta (HTTP {resp.status_code}): {error_detalle}",
            resp.status_code,
        )

    return None, _resultado_error_ia("Groq no devolvió un JSON utilizable tras dos intentos.")


def _llamar_ia_nube_groq(mensaje_usuario, api_key, timeout=20):
    analisis, error = _peticion_groq(
        _construir_system_prompt_ia(CATEGORIAS_DINAMICAS),
        mensaje_usuario,
        api_key,
        max_tokens=1024,
        timeout=timeout,
    )
    if error is not None:
        return error
    return _validar_respuesta_ia(analisis)


def _llamar_ia_local(mensaje_usuario, modelo_local):
    if modelo_local is None:
        return _resultado_error_ia("El modelo local no está cargado.")
    try:
        with _ia_lock:
            respuesta = modelo_local.create_chat_completion(
                messages=[
                    {"role": "system", "content": _construir_system_prompt_ia(CATEGORIAS_DINAMICAS)},
                    {"role": "user", "content": mensaje_usuario},
                ],
                temperature=0.0,
                max_tokens=400,
                response_format={"type": "json_object"},
            )
        try:
            contenido = respuesta["choices"][0]["message"]["content"]
        except (KeyError, IndexError, TypeError):
            return _resultado_error_ia("El modelo local respondió, pero la estructura no fue la esperada.")
        try:
            analisis = json.loads(contenido)
        except (TypeError, json.JSONDecodeError):
            return _resultado_error_ia("El modelo local respondió, pero el contenido no era un JSON válido.")
        return _validar_respuesta_ia(analisis)
    except Exception as e:
        return _resultado_error_ia(f"Error del modelo local: {type(e).__name__}.")


def _validar_respuesta_auditoria_xml(analisis):
    """Valida la respuesta de la auditoría de XML completo (esquema distinto
    al de clasificación de un concepto suelto)."""
    severidades_validas = {"CRITICO", "MEDIO", "MENOR"}
    veredictos_validos = {"ACEPTADO", "REVISION", "RECHAZADO"}

    if not isinstance(analisis, dict):
        return _resultado_error_ia("La IA devolvió una respuesta que no es un objeto JSON.")

    veredicto = str(analisis.get("veredicto", "")).strip().upper()
    resumen = str(analisis.get("resumen", "")).strip()
    hallazgos_crudos = analisis.get("hallazgos", [])

    if not isinstance(hallazgos_crudos, list):
        return _resultado_error_ia("La IA devolvió 'hallazgos' en un formato que no es una lista.")

    hallazgos = []
    for item in hallazgos_crudos:
        if not isinstance(item, dict):
            continue
        severidad = str(item.get("severidad", "")).strip().upper()
        if severidad not in severidades_validas:
            severidad = "MEDIO"
        detalle = str(item.get("detalle", "")).strip()
        if not detalle:
            continue
        hallazgos.append({
            "severidad": severidad,
            "rubro": str(item.get("rubro", "Otro")).strip() or "Otro",
            "detalle": detalle,
            "accion": str(item.get("accion", "")).strip() or "Revisar con el proveedor.",
        })

    # Si la IA no dio un veredicto usable, se deriva de la severidad de lo
    # que ella misma encontró, sin inventar una lógica fiscal distinta.
    if veredicto not in veredictos_validos:
        if any(h["severidad"] == "CRITICO" for h in hallazgos):
            veredicto = "RECHAZADO"
        elif hallazgos:
            veredicto = "REVISION"
        else:
            veredicto = "ACEPTADO"

    if not resumen:
        resumen = (
            f"Se detectaron {len(hallazgos)} hallazgo(s) en el comprobante."
            if hallazgos else "Sin hallazgos relevantes en la revisión del XML."
        )

    return {
        "veredicto": veredicto,
        "resumen": resumen,
        "hallazgos": hallazgos,
        "_ok": True,
    }


def _llamar_ia_generica(mensaje_usuario, system_prompt, backend, modelo_local=None,
                        max_tokens=1500, timeout=45):
    """
    Envía un prompt arbitrario al backend de IA elegido y devuelve el JSON ya
    parseado (sin validar el esquema: eso lo hace quien llama). Reutiliza la
    misma infraestructura de errores que el resto de la app.
    """
    if backend == "local":
        if modelo_local is None:
            return _resultado_error_ia("El modelo local no está cargado.")
        try:
            with _ia_lock:
                respuesta = modelo_local.create_chat_completion(
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": mensaje_usuario},
                    ],
                    temperature=0.0,
                    max_tokens=max_tokens,
                    response_format={"type": "json_object"},
                )
            contenido = respuesta["choices"][0]["message"]["content"]
        except (KeyError, IndexError, TypeError):
            return _resultado_error_ia("El modelo local respondió, pero la estructura no fue la esperada.")
        except Exception as e:
            return _resultado_error_ia(f"Error del modelo local: {type(e).__name__}.")
    else:
        api_key = _obtener_groq_api_key()
        if not api_key:
            return _resultado_error_ia("No hay API Key de Groq configurada.")
        analisis, error = _peticion_groq(
            system_prompt, mensaje_usuario, api_key,
            max_tokens=max_tokens, timeout=timeout,
        )
        if error is not None:
            # Se adjunta lo que se envió para que el analista pueda verlo en
            # "🔎 Ver datos de la consulta" (antes salía vacío).
            error["_datos_enviados"] = mensaje_usuario
            return error
        return {"_ok": True, "_json": analisis, "_datos_enviados": mensaje_usuario}

    analisis = _extraer_json_tolerante(contenido)
    if analisis is None:
        return _resultado_error_ia("La IA respondió, pero el contenido no era un JSON válido.")
    return {"_ok": True, "_json": analisis, "_datos_enviados": mensaje_usuario}


# ---------------------------------------------------------
# AHORRO DE TOKENS
# ---------------------------------------------------------
# Un CFDI en XML crudo gasta entre 3,000 y 8,000 tokens solo de entrada,
# porque el modelo paga por cada nombre de atributo, cada comilla y cada
# etiqueta de cierre. El dictamen no necesita nada de eso: necesita los
# DATOS. `_compactar_xml_para_ia` reescribe el mismo comprobante en formato
# de líneas clave=valor, conservando todos los atributos fiscales pero
# recortando entre 60% y 75% de los tokens de entrada.

_ATRIBUTOS_IRRELEVANTES_IA = {
    "xmlns", "schemaLocation", "Version", "FechaTimbrado", "Leyenda",
    "RfcProvCertif", "NoCertificado", "Certificado", "Sello",
}


def _compactar_xml_para_ia(xml_saneado):
    """Convierte el XML saneado en texto compacto clave=valor. Si por alguna
    razón no se puede parsear, devuelve el XML original sin romper nada."""
    try:
        root = ET.fromstring(xml_saneado)
    except Exception:
        return xml_saneado

    for elem in root.iter():
        if "}" in elem.tag:
            elem.tag = elem.tag.split("}", 1)[1]

    def atributos(elem, omitir=()):
        return " ".join(
            f"{k}={v}"
            for k, v in elem.attrib.items()
            if k not in _ATRIBUTOS_IRRELEVANTES_IA and k not in omitir and str(v).strip()
        )

    lineas = [f"COMPROBANTE {atributos(root)}"]

    emisor = root.find("Emisor")
    if emisor is not None:
        lineas.append(f"EMISOR {atributos(emisor)}")
    receptor = root.find("Receptor")
    if receptor is not None:
        lineas.append(f"RECEPTOR {atributos(receptor)}")

    conceptos = root.find("Conceptos")
    if conceptos is not None:
        for i, concepto in enumerate(conceptos.findall("Concepto"), start=1):
            lineas.append(f"CONCEPTO#{i} {atributos(concepto)}")
            for traslado in concepto.iter("Traslado"):
                lineas.append(f"  CONCEPTO#{i}.Traslado {atributos(traslado)}")
            for retencion in concepto.iter("Retencion"):
                lineas.append(f"  CONCEPTO#{i}.Retencion {atributos(retencion)}")

    impuestos = root.find("Impuestos")
    if impuestos is not None:
        lineas.append(f"IMPUESTOS_TOTALES {atributos(impuestos)}")
        for traslado in impuestos.iter("Traslado"):
            lineas.append(f"  Traslado_total {atributos(traslado)}")
        for retencion in impuestos.iter("Retencion"):
            lineas.append(f"  Retencion_total {atributos(retencion)}")

    for nodo in root.iter():
        if nodo.tag in ("CfdiRelacionados", "CfdiRelacionado", "InformacionAduanera", "Parte"):
            lineas.append(f"{nodo.tag} {atributos(nodo)}")

    compacto = "\n".join(l for l in lineas if l.strip())
    # Si por lo que sea el compactado saliera más largo, se usa el original.
    return compacto if len(compacto) < len(xml_saneado) else xml_saneado


# Antes vivía en cache_auditorias_ia.json (se perdía en cada redespliegue
# de Streamlit Cloud). Ahora se guarda en Supabase — ver bd_supabase.py.
def _huella_auditoria(xml_saneado, contexto_reglas, modo_ahorro):
    import hashlib
    base = f"{GROQ_MODEL}|{modo_ahorro}|{contexto_reglas}|{xml_saneado}"
    return hashlib.sha256(base.encode("utf-8", errors="ignore")).hexdigest()[:32]


def _cache_auditoria_leer(huella):
    try:
        return bd.cache_auditoria_leer(huella)
    except Exception:
        return None


def _cache_auditoria_guardar(huella, dictamen, maximo_entradas=500):
    try:
        bd.cache_auditoria_guardar(huella, dictamen, maximo_entradas=maximo_entradas)
    except Exception:
        pass


@st.cache_data(show_spinner=False, ttl=3600)
def auditar_xml_completo_ia(xml_saneado, backend, contexto_reglas="", modo_ahorro=True,
                            _modelo_local=None):
    """
    Manda el XML COMPLETO (ya saneado) a la IA para un dictamen de
    especialista fiscal sobre todo el comprobante, no solo un concepto.

    Se cachea por contenido del XML: auditar dos veces la misma factura no
    vuelve a gastar tokens.
    """
    if not xml_saneado:
        return _resultado_error_ia("No hay XML para auditar.")

    # Caché EN DISCO: sobrevive a reinicios de Streamlit, a diferencia del
    # @st.cache_data (que se pierde al reiniciar). Con ~100 XML diarios y
    # reprocesos frecuentes del mismo lote, esto es lo que más tokens ahorra.
    huella = _huella_auditoria(xml_saneado, contexto_reglas, modo_ahorro)
    en_cache = _cache_auditoria_leer(huella)
    if en_cache is not None:
        en_cache["_desde_cache"] = True
        return en_cache

    cuerpo, nota_formato = (
        (_compactar_xml_para_ia(xml_saneado), "Comprobante (formato compacto, mismos datos fiscales del XML):")
        if modo_ahorro else
        (xml_saneado, "XML del comprobante:")
    )

    # Límite defensivo: un CFDI con cientos de conceptos podría exceder el
    # contexto del modelo. Se trunca avisando, en vez de fallar en silencio.
    limite_caracteres = 10000 if modo_ahorro else 24000
    aviso_truncado = ""
    if len(cuerpo) > limite_caracteres:
        cuerpo = cuerpo[:limite_caracteres]
        aviso_truncado = (
            "\n\n[AVISO: el comprobante fue truncado por tamaño. Dictamina solo sobre la parte "
            "visible y señálalo como hallazgo MENOR.]"
        )

    mensaje_usuario = (
        "Audita este CFDI completo y reporta únicamente los problemas reales que encuentres.\n\n"
        f"{contexto_reglas}\n\n"
        f"{nota_formato}\n"
        f"{cuerpo}"
        f"{aviso_truncado}"
    )

    respuesta = _llamar_ia_generica(
        mensaje_usuario,
        SYSTEM_PROMPT_AUDITORIA_XML,
        backend,
        modelo_local=_modelo_local,
        max_tokens=1200 if modo_ahorro else 2000,
        timeout=60,
    )
    if not respuesta.get("_ok"):
        return respuesta
    dictamen = _validar_respuesta_auditoria_xml(respuesta["_json"])
    dictamen["_datos_enviados"] = mensaje_usuario
    if dictamen.get("_ok"):
        _cache_auditoria_guardar(huella, dictamen)
    return dictamen


@st.cache_data(show_spinner=False, ttl=3600)
def analizar_concepto_ambiguo_ia(descripcion, clave_prodserv, uso_cfdi, importe, backend,
                                  forma_pago="N/A", metodo_pago="N/A", _modelo_local=None):
    mensaje_usuario = _construir_mensaje_ia(
        descripcion, clave_prodserv, uso_cfdi, importe, forma_pago=forma_pago, metodo_pago=metodo_pago
    )
    if backend == "nube":
        api_key = _obtener_groq_api_key()
        if not api_key:
            return _resultado_error_ia("No se encontró la API key de Groq en groq_api.txt.")
        resultado = _llamar_ia_nube_groq(mensaje_usuario, api_key)
    elif backend == "local":
        resultado = _llamar_ia_local(mensaje_usuario, _modelo_local)
    else:
        resultado = _resultado_error_ia("No hay ningún backend de IA disponible.")

    if isinstance(resultado, dict):
        resultado["_backend"] = "Groq (nube)" if backend == "nube" else "Modelo local"
        resultado["_modelo"] = GROQ_MODEL if backend == "nube" else "Modelo .gguf local"
        resultado["_datos_enviados"] = mensaje_usuario
    return resultado


_ICONOS_VEREDICTO = {"ACEPTADO": "🟢", "REVISION": "🟡", "RECHAZADO": "🔴"}
_ICONOS_CONFIANZA = {"alta": "🟢", "media": "🟡", "baja": "🔴"}


def _render_dictamen_ia(sugerencia, mostrar_datos_enviados=True):
    """Muestra el dictamen de la IA como un mensaje concreto y accionable
    (veredicto + justificación técnico-fiscal + recomendación para
    Contabilidad / CxP de AUTOCOM), en vez de JSON crudo."""
    confianza = str(sugerencia.get("confianza", "n/d")).lower()
    veredicto = str(sugerencia.get("veredicto", "REVISION")).upper()
    tipo_badge = {"ACEPTADO": "ok", "REVISION": "warn", "RECHAZADO": "danger"}.get(veredicto, "neutral")
    icono_veredicto = _ICONOS_VEREDICTO.get(veredicto, "⚪")
    icono_conf = _ICONOS_CONFIANZA.get(confianza, "⚪")
    es_nueva = sugerencia.get("categoria_es_nueva")

    st.success("🧠 Consulta completada correctamente")
    st.markdown(
        f'{_badge(veredicto, tipo_badge, icono_veredicto)} &nbsp; '
        f'{_badge(sugerencia["categoria"] + (" (nueva)" if es_nueva else ""), "neutral", "🏷️")}',
        unsafe_allow_html=True,
    )
    if es_nueva:
        st.caption("Esta categoría no existe todavía — si la aplicas, se propondrá al administrador para su aprobación.")
    st.write(f"**Justificación técnica/fiscal:** {sugerencia.get('justificacion', sugerencia.get('motivo', ''))}")
    st.write(f"**Recomendación para Contabilidad / CxP:** {sugerencia.get('recomendacion', '')}")
    st.caption(
        f"{icono_conf} Confianza de la IA: {confianza.upper()} · "
        f"Backend: {sugerencia.get('_backend', 'n/d')} · Modelo: {sugerencia.get('_modelo', 'n/d')}"
    )
    if confianza == "baja" or veredicto != "ACEPTADO":
        st.warning("La IA recomienda revisión manual antes de aplicar la clasificación.")

    if mostrar_datos_enviados:
        with st.expander("🔎 Ver qué analizó la IA"):
            st.code(sugerencia.get("_datos_enviados", ""), language="text")


_ICONOS_SEVERIDAD = {"CRITICO": "🔴", "MEDIO": "🟠", "MENOR": "🟡"}


def construir_hechos_verificados(fila_cab, conceptos_de_esta_factura):
    """
    Bloque de "verdad terreno" calculado en Python (determinista, sin
    posibilidad de alucinación) que se le entrega a la IA junto con el XML.
    Cubre justo los puntos donde un LLM se equivoca con frecuencia: contar
    caracteres de un RFC, saber a qué tipo de persona aplica cada régimen, y
    aplicar tasas de retención exactas.

    La IA recibe la instrucción explícita de no contradecir este bloque —
    ver SYSTEM_PROMPT_AUDITORIA_XML — y de concentrarse en lo que el motor
    de reglas no puede evaluar (coherencia semántica, unidades, etc.).
    """
    regimen_texto = fila_cab.get("Régimen Fiscal Emisor", "N/A")
    regimen_codigo = regimen_texto.split(" - ")[0].strip() if " - " in regimen_texto else regimen_texto

    rfc_emisor = fila_cab.get("_rfc_emisor", "")
    es_pf = len(rfc_emisor) == 13 if rfc_emisor else None
    tipo_persona = "PERSONA FÍSICA" if es_pf else ("PERSONA MORAL" if es_pf is False else "NO DETERMINADO")

    aplica_ambos = regimen_codigo in ("601", "603", "620", "623", "624", "626")

    lineas = [
        "HECHOS YA VERIFICADOS por el motor de reglas de AUTOCOM (deterministas, NO los recalcules "
        "ni generes un hallazgo que los contradiga):",
        f"- RFC del emisor: {rfc_emisor or 'N/A'} ({len(rfc_emisor) if rfc_emisor else '?'} caracteres) "
        f"→ tipo de persona: {tipo_persona}.",
        f"- Régimen fiscal declarado: {regimen_texto}."
        + (" Este régimen aplica tanto a Persona Física como a Persona Moral: NO es evidencia de "
           "inconsistencia que un RFC de Persona Moral lo use." if aplica_ambos else ""),
    ]

    categorias_conceptos = conceptos_de_esta_factura["Categoría"].tolist() if conceptos_de_esta_factura is not None and not conceptos_de_esta_factura.empty else []
    es_flete = "FLETE" in categorias_conceptos
    retenciones_ok, obs_retenciones = sat.validar_retenciones(
        fila_cab.get("_retenciones_xml", {}), regimen_codigo, bool(es_pf),
        categorias_factura=tuple(categorias_conceptos), es_flete=es_flete,
    )
    if es_pf is None:
        lineas.append("- Retenciones ISR/IVA: no se pudo determinar el tipo de persona; verifícalo tú manualmente.")
    elif retenciones_ok:
        lineas.append(
            "- Retenciones ISR/IVA: CORRECTAS conforme a la normativa vigente para este emisor y estos "
            "conceptos. No reportes ningún hallazgo de retención faltante o incorrecta."
        )
    else:
        for obs in obs_retenciones:
            lineas.append(f"- Retención con problema (repórtalo tal cual, no lo reformules): {obs}")

    forma_pago_ok, motivo_forma_pago = sat.validar_forma_pago_vs_metodo_pago(
        fila_cab.get("_metodo_pago_raw", ""), fila_cab.get("_forma_pago_raw", "")
    )
    if forma_pago_ok:
        lineas.append("- Coherencia MetodoPago/FormaPago: CORRECTA. No reportes hallazgo sobre esto.")
    else:
        lineas.append(f"- Coherencia MetodoPago/FormaPago con problema (repórtalo tal cual): {motivo_forma_pago}")

    if categorias_conceptos:
        lineas.append(
            "- Categoría fiscal que el motor de reglas ya asignó a cada concepto (tómala como válida; "
            "tu labor es revisar coherencia semántica de clave/unidad, no reclasificar): "
            + "; ".join(f"concepto #{i+1} → {c}" for i, c in enumerate(categorias_conceptos))
        )

    return "\n".join(lineas)


def _contexto_reglas_para_ia():
    """Le pasa a la IA los parámetros fiscales VIGENTES de AUTOCOM (los de
    reglas_sat.py), para que dictamine con las mismas cifras que el motor de
    reglas y no con las que traiga memorizadas de su entrenamiento."""
    return (
        "Parámetros fiscales vigentes de AUTOCOM (úsalos como referencia obligatoria):\n"
        f"- Importe mínimo por concepto para tratarlo como Activo Fijo: ${sat.UMBRAL_ACTIVO_FIJO:,.2f} MXN.\n"
        f"- Retención ISR RESICO Persona Física (régimen 626): {sat.TASA_RET_ISR_RESICO_PF:.4%}.\n"
        f"- Retención ISR servicios profesionales PF: {sat.TASA_RET_ISR_SERVICIOS_PF:.2%}.\n"
        f"- Retención ISR arrendamiento PF (régimen 606): {sat.TASA_RET_ISR_ARRENDAMIENTO_PF:.2%}.\n"
        f"- Retención IVA servicios profesionales / arrendamiento PF: {sat.TASA_RET_IVA_SERVICIOS_PF:.4%}.\n"
        f"- Retención IVA autotransporte terrestre de carga (fletes): {sat.TASA_RET_IVA_FLETES:.2%}.\n"
        "- Facturas PUE de meses anteriores no se aceptan: deben refacturarse como PPD.\n"
        "- El receptor viene enmascarado y los sellos removidos por política de datos; NO es un error."
    )


def _render_auditoria_xml_ia(dictamen):
    """Muestra el dictamen de la auditoría del XML completo: veredicto
    global + lista de hallazgos ordenados por severidad."""
    if not dictamen.get("_ok"):
        _render_error_ia(dictamen)
        return

    veredicto = str(dictamen.get("veredicto", "REVISION")).upper()
    tipo_badge = {"ACEPTADO": "ok", "REVISION": "warn", "RECHAZADO": "danger"}.get(veredicto, "neutral")
    icono = _ICONOS_VEREDICTO.get(veredicto, "⚪")
    hallazgos = dictamen.get("hallazgos", [])

    st.markdown(f"#### Dictamen del XML completo &nbsp; {_badge(veredicto, tipo_badge, icono)}", unsafe_allow_html=True)
    st.write(dictamen.get("resumen", ""))

    if not hallazgos:
        st.success("✅ La IA no encontró irregularidades fiscales en el comprobante.")
        return

    orden = {"CRITICO": 0, "MEDIO": 1, "MENOR": 2}
    hallazgos = sorted(hallazgos, key=lambda h: orden.get(h["severidad"], 3))

    criticos = sum(1 for h in hallazgos if h["severidad"] == "CRITICO")
    if criticos:
        st.error(f"🔴 {criticos} hallazgo(s) CRÍTICO(s): no programar el pago sin resolverlos.")

    st.markdown(f"**{len(hallazgos)} hallazgo(s):**")
    tipo_por_severidad = {"CRITICO": "danger", "MEDIO": "warn", "MENOR": "neutral"}
    for h in hallazgos:
        icono_sev = _ICONOS_SEVERIDAD.get(h["severidad"], "⚪")
        st.markdown(
            f'{_badge(h["severidad"], tipo_por_severidad.get(h["severidad"], "neutral"), icono_sev)} '
            f'**{h["rubro"]}** — {h["detalle"]}  \n'
            f'↳ *Acción:* {h["accion"]}',
            unsafe_allow_html=True,
        )

    st.caption(
        "Este dictamen es apoyo de la IA, no sustituye al motor de reglas ni el criterio del analista. "
        "Verifica cada hallazgo contra el XML antes de rechazar la factura."
    )


def _render_error_ia(sugerencia):
    """Muestra el error de una consulta a la IA. Si el error es por
    agotamiento de tokens/cuota, lo notifica de forma clara y distinta a un
    error genérico, para que el usuario sepa que debe esperar o revisar su
    plan, en vez de pensar que la consulta o el concepto están mal."""
    if sugerencia.get("_cuota_agotada"):
        st.error(
            "🚫 **Se agotaron los tokens/cuota disponibles de la IA.** "
            "Espera a que se restablezca el límite (Groq) o revisa tu plan/API key antes de "
            "seguir consultando. Mientras tanto, puedes clasificar este concepto manualmente."
        )
    else:
        st.error("🔴 La consulta a la IA terminó con error")
    st.write(f"**Qué ocurrió:** {sugerencia.get('_error', 'No se obtuvo una respuesta válida.')}")
    if sugerencia.get("_codigo"):
        st.caption(f"Código HTTP: {sugerencia['_codigo']}")
    with st.expander("🔎 Ver datos de la consulta"):
        st.code(sugerencia.get("_datos_enviados", ""), language="text")


# Sidebar de IA
# La DETECCIÓN del backend (Groq / modelo local) corre siempre, para
# cualquier rol: los Analistas SÍ usan la IA en la pestaña "Análisis
# Avanzado", solo que no ven ni tocan estos controles en la barra lateral.
_groq_key = _obtener_groq_api_key()
modelos_gguf = _listar_modelos_gguf()
modelo_local_cargado = None

if _groq_key:
    ia_backend = "nube"
    ia_disponible = True
elif modelos_gguf:
    modelo_local_cargado = _cargar_modelo_ia_local(modelos_gguf[0])
    if modelo_local_cargado is not None:
        ia_backend = "local"
        ia_disponible = True
    else:
        ia_backend = None
        ia_disponible = False
else:
    ia_backend = None
    ia_disponible = False

_ajustes = _ajustes_globales()

if ES_ADMIN:
    with st.sidebar:
        st.subheader("🧠 IA para conceptos (bajo demanda)")
        st.caption("La IA se activa solo cuando seleccionas un concepto individual.")

        if ia_backend == "nube":
            st.success(f"☁️ IA en la nube activa (Groq · modelo `{GROQ_MODEL}`)")
        elif ia_backend == "local":
            st.success(f"💻 IA local activa (`{modelos_gguf[0]}`)")
        elif modelos_gguf and modelo_local_cargado is None:
            st.error("Se encontró un modelo local pero no se pudo cargar.")
        else:
            st.warning("IA no configurada. Agrega groq_api.txt o un modelo .gguf.")

        # Ajuste GLOBAL: lo que el Admin decida aquí aplica para todos los
        # Analistas, aunque ellos no vean este control.
        _ajustes["modo_ahorro_tokens"] = st.checkbox(
            "💰 Modo ahorro de tokens",
            value=_ajustes.get("modo_ahorro_tokens", True),
            key="_modo_ahorro_tokens",
            help=(
                "Envía el comprobante en formato compacto en vez del XML crudo y limita el "
                "razonamiento del modelo. Reduce entre 60% y 75% los tokens por auditoría sin "
                "quitarle ningún dato fiscal. Desactívalo solo si necesitas que la IA vea el XML "
                "literal. Este ajuste es global: aplica también para los Analistas."
            ),
        )
        try:
            st.caption(f"♻️ {bd.contar_cache_auditorias()} dictamen(es) en caché (no vuelven a gastar tokens).")
        except Exception:
            pass
else:
    # El Analista no ve ni toca este panel: solo hereda el valor global.
    st.session_state["_modo_ahorro_tokens"] = _ajustes.get("modo_ahorro_tokens", True)


# ---------------------------------------------------------
# 2. AUDITORÍA DE CFDI XML
# ---------------------------------------------------------

ICONOS_ESTATUS = {
    "ACEPTADO": "✅",
    "REVISIÓN MANUAL": "🟡",
    "RECHAZADO DE ENTRADA": "🔴",
    "FACTURA CANCELADA": "⛔",
}

# ---------------------------------------------------------
# 1.6 OPINIÓN DEL ANALISTA (acuerdo / desacuerdo por concepto)
# ---------------------------------------------------------
# Bitácora persistente (no solo de la sesión) de cuándo un analista está de
# acuerdo o en desacuerdo con lo que dictaminó el motor de reglas o la IA
# para un concepto concreto, y por qué. Es la base para: (a) dejar rastro de
# auditoría de quién anuló qué y con qué justificación, y (b) alimentar,
# bajo revisión humana, al motor de reglas (pestaña "🔔 Notificaciones").
# Antes vivía en decisiones_analista.csv (se perdía en cada redespliegue de
# Streamlit Cloud). Ahora se guarda en Supabase — ver bd_supabase.py.
COLUMNAS_DECISIONES_ANALISTA = bd.COLUMNAS_DECISIONES_ANALISTA


def _cargar_decisiones_analista():
    try:
        filas = bd.cargar_decisiones_analista()
    except Exception:
        filas = []
    if not filas:
        return pd.DataFrame(columns=COLUMNAS_DECISIONES_ANALISTA)
    return pd.DataFrame(filas, columns=COLUMNAS_DECISIONES_ANALISTA).fillna("")


def _guardar_decision_analista(registro):
    """Agrega una fila a la bitácora sin borrar el historial previo — si el
    mismo concepto ya tenía una opinión registrada, ambas quedan (se conserva
    quién opinó qué y cuándo, en vez de sobrescribir)."""
    bd.agregar_decision_analista(registro)


def _id_decision_concepto(folio, descripcion, importe):
    return f"{folio}::{descripcion}::{importe}"


# ---------------------------------------------------------
# SISTEMA DE NOTIFICACIONES: propuestas de categoría pendientes de Admin
# ---------------------------------------------------------
# Reemplaza la antigua pestaña "Enseñar Reglas" (instrucción en texto libre).
# Ahora, cada vez que un analista da clic en "✅ Aplicar clasificación de la
# IA", se genera automáticamente una propuesta aquí — de categoría nueva o
# de refuerzo de una existente — que un Admin debe aceptar o rechazar en la
# pestaña "🔔 Notificaciones" antes de que quede aprendida de forma
# permanente en categorias_dinamicas.json. Nada se aplica sin esa aprobación.
# Antes vivía en propuestas_categorias.json (se perdía en cada redespliegue
# de Streamlit Cloud). Ahora se guarda en Supabase — ver bd_supabase.py.
def _cargar_propuestas_categorias():
    try:
        return bd.cargar_propuestas_categorias()
    except Exception:
        return []


def _guardar_propuestas_categorias(propuestas):
    bd.guardar_propuestas_categorias(propuestas)


def _crear_propuesta_categoria(categoria, es_nueva, palabras_sugeridas, uso_cfdi_sugerido,
                               descripcion_categoria, contexto):
    propuestas = _cargar_propuestas_categorias()
    nueva = {
        "id": f"{int(time.time() * 1000)}",
        "categoria": categoria,
        "es_nueva": es_nueva,
        "palabras_sugeridas": palabras_sugeridas,
        "uso_cfdi_sugerido": uso_cfdi_sugerido,
        "descripcion_categoria": descripcion_categoria,
        "estado": "pendiente",
        "fecha_creacion": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "fecha_resolucion": None,
        "resuelto_por": None,
        **contexto,
    }
    propuestas.append(nueva)
    _guardar_propuestas_categorias(propuestas)
    return nueva


def _marcar_propuesta_resuelta(id_propuesta, estado):
    propuestas = _cargar_propuestas_categorias()
    for p in propuestas:
        if p.get("id") == id_propuesta:
            p["estado"] = estado
            p["fecha_resolucion"] = datetime.now().strftime("%Y-%m-%d %H:%M")
            p["resuelto_por"] = st.session_state.get("nombre_actual", "N/D")
    _guardar_propuestas_categorias(propuestas)


def _contar_propuestas_pendientes():
    return sum(1 for p in _cargar_propuestas_categorias() if p.get("estado") == "pendiente")


SYSTEM_PROMPT_EXPANSION_CATEGORIA = (
    "Eres un experto en catálogos fiscales de un grupo automotriz (agencias, talleres, refacciones, "
    "activos fijos). Se te da UN concepto real de una factura y la categoría a la que un analista lo "
    "asignó. Tu trabajo es generar una lista de palabras o frases CORTAS, en MAYÚSCULAS y tal como "
    "aparecerían escritas en la descripción de una factura real, que permitan reconocer AUTOMÁTICAMENTE "
    "en el futuro otros conceptos SIMILARES de esa misma categoría — no solo este concepto exacto, sino "
    "la categoría completa que representa. Incluye sinónimos, singular/plural, y variantes comerciales "
    "genéricas razonables de la industria automotriz. NO inventes marcas de un solo proveedor ni "
    "información que no puedas justificar. También indica a qué Uso CFDI del catálogo SAT corresponde "
    "normalmente esta categoría (valores válidos: G01, G02, G03, I01, I02, I03, I04, I05, I06, I07, I08, "
    "D01, D10).\n"
    "Responde ÚNICAMENTE con este JSON, sin texto adicional:\n"
    "{\n"
    '  "palabras": ["PALABRA O FRASE 1", "PALABRA O FRASE 2", "..."],\n'
    '  "uso_cfdi": ["G03"],\n'
    '  "descripcion_categoria": "Una frase corta (máx. 12 palabras) que resuma qué cubre esta categoría"\n'
    "}"
)


def _expandir_categoria_con_ia(categoria, descripcion_concepto, clave_prodserv, backend, modelo_local):
    """
    Le pide a la IA una lista amplia de palabras relacionadas con el
    concepto que el analista acaba de aceptar, para que la categoría (nueva
    o reforzada) reconozca MÁS de un concepto desde el primer día — no solo
    el que originó la propuesta.

    Si la IA falla por cualquier motivo, hace un respaldo mínimo con la
    descripción tal cual, para que la propuesta nunca se quede vacía.
    """
    mensaje = (
        f'Categoría asignada: "{categoria}"\n'
        f'Concepto real de la factura: "{descripcion_concepto}"\n'
        f'Clave ProdServ declarada: {clave_prodserv}'
    )
    respuesta = _llamar_ia_generica(
        mensaje, SYSTEM_PROMPT_EXPANSION_CATEGORIA, backend,
        modelo_local=modelo_local, max_tokens=500, timeout=30,
    )
    if not respuesta.get("_ok"):
        return {
            "palabras": [descripcion_concepto.strip().upper()],
            "uso_cfdi": ["G03"],
            "descripcion_categoria": "",
        }

    datos = respuesta.get("_json", {}) if isinstance(respuesta.get("_json"), dict) else {}
    palabras = [str(p).strip().upper() for p in datos.get("palabras", []) if str(p).strip()]
    if not palabras:
        palabras = [descripcion_concepto.strip().upper()]
    uso_cfdi = [str(u).strip().upper() for u in datos.get("uso_cfdi", []) if str(u).strip()] or ["G03"]
    return {
        "palabras": palabras,
        "uso_cfdi": uso_cfdi,
        "descripcion_categoria": str(datos.get("descripcion_categoria", "")).strip(),
    }


# ---------------------------------------------------------
# LA IA TRADUCE LA OPINIÓN DEL ANALISTA A UNA REGLA
# ---------------------------------------------------------
# El problema anterior: el analista escribía en lenguaje natural por qué no
# estaba de acuerdo ("esta clave es de mantenimiento, no de activo fijo") y
# el código no sabía qué hacer con ese texto — como mucho lo guardaba en la
# bitácora y le pedía al analista que él mismo tecleara el nombre de la
# categoría. Ahora la IA lee el comentario, lo interpreta y arma la
# propuesta de regla completa. Sigue sin aplicarse sola: el Admin aprueba.

SYSTEM_PROMPT_OPINION_ANALISTA = (
    "Eres el motor de aprendizaje de reglas fiscales de AUTOCOM (grupo automotriz). Un analista de "
    "Cuentas por Pagar revisó un concepto de una factura y escribió, en lenguaje natural, por qué "
    "NO está de acuerdo con la clasificación o el veredicto del sistema.\n\n"
    "Tu trabajo es traducir esa opinión a una PROPUESTA DE REGLA concreta y accionable para el "
    "motor de reglas (reglas_sat.py). No opines si el analista tiene razón: interpreta lo que pide "
    "y conviértelo en parámetros. Si el comentario es vago o no propone nada aplicable, devuelve "
    '"aplicable": false y explica qué falta.\n\n'
    "Criterios:\n"
    "- \"categoria_propuesta\": nombre corto en MAYÚSCULAS, una palabra o dos unidas con guion bajo "
    "(ej. MANTENIMIENTO, REFACCION, AGUA). Prefiere una categoría existente si el analista está "
    "describiendo una que ya existe.\n"
    "- \"palabras_clave\": 8 a 15 términos en MAYÚSCULAS que deberían disparar esa categoría, "
    "derivados del concepto y del comentario. Sin acentos, sin frases largas.\n"
    "- \"uso_cfdi_sugerido\": lista de claves de UsoCFDI válidas para esa categoría (G01, G03, "
    "I01-I08, D01...).\n"
    "- \"tipo_cambio\": \"CATEGORIA_NUEVA\" si propone un tipo de gasto que no existe; "
    "\"REFUERZO_CATEGORIA\" si solo quiere que un concepto caiga en una categoría ya existente; "
    "\"AJUSTE_UMBRAL\" si lo que pide es cambiar un monto o porcentaje.\n"
    "- \"resumen_regla\": UNA oración en imperativo describiendo la regla, máx. 20 palabras "
    "(ej. \"Clasificar mangueras y refacciones menores como REFACCION aunque superen el umbral de "
    "activo fijo.\").\n"
    "- \"riesgo_fiscal\": \"bajo\" | \"medio\" | \"alto\" — qué tan riesgoso sería aplicar esta "
    "regla sin revisión adicional.\n\n"
    "Responde ÚNICAMENTE con el objeto JSON, sin texto antes ni después:\n"
    "{\n"
    '  "aplicable": true | false,\n'
    '  "categoria_propuesta": "NOMBRE",\n'
    '  "tipo_cambio": "CATEGORIA_NUEVA" | "REFUERZO_CATEGORIA" | "AJUSTE_UMBRAL",\n'
    '  "palabras_clave": ["..."],\n'
    '  "uso_cfdi_sugerido": ["G03"],\n'
    '  "descripcion_categoria": "Qué gastos abarca, máx. 20 palabras",\n'
    '  "resumen_regla": "Regla en imperativo, máx. 20 palabras",\n'
    '  "riesgo_fiscal": "bajo" | "medio" | "alto",\n'
    '  "motivo_no_aplicable": "Solo si aplicable=false: qué le falta al comentario, máx. 15 palabras"\n'
    "}"
)


def interpretar_opinion_analista_ia(comentario, descripcion_concepto, clave_prodserv,
                                    categoria_sistema, backend, modelo_local=None,
                                    categorias_existentes=None):
    """Lee el comentario libre del analista y devuelve una propuesta de regla
    ya estructurada, lista para mandarse al Admin."""
    if not str(comentario).strip():
        return _resultado_error_ia("No hay comentario del analista que interpretar.")

    listado = ", ".join(sorted(categorias_existentes or [])) or "ninguna"
    mensaje = (
        f'Concepto de la factura: "{descripcion_concepto}"\n'
        f"Clave ProdServ declarada: {clave_prodserv}\n"
        f"Categoría que asignó el sistema: {categoria_sistema}\n"
        f"Categorías ya existentes en el motor de reglas: {listado}\n"
        f'Opinión textual del analista (está EN DESACUERDO): "{str(comentario).strip()}"'
    )

    respuesta = _llamar_ia_generica(
        mensaje, SYSTEM_PROMPT_OPINION_ANALISTA, backend,
        modelo_local=modelo_local, max_tokens=700, timeout=40,
    )
    if not respuesta.get("_ok"):
        return respuesta

    datos = respuesta.get("_json", {}) if isinstance(respuesta.get("_json"), dict) else {}
    categoria = str(datos.get("categoria_propuesta", "")).strip().upper().replace(" ", "_")
    aplicable = bool(datos.get("aplicable", True)) and bool(categoria)

    palabras = [str(p).strip().upper() for p in datos.get("palabras_clave", []) if str(p).strip()]
    if not palabras:
        palabras = [str(descripcion_concepto).strip().upper()]
    usos = [str(u).strip().upper() for u in datos.get("uso_cfdi_sugerido", []) if str(u).strip()] or ["G03"]

    return {
        "_ok": True,
        "aplicable": aplicable,
        "categoria_propuesta": categoria,
        "tipo_cambio": str(datos.get("tipo_cambio", "REFUERZO_CATEGORIA")).strip().upper(),
        "palabras_clave": palabras,
        "uso_cfdi_sugerido": usos,
        "descripcion_categoria": str(datos.get("descripcion_categoria", "")).strip(),
        "resumen_regla": str(datos.get("resumen_regla", "")).strip(),
        "riesgo_fiscal": str(datos.get("riesgo_fiscal", "medio")).strip().lower(),
        "motivo_no_aplicable": str(datos.get("motivo_no_aplicable", "")).strip(),
        "_datos_enviados": mensaje,
    }


@st.cache_data(show_spinner="Descargando listado de EFOS del SAT...", ttl=43200)
def _obtener_listado_69b_cacheado():
    return sat.descargar_listado_69b()


@st.cache_data(show_spinner=False, ttl=600)
def _consultar_estatus_sat_cacheado(rfc_emisor, rfc_receptor, total, uuid, sello):
    return sat.consultar_estatus_cfdi_sat(rfc_emisor, rfc_receptor, total, uuid, sello)


def obtener_listado_69b_bajo_demanda():
    """
    Descarga (o reutiliza del caché de Streamlit) el listado 69-B SOLO
    cuando alguien realmente va a auditar. Antes de esto, la app nunca
    debe bloquearse por una descarga que el usuario ni siquiera pidió
    todavía (ej. justo al terminar de loguearse).
    """
    try:
        listado = _obtener_listado_69b_cacheado()
        st.session_state["_listado_69b_ok"] = True
        st.session_state["_listado_69b_len"] = len(listado)
        return listado
    except Exception as e:
        st.session_state["_listado_69b_ok"] = False
        st.session_state["_listado_69b_error"] = str(e)
        return {}


@st.cache_resource(show_spinner="Cargando catálogo oficial del SAT...")
def _cargar_catalogo_sat_cacheado(marca_de_tiempo_archivo):
    """
    Carga el catálogo completo de Claves ProdServ a memoria UNA sola vez por
    sesión/proceso. El parámetro `marca_de_tiempo_archivo` no se usa dentro:
    existe para que forme parte de la llave del caché de Streamlit, de modo
    que cuando sat_catalogos.db se reemplaza con la versión nueva del mes,
    Streamlit detecta el cambio y recarga el diccionario solo, sin que nadie
    tenga que reiniciar la app.

    Se devuelve un dict {clave: texto}; todas las búsquedas posteriores son
    en memoria (instantáneas), nunca una query por concepto.
    """
    return sat.cargar_catalogo_prodserv()


def obtener_catalogo_sat():
    """Devuelve (catalogo, mensaje) usando la fecha de modificación del
    archivo como llave de caché."""
    ruta = sat.ruta_catalogo_por_defecto()
    try:
        marca = os.path.getmtime(ruta)
    except OSError:
        marca = 0.0
    return _cargar_catalogo_sat_cacheado(marca)


def actualizar_catalogo_sat_si_toca(permitir_descarga=True):
    """
    Ejecuta el flujo de arranque: si ya pasó un mes (o no existe la base),
    intenta descargar la versión nueva desde GitHub; si no hay internet,
    sigue con la copia local. Luego carga a memoria.
    """
    ruta = sat.ruta_catalogo_por_defecto()

    if permitir_descarga and sat.catalogo_necesita_actualizacion(ruta):
        with st.spinner("Buscando actualización del catálogo SAT..."):
            exito, mensaje = sat.descargar_catalogo_sat(ruta)
        st.session_state["_catalogo_mensaje_descarga"] = mensaje
        if exito:
            _cargar_catalogo_sat_cacheado.clear()

    return obtener_catalogo_sat()


# Flujo de arranque del diagrama: se resuelve una vez al iniciar la app.
if "_catalogo_sat_inicializado" not in st.session_state:
    CATALOGO_SAT, CATALOGO_SAT_MENSAJE = actualizar_catalogo_sat_si_toca()
    st.session_state["_catalogo_sat_inicializado"] = True
    st.session_state["_catalogo_sat_mensaje"] = CATALOGO_SAT_MENSAJE
else:
    CATALOGO_SAT, CATALOGO_SAT_MENSAJE = obtener_catalogo_sat()


with st.sidebar:
    st.markdown("---")
    st.subheader("📚 Catálogo oficial SAT (ProdServ)")

    if CATALOGO_SAT:
        st.success(f"✅ {len(CATALOGO_SAT):,} claves cargadas en memoria.")
    else:
        st.warning(
            "⚠️ No hay catálogo oficial cargado; se está usando el catálogo interno "
            "reducido y muchas claves saldrán como NO CATALOGADA."
        )

    st.caption(CATALOGO_SAT_MENSAJE)
    if st.session_state.get("_catalogo_mensaje_descarga"):
        st.caption(st.session_state["_catalogo_mensaje_descarga"])

    if st.button("🔄 Actualizar catálogo SAT ahora"):
        _cargar_catalogo_sat_cacheado.clear()
        with st.spinner("Descargando catálogo del SAT..."):
            exito, mensaje = sat.descargar_catalogo_sat(sat.ruta_catalogo_por_defecto())
        st.session_state["_catalogo_mensaje_descarga"] = mensaje
        st.rerun()

    st.caption(
        f"El catálogo se revisa automáticamente cada {sat.DIAS_ENTRE_ACTUALIZACIONES_CATALOGO} días "
        "al iniciar la app. Si no hay internet, se sigue trabajando con la copia local."
    )


# ---------------------------------------------------------
# CATEGORÍAS DINÁMICAS (aprendidas en caliente, aprobadas por Admin)
# ---------------------------------------------------------
@st.cache_data(show_spinner=False, ttl=60)
def _cargar_categorias_dinamicas_cacheado():
    """Ahora las categorías viven en Supabase, no en un archivo local, así que
    ya no hay 'marca de tiempo de archivo' que vigilar: se refresca sola
    cada 60s, y de inmediato en cuanto se acepta una propuesta gracias al
    _cargar_categorias_dinamicas_cacheado.clear() de más abajo."""
    return sat.cargar_categorias_dinamicas()


def obtener_categorias_dinamicas():
    return _cargar_categorias_dinamicas_cacheado()


CATEGORIAS_DINAMICAS = obtener_categorias_dinamicas()

if ES_ADMIN:
    with st.sidebar:
        st.markdown("---")
        st.subheader("🏷️ Categorías aprendidas")
        if CATEGORIAS_DINAMICAS:
            st.success(f"✅ {len(CATEGORIAS_DINAMICAS)} categoría(s) dinámica(s) activa(s).")
            with st.expander("Ver categorías aprendidas"):
                for nombre, datos in CATEGORIAS_DINAMICAS.items():
                    st.caption(
                        f"**{nombre}** — {datos.get('descripcion', 'Sin descripción')}  \n"
                        f"{len(datos.get('palabras', []))} palabra(s), "
                        f"{len(datos.get('prefijos', []))} prefijo(s) de clave SAT."
                    )
        else:
            st.caption("Ninguna categoría dinámica dada de alta todavía.")


with st.sidebar:
    st.markdown("---")
    st.subheader("🛡️ Validaciones oficiales del SAT")

    if "_listado_69b_ok" not in st.session_state:
        st.info(
            "La lista negra 69-B (EFOS) aún no se ha cargado. Se descargará "
            "automáticamente la primera vez que audites un lote, para no "
            "retrasar el inicio de sesión."
        )
    elif st.session_state["_listado_69b_ok"]:
        st.success(f"✅ Lista negra 69-B cargada: {st.session_state['_listado_69b_len']:,} RFC's.")
    else:
        st.error(f"⚠️ Error listado 69-B: {st.session_state.get('_listado_69b_error', '')}")

    # Actualizar el 69-B queda disponible para todos los roles: es una de
    # las dos únicas acciones que un Analista puede hacer en esta barra.
    if st.button("🔄 Cargar / actualizar listado 69-B ahora"):
        _obtener_listado_69b_cacheado.clear()
        with st.spinner("Descargando listado de EFOS del SAT..."):
            obtener_listado_69b_bajo_demanda()
        st.rerun()

    # Ajuste GLOBAL, igual que "Modo ahorro de tokens": solo el Admin lo ve
    # y lo cambia; los Analistas heredan el valor sin poder tocarlo.
    if ES_ADMIN:
        _ajustes["validar_estatus_sat_tiempo_real"] = st.checkbox(
            "Validar vigencia de cada CFDI en tiempo real (Web Service SAT)",
            value=_ajustes.get("validar_estatus_sat_tiempo_real", True),
            key="_validar_estatus_sat_admin",
            help=(
                "Cada factura tarda hasta ~15s adicionales si el SAT responde lento. "
                "Con lotes grandes, desactívalo para acelerar la auditoría y actívalo "
                "solo cuando necesites confirmar vigencia/cancelación. Este ajuste es "
                "global: aplica también para los Analistas."
            ),
        )
    validar_estatus_sat = _ajustes.get("validar_estatus_sat_tiempo_real", True)


def parse_cfdi_xml(xml_content, indice_lote=0, listado_69b=None, validar_estatus_sat=True,
                   catalogo_sat=None):
    listado_69b = listado_69b or {}
    catalogo_sat = catalogo_sat if catalogo_sat is not None else CATALOGO_SAT
    try:
        if isinstance(xml_content, bytes):
            xml_content = xml_content.decode('utf-8-sig', errors='ignore')

        root = ET.fromstring(xml_content)

        # Copia saneada del XML (sin datos de AUTOCOM ni sellos) por si la
        # analista pide después la auditoría completa con IA. Se guarda aquí
        # porque es el único momento en que tenemos el XML original a la mano.
        xml_saneado_para_ia, _nota_saneo = sanitizar_xml_para_ia(xml_content)

        for elem in root.iter():
            if '}' in elem.tag:
                elem.tag = elem.tag.split('}', 1)[1]

        folio = root.attrib.get('Folio', root.attrib.get('Serie', 'S/F'))
        fecha_emision_str = root.attrib.get('Fecha', '')[:10]
        metodo_pago = root.attrib.get('MetodoPago', 'N/A')
        forma_pago = root.attrib.get('FormaPago', 'N/A')
        subtotal = float(root.attrib.get('SubTotal', 0.0) or 0.0)

        emisor_elem = root.find('Emisor')
        emisor_nombre_real = emisor_elem.attrib.get('Nombre', 'DESCONOCIDO') if emisor_elem is not None else 'DESCONOCIDO'
        emisor_rfc_real = emisor_elem.attrib.get('Rfc', '') if emisor_elem is not None else ''
        regimen_emisor = emisor_elem.attrib.get('RegimenFiscal', '') if emisor_elem is not None else ''
        es_persona_fisica = len(emisor_rfc_real) == 13

        receptor_elem = root.find('Receptor')
        uso_cfdi = receptor_elem.attrib.get('UsoCFDI', 'N/A') if receptor_elem is not None else 'N/A'

        tfd_elem = root.find('.//TimbreFiscalDigital')
        uuid_cfdi = tfd_elem.attrib.get('UUID', '') if tfd_elem is not None else ''
        id_interno = uuid_cfdi or f"{folio}-{indice_lote}"

        conceptos_elem = root.find('Conceptos')

        es_flete = False
        es_honorario_profesional = False
        tiene_activo_fijo = False
        tiene_concepto_en_revision = False
        filas_detalle = []

        if conceptos_elem is not None:
            for concepto in conceptos_elem.findall('Concepto'):
                clave_prod = concepto.attrib.get('ClaveProdServ', '')
                desc = concepto.attrib.get('Descripcion', '').strip()
                importe_concepto = float(concepto.attrib.get('Importe', 0.0) or 0.0)
                objeto_imp_concepto = concepto.attrib.get('ObjetoImp', '')
                tiene_traslados_concepto = bool(concepto.findall('./Impuestos/Traslados/Traslado'))

                categoria = sat.clasificar_concepto(clave_prod, desc, importe_concepto, catalogo_sat, CATEGORIAS_DINAMICAS)
                origen_categoria = "Motor de reglas"
                motivo_linea = ""

                if categoria == "FLETE":
                    es_flete = True
                elif categoria == "HONORARIO_PROFESIONAL":
                    es_honorario_profesional = True
                elif categoria == "ACTIVO_FIJO":
                    tiene_activo_fijo = True
                elif categoria == "REVISION":
                    tiene_concepto_en_revision = True

                cumple_linea = "SI"
                if categoria == "ACTIVO_FIJO" and not uso_cfdi.startswith('I'):
                    cumple_linea = "NO"
                    motivo_linea = (motivo_linea + " " if motivo_linea else "") + \
                        "Requiere Uso CFDI de Inversión/Activo Fijo (I01-I08)."
                elif categoria == "REVISION":
                    cumple_linea = "REVISAR"
                    motivo_linea = motivo_linea or "El motor de reglas no pudo clasificar este concepto con certeza."

                # Comparación explícita: Clave ProdServ del SAT (fuente de
                # verdad) vs. la descripción del concepto. Responde
                # directamente "¿el concepto está correcto o no respecto a
                # la clave declarada?".
                clave_vs_desc_ok, motivo_clave_vs_desc = sat.validar_clave_vs_descripcion(clave_prod, desc)
                if not clave_vs_desc_ok:
                    tiene_concepto_en_revision = True
                    if cumple_linea == "SI":
                        cumple_linea = "REVISAR"
                    motivo_linea = (motivo_linea + " " if motivo_linea else "") + motivo_clave_vs_desc

                # Comparación: categoría del concepto vs. el Uso CFDI (c_UsoCFDI)
                # declarado por el receptor para toda la factura.
                uso_vs_concepto_ok, motivo_uso_vs_concepto = sat.validar_concepto_vs_uso_cfdi(categoria, uso_cfdi, CATEGORIAS_DINAMICAS)
                if not uso_vs_concepto_ok:
                    tiene_concepto_en_revision = True
                    if cumple_linea == "SI":
                        cumple_linea = "REVISAR"
                    motivo_linea = (motivo_linea + " " if motivo_linea else "") + motivo_uso_vs_concepto

                # Coherencia: ¿el concepto lleva o no impuesto trasladado, y
                # eso hace sentido con lo que es (refacción, activo, flete,
                # honorario) y con lo que el propio XML declara en ObjetoImp?
                objeto_imp_ok, motivo_objeto_imp = sat.validar_objeto_impuesto(
                    categoria, objeto_imp_concepto, tiene_traslados_concepto
                )
                if not objeto_imp_ok:
                    tiene_concepto_en_revision = True
                    if cumple_linea == "SI":
                        cumple_linea = "REVISAR"
                    motivo_linea = (motivo_linea + " " if motivo_linea else "") + motivo_objeto_imp

                filas_detalle.append({
                    "_id_interno": id_interno,
                    "Factura / Folio": folio,
                    "Clave Prod/Serv SAT": sat.obtener_nombre_clave_sat(clave_prod, catalogo_sat),
                    "Concepto / Descripción": desc,
                    "Importe": importe_concepto,
                    "Categoría": categoria,
                    "Concepto vs Clave SAT": "CORRECTO" if clave_vs_desc_ok else "REVISAR",
                    "Objeto Imp.": sat.CLAVES_OBJETO_IMP.get(objeto_imp_concepto, objeto_imp_concepto or "N/A"),
                    "Origen Clasificación": origen_categoria,
                    "Cumple": cumple_linea,
                    "Motivo": motivo_linea or "Correcto",
                })

        retenciones_xml = {}
        impuestos_elem = root.find('Impuestos')
        if impuestos_elem is not None:
            ret_node = impuestos_elem.find('Retenciones')
            if ret_node is not None:
                for ret in ret_node.findall('Retencion'):
                    imp = ret.attrib.get('Impuesto', '')
                    importe = float(ret.attrib.get('Importe', 0.0) or 0.0)
                    tasa = float(ret.attrib.get('TasaOCuota', 0.0)) if 'TasaOCuota' in ret.attrib else (importe / subtotal if subtotal > 0 else 0)
                    retenciones_xml[imp] = round(tasa, 4)

        observaciones = []

        # --- FILTRO PRINCIPAL: retenciones ISR / IVA ---
        # Toda la parametrización (tasas, regímenes, tolerancias) vive en
        # reglas_sat.py para que sea auditable y editable desde "Enseñar
        # reglas", en vez de estar escrita a mano aquí.
        categorias_factura = tuple(f["Categoría"] for f in filas_detalle)
        retenciones_ok, observaciones_retenciones = sat.validar_retenciones(
            retenciones_xml,
            regimen_emisor,
            es_persona_fisica,
            categorias_factura=categorias_factura,
            es_flete=es_flete,
        )
        if not retenciones_ok:
            observaciones.extend(observaciones_retenciones)

        forma_pago_ok, motivo_forma_pago = sat.validar_forma_pago_vs_metodo_pago(metodo_pago, forma_pago)
        if not forma_pago_ok:
            observaciones.append(motivo_forma_pago)

        if metodo_pago == 'PUE' and fecha_emision_str:
            fecha_dt = datetime.strptime(fecha_emision_str, '%Y-%m-%d')
            hoy = datetime.now()

            if (fecha_dt.year < hoy.year) or (fecha_dt.year == hoy.year and fecha_dt.month < hoy.month):
                observaciones.append("Factura PUE de meses anteriores no permitida. Solicitar refacturación en PPD.")
            elif fecha_dt.year == hoy.year and fecha_dt.month == hoy.month:
                dia_ultimo_jueves = sat.obtener_ultimo_jueves(hoy.year, hoy.month)
                if hoy.day >= dia_ultimo_jueves:
                    observaciones.append(
                        f"PUE posterior al corte del último jueves del mes (Día {dia_ultimo_jueves}). "
                        "Requiere autorización de analista."
                    )

        if tiene_activo_fijo and not uso_cfdi.startswith('I'):
            observaciones.append(
                "Activo Fijo / Equipo >= $12,000 MXN por concepto requiere Uso CFDI de Inversión (I01-I08)."
            )

        en_lista_negra, situacion_69b = sat.verificar_efos(emisor_rfc_real, listado_69b)

        estatus_sat_vigencia = "No Verificado"
        if validar_estatus_sat and uuid_cfdi:
            sello_cfdi = root.attrib.get('Sello', '')
            receptor_rfc = receptor_elem.attrib.get('Rfc', '') if receptor_elem is not None else ''
            total_cfdi = float(root.attrib.get('Total', subtotal) or subtotal)
            resultado_sat = _consultar_estatus_sat_cacheado(
                emisor_rfc_real, receptor_rfc, total_cfdi, uuid_cfdi, sello_cfdi
            )
            estatus_sat_vigencia = resultado_sat["estado"]

        bloqueo_critico = False

        # Regla de mayoría: % de conceptos correctos ("Cumple" == "SI") sobre
        # el total. Solo se usa para decidir el caso "hay conceptos dudosos"
        # más abajo — nunca para pasar por alto los bloqueos críticos (69-B,
        # CFDI cancelado, retenciones fuera de norma, corte de PUE).
        total_conceptos_factura = len(filas_detalle)
        conceptos_correctos_factura = sum(1 for f in filas_detalle if f["Cumple"] == "SI")
        cumple_mayoria, porcentaje_correcto = sat.evaluar_regla_mayoria(
            total_conceptos_factura, conceptos_correctos_factura
        )

        if en_lista_negra:
            estatus = "RECHAZADO DE ENTRADA"
            ingresa_cxp = "NO"
            observaciones.insert(
                0,
                f"ALERTA CRÍTICA: Emisor en lista negra del SAT (Art. 69-B CFF) — situación: {situacion_69b}."
            )
            bloqueo_critico = True
        elif estatus_sat_vigencia in ("Cancelado", "No Encontrado"):
            estatus = "FACTURA CANCELADA"
            ingresa_cxp = "NO"
            observaciones.insert(
                0,
                f"El comprobante aparece como '{estatus_sat_vigencia}' en el Web Service del SAT."
            )
            bloqueo_critico = True
        elif observaciones:
            estatus = "RECHAZADO DE ENTRADA"
            ingresa_cxp = "NO"
        elif tiene_concepto_en_revision:
            if cumple_mayoria:
                estatus = "ACEPTADO"
                ingresa_cxp = "SI"
                observaciones.append(
                    f"Aceptada por regla de mayoría: {conceptos_correctos_factura}/{total_conceptos_factura} "
                    f"conceptos ({porcentaje_correcto:.0%}) cumplen (≥{sat.UMBRAL_PORCENTAJE_CONCEPTOS_CORRECTOS:.0%} "
                    "requerido). Revisar igualmente el/los concepto(s) marcados en el detalle."
                )
            else:
                estatus = "REVISIÓN MANUAL"
                ingresa_cxp = "NO"
                observaciones.append(
                    f"Solo {conceptos_correctos_factura}/{total_conceptos_factura} conceptos "
                    f"({porcentaje_correcto:.0%}) cumplen, por debajo del "
                    f"{sat.UMBRAL_PORCENTAJE_CONCEPTOS_CORRECTOS:.0%} requerido para aceptar por mayoría; "
                    "ver detalle por línea."
                )
        else:
            estatus = "ACEPTADO"
            ingresa_cxp = "SI"

        mensaje = " | ".join(observaciones) if observaciones else "Factura cumple con todos los requisitos fiscales."

        fila_cabecera = {
            "_id_interno": id_interno,
            "_bloqueo_critico": bloqueo_critico,
            "Factura / Folio": folio,
            "Fecha": fecha_emision_str,
            "Emisor": emisor_nombre_real,
            "RFC Emisor": emisor_rfc_real,
            "Método Pago": metodo_pago,
            "Forma Pago": sat.obtener_nombre_forma_pago(forma_pago),
            "Régimen Fiscal Emisor": sat.obtener_nombre_regimen(regimen_emisor),
            "_xml_saneado": xml_saneado_para_ia,
            "_rfc_emisor": emisor_rfc_real,
            "_retenciones_xml": retenciones_xml,
            "_metodo_pago_raw": metodo_pago,
            "_forma_pago_raw": forma_pago,
            "Uso CFDI": uso_cfdi,
            "Núm. Conceptos": len(filas_detalle),
            "% Conceptos Correctos": round(porcentaje_correcto * 100, 1),
            "Situación EFOS (Art. 69-B)": situacion_69b if situacion_69b else "Sin coincidencia",
            "Estatus SAT (Vigencia)": estatus_sat_vigencia,
            "Estatus": estatus,
            "Ingresa CxP": ingresa_cxp,
            "Mensaje Devuelto al Proveedor / Acción": mensaje
        }
        return fila_cabecera, filas_detalle

    except Exception as e:
        id_interno_error = f"ERROR-{indice_lote}"
        fila_cabecera = {
            "_id_interno": id_interno_error,
            "_bloqueo_critico": True,
            "Factura / Folio": "ERROR",
            "Fecha": "-",
            "Emisor": "XML Inválido",
            "RFC Emisor": "-",
            "Método Pago": "-",
            "Forma Pago": "-",
            "Régimen Fiscal Emisor": "-",
            "_xml_saneado": "",
            "_rfc_emisor": "",
            "_retenciones_xml": {},
            "_metodo_pago_raw": "",
            "_forma_pago_raw": "",
            "Uso CFDI": "-",
            "Núm. Conceptos": 0,
            "% Conceptos Correctos": 0.0,
            "Situación EFOS (Art. 69-B)": "-",
            "Estatus SAT (Vigencia)": "-",
            "Estatus": "RECHAZADO DE ENTRADA",
            "Ingresa CxP": "NO",
            "Mensaje Devuelto al Proveedor / Acción": f"Estructura XML inválida: {str(e)}"
        }
        return fila_cabecera, []

# ---------------------------------------------------------
# 3. EXPORTAR A EXCEL
# ---------------------------------------------------------
def generar_excel_presentable(df_cabecera, df_detalle):
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    font_family = "Segoe UI"

    autocom_red = "C8102E"
    header_fill = PatternFill(start_color=autocom_red, end_color=autocom_red, fill_type="solid")
    header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    title_font = Font(name=font_family, size=16, bold=True, color=autocom_red)
    subtitle_font = Font(name=font_family, size=10, italic=True, color="595959")

    thin_side = Side(border_style="thin", color="D9D9D9")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    df_cabecera_export = df_cabecera.drop(columns=["_id_interno", "_bloqueo_critico", "_xml_saneado", "_rfc_emisor", "_retenciones_xml", "_metodo_pago_raw", "_forma_pago_raw"], errors="ignore")
    df_detalle_export = df_detalle.drop(columns=["_id_interno"], errors="ignore") if df_detalle is not None else pd.DataFrame()

    ws_summary = wb.active
    ws_summary.title = "Resumen de Auditoría"
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary["A2"] = "AUTOCOM - REPORTE DE AUDITORÍA FISCAL CFDI 4.0"
    ws_summary["A2"].font = title_font
    ws_summary["A3"] = f"Fecha: {datetime.now().strftime('%d/%m/%Y %H:%M')} | Desarrollado por: Juan Rogelio Cruz García"
    ws_summary["A3"].font = subtitle_font

    total = len(df_cabecera_export)
    aceptadas = len(df_cabecera_export[df_cabecera_export["Estatus"] == "ACEPTADO"]) if "Estatus" in df_cabecera_export.columns else 0
    revision = len(df_cabecera_export[df_cabecera_export["Estatus"] == "REVISIÓN MANUAL"]) if "Estatus" in df_cabecera_export.columns else 0
    rechazadas = len(df_cabecera_export[df_cabecera_export["Estatus"] == "RECHAZADO DE ENTRADA"]) if "Estatus" in df_cabecera_export.columns else 0
    canceladas = len(df_cabecera_export[df_cabecera_export["Estatus"] == "FACTURA CANCELADA"]) if "Estatus" in df_cabecera_export.columns else 0

    kpis = [
        ("TOTAL FACTURAS", total, "B5", "B6", "1F2937"),
        ("ACEPTADAS", aceptadas, "D5", "D6", "2E75B6"),
        ("REVISIÓN MANUAL", revision, "F5", "F6", "C55A11"),
        ("RECHAZADAS DE ENTRADA", rechazadas, "H5", "H6", autocom_red),
        ("CANCELADAS ANTE EL SAT", canceladas, "J5", "J6", "595959"),
    ]

    for title, val, pos_t, pos_v, color in kpis:
        ws_summary[pos_t] = title
        ws_summary[pos_t].font = Font(name=font_family, size=9, bold=True, color="595959")
        ws_summary[pos_t].alignment = Alignment(horizontal="center", vertical="center")

        ws_summary[pos_v] = val
        ws_summary[pos_v].font = Font(name=font_family, size=22, bold=True, color=color)
        ws_summary[pos_v].alignment = Alignment(horizontal="center", vertical="center")

        card_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
        ws_summary[pos_t].fill = card_fill
        ws_summary[pos_v].fill = card_fill
        ws_summary[pos_t].border = Border(left=thin_side, right=thin_side, top=thin_side)
        ws_summary[pos_v].border = Border(left=thin_side, right=thin_side, bottom=thin_side)

    for col, txt in [("A9", "Estatus"), ("B9", "Cantidad"), ("C9", "Porcentaje")]:
        ws_summary[col] = txt
        ws_summary[col].fill = header_fill
        ws_summary[col].font = header_font
        ws_summary[col].alignment = Alignment(horizontal="center", vertical="center")

    table_data = [
        ("ACEPTADO", aceptadas),
        ("REVISIÓN MANUAL", revision),
        ("RECHAZADO DE ENTRADA", rechazadas),
        ("FACTURA CANCELADA", canceladas),
    ]
    for idx, (est, cnt) in enumerate(table_data, start=10):
        ws_summary[f"A{idx}"] = est
        ws_summary[f"B{idx}"] = cnt
        ws_summary[f"C{idx}"] = f"=B{idx}/SUM($B$10:$B$13)"

        ws_summary[f"A{idx}"].alignment = Alignment(horizontal="left", vertical="center")
        ws_summary[f"B{idx}"].alignment = Alignment(horizontal="center", vertical="center")
        ws_summary[f"C{idx}"].alignment = Alignment(horizontal="right", vertical="center")
        ws_summary[f"C{idx}"].number_format = "0.0%"

        for c in [f"A{idx}", f"B{idx}", f"C{idx}"]:
            ws_summary[c].font = Font(name=font_family, size=10)
            ws_summary[c].border = thin_border

    ws_summary["A14"] = "Total"
    ws_summary["B14"] = "=SUM(B10:B13)"
    ws_summary["C14"] = "=SUM(C10:C13)"
    for c in ["A14", "B14", "C14"]:
        ws_summary[c].font = Font(name=font_family, size=10, bold=True)
        ws_summary[c].border = Border(top=thin_side, bottom=Side(border_style="double", color=autocom_red))
    ws_summary["C14"].number_format = "0.0%"

    pie = PieChart()
    labels = Reference(ws_summary, min_col=1, min_row=10, max_row=13)
    data = Reference(ws_summary, min_col=2, min_row=9, max_row=13)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Validación Fiscal CFDI 4.0"
    pie.width, pie.height = 14, 7.5
    ws_summary.add_chart(pie, "E9")

    ws_detail = wb.create_sheet(title="Detalle Auditoría")
    ws_detail.views.sheetView[0].showGridLines = True

    ws_detail["A2"] = "AUTOCOM - DETALLE DE COMPROBANTES AUDITADOS"
    ws_detail["A2"].font = title_font
    ws_detail["A3"] = f"Total registros: {len(df_cabecera_export)} | Auditoría desarrollada por Juan Rogelio Cruz García"
    ws_detail["A3"].font = subtitle_font

    headers = list(df_cabecera_export.columns)
    for col_num, header_title in enumerate(headers, 1):
        cell = ws_detail.cell(row=5, column=col_num, value=header_title)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws_detail.row_dimensions[5].height = 28

    fill_approved = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_revision = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_rejected = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_cancelada = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    fill_zebra = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    font_approved = Font(name=font_family, size=10, bold=True, color="375623")
    font_revision = Font(name=font_family, size=10, bold=True, color="7F6000")
    font_rejected = Font(name=font_family, size=10, bold=True, color="C00000")
    font_cancelada = Font(name=font_family, size=10, bold=True, color="404040")
    font_regular = Font(name=font_family, size=10)

    for row_idx, row_data in df_cabecera_export.iterrows():
        excel_row = row_idx + 6
        is_even = (row_idx % 2 == 1)

        for col_num, value in enumerate(row_data, 1):
            cell = ws_detail.cell(row=excel_row, column=col_num, value=value)
            cell.font = font_regular
            cell.border = thin_border

            col_name = headers[col_num - 1]
            if col_name in [
                "Factura / Folio", "Fecha", "RFC Emisor", "Método Pago", "Forma Pago", "Régimen Fiscal Emisor", "Uso CFDI", "Núm. Conceptos",
                "% Conceptos Correctos", "Ingresa CxP", "Situación EFOS (Art. 69-B)", "Estatus SAT (Vigencia)",
            ]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == "Estatus":
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if value == "ACEPTADO":
                    cell.fill, cell.font = fill_approved, font_approved
                elif value == "REVISIÓN MANUAL":
                    cell.fill, cell.font = fill_revision, font_revision
                elif value == "RECHAZADO DE ENTRADA":
                    cell.fill, cell.font = fill_rejected, font_rejected
                elif value == "FACTURA CANCELADA":
                    cell.fill, cell.font = fill_cancelada, font_cancelada
            elif col_name == "Situación EFOS (Art. 69-B)" and value not in ("Sin coincidencia", "-", "", None):
                cell.fill, cell.font = fill_rejected, font_rejected
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                if is_even and cell.fill.fill_type is None:
                    cell.fill = fill_zebra

        ws_detail.row_dimensions[excel_row].height = 35

    col_widths = {
        "Factura / Folio": 18, "Fecha": 14, "Emisor": 26, "RFC Emisor": 18,
        "Método Pago": 14, "Forma Pago": 26, "Régimen Fiscal Emisor": 38, "Uso CFDI": 12, "Núm. Conceptos": 14, "% Conceptos Correctos": 16,
        "Situación EFOS (Art. 69-B)": 20, "Estatus SAT (Vigencia)": 18, "Estatus": 20,
        "Ingresa CxP": 12, "Mensaje Devuelto al Proveedor / Acción": 65
    }
    for col_num, header_title in enumerate(headers, 1):
        col_letter = get_column_letter(col_num)
        ws_detail.column_dimensions[col_letter].width = col_widths.get(header_title, 20)

    ws_detail.auto_filter.ref = f"A5:{get_column_letter(len(headers))}{len(df_cabecera_export) + 5}"
    ws_detail.freeze_panes = "A6"

    ws_summary.column_dimensions["A"].width = 24
    for c_let in ["B", "C", "D", "E", "F", "G", "H"]:
        ws_summary.column_dimensions[c_let].width = 16

    if not df_detalle_export.empty:
        ws_lineas = wb.create_sheet(title="Detalle por Concepto")
        ws_lineas.views.sheetView[0].showGridLines = True

        ws_lineas["A2"] = "AUTOCOM - DETALLE POR CONCEPTO / PARTIDA"
        ws_lineas["A2"].font = title_font
        ws_lineas["A3"] = f"Total conceptos: {len(df_detalle_export)}"
        ws_lineas["A3"].font = subtitle_font

        headers_l = list(df_detalle_export.columns)
        for col_num, header_title in enumerate(headers_l, 1):
            cell = ws_lineas.cell(row=5, column=col_num, value=header_title)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws_lineas.row_dimensions[5].height = 28

        for row_idx, row_data in df_detalle_export.iterrows():
            excel_row = row_idx + 6
            is_even = (row_idx % 2 == 1)
            for col_num, value in enumerate(row_data, 1):
                cell = ws_lineas.cell(row=excel_row, column=col_num, value=value)
                cell.font = font_regular
                cell.border = thin_border
                col_name = headers_l[col_num - 1]
                if col_name in ["Factura / Folio", "Importe", "Categoría", "Cumple", "Concepto vs Clave SAT"]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    if col_name == "Cumple":
                        if value == "SI":
                            cell.fill, cell.font = fill_approved, font_approved
                        elif value == "REVISAR":
                            cell.fill, cell.font = fill_revision, font_revision
                        elif value == "NO":
                            cell.fill, cell.font = fill_rejected, font_rejected
                    if col_name == "Concepto vs Clave SAT":
                        if value == "CORRECTO":
                            cell.fill, cell.font = fill_approved, font_approved
                        elif value == "REVISAR":
                            cell.fill, cell.font = fill_revision, font_revision
                    if col_name == "Importe":
                        cell.number_format = '"$"#,##0.00'
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    if is_even and cell.fill.fill_type is None:
                        cell.fill = fill_zebra
            ws_lineas.row_dimensions[excel_row].height = 32

        col_widths_l = {
            "Factura / Folio": 16, "Clave Prod/Serv SAT": 34, "Concepto / Descripción": 42, "Objeto Imp.": 30,
            "Importe": 14, "Categoría": 20, "Concepto vs Clave SAT": 18, "Origen Clasificación": 20,
            "Cumple": 10, "Motivo": 50, "Observación IA": 55,
        }
        for col_num, header_title in enumerate(headers_l, 1):
            col_letter = get_column_letter(col_num)
            ws_lineas.column_dimensions[col_letter].width = col_widths_l.get(header_title, 20)

        ws_lineas.auto_filter.ref = f"A5:{get_column_letter(len(headers_l))}{len(df_detalle_export) + 5}"
        ws_lineas.freeze_panes = "A6"

    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def _recalcular_estatus_factura(id_interno_factura, avisos_extra=None):
    """
    ÚNICO lugar donde se decide el estatus de una factura después de que algo
    cambió (clasificación aplicada por IA, auditoría del XML completo,
    opinión del analista). Antes esta lógica estaba enterrada dentro de
    `_aplicar_clasificacion_ia` y la regla de mayoría (80%) solo se evaluaba
    en una rama muy específica, por eso una factura podía quedarse en
    "REVISIÓN MANUAL" aunque ya cumpliera el porcentaje.

    Orden de decisión (de mayor a menor autoridad):
      1. Bloqueo crítico (69-B, CFDI cancelado, retenciones, corte PUE):
         NUNCA se levanta por mayoría.
      2. Activo fijo sin Uso CFDI de Inversión: rechazo duro por línea.
      3. Regla de mayoría: si el % de conceptos que cumplen alcanza
         UMBRAL_PORCENTAJE_CONCEPTOS_CORRECTOS, la factura se ACEPTA aunque
         queden líneas en REVISION.
      4. Si no alcanza, revisión manual.
    """
    df_det = st.session_state["df_detalle"]
    df_cab = st.session_state["df_cabecera"]

    idxs_cab = df_cab.index[df_cab["_id_interno"] == id_interno_factura]
    if len(idxs_cab) == 0:
        return
    i_cab = idxs_cab[0]

    valor_bloqueo = df_cab.at[i_cab, "_bloqueo_critico"]
    if bool(valor_bloqueo) and pd.notna(valor_bloqueo):
        return  # un bloqueo crítico no se levanta por mayoría ni por la IA

    uso_cfdi = str(df_cab.at[i_cab, "Uso CFDI"])
    conceptos = df_det[df_det["_id_interno"] == id_interno_factura]

    hay_rechazo_por_linea = False
    avisos_ia = list(avisos_extra or [])

    for i_det, fila in conceptos.iterrows():
        categoria = fila["Categoría"]
        if categoria == "ACTIVO_FIJO" and not uso_cfdi.startswith("I"):
            df_det.at[i_det, "Cumple"] = "NO"
            motivo_extra = "Requiere Uso CFDI de Inversión/Activo Fijo (I01-I08)."
            motivo_actual = str(df_det.at[i_det, "Motivo"] or "")
            df_det.at[i_det, "Motivo"] = (
                f"{motivo_actual} {motivo_extra}".strip()
                if motivo_actual not in ("", "Correcto") else motivo_extra
            )
            hay_rechazo_por_linea = True
        elif categoria in ("FLETE", "HONORARIO_PROFESIONAL") and str(fila["Origen Clasificación"]).startswith("IA"):
            avisos_ia.append(
                f"La IA reclasificó un concepto como {categoria}: verifique manualmente la retención de "
                "ISR/IVA correspondiente, ya que el cálculo de retenciones se realizó antes de esta "
                "reclasificación."
            )

    # La mayoría se evalúa SIEMPRE, con el estado más reciente del detalle.
    conceptos = df_det[df_det["_id_interno"] == id_interno_factura]
    total_conceptos = len(conceptos)
    conceptos_correctos = int((conceptos["Cumple"] == "SI").sum())
    cumple_mayoria, porcentaje_correcto = sat.evaluar_regla_mayoria(total_conceptos, conceptos_correctos)

    if "% Conceptos Correctos" in df_cab.columns:
        df_cab.at[i_cab, "% Conceptos Correctos"] = round(porcentaje_correcto * 100, 1)

    detalle_mayoria = (
        f"{conceptos_correctos}/{total_conceptos} conceptos ({porcentaje_correcto:.0%}) cumplen "
        f"(umbral {sat.UMBRAL_PORCENTAJE_CONCEPTOS_CORRECTOS:.0%})"
    )

    if hay_rechazo_por_linea:
        nuevo_estatus, nuevo_cxp = "RECHAZADO DE ENTRADA", "NO"
        mensaje = ("Requiere Uso CFDI de Inversión para el Activo Fijo reclasificado por IA; "
                   "ver detalle por línea.")
    elif cumple_mayoria:
        nuevo_estatus, nuevo_cxp = "ACEPTADO", "SI"
        mensaje = f"Aceptada por regla de mayoría: {detalle_mayoria}."
        if conceptos_correctos < total_conceptos:
            mensaje += " Revisar igualmente el/los concepto(s) marcados en el detalle."
        if avisos_ia:
            mensaje += " " + " | ".join(dict.fromkeys(avisos_ia))
    else:
        nuevo_estatus, nuevo_cxp = "REVISIÓN MANUAL", "NO"
        mensaje = f"Solo {detalle_mayoria}; ver detalle por línea."
        if avisos_ia:
            mensaje += " " + " | ".join(dict.fromkeys(avisos_ia))

    df_cab.at[i_cab, "Estatus"] = nuevo_estatus
    df_cab.at[i_cab, "Ingresa CxP"] = nuevo_cxp
    df_cab.at[i_cab, "Mensaje Devuelto al Proveedor / Acción"] = mensaje


def _aplicar_clasificacion_ia(idx_detalle, id_interno_factura, sugerencia):
    df_det = st.session_state["df_detalle"]

    categoria_ia = sugerencia.get("categoria", "GENERAL")
    df_det.at[idx_detalle, "Categoría"] = categoria_ia
    df_det.at[idx_detalle, "Origen Clasificación"] = f"IA (confianza {sugerencia.get('confianza', 'n/d')})"
    df_det.at[idx_detalle, "Motivo"] = sugerencia.get("motivo", "")
    # El veredicto de la IA manda sobre el "Cumple" de esa línea: si la IA
    # dice RECHAZADO, la línea NO cumple aunque se haya clasificado.
    veredicto_ia = str(sugerencia.get("veredicto", "ACEPTADO")).upper()
    df_det.at[idx_detalle, "Cumple"] = "SI" if veredicto_ia == "ACEPTADO" else (
        "NO" if veredicto_ia == "RECHAZADO" else "REVISAR"
    )

    _recalcular_estatus_factura(id_interno_factura)


def _aplicar_dictamen_xml_ia(id_interno_factura, dictamen):
    """
    Traduce el dictamen de la auditoría del XML completo a los mismos datos
    que usa el motor de reglas, y vuelve a evaluar la factura con la regla
    de mayoría. Así el veredicto de la IA sí mueve el estatus, en vez de
    quedarse como un texto informativo al lado.
    """
    if not dictamen or not dictamen.get("_ok"):
        return

    df_cab = st.session_state["df_cabecera"]
    idxs_cab = df_cab.index[df_cab["_id_interno"] == id_interno_factura]
    if len(idxs_cab) == 0:
        return
    i_cab = idxs_cab[0]

    valor_bloqueo = df_cab.at[i_cab, "_bloqueo_critico"]
    if bool(valor_bloqueo) and pd.notna(valor_bloqueo):
        return

    hallazgos = dictamen.get("hallazgos", [])
    criticos = [h for h in hallazgos if h.get("severidad") == "CRITICO"]

    if criticos:
        # Un hallazgo CRÍTICO sí bloquea: no se acepta por mayoría.
        df_cab.at[i_cab, "Estatus"] = "REVISIÓN MANUAL"
        df_cab.at[i_cab, "Ingresa CxP"] = "NO"
        df_cab.at[i_cab, "Mensaje Devuelto al Proveedor / Acción"] = (
            f"La IA detectó {len(criticos)} hallazgo(s) CRÍTICO(s) en el comprobante: "
            f"{criticos[0].get('detalle', '')}"
        )
        return

    avisos = []
    if dictamen.get("veredicto") == "REVISION" and hallazgos:
        avisos.append(f"IA: {len(hallazgos)} observación(es) no crítica(s) en el XML.")

    _recalcular_estatus_factura(id_interno_factura, avisos_extra=avisos)



# ---------------------------------------------------------
# 4. INTERFAZ INTERACTIVA STREAMLIT (PESTAÑAS)
# ---------------------------------------------------------

# ---------------------------------------------------------
# QUÉ FACTURA ESTÁ ABIERTA (y por qué hace falta un callback)
# ---------------------------------------------------------
# Streamlit dibuja los expanders DE ARRIBA HACIA ABAJO. Cuando el analista
# aprieta un botón que está DENTRO de un expander, el script se reejecuta
# completo: para cuando la ejecución llega al `if st.button(...)`, el
# expander YA se dibujó — y se dibujó cerrado, porque en ese momento todavía
# no se había marcado cuál era la factura abierta. Por eso "se cerraba justo
# al analizar" aunque el estado se guardara correctamente: se guardaba tarde.
#
# La solución es `on_click`: Streamlit ejecuta los callbacks ANTES del cuerpo
# del script, así que cuando se dibuja el expander el estado ya está puesto.
# Nunca dentro del `if st.button(...)`.

def _abrir_factura(id_interno_factura):
    """Callback: marca esta factura como la que debe quedar abierta."""
    st.session_state["_factura_abierta"] = id_interno_factura


def _abrir_factura_y_opinion(id_interno_factura, key_opinion):
    """Igual que _abrir_factura, pero además deja abierto el sub-expander de
    'Opinión del analista' (si no, se cerraba el de adentro)."""
    st.session_state["_factura_abierta"] = id_interno_factura
    st.session_state["_opinion_abierta"] = key_opinion


def render_tab_auditoria():
    st.markdown("### 📥 Carga de Archivos para Auditoría")
    uploaded_files = st.file_uploader(
        "Arrastra o selecciona tus facturas (XML, ZIP) o reportes masivos (CSV, XLSX)",
        type=["xml", "zip", "csv", "xlsx"],
        accept_multiple_files=True
    )

    if uploaded_files:
        firma_archivos = tuple((f.name, f.size) for f in uploaded_files)

        if st.session_state.get("_firma_archivos_procesados") != firma_archivos:
            # Aquí, y solo aquí, se dispara la descarga del listado 69-B si
            # todavía no estaba en caché — es decir, cuando el usuario ya
            # decidió auditar algo, nunca antes.
            with st.spinner("Verificando listado 69-B del SAT (solo la primera vez, luego queda en caché)..."):
                listado_69b = obtener_listado_69b_bajo_demanda()

            cabeceras = []
            detalles = []
            contador = 0

            # Control de duplicados: si el mismo comprobante (mismo UUID) llega
            # más de una vez (p. ej. porque se referencia en varias OC y el
            # archivo se subió por cada una), solo se cuenta una vez.
            uuids_vistos = {}
            duplicados_detectados = []

            # Muestra una barra de progreso visual de procesamiento activo
            progress_bar = st.progress(0)
            status_box = st.empty()

            num_archivos = len(uploaded_files)

            for i, uploaded_file in enumerate(uploaded_files):
                # Notifica al usuario en pantalla qué archivo se está leyendo en tiempo real
                status_box.markdown(
                    f"⏳ **Auditando comprobantes ({i + 1} de {num_archivos}):** `{uploaded_file.name}`..."
                )
                progress_bar.progress(int(((i + 1) / num_archivos) * 100))

                file_ext = uploaded_file.name.split(".")[-1].lower()

                if file_ext == "xml":
                    content = uploaded_file.read()
                    cab, det = parse_cfdi_xml(
                        content, indice_lote=contador,
                        listado_69b=listado_69b, validar_estatus_sat=validar_estatus_sat,
                    )
                    contador += 1
                    id_interno_nuevo = cab["_id_interno"]
                    if id_interno_nuevo in uuids_vistos:
                        duplicados_detectados.append({
                            "Folio": cab.get("Factura / Folio", "N/D"),
                            "UUID / ID": id_interno_nuevo,
                            "Archivo duplicado": uploaded_file.name,
                            "Ya cargado desde": uuids_vistos[id_interno_nuevo],
                        })
                    else:
                        uuids_vistos[id_interno_nuevo] = uploaded_file.name
                        cabeceras.append(cab)
                        detalles.extend(det)

                elif file_ext == "zip":
                    with zipfile.ZipFile(uploaded_file) as z:
                        xml_files = [f for f in z.namelist() if f.endswith(".xml")]
                        total_xmls = len(xml_files)
                        for j, filename in enumerate(xml_files):
                            status_box.markdown(
                                f"⏳ **Descomprimiendo ZIP y procesando XML ({j + 1} de {total_xmls}):** `{filename}`..."
                            )
                            with z.open(filename) as f:
                                cab, det = parse_cfdi_xml(
                                    f.read(), indice_lote=contador,
                                    listado_69b=listado_69b, validar_estatus_sat=validar_estatus_sat,
                                )
                                contador += 1
                                id_interno_nuevo = cab["_id_interno"]
                                nombre_mostrado = f"{uploaded_file.name} → {filename}"
                                if id_interno_nuevo in uuids_vistos:
                                    duplicados_detectados.append({
                                        "Folio": cab.get("Factura / Folio", "N/D"),
                                        "UUID / ID": id_interno_nuevo,
                                        "Archivo duplicado": nombre_mostrado,
                                        "Ya cargado desde": uuids_vistos[id_interno_nuevo],
                                    })
                                else:
                                    uuids_vistos[id_interno_nuevo] = nombre_mostrado
                                    cabeceras.append(cab)
                                    detalles.extend(det)

                elif file_ext in ["csv", "xlsx"]:
                    df_temp = pd.read_csv(uploaded_file) if file_ext == "csv" else pd.read_excel(uploaded_file)
                    if "Unnamed: 0" in df_temp.columns:
                        df_temp = df_temp.drop(columns=["Unnamed: 0"])
                    for registro in df_temp.to_dict(orient="records"):
                        registro.setdefault("_id_interno", f"CSV-{contador}")
                        registro.setdefault("_bloqueo_critico", False)
                        cabeceras.append(registro)
                        contador += 1

            # Limpia los indicadores visuales al terminar la carga
            status_box.empty()
            progress_bar.empty()

            if cabeceras:
                df_cabecera = pd.DataFrame(cabeceras)
                df_detalle = pd.DataFrame(detalles) if detalles else pd.DataFrame(
                    columns=["_id_interno", "Factura / Folio", "Clave Prod/Serv SAT", "Concepto / Descripción",
                             "Importe", "Categoría", "Concepto vs Clave SAT", "Objeto Imp.",
                             "Origen Clasificación", "Cumple", "Motivo"]
                )
                st.session_state["df_cabecera"] = df_cabecera
                st.session_state["df_detalle"] = df_detalle
                st.session_state["_firma_archivos_procesados"] = firma_archivos
                st.session_state["_duplicados_detectados"] = duplicados_detectados

        if "df_cabecera" in st.session_state:
            df_cabecera = st.session_state["df_cabecera"]
            df_detalle = st.session_state["df_detalle"]

            duplicados_detectados = st.session_state.get("_duplicados_detectados", [])
            if duplicados_detectados:
                with st.expander(
                    f"⚠️ Se detectaron y omitieron {len(duplicados_detectados)} comprobante(s) "
                    "duplicado(s) (mismo UUID, p. ej. por venir referenciados en más de una OC)",
                    expanded=False,
                ):
                    st.dataframe(pd.DataFrame(duplicados_detectados), use_container_width=True)

            st.success(f"¡Se auditaron exitosamente **{len(df_cabecera)}** comprobante(s)!")

            c1, c2, c3, c4, c5 = st.columns(5)
            total_count = len(df_cabecera)
            aceptadas_count = len(df_cabecera[df_cabecera["Estatus"] == "ACEPTADO"])
            revision_count = len(df_cabecera[df_cabecera["Estatus"] == "REVISIÓN MANUAL"])
            rechazadas_count = len(df_cabecera[df_cabecera["Estatus"] == "RECHAZADO DE ENTRADA"])
            canceladas_count = len(df_cabecera[df_cabecera["Estatus"] == "FACTURA CANCELADA"])

            c1.metric("Total Procesadas", total_count)
            c2.metric("Aceptadas", aceptadas_count)
            c3.metric("Revisión Manual", revision_count)
            c4.metric("Rechazadas de Entrada", rechazadas_count, delta_color="inverse")
            c5.metric("Canceladas ante el SAT", canceladas_count, delta_color="inverse")

            st.markdown("---")
            st.subheader("📋 Resultados de Auditoría (Maestro)")
            st.dataframe(
                df_cabecera.drop(columns=["_id_interno", "_bloqueo_critico", "_xml_saneado", "_rfc_emisor", "_retenciones_xml", "_metodo_pago_raw", "_forma_pago_raw"], errors="ignore"),
                use_container_width=True,
            )

            st.markdown("---")
            st.subheader("🔍 Detalle por Factura (línea por línea)")

            col_f1, col_f2 = st.columns([2, 1])
            with col_f1:
                filtro_estatus = st.multiselect(
                    "Filtrar por estatus",
                    options=["ACEPTADO", "REVISIÓN MANUAL", "RECHAZADO DE ENTRADA", "FACTURA CANCELADA"],
                    default=["REVISIÓN MANUAL", "RECHAZADO DE ENTRADA", "FACTURA CANCELADA"] if total_count > 30 else
                             ["ACEPTADO", "REVISIÓN MANUAL", "RECHAZADO DE ENTRADA", "FACTURA CANCELADA"],
                )
            with col_f2:
                busqueda_folio = st.text_input("Buscar Folio", "")

            df_a_mostrar = df_cabecera[df_cabecera["Estatus"].isin(filtro_estatus)] if filtro_estatus else df_cabecera
            if busqueda_folio:
                df_a_mostrar = df_a_mostrar[df_a_mostrar["Factura / Folio"].astype(str).str.contains(busqueda_folio, case=False, na=False)]

            if df_a_mostrar.empty:
                st.info("No hay facturas que coincidan con el filtro seleccionado.")
            else:
                idx_global_renglon = 0

                # Streamlit reinicia los expanders a su estado por defecto en
                # cada rerun, y consultar la IA provoca un rerun: por eso la
                # factura se "cerraba" justo al terminar el análisis. Se
                # recuerda cuál está abierta (vía callbacks on_click, que
                # corren ANTES de dibujar) y se fuerza expanded=True en ella.
                for idx_fac, fila_cab in df_a_mostrar.iterrows():
                    icono = ICONOS_ESTATUS.get(fila_cab["Estatus"], "⚪")
                    id_interno_factura = fila_cab["_id_interno"]
                    factura_abierta = st.session_state.get("_factura_abierta") == id_interno_factura

                    # El marcador 📂 en el título cumple dos funciones: le dice
                    # al analista de un vistazo en cuál está trabajando, y
                    # obliga a Streamlit a volver a montar el expander con el
                    # nuevo valor de `expanded` (si la etiqueta no cambia, el
                    # frontend puede conservar el estado anterior y seguir
                    # mostrándolo cerrado).
                    marcador = "📂 " if factura_abierta else ""
                    titulo = (
                        f"{marcador}{icono} Folio {fila_cab['Factura / Folio']} — "
                        f"{fila_cab['Estatus']} (Ingresa CxP: {fila_cab['Ingresa CxP']})"
                    )

                    with st.expander(titulo, expanded=factura_abierta):
                        tipo_badge_estatus = {
                            "ACEPTADO": "ok", "REVISIÓN MANUAL": "warn",
                            "RECHAZADO DE ENTRADA": "danger", "FACTURA CANCELADA": "danger",
                        }.get(fila_cab["Estatus"], "neutral")
                        tipo_badge_cxp = "ok" if fila_cab["Ingresa CxP"] == "SI" else "danger"
                        st.markdown(
                            f'{_badge(fila_cab["Estatus"], tipo_badge_estatus, icono)} &nbsp; '
                            f'{_badge("Ingresa CxP: " + fila_cab["Ingresa CxP"], tipo_badge_cxp, "💳")}',
                            unsafe_allow_html=True,
                        )

                        if bool(fila_cab.get("_bloqueo_critico")) and pd.notna(fila_cab.get("_bloqueo_critico")):
                            st.error(f"🚫 {fila_cab['Mensaje Devuelto al Proveedor / Acción']}")
                        else:
                            st.caption(f"**Acción:** {fila_cab['Mensaje Devuelto al Proveedor / Acción']}")

                        conceptos_de_esta_factura = st.session_state["df_detalle"][
                            st.session_state["df_detalle"]["_id_interno"] == id_interno_factura
                        ]

                        # -----------------------------------------------
                        # Auditoría del XML COMPLETO con IA (bajo demanda).
                        # Manda todo el comprobante — saneado — para que la
                        # IA lo revise como especialista fiscal, en vez de
                        # ver un concepto aislado.
                        # -----------------------------------------------
                        xml_saneado_factura = fila_cab.get("_xml_saneado", "")
                        key_audit = f"audit_xml_{id_interno_factura}_{idx_fac}"
                        key_audit_btn = f"btn_audit_xml_{id_interno_factura}_{idx_fac}"

                        if ia_disponible and xml_saneado_factura:
                            col_audit, col_nota = st.columns([1.4, 3])
                            if col_audit.button(
                                "🔎 Auditar XML completo con IA",
                                key=key_audit_btn,
                                on_click=_abrir_factura,
                                args=(id_interno_factura,),
                            ):
                                with st.spinner("La IA está revisando el comprobante completo..."):
                                    hechos = construir_hechos_verificados(fila_cab, conceptos_de_esta_factura)
                                    dictamen = auditar_xml_completo_ia(
                                        xml_saneado_factura,
                                        ia_backend,
                                        contexto_reglas=_contexto_reglas_para_ia() + "\n\n" + hechos,
                                        modo_ahorro=st.session_state.get("_modo_ahorro_tokens", True),
                                        _modelo_local=modelo_local_cargado,
                                    )
                                    st.session_state[key_audit] = dictamen
                                # El dictamen de la IA ahora SÍ mueve el
                                # estatus de la factura, reevaluando el 80%.
                                _aplicar_dictamen_xml_ia(id_interno_factura, dictamen)
                                st.rerun()

                            col_nota.caption(
                                "Se envía el XML sin sellos ni datos del receptor (AUTOCOM)."
                                if ia_backend != "local" else
                                "IA local: el XML no sale de este equipo."
                            )

                            dictamen_xml = st.session_state.get(key_audit)
                            if dictamen_xml:
                                with st.container(border=True):
                                    if dictamen_xml.get("_desde_cache"):
                                        st.caption("♻️ Dictamen recuperado de caché: no se gastaron tokens.")
                                    _render_auditoria_xml_ia(dictamen_xml)

                        if conceptos_de_esta_factura.empty:
                            st.caption("Sin conceptos detallados (registro proveniente de CSV/Excel).")
                            continue

                        encabezados = st.columns([2.2, 3.5, 1.3, 1.7, 0.9, 1.6])
                        for col, txt in zip(encabezados, ["Clave SAT", "Concepto", "Importe", "Categoría", "Cumple", ""]):
                            col.markdown(f"**{txt}**")

                        for idx_det, fila_det in conceptos_de_esta_factura.iterrows():
                            idx_global_renglon += 1
                            c = st.columns([2.2, 3.5, 1.3, 1.7, 0.9, 1.6])
                            c[0].write(fila_det["Clave Prod/Serv SAT"])
                            c[1].write(fila_det["Concepto / Descripción"])
                            c[2].write(f"${fila_det['Importe']:,.2f}")
                            c[3].write(fila_det["Categoría"])
                            c[4].write(fila_det["Cumple"])

                            key_boton = f"btn_ia_{id_interno_factura}_{idx_fac}_{idx_det}_{idx_global_renglon}"
                            key_estado = f"sugerencia_ia_{id_interno_factura}_{idx_fac}_{idx_det}_{idx_global_renglon}"
                            key_aplicar = f"aplicar_ia_{id_interno_factura}_{idx_fac}_{idx_det}_{idx_global_renglon}"

                            if ia_disponible:
                                if c[5].button(
                                    "🧠 Consultar IA",
                                    key=key_boton,
                                    on_click=_abrir_factura,
                                    args=(id_interno_factura,),
                                ):
                                    loading_placeholder = st.empty()
                                    loading_placeholder.markdown("""
                                        <div class="ai-loading-box">
                                            <div class="ai-spinner"></div>
                                            <span style="color:#C8102E; font-weight:600; font-size:0.95rem;">
                                                Procesando matriz fiscal con IA...
                                            </span>
                                        </div>
                                    """, unsafe_allow_html=True)

                                    resultado_ia = analizar_concepto_ambiguo_ia(
                                        fila_det["Concepto / Descripción"],
                                        fila_det["Clave Prod/Serv SAT"],
                                        fila_cab["Uso CFDI"],
                                        float(fila_det["Importe"]),
                                        ia_backend,
                                        forma_pago=fila_cab.get("Forma Pago", "N/A"),
                                        metodo_pago=fila_cab.get("Método Pago", "N/A"),
                                        _modelo_local=modelo_local_cargado,
                                    )

                                    loading_placeholder.empty()
                                    st.session_state[key_estado] = resultado_ia
                            else:
                                c[5].caption("IA no disp.")

                            sugerencia = st.session_state.get(key_estado)
                            if sugerencia:
                                with st.container(border=True):
                                    if sugerencia.get("_ok") and sugerencia.get("categoria"):
                                        _render_dictamen_ia(sugerencia)

                                        if st.button(
                                            "✅ Aplicar clasificación de la IA",
                                            key=key_aplicar,
                                            on_click=_abrir_factura,
                                            args=(id_interno_factura,),
                                        ):
                                            _aplicar_clasificacion_ia(idx_det, id_interno_factura, sugerencia)

                                            categoria_aplicada = sugerencia.get("categoria", "GENERAL")
                                            if categoria_aplicada not in ("GENERAL", "REVISION"):
                                                es_nueva = bool(sugerencia.get("categoria_es_nueva")) or (
                                                    categoria_aplicada not in sat.CATEGORIAS_VALIDAS
                                                    and categoria_aplicada not in CATEGORIAS_DINAMICAS
                                                )
                                                with st.spinner("Generando propuesta para el administrador..."):
                                                    expansion = _expandir_categoria_con_ia(
                                                        categoria_aplicada,
                                                        fila_det["Concepto / Descripción"],
                                                        fila_det["Clave Prod/Serv SAT"],
                                                        ia_backend,
                                                        modelo_local_cargado,
                                                    )
                                                _crear_propuesta_categoria(
                                                    categoria_aplicada, es_nueva,
                                                    expansion["palabras"], expansion["uso_cfdi"],
                                                    expansion["descripcion_categoria"] or f"Categoría {categoria_aplicada}",
                                                    contexto={
                                                        "concepto_origen": fila_det["Concepto / Descripción"],
                                                        "clave_origen": fila_det["Clave Prod/Serv SAT"],
                                                        "factura_origen": fila_cab["Factura / Folio"],
                                                        "importe_origen": float(fila_det["Importe"]),
                                                        "analista": st.session_state.get("nombre_actual", "N/D"),
                                                    },
                                                )
                                                st.info(
                                                    f"📨 Se envió al administrador una propuesta de "
                                                    f"{'categoría nueva' if es_nueva else 'refuerzo de categoría'} "
                                                    f"**{categoria_aplicada}** para su revisión."
                                                )

                                            del st.session_state[key_estado]
                                            st.rerun()
                                    else:
                                        _render_error_ia(sugerencia)

                            # -----------------------------------------------
                            # Opinión del analista: de acuerdo / en desacuerdo,
                            # con justificación — y, si aplica, convertirlo en
                            # una propuesta de regla para revisión posterior.
                            # -----------------------------------------------
                            id_decision = _id_decision_concepto(
                                fila_cab["Factura / Folio"], fila_det["Concepto / Descripción"], fila_det["Importe"]
                            )
                            decisiones_df = _cargar_decisiones_analista()
                            decisiones_previas = decisiones_df[decisiones_df["id_decision"] == id_decision] \
                                if not decisiones_df.empty else decisiones_df

                            titulo_opinion = "🗣️ Opinión del analista"
                            if not decisiones_previas.empty:
                                titulo_opinion += f" ({len(decisiones_previas)} registrada(s))"

                            key_pref = f"opinion_{id_interno_factura}_{idx_fac}_{idx_det}_{idx_global_renglon}"
                            opinion_abierta = st.session_state.get("_opinion_abierta") == key_pref

                            with st.expander(titulo_opinion, expanded=opinion_abierta):
                                if not decisiones_previas.empty:
                                    for _, reg in decisiones_previas.iterrows():
                                        icono_reg = "✅" if reg["decision_analista"] == "DE_ACUERDO" else "❌"
                                        st.caption(f"{icono_reg} **{reg['usuario']}** · {reg['fecha']}")
                                        if reg["comentario"]:
                                            st.write(reg["comentario"])
                                    st.markdown("---")

                                # Los widgets también provocan rerun: con
                                # on_change se marca el estado antes de que se
                                # vuelvan a dibujar los expanders.
                                decision_radio = st.radio(
                                    "¿Estás de acuerdo con la clasificación / veredicto del sistema para este concepto?",
                                    ["De acuerdo", "En desacuerdo"],
                                    key=f"{key_pref}_radio",
                                    horizontal=True,
                                    on_change=_abrir_factura_y_opinion,
                                    args=(id_interno_factura, key_pref),
                                )
                                comentario_analista = st.text_area(
                                    "Justificación (obligatoria si estás en desacuerdo)",
                                    key=f"{key_pref}_comentario",
                                    placeholder="Ej: Esta clave SAT es de mantenimiento, no de activo fijo, aunque el importe sea alto.",
                                    on_change=_abrir_factura_y_opinion,
                                    args=(id_interno_factura, key_pref),
                                )

                                col_guardar, col_regla = st.columns(2)
                                with col_guardar:
                                    if st.button(
                                        "💾 Guardar opinión",
                                        key=f"{key_pref}_guardar",
                                        on_click=_abrir_factura_y_opinion,
                                        args=(id_interno_factura, key_pref),
                                    ):
                                        if decision_radio == "En desacuerdo" and not comentario_analista.strip():
                                            st.error("Escribe una justificación para registrar un desacuerdo.")
                                        else:
                                            _guardar_decision_analista({
                                                "id_decision": id_decision,
                                                "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                                "usuario": st.session_state.get("nombre_actual", "N/D"),
                                                "factura": fila_cab["Factura / Folio"],
                                                "concepto": fila_det["Concepto / Descripción"],
                                                "clave_prodserv": fila_det["Clave Prod/Serv SAT"],
                                                "categoria_sistema": fila_det["Categoría"],
                                                "cumple_sistema": fila_det["Cumple"],
                                                "decision_analista": (
                                                    "DE_ACUERDO" if decision_radio == "De acuerdo" else "EN_DESACUERDO"
                                                ),
                                                "comentario": comentario_analista.strip(),
                                            })
                                            st.success("Opinión guardada en la bitácora.")
                                            st.rerun()

                                if decision_radio == "En desacuerdo" and comentario_analista.strip():
                                    with col_regla:
                                        if not ia_disponible:
                                            st.caption("Configura la IA para poder proponer una categoría a partir de esto.")
                                        else:
                                            key_interp = f"{key_pref}_interpretacion"

                                            # La IA lee el comentario del
                                            # analista y arma la regla sola:
                                            # el analista ya no tiene que
                                            # adivinar el nombre de categoría.
                                            if st.button(
                                                "🤖 Que la IA interprete mi comentario",
                                                key=f"{key_pref}_interpretar",
                                                on_click=_abrir_factura_y_opinion,
                                                args=(id_interno_factura, key_pref),
                                            ):
                                                with st.spinner("La IA está traduciendo tu comentario a una regla..."):
                                                    st.session_state[key_interp] = interpretar_opinion_analista_ia(
                                                        comentario_analista,
                                                        fila_det["Concepto / Descripción"],
                                                        fila_det["Clave Prod/Serv SAT"],
                                                        fila_det["Categoría"],
                                                        ia_backend,
                                                        modelo_local=modelo_local_cargado,
                                                        categorias_existentes=sat.obtener_categorias_validas(
                                                            CATEGORIAS_DINAMICAS
                                                        ),
                                                    )
                                                st.rerun()

                                            interpretacion = st.session_state.get(key_interp)

                                            if interpretacion and not interpretacion.get("_ok"):
                                                _render_error_ia(interpretacion)
                                            elif interpretacion and not interpretacion.get("aplicable"):
                                                st.warning(
                                                    "La IA no pudo convertir tu comentario en una regla: "
                                                    f"{interpretacion.get('motivo_no_aplicable', 'sé más específico.')}"
                                                )
                                            elif interpretacion:
                                                riesgo = interpretacion.get("riesgo_fiscal", "medio")
                                                icono_riesgo = {"bajo": "🟢", "medio": "🟡", "alto": "🔴"}.get(riesgo, "⚪")
                                                st.markdown(
                                                    f"**Regla que entendió la IA:** {interpretacion['resumen_regla']}"
                                                )
                                                st.caption(
                                                    f"Categoría: **{interpretacion['categoria_propuesta']}** · "
                                                    f"Tipo: {interpretacion['tipo_cambio']} · "
                                                    f"Riesgo fiscal: {icono_riesgo} {riesgo}"
                                                )
                                                st.caption(
                                                    "Palabras clave: "
                                                    + ", ".join(interpretacion["palabras_clave"][:12])
                                                )

                                            # El analista puede corregir el
                                            # nombre que propuso la IA antes
                                            # de mandarlo al Admin.
                                            categoria_propuesta_input = st.text_input(
                                                "Categoría a proponer (puedes corregir lo que sugirió la IA)",
                                                value=(interpretacion or {}).get("categoria_propuesta", ""),
                                                key=f"{key_pref}_cat_propuesta",
                                                placeholder="Ej: AGUA, REFACCION, LIMPIEZA...",
                                                on_change=_abrir_factura_y_opinion,
                                                args=(id_interno_factura, key_pref),
                                            ).strip().upper().replace(" ", "_")

                                            if st.button(
                                                "📨 Proponer al administrador",
                                                key=f"{key_pref}_regla",
                                                on_click=_abrir_factura_y_opinion,
                                                args=(id_interno_factura, key_pref),
                                            ):
                                                if not categoria_propuesta_input:
                                                    st.error(
                                                        "Escribe el nombre de la categoría propuesta, o deja "
                                                        "que la IA interprete tu comentario primero."
                                                    )
                                                else:
                                                    es_nueva = (
                                                        categoria_propuesta_input not in sat.CATEGORIAS_VALIDAS
                                                        and categoria_propuesta_input not in CATEGORIAS_DINAMICAS
                                                    )
                                                    # Si la IA ya interpretó el
                                                    # comentario, se reutiliza
                                                    # su salida: cero tokens
                                                    # extra.
                                                    if (interpretacion and interpretacion.get("_ok")
                                                            and interpretacion.get("categoria_propuesta") == categoria_propuesta_input):
                                                        palabras = interpretacion["palabras_clave"]
                                                        usos = interpretacion["uso_cfdi_sugerido"]
                                                        descripcion_cat = interpretacion["descripcion_categoria"]
                                                    else:
                                                        with st.spinner("Generando propuesta..."):
                                                            expansion = _expandir_categoria_con_ia(
                                                                categoria_propuesta_input,
                                                                fila_det["Concepto / Descripción"],
                                                                fila_det["Clave Prod/Serv SAT"],
                                                                ia_backend,
                                                                modelo_local_cargado,
                                                            )
                                                        palabras = expansion["palabras"]
                                                        usos = expansion["uso_cfdi"]
                                                        descripcion_cat = expansion["descripcion_categoria"]

                                                    _crear_propuesta_categoria(
                                                        categoria_propuesta_input, es_nueva,
                                                        palabras, usos,
                                                        descripcion_cat or f"Categoría {categoria_propuesta_input}",
                                                        contexto={
                                                            "concepto_origen": fila_det["Concepto / Descripción"],
                                                            "clave_origen": fila_det["Clave Prod/Serv SAT"],
                                                            "factura_origen": fila_cab["Factura / Folio"],
                                                            "importe_origen": float(fila_det["Importe"]),
                                                            "analista": st.session_state.get("nombre_actual", "N/D"),
                                                            "comentario_analista": comentario_analista.strip(),
                                                            "regla_interpretada": (interpretacion or {}).get("resumen_regla", ""),
                                                        },
                                                    )
                                                    st.info(
                                                        "📨 Propuesta enviada — revísala en la pestaña "
                                                        "**🔔 Notificaciones**."
                                                    )

            st.markdown("---")
            st.subheader("📥 Exportar Reporte Ejecutivo")
            excel_bytes = generar_excel_presentable(st.session_state["df_cabecera"], st.session_state["df_detalle"])

            st.download_button(
                label="📥 Descargar Reporte Ejecutivo AUTOCOM (.xlsx)",
                data=excel_bytes,
                file_name=f"Auditoria_CFDI_AUTOCOM_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ---------------------------------------------------------
# 5. ESCANEO Y ANÁLISIS MASIVO DE XML CON IA
# ---------------------------------------------------------
def _generar_excel_con_observaciones_ia(df_cabecera, df_detalle, dictamenes_ia):
    """Envuelve generar_excel_presentable (sin tocar su lógica ni la del
    motor de reglas) agregando una columna con las observaciones concretas
    que generó la IA para cada concepto analizado."""
    df_detalle_con_ia = df_detalle.copy()
    df_detalle_con_ia["Observación IA"] = ""
    for idx_det, sugerencia in dictamenes_ia.items():
        if idx_det not in df_detalle_con_ia.index:
            continue
        if sugerencia.get("_ok"):
            texto = (
                f"[{sugerencia.get('veredicto', 'REVISION')}] "
                f"{sugerencia.get('justificacion', sugerencia.get('motivo', ''))} "
                f"Recomendación: {sugerencia.get('recomendacion', '')}"
            )
        else:
            texto = f"[ERROR IA] {sugerencia.get('_error', 'Sin respuesta válida.')}"
        df_detalle_con_ia.at[idx_det, "Observación IA"] = texto
    return generar_excel_presentable(df_cabecera, df_detalle_con_ia)


def render_tab_ia():
    st.markdown("### 🧠 Análisis Avanzado con IA — Escaneo Masivo")
    st.caption(
        "La IA analiza únicamente los conceptos que el motor de reglas fiscales no pudo "
        "clasificar con certeza (categoría REVISION / Cumple = REVISAR). El motor de reglas "
        "de AUTOCOM sigue siendo la autoridad final; la IA solo entrega un dictamen de apoyo."
    )

    if "df_detalle" not in st.session_state or "df_cabecera" not in st.session_state:
        st.info("Primero carga y audita tus comprobantes en la pestaña **Auditoría Masiva**.")
        return

    if not ia_disponible:
        st.warning("La IA no está configurada (agrega `groq_api.txt` o un modelo `.gguf`). "
                   "Configúrala desde la barra lateral para habilitar el escaneo masivo.")
        return

    df_cab = st.session_state["df_cabecera"]
    df_det = st.session_state["df_detalle"]
    conceptos_ambiguos = df_det[df_det["Categoría"] == "REVISION"]

    c1, c2 = st.columns(2)
    c1.metric("Conceptos ambiguos detectados", len(conceptos_ambiguos))
    c2.metric("Dictámenes de IA generados", len(st.session_state.get("_dictamenes_ia_masivo", {})))

    if conceptos_ambiguos.empty:
        st.success("No hay conceptos ambiguos pendientes de dictamen por IA.")
    else:
        if st.button("🧠 Escanear todos los XML con IA", use_container_width=True):
            resultados = {}
            progreso = st.progress(0)
            status = st.empty()
            total = len(conceptos_ambiguos)
            cuota_agotada_detectada = False

            for i, (idx_det, fila_det) in enumerate(conceptos_ambiguos.iterrows()):
                status.markdown(f"⏳ **Analizando concepto ({i + 1} de {total}):** `{fila_det['Concepto / Descripción']}`...")
                progreso.progress(int(((i + 1) / total) * 100))

                fila_cab_match = df_cab[df_cab["_id_interno"] == fila_det["_id_interno"]]
                uso_cfdi_factura = fila_cab_match.iloc[0]["Uso CFDI"] if not fila_cab_match.empty else "N/A"
                forma_pago_factura = fila_cab_match.iloc[0].get("Forma Pago", "N/A") if not fila_cab_match.empty else "N/A"
                metodo_pago_factura = fila_cab_match.iloc[0].get("Método Pago", "N/A") if not fila_cab_match.empty else "N/A"

                resultado_ia = analizar_concepto_ambiguo_ia(
                    fila_det["Concepto / Descripción"],
                    fila_det["Clave Prod/Serv SAT"],
                    uso_cfdi_factura,
                    float(fila_det["Importe"]),
                    ia_backend,
                    forma_pago=forma_pago_factura,
                    metodo_pago=metodo_pago_factura,
                    _modelo_local=modelo_local_cargado,
                )
                resultados[idx_det] = resultado_ia

                if resultado_ia.get("_cuota_agotada"):
                    cuota_agotada_detectada = True
                    break

            status.empty()
            progreso.empty()
            st.session_state["_dictamenes_ia_masivo"] = resultados

            if cuota_agotada_detectada:
                st.session_state["_ia_cuota_agotada_aviso"] = (
                    f"Se detuvo el escaneo masivo en el concepto {len(resultados)} de {total} porque "
                    "se agotaron los tokens/cuota disponibles de la IA. Los conceptos restantes quedaron "
                    "pendientes; vuelve a intentar el escaneo más tarde o revisa tu plan/API key de Groq."
                )
            else:
                st.session_state.pop("_ia_cuota_agotada_aviso", None)
            st.rerun()

    if st.session_state.get("_ia_cuota_agotada_aviso"):
        st.error(f"🚫 {st.session_state['_ia_cuota_agotada_aviso']}")

    dictamenes = st.session_state.get("_dictamenes_ia_masivo", {})
    if dictamenes:
        st.markdown("---")
        st.subheader("📋 Dictámenes de IA consolidados")

        filas_resumen = []
        for idx_det, sugerencia in dictamenes.items():
            if idx_det not in df_det.index:
                continue
            fila_det = df_det.loc[idx_det]
            if sugerencia.get("_ok"):
                filas_resumen.append({
                    "Folio": df_cab.loc[df_cab["_id_interno"] == fila_det["_id_interno"], "Factura / Folio"].iloc[0]
                        if (df_cab["_id_interno"] == fila_det["_id_interno"]).any() else "N/D",
                    "Concepto": fila_det["Concepto / Descripción"],
                    "Veredicto IA": sugerencia.get("veredicto", "REVISION"),
                    "Categoría sugerida": sugerencia.get("categoria", ""),
                    "Confianza": str(sugerencia.get("confianza", "n/d")).upper(),
                    "Justificación": sugerencia.get("justificacion", sugerencia.get("motivo", "")),
                    "Recomendación CxP": sugerencia.get("recomendacion", ""),
                })
            else:
                filas_resumen.append({
                    "Folio": "N/D",
                    "Concepto": fila_det["Concepto / Descripción"],
                    "Veredicto IA": "CUOTA AGOTADA" if sugerencia.get("_cuota_agotada") else "ERROR",
                    "Categoría sugerida": "-",
                    "Confianza": "-",
                    "Justificación": sugerencia.get("_error", "Sin respuesta válida."),
                    "Recomendación CxP": "Reintentar la consulta o clasificar manualmente.",
                })

        if filas_resumen:
            st.dataframe(pd.DataFrame(filas_resumen), use_container_width=True)

        col_aplicar, col_export = st.columns(2)
        with col_aplicar:
            if st.button("✅ Aplicar dictámenes de confianza alta/media", use_container_width=True):
                aplicados = 0
                for idx_det, sugerencia in dictamenes.items():
                    if idx_det not in df_det.index:
                        continue
                    if sugerencia.get("_ok") and sugerencia.get("confianza") in ("alta", "media"):
                        fila_det = df_det.loc[idx_det]
                        _aplicar_clasificacion_ia(idx_det, fila_det["_id_interno"], sugerencia)
                        aplicados += 1
                st.success(f"Se aplicaron {aplicados} dictámenes de IA con confianza alta/media.")
                st.rerun()

        with col_export:
            excel_bytes_ia = _generar_excel_con_observaciones_ia(df_cab, df_det, dictamenes)
            st.download_button(
                label="📥 Descargar Reporte Ejecutivo Actualizado (con IA)",
                data=excel_bytes_ia,
                file_name=f"Auditoria_CFDI_AUTOCOM_IA_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


def render_tab_notificaciones():
    st.markdown("### 🔔 Notificaciones — Propuestas de categoría")
    st.caption(
        "Cada vez que un analista da clic en \"✅ Aplicar clasificación de la IA\", el sistema genera "
        "aquí una propuesta — categoría nueva o refuerzo de una existente — con las palabras que la IA "
        "sugiere para reconocer conceptos similares en el futuro. **Nada queda aprendido de forma "
        "permanente hasta que un Admin la acepta.**"
    )

    propuestas = _cargar_propuestas_categorias()
    pendientes = [p for p in propuestas if p.get("estado") == "pendiente"]
    es_admin = st.session_state.get("rol_actual") == "Admin"

    if not pendientes:
        st.success("✅ No hay propuestas pendientes por el momento.")
    elif not es_admin:
        st.info(f"Hay {len(pendientes)} propuesta(s) pendiente(s). Solo el rol **Admin** puede aceptarlas o rechazarlas.")
        for p in pendientes:
            tipo = "🆕 Nueva" if p.get("es_nueva") else "➕ Refuerzo"
            st.write(f"- {tipo} **{p['categoria']}** — origen: \"{p.get('concepto_origen', '')}\" (propuesto por {p.get('analista', 'N/D')})")
    else:
        for p in pendientes:
            with st.container(border=True):
                tipo = "🆕 Categoría NUEVA" if p.get("es_nueva") else "➕ Refuerzo de categoría existente"
                st.markdown(f"**{tipo}: `{p['categoria']}`**")
                st.caption(
                    f"Origen: factura {p.get('factura_origen', 'N/D')} — \"{p.get('concepto_origen', '')}\" "
                    f"(Clave: {p.get('clave_origen', 'N/D')}, ${float(p.get('importe_origen', 0) or 0):,.2f}) — "
                    f"propuesto por {p.get('analista', 'N/D')} el {p.get('fecha_creacion', '')}."
                )
                if p.get("comentario_analista"):
                    st.markdown(f"🗣️ **Lo que dijo el analista:** *\"{p['comentario_analista']}\"*")
                if p.get("regla_interpretada"):
                    st.markdown(f"🤖 **Regla que entendió la IA:** {p['regla_interpretada']}")
                if p.get("descripcion_categoria"):
                    st.write(f"*{p['descripcion_categoria']}*")

                palabras_editable = st.text_area(
                    "Palabras/frases que reconocerán esta categoría (una por línea — puedes editarlas antes de aceptar)",
                    value="\n".join(p.get("palabras_sugeridas", [])),
                    key=f"palabras_{p['id']}",
                    height=120,
                )
                uso_cfdi_texto = st.text_input(
                    "Uso(s) CFDI coherente(s) (separados por coma)",
                    value=", ".join(p.get("uso_cfdi_sugerido", ["G03"])),
                    key=f"uso_{p['id']}",
                )

                col_aceptar, col_rechazar = st.columns(2)
                with col_aceptar:
                    if st.button("✅ Aceptar y aprender", key=f"aceptar_{p['id']}", use_container_width=True):
                        palabras_finales = [l.strip().upper() for l in palabras_editable.splitlines() if l.strip()]
                        usos_finales = [u.strip().upper() for u in uso_cfdi_texto.split(",") if u.strip()]
                        sat.agregar_o_reforzar_categoria_dinamica(
                            p["categoria"],
                            palabras_nuevas=palabras_finales,
                            uso_cfdi=usos_finales,
                            descripcion=p.get("descripcion_categoria", ""),
                            creado_por=st.session_state.get("nombre_actual", "Admin"),
                        )
                        _marcar_propuesta_resuelta(p["id"], "aceptada")
                        _cargar_categorias_dinamicas_cacheado.clear()
                        st.success(f"✅ Categoría '{p['categoria']}' actualizada — ya aplica desde la próxima auditoría.")
                        st.rerun()
                with col_rechazar:
                    if st.button("❌ Rechazar", key=f"rechazar_{p['id']}", use_container_width=True):
                        _marcar_propuesta_resuelta(p["id"], "rechazada")
                        st.info("Propuesta rechazada — no se modificó el motor de reglas.")
                        st.rerun()

    resueltas = [p for p in propuestas if p.get("estado") != "pendiente"]
    if resueltas:
        st.markdown("---")
        with st.expander(f"📜 Historial de propuestas resueltas ({len(resueltas)})"):
            resueltas_ordenadas = sorted(resueltas, key=lambda x: x.get("fecha_resolucion") or "", reverse=True)
            for p in resueltas_ordenadas[:100]:
                icono = "✅" if p.get("estado") == "aceptada" else "❌"
                st.caption(
                    f"{icono} **{p['categoria']}** ({'nueva' if p.get('es_nueva') else 'refuerzo'}) — "
                    f"resuelta el {p.get('fecha_resolucion', 'N/D')} por {p.get('resuelto_por', 'N/D')}. "
                    f"Origen: \"{p.get('concepto_origen', '')}\" ({p.get('analista', 'N/D')})."
                )


# ---------------------------------------------------------
# 6. GESTIÓN DE USUARIOS / AJUSTES
# ---------------------------------------------------------
def render_tab_ajustes():
    st.markdown("### ⚙️ Gestión de Usuarios / Ajustes")

    st.subheader("👤 Sesión actual")
    st.write(f"**Usuario:** {st.session_state.get('nombre_actual', 'N/D')}")
    st.write(f"**Rol:** {st.session_state.get('rol_actual', 'N/D')}")

    st.markdown("---")
    st.subheader("🧑‍🤝‍🧑 Usuarios del sistema")
    if st.session_state.get("rol_actual") == "Admin":
        usuarios_actuales = obtener_usuarios()
        tabla_usuarios = pd.DataFrame([
            {
                "Usuario": u,
                "Nombre": d["nombre"],
                "Rol": d["rol"],
                "Estado": "✅ Activo" if d.get("password_hash") else "⏳ Pendiente (aún no crea su contraseña)",
            }
            for u, d in usuarios_actuales.items()
        ])
        st.dataframe(tabla_usuarios, use_container_width=True)

        with st.expander("➕ Dar de alta a alguien nuevo"):
            st.caption(
                "Solo defines usuario, nombre y rol — la persona crea su propia "
                "contraseña la primera vez que entra a la app, en la pantalla de login."
            )
            with st.form("form_nuevo_usuario", clear_on_submit=True):
                nuevo_u = st.text_input("Usuario (sin espacios, ej. jperez)")
                nuevo_nombre = st.text_input("Nombre completo")
                nuevo_rol = st.selectbox("Rol", ["Analista", "Admin"])
                agregar = st.form_submit_button("Dar de alta")
            if agregar:
                nuevo_u_key = nuevo_u.strip().lower()
                if not nuevo_u_key or not nuevo_nombre.strip():
                    st.error("🚫 Completa usuario y nombre.")
                elif nuevo_u_key in usuarios_actuales:
                    st.error("🚫 Ese usuario ya existe.")
                else:
                    bd.crear_usuario_pendiente(nuevo_u_key, nuevo_nombre.strip(), nuevo_rol)
                    _cargar_usuarios_cacheado.clear()
                    st.success(f"✅ {nuevo_u_key} dado de alta. Ya puede crear su contraseña al entrar.")
                    st.rerun()

        if usuarios_actuales:
            with st.expander("🔁 Restablecer contraseña de alguien (la olvidó)"):
                st.caption(
                    "Esto borra su contraseña actual para que pueda crear una nueva. "
                    "Tú (Admin) nunca ves ni defines la contraseña de nadie."
                )
                usuario_a_resetear = st.selectbox(
                    "Usuario", list(usuarios_actuales.keys()), key="select_reset_password"
                )
                if st.button("Restablecer contraseña", key="btn_reset_password"):
                    bd.restablecer_password_usuario(usuario_a_resetear)
                    _cargar_usuarios_cacheado.clear()
                    st.success(f"✅ Ahora {usuario_a_resetear} puede crear una contraseña nueva al entrar.")
                    st.rerun()

            with st.expander("🗑️ Dar de baja a alguien"):
                usuario_a_borrar = st.selectbox(
                    "Usuario", list(usuarios_actuales.keys()), key="select_borrar_usuario"
                )
                if usuario_a_borrar == st.session_state.get("usuario_actual"):
                    st.caption("No puedes darte de baja a ti mismo mientras tienes la sesión abierta.")
                elif st.button("Dar de baja", key="btn_borrar_usuario"):
                    bd.eliminar_usuario(usuario_a_borrar)
                    _cargar_usuarios_cacheado.clear()
                    st.success(f"✅ {usuario_a_borrar} fue dado de baja.")
                    st.rerun()
    else:
        st.info("Solo el rol **Admin** puede ver el listado completo de usuarios.")

    st.markdown("---")
    st.subheader("🛡️ Estatus de validaciones oficiales del SAT")
    if "_listado_69b_ok" not in st.session_state:
        st.info("Listado 69-B: aún no se ha cargado (se descarga en la primera auditoría del lote).")
    elif st.session_state["_listado_69b_ok"]:
        st.success(f"✅ Lista negra 69-B cargada: {st.session_state['_listado_69b_len']:,} RFC's.")
    else:
        st.error(f"⚠️ Error listado 69-B: {st.session_state.get('_listado_69b_error', '')}")
    st.write(f"**Validación de vigencia en tiempo real (Web Service SAT):** "
             f"{'Activada' if validar_estatus_sat else 'Desactivada'}")
    st.caption("Estos ajustes se controlan desde la barra lateral.")

    st.markdown("---")
    st.subheader("🧠 Estatus del backend de IA")
    if ia_disponible:
        backend_texto = "☁️ Groq (nube)" if ia_backend == "nube" else "💻 Modelo local"
        st.success(f"IA activa — {backend_texto}")
    else:
        st.warning("IA no configurada. Agrega `groq_api.txt` (Groq) o un modelo `.gguf` en `/models`.")

    st.markdown("---")
    st.subheader("🗣️ Bitácora de opiniones del analista")
    st.caption(
        "Historial completo de acuerdos/desacuerdos que el equipo ha registrado sobre "
        "clasificaciones del motor de reglas o de la IA, con su justificación."
    )
    bitacora = _cargar_decisiones_analista()
    if bitacora.empty:
        st.info("Todavía no se ha registrado ninguna opinión.")
    else:
        filtro_desacuerdos = st.checkbox("Mostrar solo desacuerdos", value=False)
        bitacora_mostrar = bitacora[bitacora["decision_analista"] == "EN_DESACUERDO"] if filtro_desacuerdos else bitacora
        st.dataframe(
            bitacora_mostrar.drop(columns=["id_decision"], errors="ignore").sort_values("fecha", ascending=False),
            use_container_width=True,
        )
        st.download_button(
            "📥 Descargar bitácora completa (.csv)",
            data=bitacora.to_csv(index=False).encode("utf-8-sig"),
            file_name="decisiones_analista.csv",
            mime="text/csv",
        )


# ---------------------------------------------------------
# 7. CREACIÓN DE PESTAÑAS PRINCIPALES
# ---------------------------------------------------------
# Los Analistas solo trabajan la auditoría del día a día: Notificaciones
# (aprobar/rechazar categorías propuestas) y Ajustes son cosas de Admin.
if ES_ADMIN:
    tab_auditoria, tab_ia, tab_notificaciones, tab_ajustes = st.tabs([
        "📋 Auditoría Masiva",
        "🧠 Análisis Avanzado con IA",
        f"🔔 Notificaciones{f' ({_contar_propuestas_pendientes()})' if _contar_propuestas_pendientes() else ''}",
        "⚙️ Gestión de Usuarios / Ajustes",
    ])

    with tab_auditoria:
        render_tab_auditoria()

    with tab_ia:
        render_tab_ia()

    with tab_notificaciones:
        render_tab_notificaciones()

    with tab_ajustes:
        render_tab_ajustes()
else:
    tab_auditoria, tab_ia = st.tabs([
        "📋 Auditoría Masiva",
        "🧠 Análisis Avanzado con IA",
    ])

    with tab_auditoria:
        render_tab_auditoria()

    with tab_ia:
        render_tab_ia()

