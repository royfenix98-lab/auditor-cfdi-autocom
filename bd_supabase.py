# -*- coding: utf-8 -*-
"""
bd_supabase.py — Persistencia externa (Supabase / Postgres) para AUTOCOM.
===========================================================================

Mismo propósito que bd_turso.py (documentado ahí): mover lo que la app
"aprende" fuera del contenedor efímero de Streamlit Community Cloud. Este
archivo hace exactamente lo mismo pero contra una base Postgres de
Supabase en vez de Turso.

CONFIGURACIÓN NECESARIA (una sola vez) — ver también INSTRUCCIONES_SUPABASE.md
--------------------------------------------------------------------------
1. Crear proyecto gratis en https://supabase.com
2. En el proyecto: Project Settings -> Database -> Connection string ->
   pestaña "Transaction pooler" (¡IMPORTANTE: no la "Direct connection"! —
   Streamlit Community Cloud no soporta IPv6 y la conexión directa de
   Supabase lo requiere; el pooler sí funciona por IPv4).
3. Copiar esa cadena (empieza con postgresql://postgres.xxxx:...) y
   reemplazar [YOUR-PASSWORD] por la contraseña que pusiste al crear el
   proyecto.
4. Guardarla como secreto:
   - Local: .streamlit/secrets.toml -> SUPABASE_DB_URL = "postgresql://..."
   - Streamlit Community Cloud: Settings -> Secrets -> pegar la misma línea.

Este módulo crea las tablas solas la primera vez que se usan
(`asegurar_esquema`).
"""

import os
import json
import threading
from datetime import datetime

import psycopg2
import psycopg2.extras
from psycopg2.pool import ThreadedConnectionPool

try:
    import streamlit as st
    _SECRETS = st.secrets
except Exception:  # permite importar este módulo fuera de Streamlit (scripts sueltos)
    _SECRETS = {}


def _config(nombre):
    try:
        if nombre in _SECRETS:
            return _SECRETS[nombre]
    except Exception:
        pass
    return os.environ.get(nombre, "")


SUPABASE_DB_URL = _config("SUPABASE_DB_URL")

_lock = threading.Lock()
_pool = None
_esquema_listo = False


class ErrorConfiguracionBD(RuntimeError):
    pass


def _obtener_pool():
    global _pool
    if _pool is None:
        if not SUPABASE_DB_URL:
            raise ErrorConfiguracionBD(
                "Falta SUPABASE_DB_URL. Configúralo en .streamlit/secrets.toml "
                "(local) o en Settings > Secrets de la app en Streamlit "
                "Community Cloud, con la cadena del 'Transaction pooler' de "
                "Supabase (Project Settings > Database > Connection string)."
            )
        _pool = ThreadedConnectionPool(1, 5, dsn=SUPABASE_DB_URL)
    return _pool


def ejecutar(sql, args=None, devolver_filas=True):
    """Abre una conexión del pool, ejecuta y regresa filas (lista de tuplas)
    si la sentencia produce resultado, o [] si no."""
    pool = _obtener_pool()
    with _lock:
        conn = pool.getconn()
    try:
        with conn:
            with conn.cursor() as cur:
                cur.execute(sql, args or [])
                if devolver_filas and cur.description is not None:
                    return cur.fetchall()
                return []
    finally:
        with _lock:
            pool.putconn(conn)


_ESQUEMA = [
    """CREATE TABLE IF NOT EXISTS categorias_dinamicas (
        nombre TEXT PRIMARY KEY,
        palabras TEXT NOT NULL DEFAULT '[]',
        prefijos TEXT NOT NULL DEFAULT '[]',
        uso_cfdi TEXT NOT NULL DEFAULT '[]',
        descripcion TEXT DEFAULT '',
        creado_por TEXT DEFAULT '',
        fecha_creacion TEXT DEFAULT ''
    )""",
    """CREATE TABLE IF NOT EXISTS propuestas_categorias (
        id TEXT PRIMARY KEY,
        categoria TEXT,
        es_nueva INTEGER,
        palabras_sugeridas TEXT,
        uso_cfdi_sugerido TEXT,
        descripcion_categoria TEXT,
        estado TEXT,
        fecha_creacion TEXT,
        fecha_resolucion TEXT,
        resuelto_por TEXT,
        contexto_extra TEXT DEFAULT '{}'
    )""",
    """CREATE TABLE IF NOT EXISTS decisiones_analista (
        id BIGSERIAL PRIMARY KEY,
        id_decision TEXT,
        fecha TEXT,
        usuario TEXT,
        factura TEXT,
        concepto TEXT,
        clave_prodserv TEXT,
        categoria_sistema TEXT,
        cumple_sistema TEXT,
        decision_analista TEXT,
        comentario TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS cache_auditorias_ia (
        huella TEXT PRIMARY KEY,
        dictamen TEXT,
        fecha_creacion TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS usuarios (
        usuario TEXT PRIMARY KEY,
        nombre TEXT NOT NULL,
        rol TEXT NOT NULL,
        password_hash TEXT,
        password_salt TEXT,
        fecha_creacion TEXT,
        fecha_actualizacion TEXT
    )""",
]


def asegurar_esquema():
    global _esquema_listo
    if _esquema_listo:
        return
    with _lock:
        if _esquema_listo:
            return
    for sentencia in _ESQUEMA:
        ejecutar(sentencia, devolver_filas=False)
    with _lock:
        _esquema_listo = True


# ---------------------------------------------------------
# CATEGORÍAS DINÁMICAS  (reemplaza categorias_dinamicas.json)
# ---------------------------------------------------------
def cargar_categorias_dinamicas():
    asegurar_esquema()
    try:
        filas = ejecutar(
            "SELECT nombre, palabras, prefijos, uso_cfdi, descripcion, creado_por, "
            "fecha_creacion FROM categorias_dinamicas"
        )
    except Exception:
        return {}
    categorias = {}
    for nombre, palabras, prefijos, uso_cfdi, descripcion, creado_por, fecha_creacion in filas:
        categorias[nombre] = {
            "palabras": json.loads(palabras or "[]"),
            "prefijos": json.loads(prefijos or "[]"),
            "uso_cfdi": json.loads(uso_cfdi or "[]"),
            "descripcion": descripcion or "",
            "creado_por": creado_por or "",
            "fecha_creacion": fecha_creacion or "",
        }
    return categorias


def guardar_categoria_dinamica(nombre, entrada):
    """Upsert de UNA categoría (no reescribe la tabla completa)."""
    asegurar_esquema()
    ejecutar(
        """INSERT INTO categorias_dinamicas
             (nombre, palabras, prefijos, uso_cfdi, descripcion, creado_por, fecha_creacion)
           VALUES (%s, %s, %s, %s, %s, %s, %s)
           ON CONFLICT (nombre) DO UPDATE SET
             palabras=excluded.palabras,
             prefijos=excluded.prefijos,
             uso_cfdi=excluded.uso_cfdi,
             descripcion=excluded.descripcion,
             creado_por=excluded.creado_por""",
        [
            nombre,
            json.dumps(entrada.get("palabras", []), ensure_ascii=False),
            json.dumps(entrada.get("prefijos", []), ensure_ascii=False),
            json.dumps(entrada.get("uso_cfdi", []), ensure_ascii=False),
            entrada.get("descripcion", ""),
            entrada.get("creado_por", ""),
            entrada.get("fecha_creacion", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ],
        devolver_filas=False,
    )


def guardar_categorias_dinamicas(categorias):
    """Compatibilidad con la firma antigua (dict completo) — hace upsert
    entrada por entrada en vez de sobreescribir un archivo."""
    for nombre, entrada in categorias.items():
        guardar_categoria_dinamica(nombre, entrada)


# ---------------------------------------------------------
# PROPUESTAS DE CATEGORÍA  (reemplaza propuestas_categorias.json)
# ---------------------------------------------------------
_CAMPOS_PROPUESTA_FIJOS = {
    "id", "categoria", "es_nueva", "palabras_sugeridas", "uso_cfdi_sugerido",
    "descripcion_categoria", "estado", "fecha_creacion", "fecha_resolucion", "resuelto_por",
}


def cargar_propuestas_categorias():
    asegurar_esquema()
    filas = ejecutar(
        "SELECT id, categoria, es_nueva, palabras_sugeridas, uso_cfdi_sugerido, "
        "descripcion_categoria, estado, fecha_creacion, fecha_resolucion, resuelto_por, "
        "contexto_extra FROM propuestas_categorias ORDER BY fecha_creacion"
    )
    propuestas = []
    for (id_, categoria, es_nueva, palabras_sugeridas, uso_cfdi_sugerido, descripcion_categoria,
         estado, fecha_creacion, fecha_resolucion, resuelto_por, contexto_extra) in filas:
        propuesta = {
            "id": id_,
            "categoria": categoria,
            "es_nueva": bool(es_nueva),
            "palabras_sugeridas": json.loads(palabras_sugeridas or "[]"),
            "uso_cfdi_sugerido": json.loads(uso_cfdi_sugerido or "[]"),
            "descripcion_categoria": descripcion_categoria or "",
            "estado": estado,
            "fecha_creacion": fecha_creacion,
            "fecha_resolucion": fecha_resolucion,
            "resuelto_por": resuelto_por,
        }
        propuesta.update(json.loads(contexto_extra or "{}"))
        propuestas.append(propuesta)
    return propuestas


def guardar_propuesta_categoria(propuesta):
    """Upsert de UNA propuesta completa (mismo dict que ya maneja app.py)."""
    asegurar_esquema()
    extra = {k: v for k, v in propuesta.items() if k not in _CAMPOS_PROPUESTA_FIJOS}
    ejecutar(
        """INSERT INTO propuestas_categorias
             (id, categoria, es_nueva, palabras_sugeridas, uso_cfdi_sugerido,
              descripcion_categoria, estado, fecha_creacion, fecha_resolucion,
              resuelto_por, contexto_extra)
           VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
           ON CONFLICT (id) DO UPDATE SET
             categoria=excluded.categoria,
             es_nueva=excluded.es_nueva,
             palabras_sugeridas=excluded.palabras_sugeridas,
             uso_cfdi_sugerido=excluded.uso_cfdi_sugerido,
             descripcion_categoria=excluded.descripcion_categoria,
             estado=excluded.estado,
             fecha_creacion=excluded.fecha_creacion,
             fecha_resolucion=excluded.fecha_resolucion,
             resuelto_por=excluded.resuelto_por,
             contexto_extra=excluded.contexto_extra""",
        [
            propuesta.get("id"),
            propuesta.get("categoria"),
            int(bool(propuesta.get("es_nueva"))),
            json.dumps(propuesta.get("palabras_sugeridas", []), ensure_ascii=False),
            json.dumps(propuesta.get("uso_cfdi_sugerido", []), ensure_ascii=False),
            propuesta.get("descripcion_categoria", ""),
            propuesta.get("estado"),
            propuesta.get("fecha_creacion"),
            propuesta.get("fecha_resolucion"),
            propuesta.get("resuelto_por"),
            json.dumps(extra, ensure_ascii=False),
        ],
        devolver_filas=False,
    )


def guardar_propuestas_categorias(propuestas):
    """Compatibilidad con la firma antigua (lista completa)."""
    for p in propuestas:
        guardar_propuesta_categoria(p)


# ---------------------------------------------------------
# BITÁCORA DE DECISIONES DEL ANALISTA  (reemplaza decisiones_analista.csv)
# ---------------------------------------------------------
COLUMNAS_DECISIONES_ANALISTA = [
    "id_decision", "fecha", "usuario", "factura", "concepto", "clave_prodserv",
    "categoria_sistema", "cumple_sistema", "decision_analista", "comentario",
]


def cargar_decisiones_analista():
    asegurar_esquema()
    filas = ejecutar(
        "SELECT id_decision, fecha, usuario, factura, concepto, clave_prodserv, "
        "categoria_sistema, cumple_sistema, decision_analista, comentario "
        "FROM decisiones_analista ORDER BY fecha"
    )
    return [dict(zip(COLUMNAS_DECISIONES_ANALISTA, fila)) for fila in filas]


def agregar_decision_analista(registro):
    """Siempre INSERTA una fila nueva — el mismo concepto puede acumular
    varias opiniones a lo largo del tiempo, así que nunca se sobreescribe."""
    asegurar_esquema()
    ejecutar(
        """INSERT INTO decisiones_analista
             (id_decision, fecha, usuario, factura, concepto, clave_prodserv,
              categoria_sistema, cumple_sistema, decision_analista, comentario)
           VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)""",
        [registro.get(c, "") for c in COLUMNAS_DECISIONES_ANALISTA],
        devolver_filas=False,
    )


# ---------------------------------------------------------
# CACHÉ DE AUDITORÍAS IA  (reemplaza cache_auditorias_ia.json)
# ---------------------------------------------------------
def cache_auditoria_leer(huella):
    asegurar_esquema()
    try:
        filas = ejecutar("SELECT dictamen FROM cache_auditorias_ia WHERE huella = %s", [huella])
    except Exception:
        return None
    if not filas:
        return None
    try:
        return json.loads(filas[0][0])
    except (ValueError, TypeError):
        return None


def cache_auditoria_guardar(huella, dictamen, maximo_entradas=500):
    asegurar_esquema()
    datos = {k: v for k, v in dictamen.items() if k != "_datos_enviados"}
    try:
        ejecutar(
            """INSERT INTO cache_auditorias_ia (huella, dictamen, fecha_creacion)
               VALUES (%s, %s, %s)
               ON CONFLICT (huella) DO UPDATE SET
                 dictamen=excluded.dictamen, fecha_creacion=excluded.fecha_creacion""",
            [huella, json.dumps(datos, ensure_ascii=False), datetime.now().isoformat()],
            devolver_filas=False,
        )
        # Poda: conserva solo las N entradas más recientes.
        ejecutar(
            """DELETE FROM cache_auditorias_ia WHERE huella NOT IN (
                 SELECT huella FROM cache_auditorias_ia
                 ORDER BY fecha_creacion DESC LIMIT %s)""",
            [maximo_entradas],
            devolver_filas=False,
        )
    except Exception:
        pass


def contar_cache_auditorias():
    asegurar_esquema()
    try:
        filas = ejecutar("SELECT COUNT(*) FROM cache_auditorias_ia")
        return filas[0][0]
    except Exception:
        return 0


# ---------------------------------------------------------
# USUARIOS  (reemplaza el diccionario USUARIOS_AUTOCOM en texto plano)
# ---------------------------------------------------------
def cargar_usuarios():
    asegurar_esquema()
    filas = ejecutar(
        "SELECT usuario, nombre, rol, password_hash, password_salt FROM usuarios"
    )
    usuarios = {}
    for usuario, nombre, rol, password_hash, password_salt in filas:
        usuarios[usuario] = {
            "nombre": nombre,
            "rol": rol,
            "password_hash": password_hash,
            "password_salt": password_salt,
        }
    return usuarios


def contar_usuarios():
    asegurar_esquema()
    filas = ejecutar("SELECT COUNT(*) FROM usuarios")
    return filas[0][0]


def crear_usuario_pendiente(usuario, nombre, rol):
    """Da de alta usuario/nombre/rol SIN contraseña — la persona crea la
    suya la primera vez que entra a la app. Si el usuario ya existe, no
    hace nada (no se sobreescribe una cuenta activa)."""
    asegurar_esquema()
    ejecutar(
        """INSERT INTO usuarios (usuario, nombre, rol, fecha_creacion)
           VALUES (%s, %s, %s, %s)
           ON CONFLICT (usuario) DO NOTHING""",
        [usuario, nombre, rol, datetime.now().isoformat()],
        devolver_filas=False,
    )


def crear_primer_admin(usuario, nombre, password_hash, password_salt):
    """Solo para el arranque inicial, cuando la tabla usuarios está
    completamente vacía: crea la primera cuenta, con rol Admin y
    contraseña ya puesta por la propia persona."""
    asegurar_esquema()
    ejecutar(
        """INSERT INTO usuarios (usuario, nombre, rol, password_hash, password_salt, fecha_creacion)
           VALUES (%s, %s, 'Admin', %s, %s, %s)
           ON CONFLICT (usuario) DO NOTHING""",
        [usuario, nombre, password_hash, password_salt, datetime.now().isoformat()],
        devolver_filas=False,
    )


def establecer_password_usuario(usuario, password_hash, password_salt):
    asegurar_esquema()
    ejecutar(
        """UPDATE usuarios SET password_hash=%s, password_salt=%s, fecha_actualizacion=%s
           WHERE usuario=%s""",
        [password_hash, password_salt, datetime.now().isoformat(), usuario],
        devolver_filas=False,
    )


def restablecer_password_usuario(usuario):
    """El Admin usa esto si alguien olvidó su contraseña: la borra para que
    la persona vuelva a crear una nueva la próxima vez que entre. El Admin
    NUNCA ve ni define la contraseña de nadie."""
    asegurar_esquema()
    ejecutar(
        "UPDATE usuarios SET password_hash=NULL, password_salt=NULL WHERE usuario=%s",
        [usuario],
        devolver_filas=False,
    )


def eliminar_usuario(usuario):
    asegurar_esquema()
    ejecutar("DELETE FROM usuarios WHERE usuario=%s", [usuario], devolver_filas=False)
