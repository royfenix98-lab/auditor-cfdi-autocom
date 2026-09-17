# -*- coding: utf-8 -*-
"""
auth.py — Hasheo de contraseñas para AUTOCOM.
================================================
Usa PBKDF2-HMAC-SHA256, incluido en la librería estándar de Python (no hay
que instalar nada nuevo), con una "sal" aleatoria distinta por usuario. Así:
  - Dos personas con la misma contraseña no se ven igual en la base.
  - Nadie con acceso de lectura a Supabase (ni un respaldo de la base) puede
    leer las contraseñas — solo ve el hash, que no se puede revertir.
  - Cada usuario CREA su propia contraseña (nunca la ve ni la define el
    Admin), así que ni siquiera el Admin puede saber la contraseña de nadie.
"""
import os
import hmac
import hashlib

_ITERACIONES = 200_000


def generar_salt():
    return os.urandom(16).hex()


def hash_password(password, salt):
    return hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), bytes.fromhex(salt), _ITERACIONES
    ).hex()


def verificar_password(password, salt, hash_guardado):
    if not salt or not hash_guardado or not password:
        return False
    calculado = hash_password(password, salt)
    return hmac.compare_digest(calculado, hash_guardado)


def contrasena_es_valida(password):
    """Regla mínima: 8+ caracteres. Súbele aquí si quieres exigir más."""
    return isinstance(password, str) and len(password) >= 8
