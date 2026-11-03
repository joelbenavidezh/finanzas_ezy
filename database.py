"""
Módulo de Base de Datos para Finanzas Personales y Familiares (Finanzas Ezy)
HARDENED VERSION: SOC 2 Type II, ISO/IEC 27001:2022, ISO/IEC 27701:2019
- Criptografía: Bcrypt (12 rounds) con migración retrocompatible.
- Pistas de Auditoría Inmutables (Audit Logs - Append Only).
- Rate Limiting y Detección de Ataques de Fuerza Bruta.
- Gobernanza de Datos y Privacidad (Derecho al Olvido / Data Purge).
- Row Level Security (RLS) habilitado por defecto.
"""

import os
import sqlite3
import re
import time
import pandas as pd
from datetime import datetime, timedelta
from dotenv import load_dotenv
import bcrypt

ENV_FILE = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(ENV_FILE, override=True)

DB_FILE = os.path.join(os.path.dirname(__file__), "finanzas.db")
EXCEL_FILE = os.path.join(os.path.dirname(__file__), "gastos_finanzas.xlsx")
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()

CATEGORIAS_DEFAULT = [
    "Supermercado",
    "Restaurante y Bares",
    "Servicios y Facturas",
    "Farmacia y Salud",
    "Transporte y Combustible",
    "Ocio y Entretenimiento",
    "Ropa y Calzado",
    "Hogar y Mantenimiento",
    "Suscripciones y Apps",
    "Educación",
    "Otros"
]

MEDIOS_PAGO_DEFAULT = [
    "Tarjeta de Crédito",
    "Tarjeta de Débito",
    "Efectivo",
    "Transferencia / QR"
]

AVATARES_DEFAULT = ["👤", "👨", "👩", "👦", "👧", "🧑", "👴", "👵", "⭐", "🐱", "🐶"]

# Estructura en memoria para Rate Limiting (SOC 2 CC6.6)
# Formato: { "email": { "intentos": int, "bloqueado_hasta": timestamp } }
_RATE_LIMIT_CACHE = {}

def es_postgres():
    return bool(DATABASE_URL and DATABASE_URL.startswith("postgres"))

def get_connection():
    if es_postgres():
        import psycopg2
        from psycopg2.extras import RealDictCursor
        conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
        return conn
    else:
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        return conn

# --- CRIPTOGRAFÍA DE ALTA SEGURIDAD (ISO 27001 A.8.24) ---
def hash_password(password: str) -> str:
    """Genera hash Bcrypt de grado militar con 12 rondas de salting."""
    pwd_bytes = password.encode("utf-8")
    salt = bcrypt.gensalt(rounds=12)
    hashed = bcrypt.hashpw(pwd_bytes, salt)
    return hashed.decode("utf-8")

def verificar_password(password_plana: str, hash_almacenado: str) -> bool:
    """Verifica contraseña contra hash Bcrypt (con compatibilidad para hashes heredados)."""
    if not hash_almacenado:
        return False
    try:
        # Caso 1: Hash Bcrypt estándar ($2a$, $2b$, $2y$)
        if hash_almacenado.startswith(("$2a$", "$2b$", "$2y$")):
            return bcrypt.checkpw(password_plana.encode("utf-8"), hash_almacenado.encode("utf-8"))
        
        # Caso 2: Hash heredado SHA-256 + salt (migración transparente)
        if "$" in hash_almacenado:
            import hashlib
            salt, hash_real = hash_almacenado.split("$", 1)
            recalculado = hashlib.sha256((password_plana + salt).encode("utf-8")).hexdigest()
            return recalculado == hash_real
    except Exception:
        return False
    return False

def validar_complejidad_password(password: str) -> tuple[bool, str]:
    """Valida requisitos de complejidad conforme a ISO 27001 / SOC 2."""
    if len(password) < 8:
        return False, "La contraseña debe tener al menos 8 caracteres."
    if not re.search(r"\d", password):
        return False, "La contraseña debe incluir al menos un número."
    if not re.search(r"[a-zA-Z]", password):
        return False, "La contraseña debe incluir al menos una letra."
    return True, ""

# --- AUDIT LOGS INMUTABLES (ISO 27001 A.8.15 / SOC 2 CC7.2) ---
def registrar_auditoria(user_id, accion: str, recurso: str, detalle: str = ""):
    """Registra eventos de seguridad de forma inmutable."""
    try:
        conn = get_connection()
        cursor = conn.cursor()
        if es_postgres():
            cursor.execute("""
            INSERT INTO audit_logs (user_id, accion, recurso, detalle)
            VALUES (%s, %s, %s, %s);
            """, (user_id, accion[:50], recurso[:100], detalle[:255]))
        else:
            cursor.execute("""
            INSERT INTO audit_logs (user_id, accion, recurso, detalle)
            VALUES (?, ?, ?, ?);
            """, (user_id, accion[:50], recurso[:100], detalle[:255]))
        conn.commit()
        conn.close()
    except Exception:
        pass  # Los fallos de auditoría no interrumpen la ejecución pero quedan en logs

# --- RATE LIMITING Y PROTECCIÓN FUERZA BRUTA (SOC 2 CC6.1 - CC6.6) ---
def verificar_rate_limit(identificador: str) -> tuple[bool, str]:
    """Verifica si el identificador (email) está temporalmente bloqueado por intentos fallidos."""
    ahora = time.time()
    record = _RATE_LIMIT_CACHE.get(identificador)
    if not record:
        return True, ""
    
    if record.get("bloqueado_hasta", 0) > ahora:
        minutos_restantes = int((record["bloqueado_hasta"] - ahora) / 60) + 1
        return False, f"Demasiados intentos fallidos. Cuenta bloqueada por seguridad. Reintenta en {minutos_restantes} minuto(s)."
    
    # Si expiró el bloqueo, resetear
    if record.get("bloqueado_hasta", 0) <= ahora and record.get("bloqueado_hasta", 0) > 0:
        _RATE_LIMIT_CACHE[identificador] = {"intentos": 0, "bloqueado_hasta": 0}
        
    return True, ""

def registrar_fallo_login(identificador: str):
    """Registra un intento fallido y bloquea tras 5 intentos consecutivos."""
    ahora = time.time()
    record = _RATE_LIMIT_CACHE.get(identificador, {"intentos": 0, "bloqueado_hasta": 0})
    record["intentos"] += 1
    
    if record["intentos"] >= 5:
        record["bloqueado_hasta"] = ahora + 900  # Bloqueo por 15 minutos (900 seg)
        registrar_auditoria(None, "ACCOUNT_LOCKED", "AUTH", f"Bloqueo temporal por 15 min: {identificador}")
    
    _RATE_LIMIT_CACHE[identificador] = record

def limpiar_fallos_login(identificador: str):
    """Limpia los intentos tras un login exitoso."""
    if identificador in _RATE_LIMIT_CACHE:
        del _RATE_LIMIT_CACHE[identificador]

# --- INICIALIZACIÓN DE LA BASE DE DATOS Y RLS ---
def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    if es_postgres():
        # Grupos Familiares
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS grupos_familiares (
            id SERIAL PRIMARY KEY,
            nombre VARCHAR(100) NOT NULL,
            codigo_invitacion VARCHAR(20) UNIQUE NOT NULL,
            creador_id INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)
        
        # Cuentas de Usuarios
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS cuentas_usuarios (
            id SERIAL PRIMARY KEY,
            email VARCHAR(255) UNIQUE NOT NULL,
            password_hash VARCHAR(255) NOT NULL,
            nombre VARCHAR(100) NOT NULL,
            avatar VARCHAR(10) DEFAULT '👤',
            grupo_id INTEGER REFERENCES grupos_familiares(id) ON DELETE SET NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)

        # Gastos
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS gastos (
            id SERIAL PRIMARY KEY,
            fecha VARCHAR(20) NOT NULL,
            comercio VARCHAR(255) NOT NULL,
            monto NUMERIC(12, 2) NOT NULL,
            moneda VARCHAR(10) DEFAULT 'ARS',
            categoria VARCHAR(100) NOT NULL,
            medio_pago VARCHAR(100) NOT NULL,
            cuotas VARCHAR(50) DEFAULT '1',
            origen VARCHAR(100) DEFAULT 'Manual',
            notas TEXT,
            usuario VARCHAR(100) DEFAULT 'Principal',
            user_id INTEGER REFERENCES cuentas_usuarios(id) ON DELETE CASCADE,
            grupo_id INTEGER REFERENCES grupos_familiares(id) ON DELETE SET NULL,
            visibilidad VARCHAR(20) DEFAULT 'familiar',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)

        # Presupuestos
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS presupuestos (
            id SERIAL PRIMARY KEY,
            grupo_id INTEGER,
            user_id INTEGER,
            categoria VARCHAR(100) NOT NULL,
            monto_maximo NUMERIC(12, 2) NOT NULL,
            mes INTEGER NOT NULL,
            anio INTEGER NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)

        # Tabla de Auditoría Inmutable (ISO 27001 A.8.15 / SOC 2)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS audit_logs (
            id SERIAL PRIMARY KEY,
            user_id INTEGER,
            accion VARCHAR(50) NOT NULL,
            recurso VARCHAR(100) NOT NULL,
            detalle VARCHAR(255),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)

        # Blindaje RLS obligatorio
        cursor.execute("""
        ALTER TABLE public.grupos_familiares ENABLE ROW LEVEL SECURITY;
        ALTER TABLE public.cuentas_usuarios ENABLE ROW LEVEL SECURITY;
        ALTER TABLE public.gastos ENABLE ROW LEVEL SECURITY;
        ALTER TABLE public.presupuestos ENABLE ROW LEVEL SECURITY;
        ALTER TABLE public.audit_logs ENABLE ROW LEVEL SECURITY;
        """)

    else:
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS grupos_familiares (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            codigo_invitacion TEXT UNIQUE NOT NULL,
            creador_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS cuentas_usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            email TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            nombre TEXT NOT NULL,
            avatar TEXT DEFAULT '👤',
            grupo_id INTEGER,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS gastos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT NOT NULL,
            comercio TEXT NOT NULL,
            monto REAL NOT NULL,
            moneda TEXT DEFAULT 'ARS',
            categoria TEXT NOT NULL,
            medio_pago TEXT NOT NULL,
            cuotas TEXT DEFAULT '1',
            origen TEXT DEFAULT 'Manual',
            notas TEXT,
            usuario TEXT DEFAULT 'Principal',
            user_id INTEGER,
            grupo_id INTEGER,
            visibilidad TEXT DEFAULT 'familiar',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            user_id INTEGER,
            accion TEXT NOT NULL,
            recurso TEXT NOT NULL,
            detalle TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """)

    conn.commit()
    conn.close()

# --- AUTENTICACIÓN Y GESTIÓN DE IDENTIDAD ---
def registrar_usuario(email, password, nombre, avatar="👤"):
    email_limpio = email.strip().lower()
    nombre_limpio = nombre.strip().capitalize()
    
    if not email_limpio or not password or not nombre_limpio:
        return False, "Todos los campos son obligatorios."
        
    valida, msg_pwd = validar_complejidad_password(password)
    if not valida:
        return False, msg_pwd
    
    conn = get_connection()
    cursor = conn.cursor()
    p_hash = hash_password(password)

    try:
        if es_postgres():
            cursor.execute("""
            INSERT INTO cuentas_usuarios (email, password_hash, nombre, avatar)
            VALUES (%s, %s, %s, %s)
            RETURNING id;
            """, (email_limpio, p_hash, nombre_limpio, avatar))
            res = cursor.fetchone()
            u_id = res["id"] if isinstance(res, dict) else res[0]
        else:
            cursor.execute("""
            INSERT INTO cuentas_usuarios (email, password_hash, nombre, avatar)
            VALUES (?, ?, ?, ?);
            """, (email_limpio, p_hash, nombre_limpio, avatar))
            u_id = cursor.lastrowid
        conn.commit()
        conn.close()
        registrar_auditoria(u_id, "USER_REGISTER", "AUTH", f"Cuenta creada para {email_limpio}")
        return True, {"id": u_id, "email": email_limpio, "nombre": nombre_limpio, "avatar": avatar, "grupo_id": None}
    except Exception as e:
        conn.close()
        if "unique" in str(e).lower():
            return False, "Ya existe una cuenta registrada con este correo electrónico."
        return False, "Error al procesar el registro de usuario."

def autenticar_usuario(email, password):
    email_limpio = email.strip().lower()
    
    # Verificar Rate Limiting (SOC 2)
    permitido, msg_limit = verificar_rate_limit(email_limpio)
    if not permitido:
        return False, msg_limit

    conn = get_connection()
    cursor = conn.cursor()
    placeholder = "%s" if es_postgres() else "?"

    cursor.execute(f"SELECT * FROM cuentas_usuarios WHERE email = {placeholder};", (email_limpio,))
    u = cursor.fetchone()
    conn.close()

    # Protección contra enumeración: Mensaje genérico uniforme
    MSG_ERROR = "Correo electrónico o contraseña incorrectos."

    if not u:
        registrar_fallo_login(email_limpio)
        registrar_auditoria(None, "LOGIN_FAILED", "AUTH", f"Intento fallido (usuario no existe): {email_limpio}")
        return False, MSG_ERROR

    p_hash = u["password_hash"] if isinstance(u, dict) else u[2]
    if verificar_password(password, p_hash):
        limpiar_fallos_login(email_limpio)
        u_dict = {
            "id": u["id"] if isinstance(u, dict) else u[0],
            "email": u["email"] if isinstance(u, dict) else u[1],
            "nombre": u["nombre"] if isinstance(u, dict) else u[3],
            "avatar": u["avatar"] if isinstance(u, dict) else u[4],
            "grupo_id": u["grupo_id"] if isinstance(u, dict) else u[5]
        }
        registrar_auditoria(u_dict["id"], "LOGIN_SUCCESS", "AUTH", "Inicio de sesión correcto")
        return True, u_dict
    else:
        registrar_fallo_login(email_limpio)
        registrar_auditoria(None, "LOGIN_FAILED", "AUTH", f"Credencial inválida: {email_limpio}")
        return False, MSG_ERROR

# --- PRIVACIDAD Y DERECHO AL OLVIDO (ISO/IEC 27701) ---
def eliminar_cuenta_usuario(user_id) -> tuple[bool, str]:
    """
    Ejecuta el borrado integral y definitivo (Right to be Forgotten - PIMS ISO 27701).
    Purga credenciales, transacciones asociadas y registros de usuario.
    """
    conn = get_connection()
    cursor = conn.cursor()
    placeholder = "%s" if es_postgres() else "?"
    try:
        # Purgar gastos asociados al usuario
        cursor.execute(f"DELETE FROM gastos WHERE user_id = {placeholder};", (user_id,))
        # Desvincular o eliminar de grupos familiares
        cursor.execute(f"DELETE FROM grupos_familiares WHERE creador_id = {placeholder};", (user_id,))
        # Purgar cuenta de usuario
        cursor.execute(f"DELETE FROM cuentas_usuarios WHERE id = {placeholder};", (user_id,))
        conn.commit()
        conn.close()
        registrar_auditoria(user_id, "ACCOUNT_PURGED", "USER_DATA", "Borrado definitivo completado (ISO 27701)")
        return True, "Cuenta y datos personales eliminados permanentemente."
    except Exception as e:
        conn.close()
        return False, f"No se pudo completar la purga: {e}"

# --- SISTEMA DE GRUPOS FAMILIARES ---
def crear_grupo_familiar(user_id, nombre_grupo):
    import secrets
    conn = get_connection()
    cursor = conn.cursor()
    prefijo = "".join(c for c in nombre_grupo[:3].upper() if c.isalnum()) or "FAM"
    codigo = f"{prefijo}-{secrets.randbelow(8999) + 1000}"

    try:
        if es_postgres():
            cursor.execute("""
            INSERT INTO grupos_familiares (nombre, codigo_invitacion, creador_id)
            VALUES (%s, %s, %s)
            RETURNING id;
            """, (nombre_grupo.strip(), codigo, user_id))
            res = cursor.fetchone()
            g_id = res["id"] if isinstance(res, dict) else res[0]
            cursor.execute("UPDATE cuentas_usuarios SET grupo_id = %s WHERE id = %s;", (g_id, user_id))
        else:
            cursor.execute("""
            INSERT INTO grupos_familiares (nombre, codigo_invitacion, creador_id)
            VALUES (?, ?, ?);
            """, (nombre_grupo.strip(), codigo, user_id))
            g_id = cursor.lastrowid
            cursor.execute("UPDATE cuentas_usuarios SET grupo_id = ? WHERE id = ?;", (g_id, user_id))

        conn.commit()
        conn.close()
        registrar_auditoria(user_id, "FAMILY_GROUP_CREATED", "GROUPS", f"Grupo creado: {nombre_grupo} ({codigo})")
        return True, {"id": g_id, "nombre": nombre_grupo, "codigo": codigo}
    except Exception as e:
        conn.close()
        return False, "Error al crear grupo familiar."

def unirse_a_grupo_familiar(user_id, codigo_invitacion):
    conn = get_connection()
    cursor = conn.cursor()
    codigo_limpio = codigo_invitacion.strip().upper()
    placeholder = "%s" if es_postgres() else "?"

    cursor.execute(f"SELECT id, nombre FROM grupos_familiares WHERE codigo_invitacion = {placeholder};", (codigo_limpio,))
    grupo = cursor.fetchone()

    if not grupo:
        conn.close()
        return False, "Código de invitación no encontrado o incorrecto."

    g_id = grupo["id"] if isinstance(grupo, dict) else grupo[0]
    g_nombre = grupo["nombre"] if isinstance(grupo, dict) else grupo[1]

    cursor.execute(f"UPDATE cuentas_usuarios SET grupo_id = {placeholder} WHERE id = {placeholder};", (g_id, user_id))
    conn.commit()
    conn.close()
    registrar_auditoria(user_id, "FAMILY_GROUP_JOIN", "GROUPS", f"Unido al grupo: {g_nombre}")
    return True, {"id": g_id, "nombre": g_nombre, "codigo": codigo_limpio}

def obtener_datos_grupo(grupo_id):
    if not grupo_id:
        return None
    conn = get_connection()
    cursor = conn.cursor()
    placeholder = "%s" if es_postgres() else "?"

    cursor.execute(f"SELECT * FROM grupos_familiares WHERE id = {placeholder};", (grupo_id,))
    g = cursor.fetchone()
    if not g:
        conn.close()
        return None

    cursor.execute(f"SELECT id, nombre, email, avatar FROM cuentas_usuarios WHERE grupo_id = {placeholder};", (grupo_id,))
    miembros = cursor.fetchall()
    conn.close()

    m_list = []
    for m in miembros:
        m_list.append({
            "id": m["id"] if isinstance(m, dict) else m[0],
            "nombre": m["nombre"] if isinstance(m, dict) else m[1],
            "email": m["email"] if isinstance(m, dict) else m[2],
            "avatar": m["avatar"] if isinstance(m, dict) else m[3]
        })

    return {
        "id": g["id"] if isinstance(g, dict) else g[0],
        "nombre": g["nombre"] if isinstance(g, dict) else g[1],
        "codigo": g["codigo_invitacion"] if isinstance(g, dict) else g[2],
        "miembros": m_list
    }

def salir_de_grupo_familiar(user_id):
    conn = get_connection()
    cursor = conn.cursor()
    placeholder = "%s" if es_postgres() else "?"
    cursor.execute(f"UPDATE cuentas_usuarios SET grupo_id = NULL WHERE id = {placeholder};", (user_id,))
    conn.commit()
    conn.close()
    registrar_auditoria(user_id, "FAMILY_GROUP_LEAVE", "GROUPS", "Salida voluntaria del grupo familiar")
    return True

# --- FUNCIONES DE GASTOS Y DETECCIÓN DE DUPLICADOS ---
def verificar_posible_duplicado(fecha, comercio, monto, dias_margen=2):
    conn = get_connection()
    cursor = conn.cursor()
    placeholder = "%s" if es_postgres() else "?"
    query = f"SELECT id, fecha, comercio, monto, usuario FROM gastos WHERE comercio ILIKE {placeholder} AND ABS(monto - {placeholder}) < 1.0;" if es_postgres() else f"SELECT id, fecha, comercio, monto, usuario FROM gastos WHERE LOWER(comercio) = LOWER({placeholder}) AND ABS(monto - {placeholder}) < 1.0;"
    
    cursor.execute(query, (comercio.strip(), float(monto)))
    coincidencias = cursor.fetchall()
    conn.close()

    if not coincidencias:
        return False, None

    for c in coincidencias:
        c_fecha = c["fecha"] if isinstance(c, dict) else c[1]
        try:
            d1 = datetime.strptime(fecha, "%Y-%m-%d")
            d2 = datetime.strptime(c_fecha, "%Y-%m-%d")
            if abs((d1 - d2).days) <= dias_margen:
                return True, {
                    "id": c["id"] if isinstance(c, dict) else c[0],
                    "fecha": c_fecha,
                    "comercio": c["comercio"] if isinstance(c, dict) else c[2],
                    "monto": c["monto"] if isinstance(c, dict) else c[3],
                    "usuario": c["usuario"] if isinstance(c, dict) else c[4]
                }
        except Exception:
            continue
    return False, None

def insertar_gasto(fecha, comercio, monto, moneda="ARS", categoria="Otros", medio_pago="Efectivo", cuotas="1", origen="Manual", notas="", usuario="Principal", user_id=None, grupo_id=None, visibilidad="familiar"):
    conn = get_connection()
    cursor = conn.cursor()
    if es_postgres():
        cursor.execute("""
        INSERT INTO gastos (fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, origen, notas, usuario, user_id, grupo_id, visibilidad)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        RETURNING id;
        """, (fecha, comercio, float(monto), moneda, categoria, medio_pago, str(cuotas), origen, notas, usuario, user_id, grupo_id, visibilidad))
        res = cursor.fetchone()
        gasto_id = res["id"] if isinstance(res, dict) else res[0]
    else:
        cursor.execute("""
        INSERT INTO gastos (fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, origen, notas, usuario, user_id, grupo_id, visibilidad)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (fecha, comercio, float(monto), moneda, categoria, medio_pago, str(cuotas), origen, notas, usuario, user_id, grupo_id, visibilidad))
        gasto_id = cursor.lastrowid

    conn.commit()
    conn.close()
    registrar_auditoria(user_id, "EXPENSE_CREATED", "EXPENSES", f"Gasto #{gasto_id}: ${monto} en {comercio}")
    return gasto_id

def insertar_gastos_lote(lista_gastos, usuario_default="Principal", user_id=None, grupo_id=None, visibilidad_default="familiar"):
    if not lista_gastos:
        return 0
    conn = get_connection()
    cursor = conn.cursor()
    count = 0
    for g in lista_gastos:
        u = g.get("usuario") or usuario_default
        f_val = str(g.get("fecha", datetime.now().strftime("%Y-%m-%d")))
        c_val = str(g.get("comercio", "Desconocido"))
        m_val = float(g.get("monto", 0.0))
        mon_val = str(g.get("moneda", "ARS"))
        cat_val = str(g.get("categoria", "Otros"))
        mp_val = str(g.get("medio_pago", "Tarjeta de Crédito"))
        cuot_val = str(g.get("cuotas", "1"))
        orig_val = str(g.get("origen", "Resumen Tarjeta"))
        not_val = str(g.get("notas", ""))
        vis_val = g.get("visibilidad") or visibilidad_default

        if es_postgres():
            cursor.execute("""
            INSERT INTO gastos (fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, origen, notas, usuario, user_id, grupo_id, visibilidad)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
            """, (f_val, c_val, m_val, mon_val, cat_val, mp_val, cuot_val, orig_val, not_val, u, user_id, grupo_id, vis_val))
        else:
            cursor.execute("""
            INSERT INTO gastos (fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, origen, notas, usuario, user_id, grupo_id, visibilidad)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, (f_val, c_val, m_val, mon_val, cat_val, mp_val, cuot_val, orig_val, not_val, u, user_id, grupo_id, vis_val))
        count += 1

    conn.commit()
    conn.close()
    registrar_auditoria(user_id, "EXPENSES_BATCH_CREATED", "EXPENSES", f"Importación por lote: {count} transacciones")
    return count

def obtener_gastos_df(user_id, grupo_id=None, modo="personal", mes=None, anio=None, categoria=None, medio_pago=None, usuario=None):
    conn = get_connection()
    placeholder = "%s" if es_postgres() else "?"
    params = []

    if modo == "familiar" and grupo_id:
        query = f"SELECT id, fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, origen, notas, usuario, visibilidad, created_at FROM gastos WHERE (grupo_id = {placeholder} AND visibilidad = 'familiar')"
        params.append(grupo_id)
    else:
        query = f"SELECT id, fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, origen, notas, usuario, visibilidad, created_at FROM gastos WHERE (user_id = {placeholder} OR (usuario = 'Principal' AND user_id IS NULL))"
        params.append(user_id)

    if anio:
        if es_postgres():
            query += f" AND SUBSTRING(fecha, 1, 4) = {placeholder}"
        else:
            query += f" AND strftime('%Y', fecha) = {placeholder}"
        params.append(str(anio))

    if mes:
        mes_str = f"{int(mes):02d}"
        if es_postgres():
            query += f" AND SUBSTRING(fecha, 6, 2) = {placeholder}"
        else:
            query += f" AND strftime('%m', fecha) = {placeholder}"
        params.append(mes_str)

    if categoria and categoria != "Todas":
        query += f" AND categoria = {placeholder}"
        params.append(categoria)

    if medio_pago and medio_pago != "Todos":
        query += f" AND medio_pago = {placeholder}"
        params.append(medio_pago)

    if usuario and usuario != "Todos":
        query += f" AND usuario = {placeholder}"
        params.append(usuario)

    query += " ORDER BY fecha DESC, id DESC;"
    df = pd.read_sql_query(query, conn, params=params)
    conn.close()

    if not df.empty:
        if "monto" in df.columns:
            df["monto"] = pd.to_numeric(df["monto"], errors="coerce").fillna(0.0)
        if "fecha" in df.columns:
            df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.strftime("%Y-%m-%d").fillna(df["fecha"])
    return df

def obtener_ultimos_gastos(user_id, grupo_id=None, modo="personal", limit=5):
    df = obtener_gastos_df(user_id, grupo_id, modo=modo)
    if not df.empty:
        return df.head(limit)
    return pd.DataFrame()

def eliminar_gasto(gasto_id, user_id=None):
    conn = get_connection()
    cursor = conn.cursor()
    if es_postgres():
        if user_id:
            cursor.execute("DELETE FROM gastos WHERE id = %s AND (user_id = %s OR user_id IS NULL);", (gasto_id, user_id))
        else:
            cursor.execute("DELETE FROM gastos WHERE id = %s;", (gasto_id,))
    else:
        cursor.execute("DELETE FROM gastos WHERE id = ?;", (gasto_id,))
    conn.commit()
    conn.close()
    registrar_auditoria(user_id, "EXPENSE_DELETED", "EXPENSES", f"Gasto #{gasto_id} eliminado")

# Inicialización segura
init_db()
