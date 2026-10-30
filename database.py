"""
Módulo de Base de Datos para Finanzas Personales y Familiares (Finanzas Ezy)
- Autenticación individual por usuario (Email, Contraseña con hash SHA-256 + Salt).
- Sistema de Grupos Familiares (Crear grupo con código de invitación único y unirse).
- Visibilidad por gasto: 'Personal' (solo el dueño lo ve) o 'Familiar' (compartido con el grupo).
- Detección de duplicados.
- Presupuestos mensuales por categoría.
- Sincronización y exportación a Microsoft Excel.
"""

import os
import sqlite3
import hashlib
import secrets
import pandas as pd
from datetime import datetime
from dotenv import load_dotenv

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

def hash_password(password, salt=None):
    """Genera hash seguro para contraseñas usando SHA-256 con salt."""
    if not salt:
        salt = secrets.token_hex(16)
    hash_obj = hashlib.sha256((password + salt).encode("utf-8"))
    return f"{salt}${hash_obj.hexdigest()}"

def verificar_password(password_plana, hash_almacenado):
    if not hash_almacenado or "$" not in hash_almacenado:
        return False
    salt, hash_real = hash_almacenado.split("$", 1)
    return hash_password(password_plana, salt) == hash_almacenado

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    if es_postgres():
        # Tabla de Grupos Familiares
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS grupos_familiares (
            id SERIAL PRIMARY KEY,
            nombre VARCHAR(100) NOT NULL,
            codigo_invitacion VARCHAR(20) UNIQUE NOT NULL,
            creador_id INTEGER,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)
        
        # Tabla de Usuarios con login propio
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

        # Tabla de Gastos con enlace a usuario y visibilidad
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

        # Migración segura de columnas en gastos si la tabla ya existía
        cursor.execute("""
        DO $$
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='gastos' AND column_name='user_id') THEN
                ALTER TABLE gastos ADD COLUMN user_id INTEGER;
            END IF;
            IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='gastos' AND column_name='grupo_id') THEN
                ALTER TABLE gastos ADD COLUMN grupo_id INTEGER;
            END IF;
            IF NOT EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name='gastos' AND column_name='visibilidad') THEN
                ALTER TABLE gastos ADD COLUMN visibilidad VARCHAR(20) DEFAULT 'familiar';
            END IF;
        END $$;
        """)

        # Tabla de Presupuestos mensuales
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

    else:
        # Modo SQLite Local
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

    conn.commit()
    conn.close()

# --- AUTENTICACIÓN Y CUENTAS ---
def registrar_usuario(email, password, nombre, avatar="👤"):
    """Registra una nueva cuenta de usuario individual."""
    email_limpio = email.strip().lower()
    nombre_limpio = nombre.strip().capitalize()
    if not email_limpio or not password or not nombre_limpio:
        return False, "Completa todos los campos obligatorios."
    
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
        return True, {"id": u_id, "email": email_limpio, "nombre": nombre_limpio, "avatar": avatar, "grupo_id": None}
    except Exception as e:
        conn.close()
        if "unique" in str(e).lower():
            return False, "Ya existe una cuenta registrada con este correo electrónico."
        return False, f"Error al registrar: {e}"

def autenticar_usuario(email, password):
    """Inicia sesión verificando credenciales."""
    email_limpio = email.strip().lower()
    conn = get_connection()
    cursor = conn.cursor()
    placeholder = "%s" if es_postgres() else "?"

    cursor.execute(f"SELECT * FROM cuentas_usuarios WHERE email = {placeholder};", (email_limpio,))
    u = cursor.fetchone()
    conn.close()

    if not u:
        return False, "Correo o contraseña incorrectos."

    p_hash = u["password_hash"] if isinstance(u, dict) else u[2]
    if verificar_password(password, p_hash):
        u_dict = {
            "id": u["id"] if isinstance(u, dict) else u[0],
            "email": u["email"] if isinstance(u, dict) else u[1],
            "nombre": u["nombre"] if isinstance(u, dict) else u[3],
            "avatar": u["avatar"] if isinstance(u, dict) else u[4],
            "grupo_id": u["grupo_id"] if isinstance(u, dict) else u[5]
        }
        return True, u_dict
    else:
        return False, "Correo o contraseña incorrectos."

# --- SISTEMA DE GRUPOS FAMILIARES ---
def crear_grupo_familiar(user_id, nombre_grupo):
    """Crea un nuevo grupo familiar y genera un código de invitación único."""
    conn = get_connection()
    cursor = conn.cursor()
    # Generar código tipo BEN-4829
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
        return True, {"id": g_id, "nombre": nombre_grupo, "codigo": codigo}
    except Exception as e:
        conn.close()
        return False, str(e)

def unirse_a_grupo_familiar(user_id, codigo_invitacion):
    """Permite a un usuario unirse a una familia existente con su código."""
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
    return True, {"id": g_id, "nombre": g_nombre, "codigo": codigo_limpio}

def obtener_datos_grupo(grupo_id):
    """Obtiene información del grupo y la lista de miembros unidos."""
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
    """Permite a un usuario desvincularse de un grupo familiar."""
    conn = get_connection()
    cursor = conn.cursor()
    placeholder = "%s" if es_postgres() else "?"
    cursor.execute(f"UPDATE cuentas_usuarios SET grupo_id = NULL WHERE id = {placeholder};", (user_id,))
    conn.commit()
    conn.close()
    return True

# --- FUNCIONES DE GASTOS Y DETECCIÓN DE DUPLICADOS ---
def verificar_posible_duplicado(fecha, comercio, monto, dias_margen=2):
    """
    Detecta si ya existe un gasto con mismo monto y comercio en fechas muy cercanas.
    Retorna True y los datos del gasto duplicado si lo encuentra.
    """
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
    return count

def obtener_gastos_df(user_id, grupo_id=None, modo="personal", mes=None, anio=None, categoria=None, medio_pago=None, usuario=None):
    """
    Obtiene los gastos según el contexto:
    - modo 'personal': Solo los gastos del user_id (tanto privados como familiares de él).
    - modo 'familiar': Si está en un grupo, muestra todos los gastos marcados como 'familiar' de los miembros del grupo.
    """
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
    """Devuelve los últimos movimientos para el feed principal."""
    df = obtener_gastos_df(user_id, grupo_id, modo=modo)
    if not df.empty:
        return df.head(limit)
    return pd.DataFrame()

def eliminar_gasto(gasto_id, user_id=None):
    """Elimina un gasto asegurando permisos."""
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

# Inicializar esquema
init_db()
