"""
Módulo de Base de Datos para Finanzas Personales Familiares
Soporta PostgreSQL en la nube (Supabase) y SQLite local con conmutación automática.
Maneja persistencia, perfiles de usuarios y sincronización/exportación a Microsoft Excel.
"""

import os
import sqlite3
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

def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    
    if es_postgres():
        # Esquema PostgreSQL (Supabase)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id SERIAL PRIMARY KEY,
            nombre VARCHAR(100) UNIQUE NOT NULL,
            avatar VARCHAR(10) DEFAULT '👤',
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)
        
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
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        """)
        
        cursor.execute("SELECT COUNT(*) FROM usuarios;")
        res = cursor.fetchone()
        count = res["count"] if isinstance(res, dict) else res[0]
        if count == 0:
            cursor.execute("INSERT INTO usuarios (nombre, avatar) VALUES (%s, %s);", ("Principal", "👤"))
    else:
        # Esquema SQLite local
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            avatar TEXT DEFAULT '👤',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """)
        cursor.execute("SELECT COUNT(*) FROM usuarios;")
        if cursor.fetchone()[0] == 0:
            cursor.execute("INSERT INTO usuarios (nombre, avatar) VALUES (?, ?);", ("Principal", "👤"))
            
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
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        );
        """)
        
        cursor.execute("PRAGMA table_info(gastos);")
        columnas = [col[1] for col in cursor.fetchall()]
        if "usuario" not in columnas:
            cursor.execute("ALTER TABLE gastos ADD COLUMN usuario TEXT DEFAULT 'Principal';")

    conn.commit()
    conn.close()

# --- FUNCIONES DE USUARIOS ---
def obtener_usuarios():
    """Devuelve la lista de miembros de la familia como diccionarios."""
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre, avatar FROM usuarios ORDER BY id ASC;")
    filas = cursor.fetchall()
    conn.close()
    return [{"id": r["id"], "nombre": r["nombre"], "avatar": r["avatar"]} for r in filas]

def crear_usuario(nombre, avatar="👤"):
    """Crea un nuevo perfil de usuario."""
    nombre_limpio = nombre.strip().capitalize()
    if not nombre_limpio:
        return False
    try:
        conn = get_connection()
        cursor = conn.cursor()
        if es_postgres():
            cursor.execute("INSERT INTO usuarios (nombre, avatar) VALUES (%s, %s);", (nombre_limpio, avatar))
        else:
            cursor.execute("INSERT INTO usuarios (nombre, avatar) VALUES (?, ?);", (nombre_limpio, avatar))
        conn.commit()
        conn.close()
        return True
    except Exception:
        return False

def eliminar_usuario(nombre):
    """Elimina un perfil (sin borrar sus gastos, se reasignan a 'Principal')."""
    if nombre == "Principal":
        return False
    conn = get_connection()
    cursor = conn.cursor()
    if es_postgres():
        cursor.execute("UPDATE gastos SET usuario = 'Principal' WHERE usuario = %s;", (nombre,))
        cursor.execute("DELETE FROM usuarios WHERE nombre = %s;", (nombre,))
    else:
        cursor.execute("UPDATE gastos SET usuario = 'Principal' WHERE usuario = ?;", (nombre,))
        cursor.execute("DELETE FROM usuarios WHERE nombre = ?;", (nombre,))
    conn.commit()
    conn.close()
    sincronizar_con_excel()
    return True

# --- FUNCIONES DE GASTOS ---
def insertar_gasto(fecha, comercio, monto, moneda="ARS", categoria="Otros", medio_pago="Efectivo", cuotas="1", origen="Manual", notas="", usuario="Principal"):
    conn = get_connection()
    cursor = conn.cursor()
    if es_postgres():
        cursor.execute("""
        INSERT INTO gastos (fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, origen, notas, usuario)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
        RETURNING id;
        """, (fecha, comercio, float(monto), moneda, categoria, medio_pago, str(cuotas), origen, notas, usuario))
        res = cursor.fetchone()
        gasto_id = res["id"] if isinstance(res, dict) else res[0]
    else:
        cursor.execute("""
        INSERT INTO gastos (fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, origen, notas, usuario)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (fecha, comercio, float(monto), moneda, categoria, medio_pago, str(cuotas), origen, notas, usuario))
        gasto_id = cursor.lastrowid

    conn.commit()
    conn.close()
    sincronizar_con_excel()
    return gasto_id

def insertar_gastos_lote(lista_gastos, usuario_default="Principal"):
    """Inserta múltiples gastos asignados al usuario seleccionado."""
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

        if es_postgres():
            cursor.execute("""
            INSERT INTO gastos (fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, origen, notas, usuario)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s);
            """, (f_val, c_val, m_val, mon_val, cat_val, mp_val, cuot_val, orig_val, not_val, u))
        else:
            cursor.execute("""
            INSERT INTO gastos (fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, origen, notas, usuario)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, (f_val, c_val, m_val, mon_val, cat_val, mp_val, cuot_val, orig_val, not_val, u))
        count += 1

    conn.commit()
    conn.close()
    sincronizar_con_excel()
    return count

def obtener_gastos_df(mes=None, anio=None, categoria=None, medio_pago=None, usuario=None):
    """Obtiene los gastos con filtros, incluyendo filtro por miembro o 'Todos'."""
    conn = get_connection()
    placeholder = "%s" if es_postgres() else "?"
    query = "SELECT id, fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, origen, notas, usuario, created_at FROM gastos WHERE 1=1"
    params = []

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

def obtener_ultimos_gastos(limit=5, usuario=None):
    """Devuelve los últimos gastos filtrados por usuario o familiares."""
    conn = get_connection()
    placeholder = "%s" if es_postgres() else "?"
    if usuario and usuario != "Todos":
        query = f"SELECT id, fecha, comercio, monto, moneda, categoria, medio_pago, usuario FROM gastos WHERE usuario = {placeholder} ORDER BY fecha DESC, id DESC LIMIT {placeholder};"
        df = pd.read_sql_query(query, conn, params=[usuario, limit])
    else:
        query = f"SELECT id, fecha, comercio, monto, moneda, categoria, medio_pago, usuario FROM gastos ORDER BY fecha DESC, id DESC LIMIT {placeholder};"
        df = pd.read_sql_query(query, conn, params=[limit])
    conn.close()
    if not df.empty and "monto" in df.columns:
        df["monto"] = pd.to_numeric(df["monto"], errors="coerce").fillna(0.0)
    return df

def eliminar_gasto(gasto_id):
    conn = get_connection()
    cursor = conn.cursor()
    if es_postgres():
        cursor.execute("DELETE FROM gastos WHERE id = %s;", (gasto_id,))
    else:
        cursor.execute("DELETE FROM gastos WHERE id = ?;", (gasto_id,))
    conn.commit()
    conn.close()
    sincronizar_con_excel()

def sincronizar_con_excel():
    """Genera el Excel consolidado y con pestañas individuales por miembro."""
    try:
        conn = get_connection()
        df_todos = pd.read_sql_query(
            "SELECT id, fecha, comercio, monto, moneda, categoria, medio_pago, cuotas, usuario, origen, notas FROM gastos ORDER BY fecha DESC;",
            conn
        )
        conn.close()

        if not df_todos.empty and "monto" in df_todos.columns:
            df_todos["monto"] = pd.to_numeric(df_todos["monto"], errors="coerce").fillna(0.0)

        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            df_todos.to_excel(writer, index=False, sheet_name="Consolidado Familiar")
            if not df_todos.empty and "usuario" in df_todos.columns:
                usuarios_con_gastos = df_todos["usuario"].dropna().unique()
                for u in usuarios_con_gastos:
                    df_u = df_todos[df_todos["usuario"] == u]
                    nombre_hoja = str(u)[:30]
                    df_u.to_excel(writer, index=False, sheet_name=nombre_hoja)
    except Exception as e:
        print(f"Advertencia al sincronizar con Excel: {e}")

def importar_desde_dataframe(df, col_fecha, col_comercio, col_monto, col_categoria=None, col_medio_pago=None, col_notas=None, usuario="Principal"):
    """Importa gastos históricos asignándolos al usuario correspondiente."""
    gastos_a_insertar = []
    hoy = datetime.now().strftime("%Y-%m-%d")
    
    for _, row in df.iterrows():
        fecha_raw = row.get(col_fecha, hoy)
        try:
            fecha_val = pd.to_datetime(fecha_raw).strftime("%Y-%m-%d")
        except Exception:
            fecha_val = hoy
            
        comercio_val = str(row.get(col_comercio, "Gasto Histórico")).strip()
        if not comercio_val or comercio_val.lower() == "nan":
            comercio_val = "Sin Nombre"
            
        monto_raw = row.get(col_monto, 0.0)
        try:
            if isinstance(monto_raw, str):
                monto_clean = monto_raw.replace("$", "").replace(" ", "").replace(".", "").replace(",", ".")
                monto_val = abs(float(monto_clean))
            else:
                monto_val = abs(float(monto_raw))
        except Exception:
            monto_val = 0.0
            
        if monto_val <= 0:
            continue
            
        cat_val = str(row.get(col_categoria, "Otros")).strip() if col_categoria else "Otros"
        if cat_val not in CATEGORIAS_DEFAULT:
            cat_val = "Otros"
            
        mp_val = str(row.get(col_medio_pago, "Efectivo")).strip() if col_medio_pago else "Efectivo"
        if mp_val not in MEDIOS_PAGO_DEFAULT:
            mp_val = "Efectivo"
            
        notas_val = str(row.get(col_notas, "Importado de Excel viejo")).strip() if col_notas else "Importado de Excel viejo"
        if notas_val.lower() == "nan":
            notas_val = ""

        gastos_a_insertar.append({
            "fecha": fecha_val,
            "comercio": comercio_val,
            "monto": monto_val,
            "moneda": "ARS",
            "categoria": cat_val,
            "medio_pago": mp_val,
            "cuotas": "1",
            "origen": "Excel Histórico",
            "notas": notas_val,
            "usuario": usuario
        })
        
    return insertar_gastos_lote(gastos_a_insertar, usuario_default=usuario)

init_db()
