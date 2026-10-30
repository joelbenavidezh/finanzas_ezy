"""
====================================================================
FINANZAS EZY - APLICACIÓN FINANCIERA INTELIGENTE PARA FAMILIAS (2.0)
- Autenticación individual (Cada familiar tiene su propia cuenta).
- Espacio Personal (Privado) y Espacio Familiar (Compartido).
- Sistema de Grupos Familiares con Código de Invitación único.
- Escaneo de tickets con IA + Detección de gastos duplicados.
- Resúmenes de tarjeta de crédito (PDF/Fotos en lote).
- Dashboard con estadísticas y comparativa de aportes familiares.
- Descarga en Excel (.xlsx) estructurado.
====================================================================
"""

import os
import io
import importlib
import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from dotenv import load_dotenv

import database as db
import ai_extractor as ai

# Asegurar recarga en caliente de módulos
importlib.reload(db)
importlib.reload(ai)

# Cargar variables de entorno
ENV_PATH = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(ENV_PATH, override=True)

# Configuración de página
st.set_page_config(
    page_title="Finanzas Ezy",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS optimizados para móvil y experiencia Fintech
st.markdown("""
<style>
    header[data-testid="stHeader"] { display: none !important; }
    div[data-testid="stDecoration"] { display: none !important; }
    #MainMenu { visibility: hidden !important; }
    footer { visibility: hidden !important; }
    [data-testid="stSidebar"] { display: none !important; }
    
    .block-container {
        max-width: 580px !important;
        padding-top: 1.2rem !important;
        padding-bottom: 3.5rem !important;
        padding-left: 1.2rem !important;
        padding-right: 1.2rem !important;
    }
    
    .app-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 12px;
    }
    .app-title {
        font-size: 1.6rem;
        font-weight: 800;
        background: linear-gradient(135deg, #38bdf8 0%, #818cf8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .hero-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid #334155;
        border-radius: 18px;
        padding: 20px;
        margin-bottom: 18px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3);
    }
    .hero-label {
        font-size: 0.85rem;
        color: #94a3b8;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .hero-amount {
        font-size: 2.1rem;
        font-weight: 800;
        color: #f8fafc;
        margin: 6px 0 12px 0;
    }
    .hero-submetrics {
        display: flex;
        justify-content: space-between;
        border-top: 1px solid #334155;
        padding-top: 12px;
    }
    .submetric-val { font-size: 1rem; font-weight: 700; color: #38bdf8; }
    .submetric-lbl { font-size: 0.75rem; color: #64748b; }

    div.stButton > button {
        border-radius: 14px !important;
        padding: 16px 12px !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        transition: all 0.2s ease !important;
        border: 1px solid #334155 !important;
    }
    div.stButton > button:hover {
        border-color: #38bdf8 !important;
        transform: translateY(-2px);
    }
    
    .btn-back > div.stButton > button {
        background-color: #1e293b !important;
        color: #94a3b8 !important;
        padding: 10px 16px !important;
        font-size: 0.9rem !important;
        margin-bottom: 15px !important;
    }
    
    .tx-item {
        background: #1e293b;
        border-radius: 12px;
        padding: 12px 16px;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        border: 1px solid #2d3748;
    }
    .tx-title { font-weight: 700; font-size: 0.95rem; color: #f1f5f9; }
    .tx-cat { font-size: 0.75rem; color: #94a3b8; }
    .tx-amount { font-weight: 800; font-size: 1rem; color: #34d399; }
</style>
""", unsafe_allow_html=True)

# Variables de Sesión
if "user" not in st.session_state:
    st.session_state["user"] = None

if "vista" not in st.session_state:
    st.session_state["vista"] = "menu"

if "espacio" not in st.session_state:
    st.session_state["espacio"] = "personal"  # 'personal' o 'familiar'

def ir_a(vista_nombre):
    st.session_state["vista"] = vista_nombre
    st.rerun()

def cerrar_sesion():
    st.session_state["user"] = None
    st.session_state["vista"] = "menu"
    st.rerun()

# ==============================================================================
# PANTALLA: LOGIN / REGISTRO (SI NO HAY SESIÓN INICIADA)
# ==============================================================================
if not st.session_state["user"]:
    st.markdown("""
    <div style="text-align: center; margin-top: 1.5rem; margin-bottom: 1.5rem;">
        <div style="font-size: 3rem;">💳</div>
        <div class="app-title" style="font-size: 2rem;">Finanzas Ezy</div>
        <div style="color: #94a3b8; font-size: 0.9rem;">Tu control de gastos personal y familiar con IA</div>
    </div>
    """, unsafe_allow_html=True)

    tab_login, tab_registro = st.tabs(["🔑 Iniciar Sesión", "✨ Crear mi Cuenta"])

    with tab_login:
        with st.form("form_login"):
            st.markdown("### Bienvenido de nuevo")
            l_email = st.text_input("Correo electrónico:", placeholder="tu@email.com").strip().lower()
            l_pass = st.text_input("Contraseña:", type="password")
            
            if st.form_submit_button("Entrar a mis Finanzas", type="primary", use_container_width=True):
                if not l_email or not l_pass:
                    st.error("Por favor completa tu correo y contraseña.")
                else:
                    ok, res = db.autenticar_usuario(l_email, l_pass)
                    if ok:
                        st.session_state["user"] = res
                        st.toast(f"¡Hola de nuevo, {res['nombre']}!", icon="👋")
                        st.rerun()
                    else:
                        st.error(res)

    with tab_registro:
        with st.form("form_registro"):
            st.markdown("### Crea tu cuenta personal")
            st.caption("Cada familiar tiene su propia cuenta segura. Luego podrán unirse al grupo familiar.")
            r_nombre = st.text_input("Tu Nombre:", placeholder="Ej: Joel, María, etc.")
            r_email = st.text_input("Tu Correo:", placeholder="tu@email.com").strip().lower()
            r_pass = st.text_input("Crea una contraseña:", type="password")
            r_avatar = st.selectbox("Elige tu Avatar:", db.AVATARES_DEFAULT)

            if st.form_submit_button("Crear mi Cuenta Gratis", type="primary", use_container_width=True):
                if len(r_pass) < 4:
                    st.error("La contraseña debe tener al menos 4 caracteres.")
                else:
                    ok, res = db.registrar_usuario(r_email, r_pass, r_nombre, r_avatar)
                    if ok:
                        st.session_state["user"] = res
                        st.success("🎉 ¡Cuenta creada con éxito!")
                        st.rerun()
                    else:
                        st.error(res)
    st.stop()

# ==============================================================================
# SESIÓN ACTIVA: DATOS DEL USUARIO Y GRUPO
# ==============================================================================
usuario = st.session_state["user"]
grupo_info = db.obtener_datos_grupo(usuario.get("grupo_id"))

# ==============================================================================
# VISTA 1: MENÚ PRINCIPAL (HOME)
# ==============================================================================
if st.session_state["vista"] == "menu":
    # Header con Perfil y Salir
    col_h1, col_h2 = st.columns([3, 1])
    with col_h1:
        st.markdown(f"""
        <div style="line-height: 1.2;">
            <div class="app-title">Finanzas Ezy</div>
            <div style="font-size: 0.85rem; color: #94a3b8;">Hola, <b>{usuario['avatar']} {usuario['nombre']}</b></div>
        </div>
        """, unsafe_allow_html=True)
    with col_h2:
        if st.button("🚪 Salir", use_container_width=True):
            cerrar_sesion()

    # Selector de Espacio: Personal vs Familiar
    if grupo_info:
        espacio_elegido = st.radio(
            "Espacio:",
            [f"👤 Mis Gastos Personales", f"👨‍👩‍👧‍👦 {grupo_info['nombre']}"],
            horizontal=True,
            label_visibility="collapsed"
        )
        st.session_state["espacio"] = "personal" if "Mis Gastos" in espacio_elegido else "familiar"
    else:
        st.session_state["espacio"] = "personal"
        # Tarjeta informativa para crear o unirse a una familia
        with st.container():
            st.markdown(f"""
            <div style="background: #1e293b; border: 1px dashed #38bdf8; border-radius: 14px; padding: 12px; margin-bottom: 12px; display: flex; align-items: center; justify-content: space-between;">
                <div style="font-size: 0.85rem; color: #cbd5e1;">
                    👨‍👩‍👧‍👦 <b>¿Quieres compartir gastos con tu familia?</b>
                </div>
            </div>
            """, unsafe_allow_html=True)
            if st.button("Conectar con mi Familia (Crear o Unirme)", use_container_width=True):
                ir_a("familia")

    # Métricas del mes actual
    hoy = datetime.now()
    meses_es = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    
    df_mes = db.obtener_gastos_df(
        user_id=usuario["id"],
        grupo_id=usuario.get("grupo_id"),
        modo=st.session_state["espacio"],
        mes=hoy.month,
        anio=hoy.year
    )

    total_mes = df_mes["monto"].sum() if not df_mes.empty else 0.0
    total_tc = df_mes[df_mes["medio_pago"] == "Tarjeta de Crédito"]["monto"].sum() if not df_mes.empty else 0.0
    total_ef = df_mes[df_mes["medio_pago"] != "Tarjeta de Crédito"]["monto"].sum() if not df_mes.empty else 0.0
    cant_compras = len(df_mes) if not df_mes.empty else 0

    etiqueta_hero = f"Gastos de {usuario['nombre']} • {meses_es[hoy.month]}" if st.session_state["espacio"] == "personal" else f"Total {grupo_info['nombre']} • {meses_es[hoy.month]}"

    # Tarjeta Principal (Hero)
    st.markdown(f"""
    <div class="hero-card">
        <div class="hero-label">{etiqueta_hero}</div>
        <div class="hero-amount">${total_mes:,.2f}</div>
        <div class="hero-submetrics">
            <div>
                <div class="submetric-val">${total_tc:,.2f}</div>
                <div class="submetric-lbl">Tarjeta Crédito</div>
            </div>
            <div>
                <div class="submetric-val">${total_ef:,.2f}</div>
                <div class="submetric-lbl">Efectivo / QR</div>
            </div>
            <div>
                <div class="submetric-val">{cant_compras}</div>
                <div class="submetric-lbl">Compras</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Botones táctiles grandes (Touch Cards)
    st.markdown("### ¿Qué quieres hacer?")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📸 Escanear Ticket\n\n(Foto o Cámara)", use_container_width=True, type="primary"):
            ir_a("ticket")
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        if st.button("📊 Estadísticas\n\n(Gráficos & Métricas)", use_container_width=True):
            ir_a("dashboard")
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        if st.button("📑 Movimientos\n\n(Historial & Excel)", use_container_width=True):
            ir_a("historial")

    with c2:
        if st.button("💳 Resumen Tarjeta\n\n(PDF o Fotos en Lote)", use_container_width=True):
            ir_a("tarjeta")
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        if st.button("➕ Gasto Manual\n\n(Anotar sin ticket)", use_container_width=True):
            ir_a("manual")
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        if st.button("👨‍👩‍👧‍👦 Grupo Familiar\n\n(Invitar / Ajustes)", use_container_width=True):
            ir_a("familia")

    # Feed de últimos gastos
    st.markdown("<br>", unsafe_allow_html=True)
    feed_title = f"🕒 Últimos movimientos de {usuario['nombre']}" if st.session_state["espacio"] == "personal" else f"🕒 Últimos movimientos de {grupo_info['nombre']}"
    st.markdown(f"### {feed_title}")
    
    df_ultimos = db.obtener_ultimos_gastos(
        user_id=usuario["id"],
        grupo_id=usuario.get("grupo_id"),
        modo=st.session_state["espacio"],
        limit=4
    )

    if df_ultimos.empty:
        st.caption("Aún no tienes movimientos registrados en este espacio.")
    else:
        for _, fila in df_ultimos.iterrows():
            tag_u = f"[{fila['usuario']}] " if st.session_state["espacio"] == "familiar" else ""
            st.markdown(f"""
            <div class="tx-item">
                <div>
                    <div class="tx-title">{tag_u}{fila['comercio']}</div>
                    <div class="tx-cat">{fila['fecha']} • {fila['categoria']} • {fila['medio_pago']}</div>
                </div>
                <div class="tx-amount">${fila['monto']:,.2f}</div>
            </div>
            """, unsafe_allow_html=True)

# ==============================================================================
# VISTA 2: ESCANEAR TICKET (IA + DETECCIÓN DE DUPLICADOS)
# ==============================================================================
elif st.session_state["vista"] == "ticket":
    st.markdown('<div class="btn-back">', unsafe_allow_html=True)
    if st.button("⬅ Volver al Menú Principal", use_container_width=True):
        ir_a("menu")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("## 📸 Escanear Ticket")
    st.caption(f"Registrando gasto desde la cuenta de: **{usuario['avatar']} {usuario['nombre']}**")

    metodo = st.radio("Modo de captura:", ["📷 Cámara", "📁 Subir Archivo"], horizontal=True)
    img_ticket = st.camera_input("Capturar") if metodo == "📷 Cámara" else st.file_uploader("Seleccionar imagen del ticket:", type=["jpg", "jpeg", "png", "webp"])

    if img_ticket is not None:
        st.image(img_ticket, use_container_width=True)

        if st.button("✨ Analizar Ticket con IA", type="primary", use_container_width=True):
            with st.spinner("Analizando y leyendo ticket con IA..."):
                try:
                    resultado = ai.extraer_ticket(img_ticket.getvalue(), img_ticket.type or "image/jpeg")
                    st.session_state["ticket_temp"] = resultado
                    st.success("✅ ¡Datos detectados!")
                except Exception as e:
                    st.error(f"Error: {e}")

    if "ticket_temp" in st.session_state and st.session_state["ticket_temp"]:
        d = st.session_state["ticket_temp"]
        
        # Alerta Inteligente de Duplicados
        es_duplicado, dup_data = db.verificar_posible_duplicado(
            fecha=d.get("fecha", datetime.now().strftime("%Y-%m-%d")),
            comercio=d.get("comercio", ""),
            monto=float(d.get("monto", 0.0))
        )
        if es_duplicado:
            st.warning(f"⚠️ **Atención:** Parece que este gasto ya fue registrado antes (${dup_data['monto']:,.2f} en {dup_data['comercio']} el {dup_data['fecha']} por {dup_data['usuario']}). Si es una compra distinta, puedes continuar.")

        with st.form("form_ticket_mobile"):
            st.markdown("### Datos extraídos:")
            
            try:
                f_val = datetime.strptime(d.get("fecha", ""), "%Y-%m-%d").date()
            except:
                f_val = datetime.now().date()

            n_fecha = st.date_input("Fecha:", value=f_val)
            n_comercio = st.text_input("Comercio:", value=d.get("comercio", ""))
            n_monto = st.number_input("Total ($):", value=float(d.get("monto", 0.0)), step=10.0, format="%.2f")
            
            cat_sug = d.get("categoria", "Otros")
            cat_idx = db.CATEGORIAS_DEFAULT.index(cat_sug) if cat_sug in db.CATEGORIAS_DEFAULT else len(db.CATEGORIAS_DEFAULT)-1
            n_cat = st.selectbox("Categoría:", db.CATEGORIAS_DEFAULT, index=cat_idx)
            
            mp_sug = d.get("medio_pago", "Efectivo")
            mp_idx = db.MEDIOS_PAGO_DEFAULT.index(mp_sug) if mp_sug in db.MEDIOS_PAGO_DEFAULT else 0
            n_mp = st.selectbox("Medio de Pago:", db.MEDIOS_PAGO_DEFAULT, index=mp_idx)
            
            # Selector de Visibilidad (Privado vs Familiar)
            opciones_vis = ["👨‍👩‍👧‍👦 Compartir con la Familia", "🔒 Gasto Privado (Solo para mí)"]
            sel_vis = st.selectbox("Visibilidad del Gasto:", opciones_vis, index=0 if grupo_info else 1)
            vis_val = "familiar" if "Compartir" in sel_vis else "personal"

            n_cuotas = st.text_input("Cuotas:", value=str(d.get("cuotas", "1")))
            n_notas = st.text_input("Notas / Detalle:", value=d.get("notas", ""))

            if st.form_submit_button("💾 Confirmar y Guardar Gasto", type="primary", use_container_width=True):
                db.insertar_gasto(
                    fecha=n_fecha.strftime("%Y-%m-%d"),
                    comercio=n_comercio,
                    monto=n_monto,
                    moneda="ARS",
                    categoria=n_cat,
                    medio_pago=n_mp,
                    cuotas=n_cuotas,
                    origen="Ticket",
                    notas=n_notas,
                    usuario=usuario["nombre"],
                    user_id=usuario["id"],
                    grupo_id=usuario.get("grupo_id"),
                    visibilidad=vis_val
                )
                st.session_state["ticket_temp"] = None
                st.toast("🎉 ¡Gasto guardado con éxito!", icon="✅")
                ir_a("menu")

# ==============================================================================
# VISTA 3: RESUMEN DE TARJETA EN LOTE
# ==============================================================================
elif st.session_state["vista"] == "tarjeta":
    st.markdown('<div class="btn-back">', unsafe_allow_html=True)
    if st.button("⬅ Volver al Menú Principal", use_container_width=True):
        ir_a("menu")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("## 💳 Resumen de Tarjeta (Lote)")
    st.caption("Sube tu resumen bancario (PDF o foto). La IA extraerá todos los consumos:")

    archivo_tc = st.file_uploader("Subir PDF o imagen del extracto:", type=["pdf", "jpg", "jpeg", "png"])

    if archivo_tc is not None:
        if st.button("✨ Procesar Resumen Completo con IA", type="primary", use_container_width=True):
            with st.spinner("Extrayendo consumos en lote con IA..."):
                try:
                    mime = archivo_tc.type or ("application/pdf" if archivo_tc.name.endswith(".pdf") else "image/jpeg")
                    consumos = ai.extraer_resumen_tarjeta(archivo_tc.getvalue(), mime)
                    if consumos:
                        st.session_state["consumos_lote"] = consumos
                        st.success(f"✅ ¡Se detectaron {len(consumos)} consumos!")
                    else:
                        st.warning("No se encontraron consumos claros en el documento.")
                except Exception as e:
                    st.error(f"Error procesando resumen: {e}")

    if "consumos_lote" in st.session_state and st.session_state["consumos_lote"]:
        df_tc = pd.DataFrame(st.session_state["consumos_lote"])
        df_tc["usuario"] = usuario["nombre"]
        st.markdown(f"**Total detectado:** ${df_tc['monto'].sum():,.2f}")
        
        df_edit = st.data_editor(
            df_tc,
            use_container_width=True,
            num_rows="dynamic"
        )
        
        vis_tc = st.radio("¿Cómo guardar estos consumos?", ["👨‍👩‍👧‍👦 Compartir con la Familia", "🔒 Guardar como Personales"], horizontal=True)
        vis_final = "familiar" if "Compartir" in vis_tc else "personal"

        if st.button("💾 Guardar TODOS los consumos", type="primary", use_container_width=True):
            total_guardados = db.insertar_gastos_lote(
                lista_gastos=df_edit.to_dict(orient="records"),
                usuario_default=usuario["nombre"],
                user_id=usuario["id"],
                grupo_id=usuario.get("grupo_id"),
                visibilidad_default=vis_final
            )
            st.session_state["consumos_lote"] = None
            st.toast(f"¡{total_guardados} consumos registrados!", icon="✅")
            ir_a("menu")

# ==============================================================================
# VISTA 4: ESTADÍSTICAS & DASHBOARD
# ==============================================================================
elif st.session_state["vista"] == "dashboard":
    st.markdown('<div class="btn-back">', unsafe_allow_html=True)
    if st.button("⬅ Volver al Menú Principal", use_container_width=True):
        ir_a("menu")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("## 📊 Estadísticas y Análisis")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        sel_anio = st.selectbox("Año:", [datetime.now().year, datetime.now().year - 1])
    with col_f2:
        meses_n = ["Todos", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                   "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        sel_mes = st.selectbox("Mes:", meses_n, index=datetime.now().month)
        num_m = None if sel_mes == "Todos" else meses_n.index(sel_mes)

    df_d = db.obtener_gastos_df(
        user_id=usuario["id"],
        grupo_id=usuario.get("grupo_id"),
        modo=st.session_state["espacio"],
        mes=num_m,
        anio=sel_anio
    )

    if df_d.empty:
        st.info("No hay gastos registrados en este período y espacio.")
    else:
        st.metric("Total Gastado:", f"${df_d['monto'].sum():,.2f}")

        # Comparativa por Miembro si es espacio familiar
        if st.session_state["espacio"] == "familiar" and "usuario" in df_d.columns:
            st.subheader("👥 Aportes por Integrante del Hogar")
            df_u_monto = df_d.groupby("usuario")["monto"].sum().reset_index().sort_values(by="monto", ascending=False)
            fig_fam = px.bar(
                df_u_monto,
                x="usuario",
                y="monto",
                color="usuario",
                text_auto=".2s",
                labels={"usuario": "Familiar", "monto": "Total Gastado ($)"}
            )
            fig_fam.update_layout(margin=dict(t=5, b=5, l=5, r=5), showlegend=False)
            st.plotly_chart(fig_fam, use_container_width=True)

        # Gráfico Donut de Rubros
        st.subheader("Gastos por Categoría")
        df_cat = df_d.groupby("categoria")["monto"].sum().reset_index().sort_values(by="monto", ascending=False)
        fig_donut = px.pie(df_cat, values="monto", names="categoria", hole=0.5, color_discrete_sequence=px.colors.qualitative.Pastel)
        fig_donut.update_layout(margin=dict(t=5, b=5, l=5, r=5), showlegend=True)
        st.plotly_chart(fig_donut, use_container_width=True)

        # Gráfico Barras Medios de Pago
        st.subheader("Medios de Pago")
        df_mp = df_d.groupby("medio_pago")["monto"].sum().reset_index()
        fig_bar = px.bar(df_mp, x="medio_pago", y="monto", color="medio_pago", text_auto=".2s")
        fig_bar.update_layout(margin=dict(t=5, b=5, l=5, r=5), showlegend=False, xaxis_title="", yaxis_title="Monto ($)")
        st.plotly_chart(fig_bar, use_container_width=True)

# ==============================================================================
# VISTA 5: GASTO MANUAL
# ==============================================================================
elif st.session_state["vista"] == "manual":
    st.markdown('<div class="btn-back">', unsafe_allow_html=True)
    if st.button("⬅ Volver al Menú Principal", use_container_width=True):
        ir_a("menu")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("## ➕ Cargar Gasto Manual")
    with st.form("form_manual_simple"):
        f_fecha = st.date_input("Fecha:", value=datetime.now().date())
        f_comercio = st.text_input("Comercio / Detalle:")
        f_monto = st.number_input("Monto ($):", min_value=0.01, step=50.0, format="%.2f")
        f_cat = st.selectbox("Categoría:", db.CATEGORIAS_DEFAULT)
        f_mp = st.selectbox("Medio de Pago:", db.MEDIOS_PAGO_DEFAULT)
        
        opciones_vis_m = ["👨‍👩‍👧‍👦 Compartir con la Familia", "🔒 Gasto Privado (Solo para mí)"]
        sel_vis_m = st.selectbox("Visibilidad:", opciones_vis_m, index=0 if grupo_info else 1)
        vis_val_m = "familiar" if "Compartir" in sel_vis_m else "personal"

        f_cuotas = st.text_input("Cuotas:", value="1")
        f_notas = st.text_input("Notas adicionales:")

        if st.form_submit_button("💾 Guardar Gasto", type="primary", use_container_width=True):
            if not f_comercio:
                st.error("Ingresa el nombre del comercio.")
            else:
                db.insertar_gasto(
                    fecha=f_fecha.strftime("%Y-%m-%d"),
                    comercio=f_comercio,
                    monto=f_monto,
                    moneda="ARS",
                    categoria=f_cat,
                    medio_pago=f_mp,
                    cuotas=f_cuotas,
                    origen="Manual",
                    notas=f_notas,
                    usuario=usuario["nombre"],
                    user_id=usuario["id"],
                    grupo_id=usuario.get("grupo_id"),
                    visibilidad=vis_val_m
                )
                st.toast("¡Gasto guardado con éxito!", icon="✅")
                ir_a("menu")

# ==============================================================================
# VISTA 6: GESTIÓN DE GRUPOS FAMILIARES (INVITAR Y UNIRSE)
# ==============================================================================
elif st.session_state["vista"] == "familia":
    st.markdown('<div class="btn-back">', unsafe_allow_html=True)
    if st.button("⬅ Volver al Menú Principal", use_container_width=True):
        ir_a("menu")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("## 👨‍👩‍👧‍👦 Grupo Familiar")

    if grupo_info:
        st.success(f"Estás en el grupo: **{grupo_info['nombre']}**")
        
        # Tarjeta del código de invitación
        st.markdown(f"""
        <div style="background: #1e293b; border: 2px dashed #38bdf8; border-radius: 14px; padding: 18px; text-align: center; margin: 15px 0;">
            <div style="font-size: 0.85rem; color: #94a3b8; text-transform: uppercase;">Código para invitar a tus familiares:</div>
            <div style="font-size: 2rem; font-weight: 800; color: #38bdf8; letter-spacing: 2px; margin: 8px 0;">{grupo_info['codigo']}</div>
            <div style="font-size: 0.8rem; color: #cbd5e1;">Pásales este código para que se unan desde sus teléfonos.</div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown("### Integrantes unidos:")
        for m in grupo_info["miembros"]:
            es_yo = (m["id"] == usuario["id"])
            st.markdown(f"""
            <div class="tx-item">
                <div style="display: flex; align-items: center; gap: 8px;">
                    <span style="font-size: 1.4rem;">{m['avatar']}</span>
                    <div>
                        <div class="tx-title">{m['nombre']} {'(Tú)' if es_yo else ''}</div>
                        <div class="tx-cat">{m['email']}</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚪 Salir de este Grupo Familiar", type="secondary"):
            db.salir_de_grupo_familiar(usuario["id"])
            usuario["grupo_id"] = None
            st.session_state["user"]["grupo_id"] = None
            st.toast("Has salido del grupo.", icon="ℹ️")
            st.rerun()

    else:
        st.caption("Aún no estás conectado a ningún grupo familiar. Puedes crear uno nuevo o unirte a uno existente:")

        tab_crear_g, tab_unir_g = st.tabs(["➕ Crear Grupo Familiar", "🔗 Unirme con un Código"])

        with tab_crear_g:
            with st.form("form_crear_grupo"):
                st.markdown("### Crear un nuevo Hogar")
                nom_g = st.text_input("Nombre de la familia o grupo:", placeholder="Ej: Familia Benavidez")
                if st.form_submit_button("Crear Grupo y Obtener Código", type="primary", use_container_width=True):
                    if nom_g.strip():
                        ok, g_nuevo = db.crear_grupo_familiar(usuario["id"], nom_g)
                        if ok:
                            st.session_state["user"]["grupo_id"] = g_nuevo["id"]
                            st.success(f"🎉 ¡Grupo creado! Código de invitación: {g_nuevo['codigo']}")
                            st.rerun()
                        else:
                            st.error(g_nuevo)

        with tab_unir_g:
            with st.form("form_unir_grupo"):
                st.markdown("### Unirme con Invitación")
                cod_ingresado = st.text_input("Pega el código de invitación recibido:", placeholder="Ej: BEN-4829").strip().upper()
                if st.form_submit_button("Unirme al Grupo Familiar", type="primary", use_container_width=True):
                    if cod_ingresado:
                        ok, res = db.unirse_a_grupo_familiar(usuario["id"], cod_ingresado)
                        if ok:
                            st.session_state["user"]["grupo_id"] = res["id"]
                            st.success(f"🎉 ¡Te has unido con éxito a {res['nombre']}!")
                            st.rerun()
                        else:
                            st.error(res)

# ==============================================================================
# VISTA 7: HISTORIAL & DESCARGA EN EXCEL
# ==============================================================================
elif st.session_state["vista"] == "historial":
    st.markdown('<div class="btn-back">', unsafe_allow_html=True)
    if st.button("⬅ Volver al Menú Principal", use_container_width=True):
        ir_a("menu")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("## 📑 Movimientos & Excel")
    
    df_todos = db.obtener_gastos_df(
        user_id=usuario["id"],
        grupo_id=usuario.get("grupo_id"),
        modo=st.session_state["espacio"]
    )

    if not df_todos.empty:
        # Generar Excel en memoria para descarga instantánea
        buffer_excel = io.BytesIO()
        with pd.ExcelWriter(buffer_excel, engine="openpyxl") as writer:
            nombre_hoja = "Mis Gastos" if st.session_state["espacio"] == "personal" else str(grupo_info['nombre'])[:30]
            df_todos.to_excel(writer, index=False, sheet_name=nombre_hoja)
        buffer_excel.seek(0)

        st.download_button(
            label="📥 Descargar archivo Excel (.xlsx)",
            data=buffer_excel,
            file_name=f"finanzas_ezy_{st.session_state['espacio']}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
        st.markdown("<br>", unsafe_allow_html=True)

        st.dataframe(
            df_todos[["id", "fecha", "comercio", "monto", "usuario", "categoria", "medio_pago", "visibilidad"]],
            use_container_width=True,
            hide_index=True
        )

        with st.expander("🗑️ Eliminar un registro"):
            id_borrar = st.selectbox("Seleccionar ID del gasto a eliminar:", df_todos["id"].tolist())
            if st.button("Confirmar eliminación"):
                db.eliminar_gasto(id_borrar, user_id=usuario["id"])
                st.toast(f"Registro #{id_borrar} eliminado.", icon="🗑️")
                st.rerun()
    else:
        st.info("No hay registros en este espacio.")
