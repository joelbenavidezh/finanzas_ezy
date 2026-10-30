"""
====================================================================
APLICACIÓN DE FINANZAS PERSONALES FAMILIARES - INTERFAZ MOBILE-FIRST
- Gestión de usuarios y miembros de la familia.
- Gastos independientes por integrante y modo 'Consolidado Familiar'.
- Escaneo de tickets (foto/cámara) con IA y optimización de imagen.
- Resúmenes de tarjeta de crédito (PDF/Imagen en lote).
- Dashboard con comparativa entre miembros de la familia.
- Sincronización automática en Excel (hojas separadas por miembro).
====================================================================
"""

import os
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
    page_title="Finanzas Familiares",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilo CSS optimizado para móvil
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
        margin-bottom: 15px;
    }
    .app-title {
        font-size: 1.6rem;
        font-weight: 800;
        background: linear-gradient(135deg, #38bdf8 0%, #818cf8 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .user-pill {
        background: #1e293b;
        border: 1px solid #38bdf8;
        padding: 6px 14px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 700;
        color: #f8fafc;
        display: flex;
        align-items: center;
        gap: 6px;
    }
    
    .hero-card {
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        border: 1px solid #334155;
        border-radius: 18px;
        padding: 20px;
        margin-bottom: 20px;
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

# Manejo de usuarios
lista_usuarios = db.obtener_usuarios()
nombres_usuarios = [u["nombre"] for u in lista_usuarios]

if "usuario_actual" not in st.session_state or st.session_state["usuario_actual"] not in nombres_usuarios:
    st.session_state["usuario_actual"] = nombres_usuarios[0] if nombres_usuarios else "Principal"

if "vista" not in st.session_state:
    st.session_state["vista"] = "menu"

if "modo_vista" not in st.session_state:
    st.session_state["modo_vista"] = "personal"  # 'personal' o 'familiar'

def ir_a(vista_nombre):
    st.session_state["vista"] = vista_nombre
    st.rerun()

# ==============================================================================
# VISTA 1: MENÚ PRINCIPAL
# ==============================================================================
if st.session_state["vista"] == "menu":
    # Header con selector de miembro de la familia
    col_h1, col_h2 = st.columns([3, 2])
    with col_h1:
        st.markdown("""
        <div style="line-height: 1.2;">
            <div class="app-title">💳 Finanzas</div>
            <div style="font-size: 0.8rem; color: #64748b;">Control familiar con IA</div>
        </div>
        """, unsafe_allow_html=True)
    with col_h2:
        idx_u = nombres_usuarios.index(st.session_state["usuario_actual"]) if st.session_state["usuario_actual"] in nombres_usuarios else 0
        u_sel = st.selectbox("Perfil:", nombres_usuarios, index=idx_u, label_visibility="collapsed")
        if u_sel != st.session_state["usuario_actual"]:
            st.session_state["usuario_actual"] = u_sel
            st.rerun()

    # Selector de Modo: Mis Gastos vs Total Familiar
    modo_opcion = st.radio(
        "Modo de visualización:",
        [f"👤 Solo de {st.session_state['usuario_actual']}", "👨‍👩‍👧‍👦 Total Familiar (Todos)"],
        horizontal=True,
        label_visibility="collapsed"
    )
    st.session_state["modo_vista"] = "personal" if "Solo de" in modo_opcion else "familiar"

    # Datos según el modo
    hoy = datetime.now()
    meses_es = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    
    filtro_user = None if st.session_state["modo_vista"] == "familiar" else st.session_state["usuario_actual"]
    df_mes = db.obtener_gastos_df(mes=hoy.month, anio=hoy.year, usuario=filtro_user)
    
    total_mes = df_mes["monto"].sum() if not df_mes.empty else 0.0
    total_tc = df_mes[df_mes["medio_pago"] == "Tarjeta de Crédito"]["monto"].sum() if not df_mes.empty else 0.0
    total_ef = df_mes[df_mes["medio_pago"] != "Tarjeta de Crédito"]["monto"].sum() if not df_mes.empty else 0.0
    cant_compras = len(df_mes) if not df_mes.empty else 0

    etiqueta_hero = f"Gastos de {st.session_state['usuario_actual']} • {meses_es[hoy.month]}" if st.session_state["modo_vista"] == "personal" else f"Total del Hogar • {meses_es[hoy.month]}"

    # Hero Card
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

    # Cuadrícula de Acciones (Touch Cards)
    st.markdown("### ¿Qué quieres hacer?")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📸 Escanear Ticket\n\n(Foto o Cámara)", use_container_width=True, type="primary"):
            ir_a("ticket")
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        if st.button("📊 Estadísticas\n\n(Gráficos & Familia)", use_container_width=True):
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
        if st.button("👥 Miembros Familia\n\n(Agregar o Cambiar)", use_container_width=True):
            ir_a("familia")

    # Feed de últimos gastos
    st.markdown("<br>", unsafe_allow_html=True)
    titulo_feed = "🕒 Últimos movimientos familiares" if st.session_state["modo_vista"] == "familiar" else f"🕒 Últimos movimientos de {st.session_state['usuario_actual']}"
    st.markdown(f"### {titulo_feed}")
    
    df_ultimos = db.obtener_ultimos_gastos(limit=4, usuario=filtro_user)
    if df_ultimos.empty:
        st.caption("Aún no tienes movimientos registrados.")
    else:
        for _, fila in df_ultimos.iterrows():
            tag_u = f"[{fila['usuario']}] " if st.session_state["modo_vista"] == "familiar" else ""
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
# VISTA 2: ESCANEAR TICKET (IA)
# ==============================================================================
elif st.session_state["vista"] == "ticket":
    st.markdown('<div class="btn-back">', unsafe_allow_html=True)
    if st.button("⬅ Volver al Menú Principal", use_container_width=True):
        ir_a("menu")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("## 📸 Escanear Ticket")
    st.caption(f"Registrando gasto para: **{st.session_state['usuario_actual']}**")

    metodo = st.radio("Modo de captura:", ["📷 Cámara", "📁 Subir Archivo"], horizontal=True)
    img_ticket = st.camera_input("Capturar") if metodo == "📷 Cámara" else st.file_uploader("Seleccionar imagen:", type=["jpg", "jpeg", "png", "webp"])

    if img_ticket is not None:
        st.image(img_ticket, use_container_width=True)

        if st.button("✨ Analizar Ticket con IA", type="primary", use_container_width=True):
            with st.spinner("Analizando ticket con IA optimizada..."):
                try:
                    resultado = ai.extraer_ticket(img_ticket.getvalue(), img_ticket.type or "image/jpeg")
                    st.session_state["ticket_temp"] = resultado
                    st.success("✅ ¡Datos detectados! Revisa abajo:")
                except Exception as e:
                    st.error(f"Error: {e}")

    if "ticket_temp" in st.session_state and st.session_state["ticket_temp"]:
        d = st.session_state["ticket_temp"]
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
            
            # Selector de miembro asignado (por defecto el activo)
            idx_u_ticket = nombres_usuarios.index(st.session_state["usuario_actual"]) if st.session_state["usuario_actual"] in nombres_usuarios else 0
            n_usuario = st.selectbox("Gasto realizado por:", nombres_usuarios, index=idx_u_ticket)

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
                    usuario=n_usuario
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
    
    idx_u_tc = nombres_usuarios.index(st.session_state["usuario_actual"]) if st.session_state["usuario_actual"] in nombres_usuarios else 0
    usuario_tarjeta = st.selectbox("Tarjeta perteneciente a:", nombres_usuarios, index=idx_u_tc)
    
    archivo_tc = st.file_uploader("Subir PDF o foto del resumen bancario:", type=["pdf", "jpg", "jpeg", "png"])

    if archivo_tc is not None:
        if st.button("✨ Procesar Resumen Completo", type="primary", use_container_width=True):
            with st.spinner("Extrayendo consumos y cuotas con IA..."):
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
        df_tc["usuario"] = usuario_tarjeta
        st.markdown(f"**Total detectado:** ${df_tc['monto'].sum():,.2f}")
        
        df_edit = st.data_editor(
            df_tc,
            use_container_width=True,
            num_rows="dynamic"
        )
        
        if st.button("💾 Guardar TODOS los consumos", type="primary", use_container_width=True):
            total_guardados = db.insertar_gastos_lote(df_edit.to_dict(orient="records"), usuario_default=usuario_tarjeta)
            st.session_state["consumos_lote"] = None
            st.toast(f"¡{total_guardados} consumos asignados a {usuario_tarjeta}!", icon="✅")
            ir_a("menu")

# ==============================================================================
# VISTA 4: ESTADÍSTICAS & COMPARATIVA FAMILIAR
# ==============================================================================
elif st.session_state["vista"] == "dashboard":
    st.markdown('<div class="btn-back">', unsafe_allow_html=True)
    if st.button("⬅ Volver al Menú Principal", use_container_width=True):
        ir_a("menu")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("## 📊 Estadísticas")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        sel_anio = st.selectbox("Año:", [datetime.now().year, datetime.now().year - 1])
    with col_f2:
        meses_n = ["Todos", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                   "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
        sel_mes = st.selectbox("Mes:", meses_n, index=datetime.now().month)
        num_m = None if sel_mes == "Todos" else meses_n.index(sel_mes)

    # Filtro de usuario en estadísticas
    opciones_filtro_u = ["👨‍👩‍👧‍👦 Toda la Familia"] + nombres_usuarios
    idx_def_u = opciones_filtro_u.index(st.session_state["usuario_actual"]) if st.session_state["modo_vista"] == "personal" else 0
    sel_u_stats = st.selectbox("Ver datos de:", opciones_filtro_u, index=idx_def_u)
    
    filtro_db = None if "Toda la Familia" in sel_u_stats else sel_u_stats
    df_d = db.obtener_gastos_df(mes=num_m, anio=sel_anio, usuario=filtro_db)

    if df_d.empty:
        st.info("No hay gastos registrados para este período.")
    else:
        st.metric("Total Gastado:", f"${df_d['monto'].sum():,.2f}")

        # Si es vista familiar, mostrar comparación entre miembros
        if "Toda la Familia" in sel_u_stats and "usuario" in df_d.columns:
            st.subheader("👥 Gastos por Miembro de la Familia")
            df_usuarios_monto = df_d.groupby("usuario")["monto"].sum().reset_index().sort_values(by="monto", ascending=False)
            fig_fam = px.bar(
                df_usuarios_monto,
                x="usuario",
                y="monto",
                color="usuario",
                text_auto=".2s",
                labels={"usuario": "Miembro", "monto": "Monto ($)"}
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
        
        idx_u_manual = nombres_usuarios.index(st.session_state["usuario_actual"]) if st.session_state["usuario_actual"] in nombres_usuarios else 0
        f_user = st.selectbox("Miembro asignado:", nombres_usuarios, index=idx_u_manual)
        
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
                    usuario=f_user
                )
                st.toast("¡Gasto guardado con éxito!", icon="✅")
                ir_a("menu")

# ==============================================================================
# VISTA 6: GESTIÓN DE MIEMBROS DE LA FAMILIA
# ==============================================================================
elif st.session_state["vista"] == "familia":
    st.markdown('<div class="btn-back">', unsafe_allow_html=True)
    if st.button("⬅ Volver al Menú Principal", use_container_width=True):
        ir_a("menu")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("## 👥 Miembros de la Familia")
    st.caption("Administra los perfiles de tu hogar:")

    # Lista de perfiles actuales
    for u in lista_usuarios:
        es_activo = (u["nombre"] == st.session_state["usuario_actual"])
        col_u1, col_u2 = st.columns([4, 2])
        with col_u1:
            st.markdown(f"### {u['avatar']} {u['nombre']} {'*(Activo)*' if es_activo else ''}")
        with col_u2:
            if not es_activo:
                if st.button(f"Seleccionar", key=f"sel_{u['nombre']}", use_container_width=True):
                    st.session_state["usuario_actual"] = u["nombre"]
                    st.rerun()

    st.markdown("---")
    st.markdown("### ➕ Agregar Nuevo Integrante")
    with st.form("form_nuevo_miembro"):
        nuevo_nom = st.text_input("Nombre del familiar:")
        nuevo_av = st.selectbox("Avatar:", db.AVATARES_DEFAULT)
        if st.form_submit_button("Crear Perfil", type="primary", use_container_width=True):
            if nuevo_nom.strip():
                if db.crear_usuario(nuevo_nom, nuevo_av):
                    st.session_state["usuario_actual"] = nuevo_nom.strip().capitalize()
                    st.toast(f"¡Perfil de {nuevo_nom} creado!", icon="🎉")
                    st.rerun()
                else:
                    st.error("Ese nombre ya existe o es inválido.")
            else:
                st.error("Ingresa un nombre.")

# ==============================================================================
# VISTA 7: HISTORIAL & EXCEL
# ==============================================================================
elif st.session_state["vista"] == "historial":
    st.markdown('<div class="btn-back">', unsafe_allow_html=True)
    if st.button("⬅ Volver al Menú Principal", use_container_width=True):
        ir_a("menu")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("## 📑 Movimientos & Excel")
    
    # Botón directo para descargar Excel
    db.sincronizar_con_excel()
    if os.path.exists(db.EXCEL_FILE):
        with open(db.EXCEL_FILE, "rb") as f:
            st.download_button(
                label="📥 Descargar archivo Excel (.xlsx) [Con hojas separadas por miembro]",
                data=f,
                file_name=f"gastos_familia_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )

    st.markdown("<br>", unsafe_allow_html=True)
    
    # Filtro por usuario en la tabla
    filtro_u_hist = st.selectbox("Filtrar registros por:", ["👨‍👩‍👧‍👦 Toda la Familia"] + nombres_usuarios)
    u_query = None if "Toda la Familia" in filtro_u_hist else filtro_u_hist
    
    df_todos = db.obtener_gastos_df(usuario=u_query)
    
    if df_todos.empty:
        st.info("No hay registros en la base de datos.")
    else:
        st.dataframe(
            df_todos[["id", "fecha", "comercio", "monto", "usuario", "categoria", "medio_pago", "cuotas"]],
            use_container_width=True,
            hide_index=True
        )

        with st.expander("🗑️ Eliminar un registro"):
            id_borrar = st.selectbox("Seleccionar ID:", df_todos["id"].tolist())
            if st.button("Confirmar eliminación"):
                db.eliminar_gasto(id_borrar)
                st.toast(f"Registro #{id_borrar} eliminado.", icon="🗑️")
                st.rerun()
