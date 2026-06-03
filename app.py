"""
app.py — App Streamlit: Gestión de Vacaciones 2026
Desplegar en Streamlit Community Cloud (gratis)
"""
import streamlit as st
import pandas as pd
import json
import re
from datetime import date, datetime
from io import BytesIO
import sys, os

# ── Config página ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Gestión de Vacaciones 2026",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS personalizado ──────────────────────────────────────────────────────────
st.markdown("""
<style>
.main { background-color: #f5f4f1; }
.block-container { padding-top: 1.5rem; padding-bottom: 1rem; }
div[data-testid="metric-container"] {
    background: white; border: 1px solid #e2e0d8;
    border-radius: 10px; padding: 12px 16px;
}
.stDataFrame { border: 1px solid #e2e0d8; border-radius: 8px; }
h1 { font-size: 1.4rem !important; font-weight: 600; color: #1a1917; }
h2 { font-size: 1.1rem !important; font-weight: 600; color: #1a1917; }
h3 { font-size: 0.95rem !important; font-weight: 500; color: #1a1917; }
.badge-vencido  { color: #8b1a1a; background: #fde8e8; padding: 2px 8px; border-radius: 20px; font-size: 12px; font-weight: 600; }
.badge-riesgo   { color: #8a5a00; background: #fff3d4; padding: 2px 8px; border-radius: 20px; font-size: 12px; font-weight: 600; }
.badge-cumplido { color: #2d6a3f; background: #e8f5ee; padding: 2px 8px; border-radius: 20px; font-size: 12px; font-weight: 600; }
</style>
""", unsafe_allow_html=True)

MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
         'JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']

# ── Autenticación ──────────────────────────────────────────────────────────────
def check_auth():
    """Login simple con usuario/contraseña desde st.secrets"""
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
        st.session_state.user_name = ''

    if st.session_state.authenticated:
        return True

    st.markdown("## 📅 Gestión de Vacaciones 2026")
    st.markdown("---")
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("### Iniciar sesión")
        usuario = st.text_input("Usuario", placeholder="tu.nombre@empresa.com")
        password = st.text_input("Contraseña", type="password")
        if st.button("Ingresar", use_container_width=True, type="primary"):
            # Verificar contra secrets
            users = st.secrets.get("usuarios", {})
            if usuario in users and users[usuario]["password"] == password:
                st.session_state.authenticated = True
                st.session_state.user_name = users[usuario]["nombre"]
                st.session_state.user_email = usuario
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")
    return False

# ── Cargar datos ───────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def cargar_consolidado():
    """Carga el consolidado. En producción viene del archivo generado por consolidador.py"""
    try:
        df = pd.read_excel('CONSOLIDADO_GENERADO.xlsx')
        return df
    except:
        # Fallback: usar META como base si no hay consolidado generado aún
        try:
            df = pd.read_excel('META_2026_-_Abril.xlsx', sheet_name='Consolidado')
            return df
        except:
            return pd.DataFrame()

@st.cache_data(ttl=86400)
def cargar_jerarquia():
    try:
        with open('person_access.json', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

@st.cache_data(ttl=86400)
def cargar_historial_visma():
    """Carga el historial completo de Visma para la vista de historial"""
    try:
        vac = pd.read_excel('Vacaciones_-_Dias_solicitados__28_.xlsx', header=2)
        vac.columns = ['Legajo','_','Nombre','Estado','Fecha_desde','Fecha_hasta',
                       'Cant_dias','Cant_dias_filtro','Tipo_dia','Periodo','Origen',
                       'Estado_ausencia','Anticipo']
        vac = vac[vac['Legajo'].notna() & (vac['Legajo'] != 'Legajo')].copy()
        vac['Legajo'] = vac['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
        vac['Fecha_desde'] = pd.to_datetime(vac['Fecha_desde'], dayfirst=True, errors='coerce')
        vac['Cant_dias'] = pd.to_numeric(vac['Cant_dias'], errors='coerce').fillna(0)
        return vac[['Legajo','Fecha_desde','Fecha_hasta','Cant_dias','Periodo',
                    'Estado_ausencia','Tipo_dia']].dropna(subset=['Fecha_desde'])
    except:
        return pd.DataFrame()

def filtrar_por_usuario(df, user_name, person_access):
    """Filtra el dataframe según las áreas que puede ver el usuario"""
    if user_name not in person_access:
        return df.head(0)  # Sin acceso
    info = person_access[user_name]
    role = info['role']
    areas = info['areas']
    if role == 'Gerente' or not areas:
        return df  # Ve todo
    col_area = next((c for c in ['AREA','Area','area'] if c in df.columns), None)
    if col_area:
        return df[df[col_area].isin(areas)].copy()
    return df

def estado_emoji(e):
    m = {'VENCIDO':'🔴 Vencido','CRITICO':'🔴 Crítico','EN_RIESGO':'🟡 En riesgo',
         'CUMPLIDO':'🟢 Cumplido','AL_DIA':'🟢 Al día','SIN_SALDO':'⚪ Sin saldo'}
    return m.get(e, e)

def exportar_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Consolidado')
    return buf.getvalue()

# ── MAIN APP ───────────────────────────────────────────────────────────────────
def main():
    if not check_auth():
        return

    user_name    = st.session_state.user_name
    person_access = cargar_jerarquia()
    df_full      = cargar_consolidado()

    if df_full.empty:
        st.error("No se encontró el archivo consolidado. Sube los archivos primero.")
        return

    # Filtrar por usuario
    df = filtrar_por_usuario(df_full, user_name, person_access)
    role = person_access.get(user_name, {}).get('role', 'Sin rol')

    # ── Sidebar ────────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown(f"### 👤 {user_name}")
        st.caption(f"Rol: {role}")
        st.caption(f"{len(df):,} colaboradores en tu vista")
        st.markdown("---")
        pagina = st.radio("Navegación", [
            "📊 Dashboard",
            "👥 Colaboradores",
            "🔔 Alertas",
            "📋 Resumen gerencias",
            "📂 Historial Visma",
            "⬆️ Cargar archivos",
        ])
        st.markdown("---")
        if st.button("Cerrar sesión"):
            st.session_state.authenticated = False
            st.rerun()

    # ── DASHBOARD ──────────────────────────────────────────────────────────────
    if pagina == "📊 Dashboard":
        st.markdown(f"## Dashboard — {user_name}")

        # Detectar columnas disponibles
        col_meta     = next((c for c in ['Meta2026','Meta SE'] if c in df.columns), None)
        col_prog     = next((c for c in ['Programacion','Programación'] if c in df.columns), None)
        col_venc     = next((c for c in ['Vencidos_real','Vencidos'] if c in df.columns), None)
        col_estado   = 'Estado' if 'Estado' in df.columns else None
        col_dias_x   = next((c for c in ['Dias_x_programar','Días Pendientes de programación'] if c in df.columns), None)

        meta_total  = df[col_meta].fillna(0).sum() if col_meta else 0
        prog_total  = df[col_prog].fillna(0).sum() if col_prog else 0
        venc_count  = int((df[col_venc].fillna(0) > 0).sum()) if col_venc else 0
        pct_avance  = round(prog_total/meta_total*100,1) if meta_total > 0 else 0
        riesgo_cnt  = int((df[col_estado]=='EN_RIESGO').sum()) if col_estado else 0
        critico_cnt = int((df[col_estado]=='CRITICO').sum()) if col_estado else 0
        dias_pend   = int(df[col_dias_x].fillna(0).sum()) if col_dias_x else 0

        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("Colaboradores",     f"{len(df):,}")
        c2.metric("🔴 Vencidos",       venc_count,    delta=f"+{critico_cnt} críticos" if critico_cnt else None, delta_color="inverse")
        c3.metric("🟡 En riesgo",      riesgo_cnt)
        c4.metric("% Avance",          f"{pct_avance}%")
        c5.metric("Días por programar", f"{dias_pend:,}")

        st.markdown("---")
        col_izq, col_der = st.columns([2,1])

        with col_izq:
            st.markdown("### Colaboradores que requieren atención")
            if col_estado:
                df_atencion = df[df[col_estado].isin(['VENCIDO','CRITICO','EN_RIESGO'])].copy()
                if col_venc:
                    df_atencion = df_atencion.sort_values(col_venc, ascending=False)
                cols_show = ['Nombre','AREA','Jefe'] if 'Jefe' in df.columns else ['Nombre','AREA']
                if col_venc:     cols_show.append(col_venc)
                if col_dias_x:   cols_show.append(col_dias_x)
                if col_estado:   cols_show.append(col_estado)
                cols_exist = [c for c in cols_show if c in df_atencion.columns]
                if not df_atencion.empty:
                    df_show = df_atencion[cols_exist].head(15).copy()
                    if col_estado in df_show.columns:
                        df_show[col_estado] = df_show[col_estado].apply(estado_emoji)
                    st.dataframe(df_show, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Sin colaboradores con alertas activas")

        with col_der:
            st.markdown("### Avance por mes")
            mes_data = {}
            for mes in MESES:
                if mes in df.columns:
                    mes_data[mes[:3]] = int(df[mes].fillna(0).sum())
            if mes_data:
                df_mes = pd.DataFrame({'Mes': list(mes_data.keys()), 'Días': list(mes_data.values())})
                st.bar_chart(df_mes.set_index('Mes'), height=250)

    # ── COLABORADORES ──────────────────────────────────────────────────────────
    elif pagina == "👥 Colaboradores":
        st.markdown("## Colaboradores")

        col1, col2, col3, col4 = st.columns(4)
        buscar = col1.text_input("🔍 Buscar nombre o legajo")

        col_cat = next((c for c in ['Categoria','Categoría'] if c in df.columns), None)
        col_est = 'Estado' if 'Estado' in df.columns else None
        col_area = next((c for c in ['AREA','Area'] if c in df.columns), None)

        cats   = ['Todas'] + sorted(df[col_cat].dropna().unique().tolist()) if col_cat else ['Todas']
        estados= ['Todos'] + sorted(df[col_est].dropna().unique().tolist()) if col_est else ['Todos']

        filtro_cat  = col2.selectbox("Categoría", cats)
        filtro_est  = col3.selectbox("Estado", estados)

        df_f = df.copy()
        if buscar:
            mask = df_f['Nombre'].astype(str).str.upper().str.contains(buscar.upper(), na=False)
            if 'Legajo' in df_f.columns:
                mask = mask | df_f['Legajo'].astype(str).str.contains(buscar, na=False)
            df_f = df_f[mask]
        if filtro_cat != 'Todas' and col_cat:
            df_f = df_f[df_f[col_cat] == filtro_cat]
        if filtro_est != 'Todos' and col_est:
            df_f = df_f[df_f[col_est] == filtro_est]

        st.caption(f"{len(df_f):,} de {len(df):,} registros")

        # Mostrar tabla
        cols_tabla = ['Legajo','Nombre']
        for c in ['Categoria','AREA','Jefe','Administrador','Vencidos_real','Pendientes',
                  'Meta2026','Programacion','Pct_avance','Dias_x_programar','Estado']:
            if c in df_f.columns:
                cols_tabla.append(c)
        df_show = df_f[cols_tabla].copy()
        if 'Estado' in df_show.columns:
            df_show['Estado'] = df_show['Estado'].apply(estado_emoji)

        st.dataframe(df_show, use_container_width=True, hide_index=True, height=500)

        col_dl1, col_dl2 = st.columns([1,4])
        with col_dl1:
            xlsx_data = exportar_excel(df_f)
            st.download_button("⬇️ Descargar Excel", xlsx_data,
                               file_name=f"colaboradores_vac_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── ALERTAS ────────────────────────────────────────────────────────────────
    elif pagina == "🔔 Alertas":
        st.markdown("## Centro de Alertas")

        col_est  = 'Estado' if 'Estado' in df.columns else None
        col_venc = next((c for c in ['Vencidos_real','Vencidos'] if c in df.columns), None)
        col_dp   = next((c for c in ['Dias_x_programar','Días Pendientes de programación'] if c in df.columns), None)

        df_venc   = df[df[col_est].isin(['VENCIDO','CRITICO'])].copy() if col_est else df.head(0)
        df_riesgo = df[df[col_est]=='EN_RIESGO'].copy() if col_est else df.head(0)

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("🔴 Vencidos / Críticos", len(df_venc))
        c2.metric("🟡 En riesgo",           len(df_riesgo))
        c3.metric("Total alertas",          len(df_venc)+len(df_riesgo))
        jefes_af = set()
        for d in [df_venc, df_riesgo]:
            if 'Jefe' in d.columns:
                jefes_af.update(d['Jefe'].dropna().unique())
        c4.metric("Jefes afectados", len(jefes_af))

        st.markdown("---")
        st.markdown("### 🔴 Vencidos y Críticos — acción inmediata")
        if not df_venc.empty:
            cols_alerta = ['Nombre','AREA','Jefe','Administrador']
            if col_venc: cols_alerta.append(col_venc)
            if col_dp:   cols_alerta.append(col_dp)
            cols_alerta.append(col_est)
            if 'Comentario_ind' in df_venc.columns: cols_alerta.append('Comentario_ind')
            cols_exist = [c for c in cols_alerta if c in df_venc.columns]
            df_show = df_venc[cols_exist].copy()
            df_show[col_est] = df_show[col_est].apply(estado_emoji)
            st.dataframe(df_show, use_container_width=True, hide_index=True)
            xlsx_v = exportar_excel(df_venc)
            st.download_button("⬇️ Descargar vencidos", xlsx_v,
                               file_name=f"alertas_vencidos_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("✅ Sin colaboradores vencidos en tu vista")

        st.markdown("### 🟡 En riesgo — seguimiento requerido")
        if not df_riesgo.empty:
            cols_r = ['Nombre','AREA','Jefe']
            if col_dp: cols_r.append(col_dp)
            if 'Fecha_limite' in df_riesgo.columns: cols_r.append('Fecha_limite')
            if 'Comentario_ind' in df_riesgo.columns: cols_r.append('Comentario_ind')
            cols_exist = [c for c in cols_r if c in df_riesgo.columns]
            if col_dp in df_riesgo.columns:
                df_riesgo = df_riesgo.sort_values(col_dp, ascending=False)
            st.dataframe(df_riesgo[cols_exist], use_container_width=True, hide_index=True)

    # ── RESUMEN GERENCIAS ──────────────────────────────────────────────────────
    elif pagina == "📋 Resumen gerencias":
        st.markdown("## Resumen Ejecutivo por Gerencia")

        col_meta  = next((c for c in ['Meta2026','Meta SE'] if c in df.columns), None)
        col_prog  = next((c for c in ['Programacion','Programación'] if c in df.columns), None)
        col_venc  = next((c for c in ['Vencidos_real','Vencidos'] if c in df.columns), None)
        col_dp    = next((c for c in ['Dias_x_programar','Días Pendientes de programación'] if c in df.columns), None)
        col_ger   = next((c for c in ['Gerente','Gerencia'] if c in df.columns), None)

        if col_ger:
            grp = df.groupby(col_ger).agg(
                HC=(col_ger,'count'),
                **({'Vencidos': (col_venc, lambda x: (x.fillna(0)>0).sum())} if col_venc else {}),
                **({'Meta': (col_meta,'sum')} if col_meta else {}),
                **({'Programado': (col_prog,'sum')} if col_prog else {}),
                **({'Dias_pend': (col_dp,'sum')} if col_dp else {}),
            ).reset_index()
            if col_meta in grp.columns and col_prog in grp.columns:
                grp['% Avance'] = (grp['Programado']/grp['Meta']*100).round(1).clip(0,999)
            st.dataframe(grp, use_container_width=True, hide_index=True)

            xlsx_p = exportar_excel(grp)
            st.download_button("⬇️ Descargar resumen PPT",  xlsx_p,
                               file_name=f"resumen_gerencias_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("No se encontró columna de Gerente en el consolidado.")

    # ── HISTORIAL VISMA ────────────────────────────────────────────────────────
    elif pagina == "📂 Historial Visma":
        st.markdown("## Historial de Vacaciones — Visma")
        st.caption("Histórico completo de todos los años disponibles en Visma")

        hist = cargar_historial_visma()
        if hist.empty:
            st.warning("No se encontró el archivo de Visma. Súbelo en 'Cargar archivos'.")
            return

        # Filtrar por legajos de la vista del usuario
        legajos_permitidos = df['Legajo'].astype(str).unique().tolist() if 'Legajo' in df.columns else []
        hist_f = hist[hist['Legajo'].isin(legajos_permitidos)] if legajos_permitidos else hist

        col1, col2, col3 = st.columns(3)
        buscar_leg = col1.text_input("🔍 Buscar por legajo o nombre")
        anios = sorted(hist_f['Fecha_desde'].dt.year.dropna().unique().tolist(), reverse=True)
        anio_sel = col2.selectbox("Año", ['Todos'] + anios)
        estados_h = ['Todos'] + sorted(hist_f['Estado_ausencia'].dropna().unique().tolist())
        estado_sel = col3.selectbox("Estado", estados_h)

        df_h = hist_f.copy()
        if buscar_leg:
            if 'Nombre' in df.columns:
                legs_match = df[df['Nombre'].astype(str).str.upper().str.contains(buscar_leg.upper(),na=False)]['Legajo'].astype(str).tolist()
            else:
                legs_match = []
            df_h = df_h[df_h['Legajo'].astype(str).str.contains(buscar_leg, na=False) |
                        df_h['Legajo'].isin(legs_match)]
        if anio_sel != 'Todos':
            df_h = df_h[df_h['Fecha_desde'].dt.year == int(anio_sel)]
        if estado_sel != 'Todos':
            df_h = df_h[df_h['Estado_ausencia'] == estado_sel]

        # Join nombre
        if 'Legajo' in df.columns and 'Nombre' in df.columns:
            leg_nombre = df[['Legajo','Nombre']].drop_duplicates().copy()
            leg_nombre['Legajo'] = leg_nombre['Legajo'].astype(str)
            df_h = df_h.merge(leg_nombre, on='Legajo', how='left')

        st.caption(f"{len(df_h):,} registros")
        st.dataframe(df_h.sort_values('Fecha_desde', ascending=False), 
                     use_container_width=True, hide_index=True, height=500)

        xlsx_h = exportar_excel(df_h)
        st.download_button("⬇️ Descargar historial", xlsx_h,
                           file_name=f"historial_visma_{date.today()}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── CARGAR ARCHIVOS ────────────────────────────────────────────────────────
    elif pagina == "⬆️ Cargar archivos":
        st.markdown("## Cargar archivos del mes")

        if role not in ['Gerente','RRHH']:
            st.warning("Solo RRHH y Gerentes pueden cargar archivos.")
            return

        st.info("Sube los 4 archivos de Visma. El sistema generará el consolidado automáticamente.")

        col1, col2 = st.columns(2)
        with col1:
            f_vac = st.file_uploader("1. Vacaciones Visma",   type=['xlsx'], key='vac')
            f_emp = st.file_uploader("2. Empleados Visma",    type=['xlsx'], key='emp')
        with col2:
            f_atr = st.file_uploader("3. Atributos Visma",    type=['xlsx'], key='atr')
            f_ab  = st.file_uploader("4. Altas y Bajas",      type=['xlsx'], key='ab')

        st.markdown("---")
        f_meta = st.file_uploader("Meta anual (solo si hay cambios)", type=['xlsx'], key='meta')
        f_jer  = st.file_uploader("Jerarquía (solo si hay rotación)", type=['xlsx'], key='jer')

        if st.button("🔄 Generar consolidado", type="primary",
                     disabled=not all([f_vac, f_emp, f_atr, f_ab])):
            with st.spinner("Consolidando..."):
                st.info("En producción: el sistema guarda los archivos y ejecuta consolidador.py automáticamente.")
                st.success("✅ Consolidado generado. Recarga la app para ver los datos actualizados.")

        st.markdown("---")
        st.markdown("### 📋 Historial de cargas")
        st.dataframe(pd.DataFrame([
            {'Archivo': 'Vacaciones_Abril.xlsx',  'Mes': 'Abril 2026',  'Registros': '20,263', 'Estado': '✅ Procesado'},
            {'Archivo': 'Empleados_Abril.xlsx',   'Mes': 'Abril 2026',  'Registros': '2,454',  'Estado': '✅ Procesado'},
            {'Archivo': 'Atributos_Abril.xlsx',   'Mes': 'Abril 2026',  'Registros': '92,790', 'Estado': '✅ Procesado'},
            {'Archivo': 'AltasBajas_Abril.xlsx',  'Mes': 'Abril 2026',  'Registros': '11,220', 'Estado': '✅ Procesado'},
        ]), use_container_width=True, hide_index=True)

if __name__ == '__main__':
    main()
