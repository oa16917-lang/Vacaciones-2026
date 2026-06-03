"""
app.py — App Streamlit: Gestión de Vacaciones 2026
"""
import streamlit as st
import pandas as pd
import json
import re
from datetime import date, datetime
from io import BytesIO

st.set_page_config(
    page_title="Gestión de Vacaciones 2026",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.main { background-color: #f5f4f1; }
.block-container { padding-top: 1.5rem; padding-bottom: 1rem; }
div[data-testid="metric-container"] {
    background: white; border: 1px solid #e2e0d8;
    border-radius: 10px; padding: 12px 16px;
}
h1 { font-size: 1.4rem !important; font-weight: 600; color: #1a1917; }
h2 { font-size: 1.1rem !important; font-weight: 600; color: #1a1917; }
h3 { font-size: 0.95rem !important; font-weight: 500; color: #1a1917; }
</style>
""", unsafe_allow_html=True)

MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
         'JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']

def check_auth():
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
            users = st.secrets.get("usuarios", {})
            if usuario in users and users[usuario]["password"] == password:
                st.session_state.authenticated = True
                st.session_state.user_name = users[usuario]["nombre"]
                st.session_state.user_email = usuario
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")
    return False

@st.cache_data(ttl=3600)
def cargar_consolidado():
    rename_map = {
        'Apellidos y Nombres': 'Nombre',
        'Area': 'AREA',
        'Sede': 'SEDE',
        'Meta SE': 'Meta2026',
        'Programación': 'Programacion',
        'Días Pendientes de programación': 'Dias_x_programar',
        'COMENTARIO PARA EVITAR INDEMNIZACION': 'Comentario_ind',
        'COMENTARIOS PARA CUMPLIMIENTO META 2026': 'Comentario_meta',
    }
    # Intenta primero consolidado generado, luego META directo
    sources = [
        ('CONSOLIDADO_GENERADO.xlsx', None),
        ('META_2026_-_Abril.xlsx', 'Consolidado'),
    ]
    for filename, sheet in sources:
        try:
            if sheet:
                df = pd.read_excel(filename, sheet_name=sheet)
            else:
                df = pd.read_excel(filename)
            if len(df) > 0:
                df = df.rename(columns={k:v for k,v in rename_map.items() if k in df.columns})
                df['Legajo'] = df['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
                return df
        except Exception:
            continue
    return pd.DataFrame()

@st.cache_data(ttl=86400)
def cargar_jerarquia():
    try:
        with open('person_access.json', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}

@st.cache_data(ttl=86400)
def cargar_historial_visma():
    try:
        vac = pd.read_excel('Vacaciones_-_Dias_solicitados__28_.xlsx', header=2)
        vac.columns = ['Legajo','_','Nombre','Estado','Fecha_desde','Fecha_hasta',
                       'Cant_dias','Cant_dias_filtro','Tipo_dia','Periodo','Origen',
                       'Estado_ausencia','Anticipo']
        vac = vac[vac['Legajo'].notna() & (vac['Legajo'] != 'Legajo')].copy()
        vac['Legajo'] = vac['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
        vac['Fecha_desde'] = pd.to_datetime(vac['Fecha_desde'], dayfirst=True, errors='coerce')
        vac['Cant_dias'] = pd.to_numeric(vac['Cant_dias'], errors='coerce').fillna(0)
        return vac[['Legajo','Fecha_desde','Fecha_hasta','Cant_dias',
                    'Periodo','Estado_ausencia','Tipo_dia']].dropna(subset=['Fecha_desde'])
    except Exception:
        return pd.DataFrame()

def filtrar_por_usuario(df, user_name, person_access):
    if user_name not in person_access:
        return df.head(0)
    info = person_access[user_name]
    role = info['role']
    areas = info['areas']
    if role == 'Gerente' or not areas:
        return df
    col_area = next((c for c in ['AREA','Area','area'] if c in df.columns), None)
    if col_area:
        return df[df[col_area].isin(areas)].copy()
    return df

def estado_emoji(e):
    m = {'VENCIDO':'🔴 Vencido','CRITICO':'🔴 Crítico','EN_RIESGO':'🟡 En riesgo',
         'CUMPLIDO':'🟢 Cumplido','AL_DIA':'🟢 Al día','SIN_SALDO':'⚪ Sin saldo'}
    return m.get(str(e), str(e))

def exportar_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Consolidado')
    return buf.getvalue()

def main():
    if not check_auth():
        return

    user_name = st.session_state.user_name
    person_access = cargar_jerarquia()
    df_full = cargar_consolidado()

    if df_full.empty:
        st.error("No se encontró el archivo de datos. Sube META_2026_-_Abril.xlsx al repositorio.")
        return

    df = filtrar_por_usuario(df_full, user_name, person_access)
    role = person_access.get(user_name, {}).get('role', 'RRHH')

    # Si no está en jerarquía pero está autenticado, ve todo (admin)
    if df.empty and user_name not in person_access:
        df = df_full.copy()
        role = 'RRHH'

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

    # Columnas clave
    col_meta   = next((c for c in ['Meta2026','Meta SE'] if c in df.columns), None)
    col_prog   = next((c for c in ['Programacion','Programación'] if c in df.columns), None)
    col_venc   = next((c for c in ['Vencidos_real','Vencidos'] if c in df.columns), None)
    col_estado = 'Estado' if 'Estado' in df.columns else None
    col_dp     = next((c for c in ['Dias_x_programar','Días Pendientes de programación'] if c in df.columns), None)
    col_area   = next((c for c in ['AREA','Area'] if c in df.columns), None)
    col_cat    = next((c for c in ['Categoria','Categoría'] if c in df.columns), None)
    col_ger    = next((c for c in ['Gerente','Gerencia'] if c in df.columns), None)

    if pagina == "📊 Dashboard":
        st.markdown(f"## Dashboard — {user_name}")

        meta_total = df[col_meta].fillna(0).sum() if col_meta else 0
        prog_total = df[col_prog].fillna(0).sum() if col_prog else 0
        venc_count = int((df[col_venc].fillna(0) > 0).sum()) if col_venc else 0
        pct_avance = round(prog_total/meta_total*100,1) if meta_total > 0 else 0
        riesgo_cnt = int((df[col_estado]=='EN_RIESGO').sum()) if col_estado else 0
        dias_pend  = int(df[col_dp].fillna(0).sum()) if col_dp else 0

        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("Colaboradores",      f"{len(df):,}")
        c2.metric("🔴 Vencidos",        venc_count)
        c3.metric("🟡 En riesgo",       riesgo_cnt)
        c4.metric("% Avance",           f"{pct_avance}%")
        c5.metric("Días por programar", f"{dias_pend:,}")

        st.markdown("---")
        col_izq, col_der = st.columns([2,1])
        with col_izq:
            st.markdown("### Colaboradores que requieren atención")
            if col_estado:
                df_at = df[df[col_estado].isin(['VENCIDO','CRITICO','EN_RIESGO'])].copy()
                cols_s = [c for c in ['Nombre','Apellidos y Nombres',col_area,'Jefe',
                                       col_venc,col_dp,col_estado] if c and c in df_at.columns]
                if not df_at.empty:
                    show = df_at[cols_s].head(15).copy()
                    if col_estado in show.columns:
                        show[col_estado] = show[col_estado].apply(estado_emoji)
                    st.dataframe(show, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Sin colaboradores con alertas")
            else:
                cols_s = [c for c in ['Apellidos y Nombres','Nombre',col_area,col_dp] if c and c in df.columns]
                st.dataframe(df[cols_s].head(15), use_container_width=True, hide_index=True)

        with col_der:
            st.markdown("### Avance por mes")
            mes_data = {m[:3]: int(df[m].fillna(0).sum()) for m in MESES if m in df.columns}
            if mes_data:
                st.bar_chart(pd.DataFrame({'Días': mes_data}), height=260)

    elif pagina == "👥 Colaboradores":
        st.markdown("## Colaboradores")
        col1,col2,col3,col4 = st.columns(4)
        buscar = col1.text_input("🔍 Nombre o legajo")
        cats   = ['Todas'] + sorted(df[col_cat].dropna().unique().tolist()) if col_cat else ['Todas']
        ests   = ['Todos'] + sorted(df[col_estado].dropna().unique().tolist()) if col_estado else ['Todos']
        filtro_cat = col2.selectbox("Categoría", cats)
        filtro_est = col3.selectbox("Estado",    ests)

        df_f = df.copy()
        nombre_col = next((c for c in ['Nombre','Apellidos y Nombres'] if c in df_f.columns), None)
        if buscar and nombre_col:
            mask = df_f[nombre_col].astype(str).str.upper().str.contains(buscar.upper(), na=False)
            if 'Legajo' in df_f.columns:
                mask = mask | df_f['Legajo'].astype(str).str.contains(buscar, na=False)
            df_f = df_f[mask]
        if filtro_cat != 'Todas' and col_cat:
            df_f = df_f[df_f[col_cat] == filtro_cat]
        if filtro_est != 'Todos' and col_estado:
            df_f = df_f[df_f[col_estado] == filtro_est]

        st.caption(f"{len(df_f):,} de {len(df):,} registros")
        cols_t = [c for c in ['Legajo','Nombre','Apellidos y Nombres','Categoria','Categoría',
                               'AREA','Area','Jefe','Administrador','Vencidos_real','Vencidos',
                               'Pendientes','Meta2026','Meta SE','Programacion','Programación',
                               'Pct_avance','%','Dias_x_programar',
                               'Días Pendientes de programación','Estado']
                  if c in df_f.columns]
        show = df_f[cols_t].copy()
        if col_estado in show.columns:
            show[col_estado] = show[col_estado].apply(estado_emoji)
        st.dataframe(show, use_container_width=True, hide_index=True, height=500)
        st.download_button("⬇️ Descargar Excel", exportar_excel(df_f),
                           file_name=f"colaboradores_{date.today()}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    elif pagina == "🔔 Alertas":
        st.markdown("## Centro de Alertas")
        df_venc   = df[df[col_estado].isin(['VENCIDO','CRITICO'])].copy() if col_estado else df.head(0)
        df_riesgo = df[df[col_estado]=='EN_RIESGO'].copy() if col_estado else df.head(0)

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("🔴 Vencidos/Críticos", len(df_venc))
        c2.metric("🟡 En riesgo",         len(df_riesgo))
        c3.metric("Total alertas",         len(df_venc)+len(df_riesgo))
        jefes_af = set()
        for d in [df_venc, df_riesgo]:
            if 'Jefe' in d.columns:
                jefes_af.update(d['Jefe'].dropna().unique())
        c4.metric("Jefes afectados", len(jefes_af))

        st.markdown("---")
        st.markdown("### 🔴 Vencidos y Críticos")
        if not df_venc.empty:
            cols_a = [c for c in ['Nombre','Apellidos y Nombres','AREA','Area','Jefe',
                                   'Vencidos_real','Vencidos',col_dp,'Estado','Comentario_ind',
                                   'COMENTARIO PARA EVITAR INDEMNIZACION'] if c in df_venc.columns]
            show_v = df_venc[cols_a].copy()
            if col_estado in show_v.columns:
                show_v[col_estado] = show_v[col_estado].apply(estado_emoji)
            st.dataframe(show_v, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Descargar vencidos", exportar_excel(df_venc),
                               file_name=f"vencidos_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("✅ Sin vencidos en tu vista")

        st.markdown("### 🟡 En riesgo")
        if not df_riesgo.empty:
            cols_r = [c for c in ['Nombre','Apellidos y Nombres','AREA','Area','Jefe',
                                   col_dp,'Dias_x_programar','Comentario_ind',
                                   'COMENTARIO PARA EVITAR INDEMNIZACION'] if c in df_riesgo.columns]
            st.dataframe(df_riesgo[cols_r], use_container_width=True, hide_index=True)
        else:
            st.success("✅ Sin colaboradores en riesgo")

    elif pagina == "📋 Resumen gerencias":
        st.markdown("## Resumen Ejecutivo por Gerencia")
        if col_ger:
            agg = {}
            if col_venc:  agg['Vencidos']  = (col_venc,  lambda x: (x.fillna(0)>0).sum())
            if col_meta:  agg['Meta']       = (col_meta,  'sum')
            if col_prog:  agg['Programado'] = (col_prog,  'sum')
            if col_dp:    agg['Dias_pend']  = (col_dp,    'sum')
            grp = df.groupby(col_ger).agg(HC=(col_ger,'count'), **agg).reset_index()
            if 'Meta' in grp.columns and 'Programado' in grp.columns:
                grp['% Avance'] = (grp['Programado']/grp['Meta']*100).round(1).clip(0,999)
            st.dataframe(grp, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Descargar resumen", exportar_excel(grp),
                               file_name=f"resumen_gerencias_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            # Agrupar por Jefe si no hay gerencia
            col_jefe = next((c for c in ['Jefe','Sub_Gerente'] if c in df.columns), None)
            if col_jefe:
                grp = df.groupby(col_jefe).agg(HC=(col_jefe,'count')).reset_index()
                st.dataframe(grp, use_container_width=True, hide_index=True)
            else:
                st.info("Sube el consolidado generado para ver el resumen por gerencia.")

    elif pagina == "📂 Historial Visma":
        st.markdown("## Historial de Vacaciones — Visma")
        hist = cargar_historial_visma()
        if hist.empty:
            st.warning("Sube el archivo Vacaciones_-_Dias_solicitados__28_.xlsx al repositorio.")
            return
        legajos = df['Legajo'].astype(str).unique().tolist() if 'Legajo' in df.columns else []
        hist_f = hist[hist['Legajo'].isin(legajos)] if legajos else hist

        col1,col2,col3 = st.columns(3)
        buscar_l = col1.text_input("🔍 Legajo o nombre")
        anios    = sorted(hist_f['Fecha_desde'].dt.year.dropna().unique().tolist(), reverse=True)
        anio_s   = col2.selectbox("Año", ['Todos'] + anios)
        ests_h   = ['Todos'] + sorted(hist_f['Estado_ausencia'].dropna().unique().tolist())
        est_s    = col3.selectbox("Estado", ests_h)

        df_h = hist_f.copy()
        if buscar_l:
            df_h = df_h[df_h['Legajo'].astype(str).str.contains(buscar_l, na=False)]
        if anio_s != 'Todos':
            df_h = df_h[df_h['Fecha_desde'].dt.year == int(anio_s)]
        if est_s != 'Todos':
            df_h = df_h[df_h['Estado_ausencia'] == est_s]

        nombre_col = next((c for c in ['Nombre','Apellidos y Nombres'] if c in df.columns), None)
        if nombre_col and 'Legajo' in df.columns:
            ln = df[['Legajo', nombre_col]].drop_duplicates()
            ln['Legajo'] = ln['Legajo'].astype(str)
            df_h = df_h.merge(ln, on='Legajo', how='left')

        st.caption(f"{len(df_h):,} registros")
        st.dataframe(df_h.sort_values('Fecha_desde', ascending=False),
                     use_container_width=True, hide_index=True, height=500)
        st.download_button("⬇️ Descargar historial", exportar_excel(df_h),
                           file_name=f"historial_{date.today()}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    elif pagina == "⬆️ Cargar archivos":
        st.markdown("## Cargar archivos del mes")
        if role not in ['Gerente','RRHH']:
            st.warning("Solo RRHH y Gerentes pueden cargar archivos.")
            return
        st.info("Sube los archivos de Visma. El sistema actualizará los datos automáticamente.")
        col1,col2 = st.columns(2)
        with col1:
            st.file_uploader("1. Vacaciones Visma",  type=['xlsx'], key='vac')
            st.file_uploader("2. Empleados Visma",   type=['xlsx'], key='emp')
        with col2:
            st.file_uploader("3. Atributos Visma",   type=['xlsx'], key='atr')
            st.file_uploader("4. Altas y Bajas",     type=['xlsx'], key='ab')
        st.markdown("---")
        st.file_uploader("Meta anual (solo si hay cambios)", type=['xlsx'], key='meta')
        st.file_uploader("Jerarquía (solo si hay rotación)", type=['xlsx'], key='jer')
        st.markdown("---")
        st.markdown("### 📋 Historial de cargas")
        st.dataframe(pd.DataFrame([
            {'Archivo':'META_2026_-_Abril.xlsx','Mes':'Abril 2026','Estado':'✅ Activo'},
        ]), use_container_width=True, hide_index=True)

if __name__ == '__main__':
    main()
