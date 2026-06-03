"""
vacaciones.py — App Streamlit: Gestión de Vacaciones 2026
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
h1,h2,h3 { color: #1a1917; }
</style>
""", unsafe_allow_html=True)

MESES = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
         'JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']

# ── Auth ───────────────────────────────────────────────────────────────────────
def check_auth():
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True
    st.markdown("## 📅 Gestión de Vacaciones 2026")
    st.markdown("---")
    c1, c2, c3 = st.columns([1,1.2,1])
    with c2:
        st.markdown("### Iniciar sesión")
        usuario  = st.text_input("Usuario", placeholder="correo@empresa.com")
        password = st.text_input("Contraseña", type="password")
        if st.button("Ingresar", use_container_width=True, type="primary"):
            users = st.secrets.get("usuarios", {})
            if usuario in users and users[usuario]["password"] == password:
                st.session_state.authenticated = True
                st.session_state.user_name  = users[usuario]["nombre"]
                st.session_state.user_email = usuario
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")
    return False

# ── Carga de datos ─────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def cargar_consolidado():
    rename_map = {
        'Apellidos y Nombres': 'Nombre',
        'Area': 'AREA', 'Sede': 'SEDE',
        'Meta SE': 'Meta2026',
        'Programación': 'Programacion',
        'Días Pendientes de programación': 'Dias_x_programar',
        'COMENTARIO PARA EVITAR INDEMNIZACION': 'Comentario_ind',
        'COMENTARIOS PARA CUMPLIMIENTO META 2026': 'Comentario_meta',
    }
    for filename, sheet in [('CONSOLIDADO_GENERADO.xlsx', None),
                              ('META_2026_-_Abril.xlsx', 'Consolidado')]:
        try:
            df = pd.read_excel(filename, sheet_name=sheet)
            if len(df) > 0:
                df = df.rename(columns={k:v for k,v in rename_map.items() if k in df.columns})
                df['Legajo'] = df['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
                # Excluir personal sin área (cesados sin datos)
                col_area = next((c for c in ['AREA','Area'] if c in df.columns), None)
                if col_area:
                    df = df[df[col_area].notna() & (df[col_area].astype(str).str.strip() != '')]
                # Calcular vencidos reales desde comentario
                df = calcular_estados(df)
                return df
        except Exception:
            continue
    return pd.DataFrame()

def extraer_fecha_limite(comentario):
    if not comentario or str(comentario).strip() in ['-','nan','None','']:
        return None
    match = re.search(r'ANTES DEL (\d{2}/\d{2}/\d{4})', str(comentario).upper())
    if match:
        try:
            return datetime.strptime(match.group(1), '%d/%m/%Y').date()
        except Exception:
            return None
    return None

def calcular_estados(df):
    hoy = date.today()
    col_ind  = next((c for c in ['Comentario_ind','COMENTARIO PARA EVITAR INDEMNIZACION'] if c in df.columns), None)
    col_pend = next((c for c in ['Pendientes'] if c in df.columns), None)
    col_meta = next((c for c in ['Meta2026','Meta SE'] if c in df.columns), None)
    col_prog = next((c for c in ['Programacion','Programación'] if c in df.columns), None)

    estados, vencidos_real, dias_alerta, fechas_limite = [], [], [], []

    for _, row in df.iterrows():
        comentario = str(row[col_ind]) if col_ind and pd.notna(row[col_ind]) else ''
        pendientes = float(row[col_pend]) if col_pend and pd.notna(row[col_pend]) else 0
        meta       = float(row[col_meta]) if col_meta and pd.notna(row[col_meta]) else 0
        prog       = float(row[col_prog]) if col_prog and pd.notna(row[col_prog]) else 0
        # Cap programacion at 100% per person for KPI
        prog_capped = min(prog, meta) if meta > 0 else prog
        fecha_lim  = extraer_fecha_limite(comentario)

        # Vencidos reales
        venc = 0
        if fecha_lim and fecha_lim < hoy and pendientes > 0:
            match_d = re.search(r'DEBE GOZAR (\d+)', comentario.upper())
            dias_debia = int(match_d.group(1)) if match_d else int(pendientes)
            if prog < dias_debia:
                venc = max(0, dias_debia - prog)

        # Estado semáforo
        dias_rest = (fecha_lim - hoy).days if fecha_lim else 999
        if venc > 0:
            estado = 'VENCIDO'
        elif fecha_lim and dias_rest <= 30 and pendientes > 0:
            estado = 'CRITICO'
        elif fecha_lim and dias_rest <= 90 and pendientes > 0:
            estado = 'EN_RIESGO'
        elif meta > 0 and prog_capped >= meta:
            estado = 'CUMPLIDO'
        elif meta == 0 and pendientes == 0:
            estado = 'SIN_SALDO'
        else:
            estado = 'AL_DIA'

        estados.append(estado)
        vencidos_real.append(round(venc, 1))
        fechas_limite.append(str(fecha_lim) if fecha_lim else '')
        dias_alerta.append(dias_rest if fecha_lim else 999)

    df = df.copy()
    df['Estado']        = estados
    df['Vencidos_real'] = vencidos_real
    df['Fecha_limite']  = fechas_limite
    df['Dias_alerta']   = dias_alerta

    # Recalcular % con cap por persona
    if col_meta and col_prog:
        df['Pct_avance'] = df.apply(
            lambda r: round(min(float(r[col_prog] or 0), float(r[col_meta] or 0)) /
                            float(r[col_meta]) * 100, 1) if float(r[col_meta] or 0) > 0 else 0,
            axis=1
        )
    return df

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
        vac.columns = ['Legajo','_','Nombre_vac','Estado','Fecha_desde','Fecha_hasta',
                       'Cant_dias','_2','Tipo_dia','Periodo','Origen','Estado_ausencia','Anticipo']
        vac = vac[vac['Legajo'].notna() & (vac['Legajo'] != 'Legajo')].copy()
        vac['Legajo']      = vac['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
        vac['Fecha_desde'] = pd.to_datetime(vac['Fecha_desde'], dayfirst=True, errors='coerce')
        vac['Fecha_hasta'] = pd.to_datetime(vac['Fecha_hasta'], dayfirst=True, errors='coerce')
        vac['Cant_dias']   = pd.to_numeric(vac['Cant_dias'], errors='coerce').fillna(0)
        return vac[['Legajo','Fecha_desde','Fecha_hasta','Cant_dias',
                    'Periodo','Estado_ausencia','Tipo_dia']].dropna(subset=['Fecha_desde'])
    except Exception:
        return pd.DataFrame()

def filtrar_por_usuario(df, user_name, person_access):
    if user_name not in person_access:
        return df
    info  = person_access[user_name]
    role  = info['role']
    areas = info.get('areas', [])
    if role == 'Gerente' or not areas:
        return df
    col_area = next((c for c in ['AREA','Area'] if c in df.columns), None)
    if col_area:
        return df[df[col_area].isin(areas)].copy()
    return df

def estado_emoji(e):
    m = {'VENCIDO':'🔴 Vencido','CRITICO':'🔴 Crítico','EN_RIESGO':'🟡 En riesgo',
         'CUMPLIDO':'🟢 Cumplido','AL_DIA':'🟢 Al día','SIN_SALDO':'⚪ Sin saldo'}
    return m.get(str(e), str(e))

def exportar_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        df.to_excel(w, index=False, sheet_name='Datos')
    return buf.getvalue()

# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    if not check_auth():
        return

    user_name     = st.session_state.user_name
    person_access = cargar_jerarquia()
    df_full       = cargar_consolidado()

    if df_full.empty:
        st.error("No se encontró el archivo de datos. Sube META_2026_-_Abril.xlsx al repositorio.")
        return

    df   = filtrar_por_usuario(df_full, user_name, person_access)
    role = person_access.get(user_name, {}).get('role', 'RRHH')

    # Columnas clave
    col_meta  = next((c for c in ['Meta2026','Meta SE']                   if c in df.columns), None)
    col_prog  = next((c for c in ['Programacion','Programación']          if c in df.columns), None)
    col_pend  = 'Pendientes' if 'Pendientes' in df.columns else None
    col_dp    = next((c for c in ['Dias_x_programar',
                                   'Días Pendientes de programación']     if c in df.columns), None)
    col_area  = next((c for c in ['AREA','Area']                          if c in df.columns), None)
    col_cat   = next((c for c in ['Categoria','Categoría']                if c in df.columns), None)
    col_ger   = next((c for c in ['Gerente','Gerencia']                   if c in df.columns), None)
    col_jefe  = next((c for c in ['Jefe']                                 if c in df.columns), None)
    col_nom   = next((c for c in ['Nombre','Apellidos y Nombres']         if c in df.columns), None)

    with st.sidebar:
        st.markdown(f"### 👤 {user_name}")
        st.caption(f"Rol: {role}")
        st.caption(f"{len(df):,} colaboradores en tu vista")
        st.markdown("---")
        pagina = st.radio("Navegación", [
            "📊 Dashboard", "👥 Colaboradores", "🔔 Alertas",
            "📋 Resumen gerencias", "📂 Historial Visma", "⬆️ Cargar archivos",
        ])
        st.markdown("---")
        if st.button("Cerrar sesión"):
            st.session_state.authenticated = False
            st.rerun()

    # ── DASHBOARD ──────────────────────────────────────────────────────────────
    if pagina == "📊 Dashboard":
        st.markdown(f"## Dashboard — {user_name}")

        meta_total  = df[col_meta].fillna(0).astype(float).sum() if col_meta else 0
        prog_total  = df[col_prog].fillna(0).astype(float).apply(lambda x: x).sum() if col_prog else 0
        # Cap programacion para KPI global
        prog_capped = df.apply(lambda r: min(float(r[col_prog] or 0), float(r[col_meta] or 0))
                               if col_meta and col_prog else 0, axis=1).sum()
        pct_avance  = round(prog_capped / meta_total * 100, 1) if meta_total > 0 else 0
        venc_count  = int((df['Vencidos_real'] > 0).sum()) if 'Vencidos_real' in df.columns else 0
        critico_cnt = int((df['Estado'] == 'CRITICO').sum()) if 'Estado' in df.columns else 0
        riesgo_cnt  = int((df['Estado'] == 'EN_RIESGO').sum()) if 'Estado' in df.columns else 0
        dias_pend   = int(df[col_dp].fillna(0).astype(float).sum()) if col_dp else 0

        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("Colaboradores",      f"{len(df):,}")
        c2.metric("🔴 Vencidos",        venc_count)
        c3.metric("🟠 Críticos (≤30d)", critico_cnt)
        c4.metric("% Avance",           f"{pct_avance}%")
        c5.metric("Días por programar", f"{dias_pend:,}")

        st.markdown("---")
        col_izq, col_der = st.columns([2,1])

        with col_izq:
            st.markdown("### Colaboradores que requieren atención")
            if 'Estado' in df.columns:
                # Solo mostrar los que tienen días pendientes reales > 0
                df_at = df[
                    (df['Estado'].isin(['VENCIDO','CRITICO','EN_RIESGO'])) &
                    (df[col_pend].fillna(0).astype(float) > 0 if col_pend else True)
                ].copy()
                if not df_at.empty:
                    df_at = df_at.sort_values('Vencidos_real', ascending=False)
                    cols_s = [c for c in [col_nom, col_area, col_jefe,
                                          'Vencidos_real', col_pend, col_dp,
                                          'Fecha_limite', 'Estado']
                              if c and c in df_at.columns]
                    show = df_at[cols_s].head(15).copy()
                    show['Estado'] = show['Estado'].apply(estado_emoji)
                    st.dataframe(show, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Sin colaboradores con alertas activas")

        with col_der:
            st.markdown("### Avance por mes")
            mes_data = {m[:3]: int(df[m].fillna(0).astype(float).sum())
                        for m in MESES if m in df.columns}
            if mes_data:
                st.bar_chart(pd.DataFrame({'Días': mes_data}), height=260)

            st.markdown("### Por categoría")
            if col_cat and col_meta and col_prog:
                for cat in df[col_cat].dropna().unique():
                    cd = df[df[col_cat] == cat]
                    m  = cd[col_meta].fillna(0).astype(float).sum()
                    p  = cd.apply(lambda r: min(float(r[col_prog] or 0),
                                                float(r[col_meta] or 0)), axis=1).sum()
                    pct = round(p/m*100, 0) if m > 0 else 0
                    st.caption(f"{cat}: {pct:.0f}% ({len(cd)} personas)")

    # ── COLABORADORES ──────────────────────────────────────────────────────────
    elif pagina == "👥 Colaboradores":
        st.markdown("## Colaboradores")

        c1,c2,c3,c4,c5 = st.columns(5)
        buscar     = c1.text_input("🔍 Nombre o legajo")
        areas_list = ['Todas'] + sorted(df[col_area].dropna().unique().tolist()) if col_area else ['Todas']
        cats_list  = ['Todas'] + sorted(df[col_cat].dropna().unique().tolist())  if col_cat  else ['Todas']
        ests_list  = ['Todos'] + sorted(df['Estado'].dropna().unique().tolist())  if 'Estado' in df.columns else ['Todos']
        jefes_list = ['Todos'] + sorted(df[col_jefe].dropna().unique().tolist()) if col_jefe else ['Todos']

        f_area = c2.selectbox("Área",      areas_list)
        f_cat  = c3.selectbox("Categoría", cats_list)
        f_est  = c4.selectbox("Estado",    ests_list)
        f_jefe = c5.selectbox("Jefe",      jefes_list)

        df_f = df.copy()
        if buscar and col_nom:
            mask = df_f[col_nom].astype(str).str.upper().str.contains(buscar.upper(), na=False)
            if 'Legajo' in df_f.columns:
                mask = mask | df_f['Legajo'].astype(str).str.contains(buscar, na=False)
            df_f = df_f[mask]
        if f_area != 'Todas' and col_area:
            df_f = df_f[df_f[col_area] == f_area]
        if f_cat  != 'Todas' and col_cat:
            df_f = df_f[df_f[col_cat]  == f_cat]
        if f_est  != 'Todos' and 'Estado' in df_f.columns:
            df_f = df_f[df_f['Estado'] == f_est]
        if f_jefe != 'Todos' and col_jefe:
            df_f = df_f[df_f[col_jefe] == f_jefe]

        st.caption(f"{len(df_f):,} de {len(df):,} registros")
        cols_t = [c for c in ['Legajo', col_nom, col_cat, col_area, col_jefe,
                               'Administrador', 'Vencidos_real', col_pend,
                               col_meta, col_prog, 'Pct_avance', col_dp,
                               'Fecha_limite', 'Estado']
                  if c and c in df_f.columns]
        show = df_f[cols_t].copy()
        if 'Estado' in show.columns:
            show['Estado'] = show['Estado'].apply(estado_emoji)
        st.dataframe(show, use_container_width=True, hide_index=True, height=500)
        st.download_button("⬇️ Descargar Excel", exportar_excel(df_f),
                           file_name=f"colaboradores_{date.today()}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── ALERTAS ────────────────────────────────────────────────────────────────
    elif pagina == "🔔 Alertas":
        st.markdown("## Centro de Alertas")

        df_venc    = df[df['Estado'] == 'VENCIDO'].copy()   if 'Estado' in df.columns else df.head(0)
        df_critico = df[df['Estado'] == 'CRITICO'].copy()   if 'Estado' in df.columns else df.head(0)
        df_riesgo  = df[df['Estado'] == 'EN_RIESGO'].copy() if 'Estado' in df.columns else df.head(0)

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("🔴 Vencidos",        len(df_venc))
        c2.metric("🟠 Críticos (≤30d)", len(df_critico))
        c3.metric("🟡 En riesgo (≤90d)",len(df_riesgo))
        jefes_af = set()
        for d in [df_venc, df_critico, df_riesgo]:
            if col_jefe in d.columns:
                jefes_af.update(d[col_jefe].dropna().unique())
        c4.metric("Jefes afectados", len(jefes_af))

        cols_a = [c for c in [col_nom, col_area, col_jefe, 'Administrador',
                               'Vencidos_real', col_pend, col_dp,
                               'Fecha_limite', 'Dias_alerta', 'Estado',
                               'Comentario_ind']
                  if c and c in df.columns]

        st.markdown("---")
        st.markdown("### 🔴 Días ya vencidos — indemnización en riesgo")
        if not df_venc.empty:
            show = df_venc[cols_a].copy()
            show['Estado'] = show['Estado'].apply(estado_emoji)
            show = show.sort_values('Vencidos_real', ascending=False)
            st.dataframe(show, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Descargar vencidos", exportar_excel(df_venc),
                               file_name=f"vencidos_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("✅ Sin días vencidos en tu vista")

        st.markdown("### 🟠 Críticos — vencen en menos de 30 días")
        if not df_critico.empty:
            show = df_critico[cols_a].copy()
            show['Estado'] = show['Estado'].apply(estado_emoji)
            show = show.sort_values('Dias_alerta')
            st.dataframe(show, use_container_width=True, hide_index=True)
        else:
            st.success("✅ Sin colaboradores críticos")

        st.markdown("### 🟡 En riesgo — vencen en 30–90 días")
        if not df_riesgo.empty:
            show = df_riesgo[cols_a].copy()
            show['Estado'] = show['Estado'].apply(estado_emoji)
            show = show.sort_values('Dias_alerta')
            st.dataframe(show, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Descargar en riesgo", exportar_excel(df_riesgo),
                               file_name=f"riesgo_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("✅ Sin colaboradores en riesgo")

    # ── RESUMEN GERENCIAS ──────────────────────────────────────────────────────
    elif pagina == "📋 Resumen gerencias":
        st.markdown("## Resumen Ejecutivo por Gerencia")
        group_col = col_ger if col_ger else col_jefe
        if group_col:
            rows = []
            for grp_val in sorted(df[group_col].dropna().unique()):
                gd = df[df[group_col] == grp_val]
                meta = gd[col_meta].fillna(0).astype(float).sum() if col_meta else 0
                prog = gd.apply(lambda r: min(float(r[col_prog] or 0),
                                              float(r[col_meta] or 0))
                                if col_meta and col_prog else 0, axis=1).sum()
                pct  = round(prog/meta*100, 1) if meta > 0 else 0
                rows.append({
                    group_col:     grp_val,
                    'HC':          len(gd),
                    'Vencidos':    int((gd['Vencidos_real'] > 0).sum()) if 'Vencidos_real' in gd.columns else 0,
                    'Críticos':    int((gd['Estado'] == 'CRITICO').sum()) if 'Estado' in gd.columns else 0,
                    'En riesgo':   int((gd['Estado'] == 'EN_RIESGO').sum()) if 'Estado' in gd.columns else 0,
                    'Meta':        int(meta),
                    'Programado':  int(prog),
                    '% Avance':    pct,
                    'Días x prog': int(gd[col_dp].fillna(0).astype(float).sum()) if col_dp else 0,
                })
            grp_df = pd.DataFrame(rows)
            st.dataframe(grp_df, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Descargar resumen", exportar_excel(grp_df),
                               file_name=f"resumen_{date.today()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("No se encontró columna de agrupación.")

    # ── HISTORIAL VISMA ────────────────────────────────────────────────────────
    elif pagina == "📂 Historial Visma":
        st.markdown("## Historial de Vacaciones — Visma")
        st.caption("Registro histórico completo de todos los periodos")

        hist = cargar_historial_visma()
        if hist.empty:
            st.warning("Sube el archivo Vacaciones_-_Dias_solicitados__28_.xlsx al repositorio para ver el historial.")
            return

        # Filtrar por legajos de la vista del usuario
        if 'Legajo' in df.columns:
            legajos_ok = df['Legajo'].astype(str).unique().tolist()
            hist_f = hist[hist['Legajo'].isin(legajos_ok)].copy()
        else:
            hist_f = hist.copy()

        # Unir nombre
        if col_nom and 'Legajo' in df.columns:
            ln = df[['Legajo', col_nom]].drop_duplicates()
            ln['Legajo'] = ln['Legajo'].astype(str)
            hist_f = hist_f.merge(ln, on='Legajo', how='left')

        c1,c2,c3,c4 = st.columns(4)
        buscar_l = c1.text_input("🔍 Legajo o nombre")
        anios    = sorted(hist_f['Fecha_desde'].dt.year.dropna().astype(int).unique().tolist(), reverse=True)
        anio_s   = c2.selectbox("Año", ['Todos'] + anios)
        ests_h   = ['Todos'] + sorted(hist_f['Estado_ausencia'].dropna().unique().tolist())
        est_s    = c3.selectbox("Estado", ests_h)
        periodos = ['Todos'] + sorted(hist_f['Periodo'].dropna().unique().tolist(), reverse=True)
        per_s    = c4.selectbox("Periodo", periodos)

        df_h = hist_f.copy()
        if buscar_l:
            mask = df_h['Legajo'].astype(str).str.contains(buscar_l, na=False)
            if col_nom and col_nom in df_h.columns:
                mask = mask | df_h[col_nom].astype(str).str.upper().str.contains(buscar_l.upper(), na=False)
            df_h = df_h[mask]
        if anio_s != 'Todos':
            df_h = df_h[df_h['Fecha_desde'].dt.year == int(anio_s)]
        if est_s != 'Todos':
            df_h = df_h[df_h['Estado_ausencia'] == est_s]
        if per_s != 'Todos':
            df_h = df_h[df_h['Periodo'] == per_s]

        st.caption(f"{len(df_h):,} registros — {df_h['Cant_dias'].sum():.0f} días totales")
        st.dataframe(df_h.sort_values('Fecha_desde', ascending=False),
                     use_container_width=True, hide_index=True, height=500)
        st.download_button("⬇️ Descargar historial", exportar_excel(df_h),
                           file_name=f"historial_visma_{date.today()}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── CARGAR ARCHIVOS ────────────────────────────────────────────────────────
    elif pagina == "⬆️ Cargar archivos":
        st.markdown("## Cargar archivos del mes")
        if role not in ['Gerente','RRHH']:
            st.warning("Solo RRHH y Gerentes pueden cargar archivos.")
            return
        st.info("Sube los archivos de Visma exportados del mes. El consolidado se actualiza automáticamente.")
        c1,c2 = st.columns(2)
        with c1:
            st.file_uploader("1. Vacaciones Visma",  type=['xlsx'], key='vac')
            st.file_uploader("2. Empleados Visma",   type=['xlsx'], key='emp')
        with c2:
            st.file_uploader("3. Atributos Visma",   type=['xlsx'], key='atr')
            st.file_uploader("4. Altas y Bajas",     type=['xlsx'], key='ab')
        st.markdown("---")
        st.file_uploader("Meta anual (solo si hay cambios)", type=['xlsx'], key='meta')
        st.file_uploader("Jerarquía (solo si hay rotación)", type=['xlsx'], key='jer')
        st.markdown("---")
        st.markdown("### Historial de cargas")
        st.dataframe(pd.DataFrame([
            {'Archivo':'META_2026_-_Abril.xlsx','Mes':'Abril 2026','Registros':'2,144','Estado':'✅ Activo'},
        ]), use_container_width=True, hide_index=True)

if __name__ == '__main__':
    main()
