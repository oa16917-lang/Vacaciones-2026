"""
vacaciones.py — Sistema de Gestión de Vacaciones 2026
Streamlit Community Cloud — gratis
"""
import streamlit as st
import pandas as pd
import json, re, calendar
from datetime import date, datetime, timedelta
from io import BytesIO

st.set_page_config(page_title="Vacaciones 2026", page_icon="📅", layout="wide",
                   initial_sidebar_state="expanded")
st.markdown("""<style>
.main{background:#f5f4f1}.block-container{padding-top:1.5rem}
div[data-testid="metric-container"]{background:white;border:1px solid #e2e0d8;border-radius:10px;padding:12px 16px}
.cal-day{border:1px solid #e2e0d8;border-radius:6px;padding:4px 6px;min-height:60px;background:white;font-size:11px}
.cal-day-num{font-weight:600;color:#6b6860;font-size:12px;margin-bottom:3px}
.cal-person{background:#e5edf8;color:#1a3a6b;border-radius:3px;padding:1px 4px;margin:1px 0;font-size:10px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.cal-person-end{background:#e8f5ee;color:#2d6a3f;border-radius:3px;padding:1px 4px;margin:1px 0;font-size:10px}
.weekend{background:#f8f7f4!important}
</style>""", unsafe_allow_html=True)

MESES     = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
             'JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']
MES_NAMES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
             'Julio','Agosto','Setiembre','Octubre','Noviembre','Diciembre']
MES_NUM   = {m:i+1 for i,m in enumerate(MESES)}

# ─── AUTH ──────────────────────────────────────────────────────────────────────
def check_auth():
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True
    st.markdown("## 📅 Gestión de Vacaciones 2026")
    st.markdown("---")
    _, col, _ = st.columns([1,1.2,1])
    with col:
        st.markdown("### Iniciar sesión")
        usr = st.text_input("Usuario", placeholder="correo@empresa.com")
        pwd = st.text_input("Contraseña", type="password")
        if st.button("Ingresar", use_container_width=True, type="primary"):
            users = st.secrets.get("usuarios", {})
            if usr in users and users[usr]["password"] == pwd:
                st.session_state.update(authenticated=True,
                    user_name=users[usr]["nombre"], user_email=usr)
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")
    return False

# ─── CARGA DE DATOS ────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def cargar_meta():
    rename = {'Apellidos y Nombres':'Nombre','Area':'AREA','Sede':'SEDE',
               'Meta SE':'Meta2026','Programación':'Prog_meta',
               'Días Pendientes de programación':'Dias_x_prog',
               'COMENTARIO PARA EVITAR INDEMNIZACION':'Comentario_ind',
               'COMENTARIOS PARA CUMPLIMIENTO META 2026':'Comentario_meta'}
    for fn, sh in [('CONSOLIDADO_GENERADO.xlsx',None),('META_2026_-_Abril.xlsx','Consolidado')]:
        try:
            df = pd.read_excel(fn, sheet_name=sh)
            if len(df):
                df = df.rename(columns={k:v for k,v in rename.items() if k in df.columns})
                df['Legajo'] = df['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
                col_area = next((c for c in ['AREA','Area'] if c in df.columns), None)
                if col_area:
                    df = df[df[col_area].notna() & (df[col_area].astype(str).str.strip()!='')]
                return df
        except: continue
    return pd.DataFrame()

@st.cache_data(ttl=3600)
def cargar_visma():
    try:
        v = pd.read_excel('Vacaciones_-_Dias_solicitados__28_.xlsx', header=2)
        v.columns = ['Legajo','_','Nombre_v','Estado','Fecha_desde','Fecha_hasta',
                     'Cant_dias','_2','Tipo_dia','Periodo','Origen','Estado_aus','Anticipo']
        v = v[v['Legajo'].notna() & (v['Legajo']!='Legajo')].copy()
        v['Legajo']      = v['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
        v['Fecha_desde'] = pd.to_datetime(v['Fecha_desde'], dayfirst=True, errors='coerce')
        v['Fecha_hasta'] = pd.to_datetime(v['Fecha_hasta'], dayfirst=True, errors='coerce')
        v['Cant_dias']   = pd.to_numeric(v['Cant_dias'], errors='coerce').fillna(0)
        return v[v['Fecha_desde'].notna()].copy()
    except: return pd.DataFrame()

@st.cache_data(ttl=86400)
def cargar_jerarquia():
    try:
        with open('person_access.json', encoding='utf-8') as f: return json.load(f)
    except: return {}

def get_ignorados():
    return st.session_state.get('ignorados', set())

def ignorar_legajo(legajo):
    ig = st.session_state.get('ignorados', set())
    ig.add(str(legajo))
    st.session_state['ignorados'] = ig

def restaurar_legajo(legajo):
    ig = st.session_state.get('ignorados', set())
    ig.discard(str(legajo))
    st.session_state['ignorados'] = ig

# ─── LÓGICA ────────────────────────────────────────────────────────────────────
def fecha_limite(comentario):
    if not comentario or str(comentario).strip() in ['-','nan','None','']: return None
    m = re.search(r'ANTES DEL (\d{2}/\d{2}/\d{4})', str(comentario).upper())
    if m:
        try: return datetime.strptime(m.group(1),'%d/%m/%Y').date()
        except: return None
    return None

def dias_gozados_por_mes_visma(visma_2026, legajo):
    """Suma días Visma por mes para un legajo en 2026"""
    sub = visma_2026[visma_2026['Legajo']==legajo]
    result = {}
    for m in MESES: result[m] = 0
    for _, r in sub.iterrows():
        mes_num = r['Fecha_desde'].month
        result[MESES[mes_num-1]] += r['Cant_dias']
    return result

def construir_consolidado(df_meta, df_visma):
    """Une meta + Visma y calcula estados reales"""
    hoy = date.today()
    ignorados = get_ignorados()

    # Visma solo 2026, aprobadas+pendientes
    if not df_visma.empty:
        v2026 = df_visma[
            (df_visma['Fecha_desde'].dt.year == 2026) &
            (df_visma['Estado_aus'].isin(['Aprobada','Pendiente']))
        ].copy()
        # Pivot días por mes por legajo
        v2026['mes'] = v2026['Fecha_desde'].dt.month
        pivot = v2026.pivot_table(index='Legajo', columns='mes',
                                   values='Cant_dias', aggfunc='sum', fill_value=0).reset_index()
        for i in range(1,13):
            if i not in pivot.columns: pivot[i] = 0
        pivot.columns = ['Legajo'] + [MESES[i-1] for i in range(1,13)]
        pivot['Prog_visma'] = pivot[MESES].sum(axis=1)
        df = df_meta.merge(pivot, on='Legajo', how='left')
        for m in MESES:
            if m not in df.columns: df[m] = 0
            else: df[m] = df[m].fillna(0)
        df['Prog_visma'] = df['Prog_visma'].fillna(0)
    else:
        df = df_meta.copy()
        for m in MESES:
            if m not in df.columns: df[m] = 0
        df['Prog_visma'] = df[[m for m in MESES if m in df.columns]].sum(axis=1)

    col_meta_val = next((c for c in ['Meta2026','Meta SE'] if c in df.columns), None)

    estados, vencidos_r, fechas_lim, dias_rest_list, pct_list = [], [], [], [], []
    for _, row in df.iterrows():
        leg       = str(row['Legajo'])
        comentario= str(row.get('Comentario_ind','') or '')
        meta      = float(row[col_meta_val] or 0) if col_meta_val else 0
        prog      = float(row.get('Prog_visma', 0) or 0)
        pendientes= float(row.get('Pendientes', 0) or 0)
        fl        = fecha_limite(comentario)

        # % avance capeado
        pct = round(min(prog, meta) / meta * 100, 1) if meta > 0 else 0

        # Vencidos reales
        venc = 0
        if fl and fl < hoy and pendientes > 0 and leg not in ignorados:
            md = re.search(r'DEBE GOZAR (\d+)', comentario.upper())
            dias_debia = int(md.group(1)) if md else int(pendientes)
            if prog < dias_debia:
                venc = max(0, dias_debia - prog)

        # Estado
        dias_r = (fl - hoy).days if fl else 999
        if leg in ignorados:
            estado = 'IGNORADO'
        elif venc > 0:
            estado = 'VENCIDO'
        elif fl and dias_r <= 30 and pendientes > 0:
            estado = 'CRITICO'
        elif fl and dias_r <= 90 and pendientes > 0:
            estado = 'EN_RIESGO'
        elif meta > 0 and prog >= meta:
            estado = 'CUMPLIDO'
        elif meta == 0 and pendientes == 0:
            estado = 'SIN_SALDO'
        else:
            estado = 'AL_DIA'

        estados.append(estado)
        vencidos_r.append(round(venc,1))
        fechas_lim.append(str(fl) if fl else '')
        dias_rest_list.append(dias_r)
        pct_list.append(pct)

    df = df.copy()
    df['Estado']       = estados
    df['Vencidos_real']= vencidos_r
    df['Fecha_limite'] = fechas_lim
    df['Dias_restantes']= dias_rest_list
    df['Pct_avance']   = pct_list
    return df

def filtrar_usuario(df, user_name, person_access):
    if user_name not in person_access: return df
    info = person_access[user_name]
    if info['role'] == 'Gerente' or not info.get('areas'): return df
    col = next((c for c in ['AREA','Area'] if c in df.columns), None)
    if col: return df[df[col].isin(info['areas'])].copy()
    return df

def emoji_estado(e):
    return {'VENCIDO':'🔴 Vencido','CRITICO':'🟠 Crítico','EN_RIESGO':'🟡 En riesgo',
            'CUMPLIDO':'🟢 Cumplido','AL_DIA':'🟢 Al día','SIN_SALDO':'⚪ Sin saldo',
            'IGNORADO':'🔕 Ignorado'}.get(str(e), str(e))

def to_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w: df.to_excel(w, index=False)
    return buf.getvalue()

# ─── CALENDARIO ────────────────────────────────────────────────────────────────
def render_calendario(df_visma, df_user, mes_num, anio=2026):
    """Muestra un calendario con colaboradores que tienen vacaciones en ese mes"""
    if df_visma.empty:
        st.warning("Sube el archivo de Visma para ver el calendario.")
        return

    # Legajos permitidos para este usuario
    legajos_ok = set(df_user['Legajo'].astype(str).unique())
    col_nom    = next((c for c in ['Nombre','Apellidos y Nombres'] if c in df_user.columns), None)
    leg_nombre = {}
    if col_nom:
        for _, r in df_user[['Legajo', col_nom]].drop_duplicates().iterrows():
            leg_nombre[str(r['Legajo'])] = str(r[col_nom]).title()

    # Filtrar visma: mes seleccionado, legajos permitidos, aprobadas+pendientes
    mask = (
        (df_visma['Fecha_desde'].dt.year  == anio) &
        (df_visma['Fecha_desde'].dt.month == mes_num) &
        (df_visma['Legajo'].isin(legajos_ok)) &
        (df_visma['Estado_aus'].isin(['Aprobada','Pendiente']))
    )
    vac_mes = df_visma[mask].copy()

    # También incluir vacaciones que EMPEZARON antes pero TERMINAN en este mes
    mask2 = (
        (df_visma['Fecha_hasta'].notna()) &
        (df_visma['Fecha_hasta'].dt.year  == anio) &
        (df_visma['Fecha_hasta'].dt.month == mes_num) &
        (df_visma['Fecha_desde'].dt.month < mes_num) &
        (df_visma['Legajo'].isin(legajos_ok)) &
        (df_visma['Estado_aus'].isin(['Aprobada','Pendiente']))
    )
    vac_mes = pd.concat([vac_mes, df_visma[mask2]]).drop_duplicates()

    # Construir diccionario día → lista de nombres
    dias_personas = {d: [] for d in range(1, 32)}
    for _, r in vac_mes.iterrows():
        nombre = leg_nombre.get(str(r['Legajo']), str(r['Legajo']))
        fd = r['Fecha_desde']
        fh = r['Fecha_hasta'] if pd.notna(r['Fecha_hasta']) else fd
        # Marcar cada día del rango dentro del mes
        cur = fd.date() if hasattr(fd,'date') else fd
        end = fh.date() if hasattr(fh,'date') else fh
        while cur <= end:
            if cur.year == anio and cur.month == mes_num:
                dias_personas[cur.day].append(nombre)
            cur += timedelta(days=1)

    # Grilla HTML del calendario
    cal_mat = calendar.monthcalendar(anio, mes_num)
    dias_semana = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom']

    # Cabecera
    cols_cal = st.columns(7)
    for i, d in enumerate(dias_semana):
        cols_cal[i].markdown(f"<div style='text-align:center;font-weight:600;font-size:12px;"
                              f"color:#6b6860;padding:4px'>{d}</div>", unsafe_allow_html=True)

    for semana in cal_mat:
        cols_w = st.columns(7)
        for i, dia in enumerate(semana):
            with cols_w[i]:
                if dia == 0:
                    st.markdown("<div class='cal-day' style='background:#f8f7f4;border:none'></div>",
                                unsafe_allow_html=True)
                else:
                    personas = dias_personas.get(dia, [])
                    es_finde = i >= 5
                    bg = "#f8f7f4" if es_finde else "white"
                    html = f"<div class='cal-day' style='background:{bg}'>"
                    html += f"<div class='cal-day-num'>{dia}</div>"
                    for p in personas[:4]:  # max 4 nombres por día
                        color = "#e5edf8" if i < 5 else "#fff3d4"
                        text  = "#1a3a6b" if i < 5 else "#8a5a00"
                        html += (f"<div style='background:{color};color:{text};border-radius:3px;"
                                 f"padding:1px 4px;margin:1px 0;font-size:10px;"
                                 f"overflow:hidden;text-overflow:ellipsis;white-space:nowrap'>"
                                 f"{p.split()[0] if p else ''}</div>")
                    if len(personas) > 4:
                        html += f"<div style='font-size:9px;color:#9e9b94'>+{len(personas)-4} más</div>"
                    html += "</div>"
                    st.markdown(html, unsafe_allow_html=True)

    # Resumen del mes
    total_personas = len(set(n for ps in dias_personas.values() for n in ps))
    total_dias     = sum(len(ps) for ps in dias_personas.values())
    st.caption(f"📊 {total_personas} colaboradores con vacaciones — {len(vac_mes)} registros — {vac_mes['Cant_dias'].sum():.0f} días totales")

    # Tabla exportable
    if not vac_mes.empty:
        exp = vac_mes.copy()
        exp['Nombre'] = exp['Legajo'].map(leg_nombre)
        exp = exp[['Legajo','Nombre','Fecha_desde','Fecha_hasta','Cant_dias','Estado_aus','Periodo']]
        exp.columns  = ['Legajo','Nombre','Desde','Hasta','Días','Estado','Periodo']
        with st.expander(f"📋 Ver lista detallada — {MES_NAMES[mes_num-1]} {anio}"):
            st.dataframe(exp.sort_values('Desde'), use_container_width=True, hide_index=True)
        st.download_button(
            f"⬇️ Descargar programación {MES_NAMES[mes_num-1]}",
            to_excel(exp),
            file_name=f"programacion_{MES_NAMES[mes_num-1]}_{anio}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# ─── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    if not check_auth(): return

    user_name     = st.session_state.user_name
    person_access = cargar_jerarquia()
    df_meta       = cargar_meta()
    df_visma      = cargar_visma()

    if df_meta.empty:
        st.error("No se encontró el archivo de datos. Sube META_2026_-_Abril.xlsx al repositorio.")
        return

    # Construir consolidado con estados reales
    df_full = construir_consolidado(df_meta, df_visma)
    df      = filtrar_usuario(df_full, user_name, person_access)
    role    = person_access.get(user_name, {}).get('role', 'RRHH')

    col_meta = next((c for c in ['Meta2026','Meta SE'] if c in df.columns), None)
    col_pend = 'Pendientes' if 'Pendientes' in df.columns else None
    col_dp   = next((c for c in ['Dias_x_prog','Días Pendientes de programación'] if c in df.columns), None)
    col_area = next((c for c in ['AREA','Area'] if c in df.columns), None)
    col_cat  = next((c for c in ['Categoria','Categoría'] if c in df.columns), None)
    col_ger  = next((c for c in ['Gerente','Gerencia'] if c in df.columns), None)
    col_jefe = 'Jefe' if 'Jefe' in df.columns else None
    col_nom  = next((c for c in ['Nombre','Apellidos y Nombres'] if c in df.columns), None)

    with st.sidebar:
        st.markdown(f"### 👤 {user_name}")
        st.caption(f"Rol: {role}")
        st.caption(f"{len(df):,} colaboradores")
        st.markdown("---")
        pagina = st.radio("Navegación", [
            "📊 Dashboard", "👥 Colaboradores", "🔔 Alertas",
            "📅 Calendario", "📋 Resumen gerencias",
            "📂 Historial Visma", "⬆️ Cargar archivos",
        ])
        st.markdown("---")
        if st.button("Cerrar sesión"):
            st.session_state.authenticated = False; st.rerun()

    # ── DASHBOARD ──────────────────────────────────────────────────────────────
    if pagina == "📊 Dashboard":
        st.markdown(f"## Dashboard — {user_name}")
        meta_t  = df[col_meta].fillna(0).astype(float).sum() if col_meta else 0
        prog_t  = df['Prog_visma'].fillna(0).sum() if 'Prog_visma' in df.columns else 0
        prog_cap= df.apply(lambda r: min(float(r.get('Prog_visma',0) or 0),
                                          float(r[col_meta] or 0)) if col_meta else 0, axis=1).sum()
        pct     = round(prog_cap/meta_t*100,1) if meta_t > 0 else 0
        venc    = int((df['Vencidos_real']>0).sum()) if 'Vencidos_real' in df.columns else 0
        critico = int((df['Estado']=='CRITICO').sum()) if 'Estado' in df.columns else 0
        riesgo  = int((df['Estado']=='EN_RIESGO').sum()) if 'Estado' in df.columns else 0
        dp      = int(df[col_dp].fillna(0).astype(float).sum()) if col_dp else 0

        c1,c2,c3,c4,c5 = st.columns(5)
        c1.metric("Colaboradores",f"{len(df):,}")
        c2.metric("🔴 Vencidos",venc)
        c3.metric("🟠 Críticos ≤30d",critico)
        c4.metric("% Avance",f"{pct}%")
        c5.metric("Días por programar",f"{dp:,}")

        st.markdown("---")
        li, ri = st.columns([2,1])
        with li:
            st.markdown("### Requieren atención")
            if 'Estado' in df.columns:
                df_at = df[df['Estado'].isin(['VENCIDO','CRITICO','EN_RIESGO']) &
                           (df[col_pend].fillna(0).astype(float)>0 if col_pend else True)].copy()
                if not df_at.empty:
                    cols_s = [c for c in [col_nom,col_area,col_jefe,'Vencidos_real',
                                          col_pend,col_dp,'Fecha_limite','Estado']
                              if c and c in df_at.columns]
                    show = df_at[cols_s].head(15).copy()
                    show['Estado'] = show['Estado'].apply(emoji_estado)
                    st.dataframe(show, use_container_width=True, hide_index=True)
                else:
                    st.success("✅ Sin alertas activas")
        with ri:
            st.markdown("### Avance por mes")
            mes_data = {m[:3]:int(df[m].fillna(0).sum()) for m in MESES if m in df.columns}
            if mes_data: st.bar_chart(pd.DataFrame({'Días':mes_data}), height=220)

    # ── COLABORADORES ──────────────────────────────────────────────────────────
    elif pagina == "👥 Colaboradores":
        st.markdown("## Colaboradores")
        c1,c2,c3,c4,c5 = st.columns(5)
        buscar  = c1.text_input("🔍 Nombre o legajo")
        f_area  = c2.selectbox("Área",      ['Todas']+sorted(df[col_area].dropna().unique().tolist()) if col_area else ['Todas'])
        f_cat   = c3.selectbox("Categoría", ['Todas']+sorted(df[col_cat].dropna().unique().tolist())  if col_cat  else ['Todas'])
        f_est   = c4.selectbox("Estado",    ['Todos']+sorted(df['Estado'].dropna().unique().tolist())  if 'Estado' in df.columns else ['Todos'])
        f_jefe  = c5.selectbox("Jefe",      ['Todos']+sorted(df[col_jefe].dropna().unique().tolist()) if col_jefe else ['Todos'])

        df_f = df.copy()
        if buscar and col_nom:
            m = df_f[col_nom].astype(str).str.upper().str.contains(buscar.upper(),na=False)
            if 'Legajo' in df_f.columns: m |= df_f['Legajo'].astype(str).str.contains(buscar,na=False)
            df_f = df_f[m]
        if f_area!='Todas' and col_area: df_f = df_f[df_f[col_area]==f_area]
        if f_cat !='Todas' and col_cat:  df_f = df_f[df_f[col_cat]==f_cat]
        if f_est !='Todos' and 'Estado' in df_f.columns: df_f = df_f[df_f['Estado']==f_est]
        if f_jefe!='Todos' and col_jefe: df_f = df_f[df_f[col_jefe]==f_jefe]

        st.caption(f"{len(df_f):,} de {len(df):,} registros")
        cols_t = [c for c in ['Legajo',col_nom,col_cat,col_area,col_jefe,'Administrador',
                               'Vencidos_real',col_pend,col_meta,'Prog_visma','Pct_avance',
                               col_dp,'Fecha_limite','Estado']
                  if c and c in df_f.columns]
        show = df_f[cols_t].copy()
        if 'Estado' in show.columns: show['Estado'] = show['Estado'].apply(emoji_estado)
        st.dataframe(show, use_container_width=True, hide_index=True, height=500)
        st.download_button("⬇️ Descargar Excel", to_excel(df_f),
            file_name=f"colaboradores_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── ALERTAS ────────────────────────────────────────────────────────────────
    elif pagina == "🔔 Alertas":
        st.markdown("## Centro de Alertas")
        df_v = df[df['Estado']=='VENCIDO'].copy()   if 'Estado' in df.columns else df.head(0)
        df_c = df[df['Estado']=='CRITICO'].copy()   if 'Estado' in df.columns else df.head(0)
        df_r = df[df['Estado']=='EN_RIESGO'].copy() if 'Estado' in df.columns else df.head(0)
        df_i = df[df['Estado']=='IGNORADO'].copy()  if 'Estado' in df.columns else df.head(0)

        c1,c2,c3,c4 = st.columns(4)
        c1.metric("🔴 Vencidos",len(df_v))
        c2.metric("🟠 Críticos ≤30d",len(df_c))
        c3.metric("🟡 En riesgo ≤90d",len(df_r))
        c4.metric("🔕 Ignorados",len(df_i))

        cols_a = [c for c in [col_nom,col_area,col_jefe,'Vencidos_real',
                               col_pend,col_dp,'Fecha_limite','Dias_restantes','Estado','Comentario_ind']
                  if c and c in df.columns]

        st.markdown("---")
        st.markdown("### 🔴 Días vencidos — riesgo de indemnización")
        if not df_v.empty:
            show = df_v[cols_a].copy().sort_values('Vencidos_real',ascending=False)
            show['Estado'] = show['Estado'].apply(emoji_estado)
            st.dataframe(show, use_container_width=True, hide_index=True)
            if role == 'RRHH':
                st.markdown("**Marcar como ignorado** (RRHH — ya gestionado por fuera):")
                leg_ig = st.text_input("Legajo a ignorar", key="ig_input",
                                        placeholder="Escribe el legajo y presiona Enter")
                col_ig1, col_ig2 = st.columns(2)
                if col_ig1.button("🔕 Ignorar este legajo") and leg_ig:
                    ignorar_legajo(leg_ig); st.rerun()
                if col_ig2.button("🔔 Restaurar legajo ignorado") and leg_ig:
                    restaurar_legajo(leg_ig); st.rerun()
            st.download_button("⬇️ Descargar vencidos", to_excel(df_v),
                file_name=f"vencidos_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("✅ Sin días vencidos en tu vista")

        st.markdown("### 🟠 Críticos — vencen en menos de 30 días")
        if not df_c.empty:
            show = df_c[cols_a].copy().sort_values('Dias_restantes')
            show['Estado'] = show['Estado'].apply(emoji_estado)
            st.dataframe(show, use_container_width=True, hide_index=True)
        else:
            st.success("✅ Sin críticos")

        st.markdown("### 🟡 En riesgo — vencen en 30–90 días")
        if not df_r.empty:
            show = df_r[cols_a].copy().sort_values('Dias_restantes')
            show['Estado'] = show['Estado'].apply(emoji_estado)
            st.dataframe(show, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Descargar en riesgo", to_excel(df_r),
                file_name=f"riesgo_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("✅ Sin colaboradores en riesgo")

        if not df_i.empty and role == 'RRHH':
            st.markdown("### 🔕 Ignorados por RRHH")
            st.dataframe(df_i[[c for c in [col_nom,col_area,col_jefe,'Comentario_ind']
                                if c and c in df_i.columns]],
                         use_container_width=True, hide_index=True)

    # ── CALENDARIO ────────────────────────────────────────────────────────────
    elif pagina == "📅 Calendario":
        st.markdown("## Calendario de Vacaciones")
        c1, c2 = st.columns([1, 4])
        mes_sel  = c1.selectbox("Mes", MES_NAMES, index=date.today().month-1)
        mes_num  = MES_NAMES.index(mes_sel) + 1
        anio_sel = c1.selectbox("Año", [2025, 2026, 2027], index=1)

        # Filtros adicionales
        f_cal_jefe = c1.selectbox("Jefe",
            ['Todos'] + sorted(df[col_jefe].dropna().unique().tolist()) if col_jefe else ['Todos'])
        f_cal_area = c1.selectbox("Área",
            ['Todas'] + sorted(df[col_area].dropna().unique().tolist()) if col_area else ['Todas'])

        df_cal = df.copy()
        if f_cal_jefe != 'Todos' and col_jefe:
            df_cal = df_cal[df_cal[col_jefe] == f_cal_jefe]
        if f_cal_area != 'Todas' and col_area:
            df_cal = df_cal[df_cal[col_area] == f_cal_area]

        with c2:
            st.markdown(f"### {mes_sel} {anio_sel} — {len(df_cal):,} colaboradores en vista")
            render_calendario(df_visma, df_cal, mes_num, anio_sel)

    # ── RESUMEN GERENCIAS ──────────────────────────────────────────────────────
    elif pagina == "📋 Resumen gerencias":
        st.markdown("## Resumen Ejecutivo")
        grp_col = col_ger if col_ger else col_jefe
        if grp_col:
            rows = []
            for gv in sorted(df[grp_col].dropna().unique()):
                gd   = df[df[grp_col]==gv]
                meta = gd[col_meta].fillna(0).astype(float).sum() if col_meta else 0
                prog = gd['Prog_visma'].fillna(0).sum() if 'Prog_visma' in gd.columns else 0
                prog_cap = gd.apply(lambda r: min(float(r.get('Prog_visma',0) or 0),
                                                   float(r[col_meta] or 0)) if col_meta else 0, axis=1).sum()
                pct  = round(prog_cap/meta*100,1) if meta > 0 else 0
                rows.append({
                    grp_col: gv, 'HC': len(gd),
                    'Vencidos': int((gd['Vencidos_real']>0).sum()) if 'Vencidos_real' in gd.columns else 0,
                    'Críticos': int((gd['Estado']=='CRITICO').sum()) if 'Estado' in gd.columns else 0,
                    'En riesgo':int((gd['Estado']=='EN_RIESGO').sum()) if 'Estado' in gd.columns else 0,
                    'Meta': int(meta), 'Programado (Visma)': int(prog),
                    '% Avance': pct,
                    'Días x prog': int(gd[col_dp].fillna(0).astype(float).sum()) if col_dp else 0,
                })
            gdf = pd.DataFrame(rows)
            st.dataframe(gdf, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Descargar resumen", to_excel(gdf),
                file_name=f"resumen_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── HISTORIAL VISMA ────────────────────────────────────────────────────────
    elif pagina == "📂 Historial Visma":
        st.markdown("## Historial de Vacaciones Visma")
        if df_visma.empty:
            st.warning("Sube Vacaciones_-_Dias_solicitados__28_.xlsx al repositorio.")
            return
        leg_ok = set(df['Legajo'].astype(str).unique())
        hist = df_visma[df_visma['Legajo'].isin(leg_ok)].copy()
        if col_nom and 'Legajo' in df.columns:
            ln = df[['Legajo',col_nom]].drop_duplicates()
            ln['Legajo'] = ln['Legajo'].astype(str)
            hist = hist.merge(ln, on='Legajo', how='left')

        c1,c2,c3,c4 = st.columns(4)
        bus  = c1.text_input("🔍 Legajo o nombre")
        anio = c2.selectbox("Año",['Todos']+sorted(hist['Fecha_desde'].dt.year.dropna().astype(int).unique().tolist(),reverse=True))
        est  = c3.selectbox("Estado",['Todos']+sorted(hist['Estado_aus'].dropna().unique().tolist()))
        per  = c4.selectbox("Periodo",['Todos']+sorted(hist['Periodo'].dropna().unique().tolist(),reverse=True))

        hf = hist.copy()
        if bus:
            mk = hf['Legajo'].astype(str).str.contains(bus,na=False)
            if col_nom and col_nom in hf.columns:
                mk |= hf[col_nom].astype(str).str.upper().str.contains(bus.upper(),na=False)
            hf = hf[mk]
        if anio!='Todos': hf = hf[hf['Fecha_desde'].dt.year==int(anio)]
        if est !='Todos': hf = hf[hf['Estado_aus']==est]
        if per !='Todos': hf = hf[hf['Periodo']==per]

        st.caption(f"{len(hf):,} registros — {hf['Cant_dias'].sum():.0f} días")
        st.dataframe(hf.sort_values('Fecha_desde',ascending=False),
                     use_container_width=True, hide_index=True, height=500)
        st.download_button("⬇️ Descargar historial", to_excel(hf),
            file_name=f"historial_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── CARGAR ARCHIVOS ────────────────────────────────────────────────────────
    elif pagina == "⬆️ Cargar archivos":
        st.markdown("## Cargar archivos del mes")
        if role not in ['Gerente','RRHH']:
            st.warning("Solo RRHH y Gerentes pueden cargar archivos.")
            return
        st.info("Sube los archivos exportados de Visma. Al confirmar, el consolidado se actualiza.")
        c1,c2 = st.columns(2)
        with c1:
            st.file_uploader("1. Vacaciones Visma",type=['xlsx'],key='vac')
            st.file_uploader("2. Empleados Visma",type=['xlsx'],key='emp')
        with c2:
            st.file_uploader("3. Atributos Visma",type=['xlsx'],key='atr')
            st.file_uploader("4. Altas y Bajas",type=['xlsx'],key='ab')
        st.markdown("---")
        st.file_uploader("Meta anual (solo si hay cambios)",type=['xlsx'],key='meta')
        st.file_uploader("Jerarquía (solo si hay rotación)",type=['xlsx'],key='jer')
        st.dataframe(pd.DataFrame([
            {'Archivo':'META_2026_-_Abril.xlsx','Mes':'Abril 2026','Registros':'2,144','Estado':'✅ Activo'},
            {'Archivo':'Vacaciones_-_Dias_solicitados__28_.xlsx','Mes':'Histórico','Registros':'20,263','Estado':'✅ Activo'},
        ]), use_container_width=True, hide_index=True)

if __name__ == '__main__':
    main()
