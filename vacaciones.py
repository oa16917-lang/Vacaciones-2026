"""
vacaciones.py — Gestión de Vacaciones Apparka
"""
import streamlit as st
import pandas as pd
import json, re, calendar
from datetime import date, datetime, timedelta
from io import BytesIO

AZUL    = "#1B1462"
FUCSIA  = "#ED2579"
MORADO  = "#8A34B4"
AZUL_L  = "#E8E7F5"
GRIS    = "#F5F4F1"
BORDE   = "#E2E0D8"

st.set_page_config(page_title="Gestión de Vacaciones Apparka",
                   page_icon="📅", layout="wide",
                   initial_sidebar_state="expanded")

st.markdown(f"""<style>
.main{{background:{GRIS}}}
.block-container{{padding-top:1.5rem;padding-bottom:1rem}}
section[data-testid="stSidebar"]{{background:{AZUL}}}
section[data-testid="stSidebar"] *{{color:white!important}}
section[data-testid="stSidebar"] .stButton button{{background:{FUCSIA};color:white;border:none;border-radius:8px;width:100%}}
div[data-testid="metric-container"]{{background:white;border:1.5px solid {BORDE};border-radius:12px;padding:14px 18px}}
div[data-testid="metric-container"] [data-testid="metric-value"]{{color:{AZUL};font-size:26px;font-weight:700}}
h1,h2,h3{{color:{AZUL}}}
.stDownloadButton button{{background:{FUCSIA};color:white;border:none;border-radius:8px}}
</style>""", unsafe_allow_html=True)

MESES     = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
             'JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']
MES_NAMES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
             'Julio','Agosto','Setiembre','Octubre','Noviembre','Diciembre']

# ── Helpers numéricos ──────────────────────────────────────────────────────────
def safe_float(val, default=0.0):
    """Convierte cualquier valor a float de forma segura"""
    if val is None: return default
    if isinstance(val, (int, float)):
        import math
        return default if math.isnan(val) else float(val)
    try:
        s = str(val).strip()
        if s in ['nan','None','','-','NaN']: return default
        return float(s)
    except:
        return default

def col_sum(df, col):
    """Suma segura de una columna"""
    if col not in df.columns: return 0.0
    return float(df[col].apply(safe_float).sum())

def cap_sum(df, col_prog, col_meta):
    """Suma de min(prog, meta) por fila — para % avance capeado"""
    if not col_prog or not col_meta: return 0.0
    total = 0.0
    for _, r in df.iterrows():
        p = safe_float(r.get(col_prog))
        m = safe_float(r.get(col_meta))
        total += min(p, m)
    return total

# ── Auth ───────────────────────────────────────────────────────────────────────
def check_auth():
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if st.session_state.authenticated:
        return True
    st.markdown(f"""<div style='text-align:center;padding:40px 0 20px'>
      <div style='font-size:40px;font-weight:800;color:{AZUL}'>Apparka</div>
      <div style='font-size:17px;color:{MORADO};font-weight:500;margin-top:4px'>
        Gestión de Vacaciones</div></div>""", unsafe_allow_html=True)
    _, col, _ = st.columns([1,1.1,1])
    with col:
        st.markdown(f"""<div style='text-align:center;margin-bottom:16px'>
          <div style='font-size:15px;color:#888;font-weight:400'>Bienvenidos</div>
          <div style='font-size:20px;font-weight:700;color:{AZUL}'>Iniciar Sesión</div>
        </div>""", unsafe_allow_html=True)
        usr = st.text_input("Correo corporativo", placeholder="nombre@apparka.com")
        pwd = st.text_input("Contraseña", type="password")
        if st.button("Ingresar →", use_container_width=True, type="primary"):
            users = st.secrets.get("usuarios", {})
            if usr in users and users[usr]["password"] == pwd:
                st.session_state.update(authenticated=True,
                    user_name=users[usr]["nombre"], user_email=usr)
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")
    return False

# ── Carga ──────────────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def cargar_meta():
    for fn, sh in [('CONSOLIDADO_GENERADO.xlsx',None),
                   ('META_2026_-_Abril.xlsx','Consolidado')]:
        try:
            df = pd.read_excel(fn, sheet_name=sh)
            if len(df):
                # Renombrar META 2025 → Meta2026 (columna V, la meta real de días a gozar)
                # Hacerlo uno por uno para garantizar el orden correcto
                col_map = {}
                if 'META 2025' in df.columns:      col_map['META 2025'] = 'Meta2026'
                elif 'Meta SE' in df.columns:       col_map['Meta SE']   = 'Meta2026'
                if 'Apellidos y Nombres' in df.columns: col_map['Apellidos y Nombres'] = 'Nombre'
                if 'Area' in df.columns:            col_map['Area']  = 'AREA'
                if 'Sede' in df.columns:            col_map['Sede']  = 'SEDE'
                if 'Programación' in df.columns:    col_map['Programación'] = 'Prog_meta'
                if 'Días Pendientes de programación' in df.columns:
                    col_map['Días Pendientes de programación'] = 'Dias_x_prog'
                if 'COMENTARIO PARA EVITAR INDEMNIZACION' in df.columns:
                    col_map['COMENTARIO PARA EVITAR INDEMNIZACION'] = 'Comentario_ind'
                if 'COMENTARIOS PARA CUMPLIMIENTO META 2026' in df.columns:
                    col_map['COMENTARIOS PARA CUMPLIMIENTO META 2026'] = 'Comentario_meta'
                df = df.rename(columns=col_map)
                df['Legajo'] = df['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
                ca = next((c for c in ['AREA','Area'] if c in df.columns), None)
                if ca:
                    df = df[df[ca].notna() & (df[ca].astype(str).str.strip()!='')]
                return df
        except:
            continue
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
    except:
        return pd.DataFrame()

@st.cache_data(ttl=86400)
def cargar_jerarquia():
    try:
        with open('person_access.json', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

def get_ignorados():
    return st.session_state.get('ignorados', set())

def ignorar(leg):
    ig = st.session_state.get('ignorados', set())
    ig.add(str(leg))
    st.session_state['ignorados'] = ig

def restaurar(leg):
    ig = st.session_state.get('ignorados', set())
    ig.discard(str(leg))
    st.session_state['ignorados'] = ig

# ── Lógica vacaciones ──────────────────────────────────────────────────────────
def fecha_limite(comentario):
    if not comentario or str(comentario).strip() in ['-','nan','None','']: return None
    m = re.search(r'ANTES DEL (\d{2}/\d{2}/\d{4})', str(comentario).upper())
    if m:
        try: return datetime.strptime(m.group(1),'%d/%m/%Y').date()
        except: pass
    return None

def construir_consolidado(df_meta, df_visma):
    hoy       = date.today()
    ignorados = get_ignorados()

    if not df_visma.empty:
        v2026 = df_visma[
            (df_visma['Fecha_desde'].dt.year==2026) &
            (df_visma['Estado_aus'].isin(['Aprobada','Pendiente']))
        ].copy()
        v2026['mes'] = v2026['Fecha_desde'].dt.month
        pivot = v2026.pivot_table(index='Legajo', columns='mes',
                                   values='Cant_dias', aggfunc='sum',
                                   fill_value=0).reset_index()
        for i in range(1,13):
            if i not in pivot.columns: pivot[i] = 0
        pivot.columns = ['Legajo'] + [MESES[i-1] for i in range(1,13)]
        pivot['Prog_visma'] = pivot[MESES].sum(axis=1)
        df = df_meta.merge(pivot, on='Legajo', how='left')
        for m in MESES:
            if m not in df.columns: df[m] = 0.0
            else: df[m] = df[m].fillna(0)
        df['Prog_visma'] = df['Prog_visma'].fillna(0)
    else:
        df = df_meta.copy()
        for m in MESES:
            if m not in df.columns: df[m] = 0.0
        df['Prog_visma'] = df[[m for m in MESES if m in df.columns]].sum(axis=1)

    col_meta = next((c for c in ['Meta2026'] if c in df.columns and df[c].notna().any()), None)

    estados, vencidos_r, fechas_l, dias_rest, pcts = [], [], [], [], []
    for _, row in df.iterrows():
        leg  = str(row['Legajo'])
        com  = str(row.get('Comentario_ind','') or '')
        meta = safe_float(row.get(col_meta) if col_meta else None)
        prog = safe_float(row.get('Prog_visma'))
        pend = safe_float(row.get('Pendientes'))
        fl   = fecha_limite(com)

        pct    = round(min(prog, meta) / meta * 100, 1) if meta > 0 else 0
        dias_r = (fl - hoy).days if fl else 999

        venc = 0
        if fl and fl < hoy and pend > 0 and leg not in ignorados:
            md = re.search(r'DEBE GOZAR (\d+)', com.upper())
            dias_debia = int(md.group(1)) if md else int(pend)
            if prog < dias_debia:
                venc = max(0, dias_debia - prog)

        if leg in ignorados:               estado = 'IGNORADO'
        elif venc > 0:                     estado = 'VENCIDO'
        elif fl and dias_r <= 30 and pend > 0: estado = 'CRITICO'
        elif fl and dias_r <= 90 and pend > 0: estado = 'EN_RIESGO'
        elif meta > 0 and prog >= meta:    estado = 'CUMPLIDO'
        elif meta == 0 and pend == 0:      estado = 'SIN_SALDO'
        else:                              estado = 'AL_DIA'

        estados.append(estado); vencidos_r.append(round(venc,1))
        fechas_l.append(str(fl) if fl else '')
        dias_rest.append(dias_r); pcts.append(pct)

    df = df.copy()
    df['Estado']        = estados
    df['Vencidos_real'] = vencidos_r
    df['Fecha_limite']  = fechas_l
    df['Dias_restantes']= dias_rest
    df['Pct_avance']    = pcts
    return df

def filtrar_usuario(df, user_name, pa):
    if user_name not in pa: return df
    info = pa[user_name]
    if info['role']=='Gerente' or not info.get('areas'): return df
    col = next((c for c in ['AREA','Area'] if c in df.columns), None)
    if col: return df[df[col].isin(info['areas'])].copy()
    return df

def emo(e):
    return {'VENCIDO':'🔴 Vencido','CRITICO':'🟠 Crítico','EN_RIESGO':'🟡 En riesgo',
            'CUMPLIDO':'🟢 Cumplido','AL_DIA':'🟢 Al día','SIN_SALDO':'⚪ Sin saldo',
            'IGNORADO':'🔕 Ignorado'}.get(str(e), str(e))

def to_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w: df.to_excel(w, index=False)
    return buf.getvalue()

# ── Calendario ─────────────────────────────────────────────────────────────────
def render_calendario(df_visma, df_user, mes_num, anio=2026):
    if df_visma.empty:
        st.warning("Sube el archivo de Visma para ver el calendario.")
        return
    legajos_ok = set(df_user['Legajo'].astype(str).unique())
    col_nom    = next((c for c in ['Nombre','Apellidos y Nombres'] if c in df_user.columns), None)
    leg_nombre = {}
    if col_nom:
        for _, r in df_user[['Legajo',col_nom]].drop_duplicates().iterrows():
            leg_nombre[str(r['Legajo'])] = str(r[col_nom]).title()

    mask = ((df_visma['Fecha_desde'].dt.year==anio) &
            (df_visma['Fecha_desde'].dt.month==mes_num) &
            (df_visma['Legajo'].isin(legajos_ok)) &
            (df_visma['Estado_aus'].isin(['Aprobada','Pendiente'])))
    mask2= (df_visma['Fecha_hasta'].notna() &
            (df_visma['Fecha_hasta'].dt.year==anio) &
            (df_visma['Fecha_hasta'].dt.month==mes_num) &
            (df_visma['Fecha_desde'].dt.month<mes_num) &
            (df_visma['Legajo'].isin(legajos_ok)) &
            (df_visma['Estado_aus'].isin(['Aprobada','Pendiente'])))
    vac_mes = pd.concat([df_visma[mask], df_visma[mask2]]).drop_duplicates()

    dias_p = {d:[] for d in range(1,32)}
    for _, r in vac_mes.iterrows():
        nombre = leg_nombre.get(str(r['Legajo']), str(r['Legajo']))
        fd = r['Fecha_desde'].date() if hasattr(r['Fecha_desde'],'date') else r['Fecha_desde']
        fh = r['Fecha_hasta'].date() if pd.notna(r['Fecha_hasta']) and hasattr(r['Fecha_hasta'],'date') else fd
        cur = fd
        while cur <= fh:
            if cur.year==anio and cur.month==mes_num:
                dias_p[cur.day].append(nombre.split()[0] if nombre else '')
            cur += timedelta(days=1)

    dias_semana = ['Lun','Mar','Mié','Jue','Vie','Sáb','Dom']
    cols_h = st.columns(7)
    for i,d in enumerate(dias_semana):
        cols_h[i].markdown(f"<div style='text-align:center;font-weight:600;font-size:12px;"
                            f"color:{AZUL};padding:6px;background:{AZUL_L};border-radius:6px'>{d}</div>",
                            unsafe_allow_html=True)
    for semana in calendar.monthcalendar(anio, mes_num):
        cols_w = st.columns(7)
        for i,dia in enumerate(semana):
            with cols_w[i]:
                if dia==0:
                    st.markdown("<div style='min-height:70px'></div>", unsafe_allow_html=True)
                else:
                    personas = dias_p.get(dia,[])
                    bg  = "#f0eff8" if i>=5 else "white"
                    brd = MORADO if personas else BORDE
                    html = f"<div style='border:1.5px solid {brd};border-radius:8px;padding:5px 7px;min-height:70px;background:{bg}'>"
                    html += f"<div style='font-weight:700;color:{AZUL};font-size:13px;margin-bottom:3px'>{dia}</div>"
                    for p in personas[:3]:
                        html += f"<div style='background:{AZUL_L};color:{AZUL};border-radius:4px;padding:1px 5px;margin:2px 0;font-size:10px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap'>{p}</div>"
                    if len(personas)>3:
                        html += f"<div style='font-size:9px;color:{MORADO};font-weight:600'>+{len(personas)-3} más</div>"
                    html += "</div>"
                    st.markdown(html, unsafe_allow_html=True)

    total_p = len(set(n for ps in dias_p.values() for n in ps))
    st.caption(f"📊 {total_p} colaboradores — {vac_mes['Cant_dias'].sum():.0f} días totales")
    if not vac_mes.empty:
        exp = vac_mes.copy()
        exp['Nombre'] = exp['Legajo'].map(leg_nombre)
        exp = exp[['Legajo','Nombre','Fecha_desde','Fecha_hasta','Cant_dias','Estado_aus','Periodo']]
        exp.columns = ['Legajo','Nombre','Desde','Hasta','Días','Estado','Periodo']
        with st.expander(f"📋 Lista — {MES_NAMES[mes_num-1]} {anio}"):
            st.dataframe(exp.sort_values('Desde'), use_container_width=True, hide_index=True)
        st.download_button(f"⬇️ Descargar {MES_NAMES[mes_num-1]}", to_excel(exp),
            file_name=f"prog_{MES_NAMES[mes_num-1]}_{anio}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    if not check_auth(): return

    user_name = st.session_state.user_name
    pa        = cargar_jerarquia()
    df_meta   = cargar_meta()
    df_visma  = cargar_visma()

    if df_meta.empty:
        st.error("No se encontró el archivo de datos. Sube META_2026_-_Abril.xlsx al repositorio.")
        return

    df_full = construir_consolidado(df_meta, df_visma)
    df      = filtrar_usuario(df_full, user_name, pa)
    role    = pa.get(user_name, {}).get('role','RRHH')

    col_meta = next((c for c in ['Meta2026'] if c in df.columns and df[c].notna().any()), None)
    col_pend = 'Pendientes' if 'Pendientes' in df.columns else None
    col_dp   = next((c for c in ['Dias_x_prog','Días Pendientes de programación'] if c in df.columns), None)
    col_area = next((c for c in ['AREA','Area'] if c in df.columns), None)
    col_cat  = next((c for c in ['Categoria','Categoría'] if c in df.columns), None)
    col_ger  = next((c for c in ['Gerente','Gerencia'] if c in df.columns), None)
    col_jefe = 'Jefe' if 'Jefe' in df.columns else None
    col_nom  = next((c for c in ['Nombre','Apellidos y Nombres'] if c in df.columns), None)

    with st.sidebar:
        st.markdown(f"""<div style='text-align:center;padding:16px 0 8px'>
          <div style='font-size:22px;font-weight:800;color:white'>Apparka</div>
          <div style='font-size:11px;color:rgba(255,255,255,0.7);margin-top:2px'>Gestión de Vacaciones</div>
        </div><hr style='border-color:rgba(255,255,255,0.2)'>""", unsafe_allow_html=True)
        st.markdown(f"**{user_name}**")
        st.caption(f"Rol: {role} · {len(df):,} colaboradores")
        st.markdown(f"<hr style='border-color:rgba(255,255,255,0.2)'>", unsafe_allow_html=True)
        pagina = st.radio("", ["📊 Dashboard","👥 Colaboradores","🔔 Alertas",
            "📅 Calendario","📋 Resumen gerencias","📂 Historial Visma","⬆️ Cargar archivos"],
            label_visibility="collapsed")
        st.markdown(f"<hr style='border-color:rgba(255,255,255,0.2)'>", unsafe_allow_html=True)
        if st.button("Cerrar sesión"):
            st.session_state.authenticated = False; st.rerun()

    # ── DASHBOARD ──────────────────────────────────────────────────────────────
    if pagina == "📊 Dashboard":
        st.markdown("## Dashboard — Vacaciones")

        meta_t   = col_sum(df, col_meta)
        prog_cap = cap_sum(df, 'Prog_visma', col_meta)
        pct      = round(prog_cap / meta_t * 100, 1) if meta_t > 0 else 0
        venc_n   = int((df['Vencidos_real'] > 0).sum()) if 'Vencidos_real' in df.columns else 0
        dp       = int(col_sum(df, col_dp))

        c1,c2,c3,c4,c5 = st.columns(5)
        with c1:
            st.markdown(f"<div style='font-size:11px;color:#6b6860;text-align:center'>Meta 2026 (días)</div><div style='font-size:26px;font-weight:700;color:{AZUL};text-align:center'>{int(meta_t):,}</div>", unsafe_allow_html=True)
        with c2:
            st.markdown(f"<div style='font-size:11px;color:#6b6860;text-align:center'>N° Colaboradores</div><div style='font-size:26px;font-weight:700;color:{AZUL};text-align:center'>{len(df):,}</div>", unsafe_allow_html=True)
        with c3:
            st.markdown(f"<div style='font-size:11px;color:#6b6860;text-align:center'>N° Colab. vac. vencidas</div><div style='font-size:26px;font-weight:700;color:{FUCSIA};text-align:center'>{venc_n}</div>", unsafe_allow_html=True)
        with c4:
            st.markdown(f"<div style='font-size:11px;color:#6b6860;text-align:center'>% Avance meta 2026</div><div style='font-size:26px;font-weight:700;color:{AZUL};text-align:center'>{pct}%</div>", unsafe_allow_html=True)
        with c5:
            st.markdown(f"<div style='font-size:11px;color:#6b6860;text-align:center'>Días por programar</div><div style='font-size:26px;font-weight:700;color:{MORADO};text-align:center'>{dp:,}</div>", unsafe_allow_html=True)

        st.markdown("---")
        st.markdown("### Colaboradores que requieren atención")
        if 'Estado' in df.columns:
            cond = df['Estado'].isin(['VENCIDO','CRITICO','EN_RIESGO'])
            if col_pend:
                cond = cond & (df[col_pend].apply(safe_float) > 0)
            df_at = df[cond].copy()
            if not df_at.empty:
                df_at = df_at.sort_values('Vencidos_real', ascending=False)
                cols_s = [c for c in [col_nom,col_area,col_jefe,'Vencidos_real',
                                       col_pend,col_dp,'Fecha_limite','Estado']
                          if c and c in df_at.columns]
                show = df_at[cols_s].head(20).copy()
                show['Estado'] = show['Estado'].apply(emo)
                st.dataframe(show, use_container_width=True, hide_index=True)
            else:
                st.success("✅ Sin colaboradores con alertas activas")

    # ── COLABORADORES ──────────────────────────────────────────────────────────
    elif pagina == "👥 Colaboradores":
        st.markdown("## Colaboradores")
        c1,c2,c3,c4,c5 = st.columns(5)
        buscar = c1.text_input("🔍 Nombre o legajo")
        f_area = c2.selectbox("Área",['Todas']+sorted(df[col_area].dropna().unique().tolist()) if col_area else ['Todas'])
        f_cat  = c3.selectbox("Categoría",['Todas']+sorted(df[col_cat].dropna().unique().tolist()) if col_cat else ['Todas'])
        f_est  = c4.selectbox("Estado",['Todos']+sorted(df['Estado'].dropna().unique().tolist()) if 'Estado' in df.columns else ['Todos'])
        f_jefe = c5.selectbox("Jefe",['Todos']+sorted(df[col_jefe].dropna().unique().tolist()) if col_jefe else ['Todos'])

        df_f = df.copy()
        if buscar and col_nom:
            m = df_f[col_nom].astype(str).str.upper().str.contains(buscar.upper(),na=False)
            if 'Legajo' in df_f.columns: m |= df_f['Legajo'].astype(str).str.contains(buscar,na=False)
            df_f = df_f[m]
        if f_area!='Todas' and col_area: df_f=df_f[df_f[col_area]==f_area]
        if f_cat !='Todas' and col_cat:  df_f=df_f[df_f[col_cat]==f_cat]
        if f_est !='Todos' and 'Estado' in df_f.columns: df_f=df_f[df_f['Estado']==f_est]
        if f_jefe!='Todos' and col_jefe: df_f=df_f[df_f[col_jefe]==f_jefe]

        st.caption(f"{len(df_f):,} de {len(df):,} registros")
        cols_t = [c for c in ['Legajo',col_nom,col_cat,col_area,col_jefe,'Administrador',
                               'Vencidos_real',col_pend,col_meta,'Prog_visma','Pct_avance',
                               col_dp,'Fecha_limite','Estado'] if c and c in df_f.columns]
        show = df_f[cols_t].copy()
        if 'Estado' in show.columns: show['Estado'] = show['Estado'].apply(emo)
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
        c1.metric("🔴 Vencidos",          len(df_v))
        c2.metric("🟠 Críticos ≤30 días", len(df_c))
        c3.metric("🟡 En riesgo ≤90 días",len(df_r))
        c4.metric("🔕 Ignorados",         len(df_i))

        cols_a = [c for c in [col_nom,col_area,col_jefe,'Vencidos_real',col_pend,
                               col_dp,'Fecha_limite','Dias_restantes','Estado','Comentario_ind']
                  if c and c in df.columns]

        st.markdown("---")
        st.markdown("### 🔴 Días vencidos — riesgo de indemnización")
        if not df_v.empty:
            show = df_v[cols_a].copy().sort_values('Vencidos_real',ascending=False)
            show['Estado'] = show['Estado'].apply(emo)
            st.dataframe(show, use_container_width=True, hide_index=True)
            if role=='RRHH':
                st.markdown("**Marcar como ignorado** (ya gestionado por fuera):")
                ci1,ci2,ci3 = st.columns(3)
                leg_ig = ci1.text_input("Legajo",key="ig_in")
                if ci2.button("🔕 Ignorar") and leg_ig: ignorar(leg_ig); st.rerun()
                if ci3.button("🔔 Restaurar") and leg_ig: restaurar(leg_ig); st.rerun()
            st.download_button("⬇️ Descargar vencidos", to_excel(df_v),
                file_name=f"vencidos_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("✅ Sin días vencidos en tu vista")

        st.markdown("### 🟠 Críticos — vencen en menos de 30 días")
        if not df_c.empty:
            show = df_c[cols_a].copy().sort_values('Dias_restantes')
            show['Estado'] = show['Estado'].apply(emo)
            st.dataframe(show, use_container_width=True, hide_index=True)
        else:
            st.success("✅ Sin críticos")

        st.markdown("### 🟡 En riesgo — vencen en 30–90 días")
        if not df_r.empty:
            show = df_r[cols_a].copy().sort_values('Dias_restantes')
            show['Estado'] = show['Estado'].apply(emo)
            st.dataframe(show, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Descargar en riesgo", to_excel(df_r),
                file_name=f"riesgo_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("✅ Sin colaboradores en riesgo")

        if not df_i.empty and role=='RRHH':
            st.markdown("### 🔕 Ignorados por RRHH")
            cols_i = [c for c in [col_nom,col_area,col_jefe,'Comentario_ind'] if c and c in df_i.columns]
            st.dataframe(df_i[cols_i], use_container_width=True, hide_index=True)

    # ── CALENDARIO ─────────────────────────────────────────────────────────────
    elif pagina == "📅 Calendario":
        st.markdown("## Calendario de Vacaciones")
        c1,c2 = st.columns([1,4])
        with c1:
            mes_sel  = st.selectbox("Mes", MES_NAMES, index=date.today().month-1)
            anio_sel = st.selectbox("Año", [2025,2026,2027], index=1)
            f_jefe_c = st.selectbox("Jefe",['Todos']+sorted(df[col_jefe].dropna().unique().tolist()) if col_jefe else ['Todos'])
            f_area_c = st.selectbox("Área",['Todas']+sorted(df[col_area].dropna().unique().tolist()) if col_area else ['Todas'])
        df_cal = df.copy()
        if f_jefe_c!='Todos' and col_jefe: df_cal=df_cal[df_cal[col_jefe]==f_jefe_c]
        if f_area_c!='Todas' and col_area: df_cal=df_cal[df_cal[col_area]==f_area_c]
        with c2:
            mes_num = MES_NAMES.index(mes_sel)+1
            st.markdown(f"### {mes_sel} {anio_sel} — {len(df_cal):,} colaboradores")
            render_calendario(df_visma, df_cal, mes_num, anio_sel)

    # ── RESUMEN GERENCIAS ──────────────────────────────────────────────────────
    elif pagina == "📋 Resumen gerencias":
        st.markdown("## Resumen Ejecutivo por Gerencia")
        grp = col_ger if col_ger else col_jefe
        if grp:
            rows = []
            for gv in sorted(df[grp].dropna().unique()):
                gd      = df[df[grp]==gv]
                meta    = col_sum(gd, col_meta)
                prog    = col_sum(gd, 'Prog_visma')
                prog_c  = cap_sum(gd, 'Prog_visma', col_meta)
                pct     = round(prog_c/meta*100,1) if meta>0 else 0
                rows.append({
                    grp:             gv,
                    'HC':            len(gd),
                    'Vencidos':      int((gd['Vencidos_real']>0).sum()) if 'Vencidos_real' in gd.columns else 0,
                    'Críticos':      int((gd['Estado']=='CRITICO').sum()) if 'Estado' in gd.columns else 0,
                    'En riesgo':     int((gd['Estado']=='EN_RIESGO').sum()) if 'Estado' in gd.columns else 0,
                    'Meta 2026':     int(meta),
                    'Prog. Visma':   int(prog),
                    '% Avance':      pct,
                    'Días x prog.':  int(col_sum(gd, col_dp)),
                })
            gdf = pd.DataFrame(rows)
            st.dataframe(gdf, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Descargar resumen", to_excel(gdf),
                file_name=f"resumen_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── HISTORIAL VISMA ────────────────────────────────────────────────────────
    elif pagina == "📂 Historial Visma":
        st.markdown("## Historial de Vacaciones — Visma")
        if df_visma.empty:
            st.warning("Sube Vacaciones_-_Dias_solicitados__28_.xlsx al repositorio.")
            return
        leg_ok = set(df['Legajo'].astype(str).unique())
        hist   = df_visma[df_visma['Legajo'].isin(leg_ok)].copy()
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
        if anio!='Todos': hf=hf[hf['Fecha_desde'].dt.year==int(anio)]
        if est !='Todos': hf=hf[hf['Estado_aus']==est]
        if per !='Todos': hf=hf[hf['Periodo']==per]
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
        st.info("Sube los archivos de Visma del mes.")
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
            {'Archivo':'META_2026_-_Abril.xlsx','Mes':'Abril 2026','Estado':'✅ Activo'},
            {'Archivo':'Vacaciones_-_Dias_solicitados__28_.xlsx','Mes':'Histórico','Estado':'✅ Activo'},
        ]),use_container_width=True,hide_index=True)

if __name__ == '__main__':
    main()
