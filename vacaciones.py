"""
vacaciones.py - Gestión de Vacaciones Apparka
"""
import streamlit as st
import pandas as pd
import json, re, calendar, math
from datetime import date, datetime, timedelta
from io import BytesIO

AZUL    = "#1B1462"
FUCSIA  = "#ED2579"
MORADO  = "#8A34B4"
AZUL_L  = "#E8E7F5"
FUCSIA_L= "#FDE8F2"
MORADO_L= "#F3E8FB"
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
.stDownloadButton button{{background:{FUCSIA};color:white;border:none;border-radius:8px;padding:8px 16px}}
.stDataFrame thead tr th{{background-color:{AZUL}!important;color:white!important}}
</style>""", unsafe_allow_html=True)

MESES     = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
             'JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']
MES_NAMES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio',
             'Julio','Agosto','Setiembre','Octubre','Noviembre','Diciembre']

# ── Helpers ────────────────────────────────────────────────────────────────────
def safe_float(val, default=0.0):
    if val is None: return default
    if isinstance(val, (int, float)):
        return default if math.isnan(val) else float(val)
    try:
        s = str(val).strip()
        return default if s in ['nan','None','','-','NaN'] else float(s)
    except:
        return default

def col_sum(df, col):
    if not col or col not in df.columns: return 0.0
    return float(df[col].apply(safe_float).sum())

def cap_sum(df, col_prog, col_meta):
    if not col_prog or not col_meta: return 0.0
    total = 0.0
    for _, r in df.iterrows():
        p = safe_float(r.get(col_prog))
        m = safe_float(r.get(col_meta))
        total += min(p, m)
    return total

def fmt_num(n):
    return f"{int(n):,}".replace(",", ",")

def es_direccion(cargo):
    if not cargo or str(cargo).strip() in ['','nan','None']: return False
    c = str(cargo).upper().strip()
    if 'ASISTENTE' in c: return False
    return any(x in c for x in ['GERENTE','SUB GERENTE','SUBGERENTE'])

# ── Ignorados persistentes ─────────────────────────────────────────────────────
def get_ignorados():
    persistentes = set()
    try:
        with open('ignorados.json', encoding='utf-8') as f:
            persistentes = set(json.load(f))
    except:
        pass
    return persistentes | st.session_state.get('ignorados', set())

def ignorar(leg):
    ig = get_ignorados(); ig.add(str(leg))
    st.session_state['ignorados'] = ig
    try:
        with open('ignorados.json', 'w', encoding='utf-8') as f:
            json.dump(list(ig), f)
    except:
        pass

def restaurar(leg):
    ig = get_ignorados(); ig.discard(str(leg))
    st.session_state['ignorados'] = ig
    try:
        with open('ignorados.json', 'w', encoding='utf-8') as f:
            json.dump(list(ig), f)
    except:
        pass

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
          <div style='font-size:15px;color:#888'>Bienvenidos</div>
          <div style='font-size:20px;font-weight:700;color:{AZUL}'>Iniciar Sesión</div>
        </div>""", unsafe_allow_html=True)
        usr = st.text_input("Correo corporativo", placeholder="nombre@apparka.com")
        pwd = st.text_input("Contraseña", type="password")
        if st.button("Ingresar", use_container_width=True, type="primary"):
            users = st.secrets.get("usuarios", {})
            if usr in users and users[usr]["password"] == pwd:
                st.session_state.update(authenticated=True,
                    user_name=users[usr]["nombre"], user_email=usr)
                st.rerun()
            else:
                st.error("Usuario o contrasena incorrectos")
    return False

# ── Carga datos ────────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600)
def cargar_meta():
    for fn, sh in [('CONSOLIDADO_GENERADO.xlsx',None),
                   ('META_2026_-_Abril.xlsx','Consolidado')]:
        try:
            df = pd.read_excel(fn, sheet_name=sh)
            if not len(df): continue
            col_map = {}
            if 'META 2025'    in df.columns: col_map['META 2025']    = 'Meta2026'
            elif 'Meta SE'    in df.columns: col_map['Meta SE']      = 'Meta2026'
            if 'Apellidos y Nombres' in df.columns: col_map['Apellidos y Nombres'] = 'Nombre'
            if 'Area'         in df.columns: col_map['Area']         = 'AREA'
            if 'Sede'         in df.columns: col_map['Sede']         = 'SEDE'
            if 'Programacion' in df.columns: col_map['Programacion'] = 'Prog_meta'
            if 'Programacion' not in df.columns and 'Programación' in df.columns:
                col_map['Programación'] = 'Prog_meta'
            if 'Dias Pendientes de programacion' in df.columns:
                col_map['Dias Pendientes de programacion'] = 'Dias_x_prog'
            if 'Días Pendientes de programación' in df.columns:
                col_map['Días Pendientes de programación'] = 'Dias_x_prog'
            if 'COMENTARIO PARA EVITAR INDEMNIZACION' in df.columns:
                col_map['COMENTARIO PARA EVITAR INDEMNIZACION'] = 'Comentario_ind'
            if 'COMENTARIOS PARA CUMPLIMIENTO META 2026' in df.columns:
                col_map['COMENTARIOS PARA CUMPLIMIENTO META 2026'] = 'Comentario_meta'
            df = df.rename(columns=col_map)
            df['Legajo'] = df['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
            # Eliminar fila vacia sin nombre
            col_n = next((c for c in ['Nombre','Apellidos y Nombres'] if c in df.columns), None)
            if col_n:
                df = df[df[col_n].notna() & (df[col_n].astype(str).str.strip()!='')]
            # Marcar cesados por area vacia en META
            # (la deteccion por altas/bajas se hace en construir_consolidado)
            ca = next((c for c in ['AREA','Area'] if c in df.columns), None)
            df['es_cesado'] = False
            if ca:
                df['es_cesado'] = df[ca].isna() | (df[ca].astype(str).str.strip()=='')
            # Redondear Truncos hacia arriba
            if 'Truncos' in df.columns:
                df['Truncos'] = df['Truncos'].apply(lambda x: math.ceil(safe_float(x)))
            return df
        except Exception as e:
            continue
    return pd.DataFrame()

@st.cache_data(ttl=3600)
def cargar_visma():
    # Buscar el archivo con cualquiera de sus nombres posibles
    import os
    nombres_visma = [
        'vacaciones.xlsx',
        'Vacaciones_-_Dias_solicitados__28_.xlsx',
        'Vacaciones_Historial.xlsx',
        'Vacaciones_Dias_solicitados.xlsx',
        'vacaciones_historial.xlsx',
    ]
    archivo_visma = next((n for n in nombres_visma if os.path.exists(n)), None)
    if not archivo_visma:
        return pd.DataFrame()
    try:
        v = pd.read_excel(archivo_visma, header=2)
        v.columns = ['Legajo','_x','Apellidos y Nombre','Estado','Fecha desde',
                     'Fecha hasta','Cant dias','_2','Tipo dia','Periodo',
                     'Origen','Estado aus','Anticipo']
        v = v[v['Legajo'].notna() & (v['Legajo']!='Legajo')].copy()
        v['Legajo']      = v['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
        v['Fecha desde'] = pd.to_datetime(v['Fecha desde'], dayfirst=True, errors='coerce')
        v['Fecha hasta'] = pd.to_datetime(v['Fecha hasta'], dayfirst=True, errors='coerce')
        v['Cant dias']   = pd.to_numeric(v['Cant dias'], errors='coerce').fillna(0)
        return v[v['Fecha desde'].notna()].copy()
    except:
        return pd.DataFrame()

@st.cache_data(ttl=3600)
def cargar_altas_bajas():
    import os
    nombres_ab = [
        'Altas y bajas de colaboradores.xlsx',      # nombre exacto en GitHub
        'Altas_y_bajas_de_colaboradores.xlsx',
        'Altas_y_bajas_de_colaboradores__27_.xlsx',
        'altas_bajas.xlsx',
    ]
    archivo = next((n for n in nombres_ab if os.path.exists(n)), None)
    if not archivo:
        return pd.DataFrame()
    try:
        # Detectar la fila de encabezado buscando la columna "Legajo"
        # El archivo puede tener un titulo en fila 1, headers en fila 2
        header_row = 1  # default: fila 2 (0-indexed = 1)
        for h in [0, 1, 2]:
            try:
                test = pd.read_excel(archivo, header=h, nrows=1)
                cols_lower = [str(c).strip().lower() for c in test.columns]
                if 'legajo' in cols_lower:
                    header_row = h
                    break
            except:
                pass

        df = pd.read_excel(archivo, header=header_row)
        # Normalizar nombres de columna
        df.columns = [str(c).strip() for c in df.columns]

        # Mapear columnas por nombre (tolerante a variaciones)
        col_map = {}
        for c in df.columns:
            cl = c.lower()
            if cl == 'legajo':                                   col_map[c] = 'Legajo'
            elif 'apellido' in cl:                               col_map[c] = 'Apellidos'
            elif 'nombre' in cl and 'apellido' not in cl:        col_map[c] = 'Nombres'
            elif cl == 'estado':                                 col_map[c] = 'Estado'
            elif 'alta' in cl and 'fecha' in cl:                 col_map[c] = 'Fecha_alta'
            elif 'baja' in cl and 'fecha' in cl:                 col_map[c] = 'Fecha_baja'
            elif 'causa' in cl or 'motivo' in cl:                col_map[c] = 'Causa'
            elif 'vacacion' in cl:                               col_map[c] = 'Vacaciones'
            elif 'indemniz' in cl:                               col_map[c] = 'Indemnizacion'
        df = df.rename(columns=col_map)

        # Filtrar filas validas (Legajo numerico)
        if 'Legajo' not in df.columns:
            return pd.DataFrame()
        df = df[df['Legajo'].notna()].copy()
        df['Legajo'] = df['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
        df = df[df['Legajo'].str.match(r'^[0-9]+$')]  # solo legajos numericos

        # Parsear fechas
        if 'Fecha_baja' in df.columns:
            df['Fecha_baja'] = pd.to_datetime(df['Fecha_baja'], dayfirst=True, errors='coerce')
        else:
            df['Fecha_baja'] = pd.NaT
        if 'Fecha_alta' in df.columns:
            df['Fecha_alta'] = pd.to_datetime(df['Fecha_alta'], dayfirst=True, errors='coerce')
        else:
            df['Fecha_alta'] = pd.NaT
        if 'Estado' not in df.columns:
            df['Estado'] = ''

        return df
    except Exception as e:
        return pd.DataFrame()

@st.cache_data(ttl=86400)
def cargar_jerarquia():
    try:
        with open('person_access.json', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

# ── Logica vacaciones ──────────────────────────────────────────────────────────
def fecha_limite(comentario):
    if not comentario or str(comentario).strip() in ['-','nan','None','']: return None
    m = re.search(r'ANTES DEL (\d{2}/\d{2}/\d{4})', str(comentario).upper())
    if m:
        try: return datetime.strptime(m.group(1),'%d/%m/%Y').date()
        except: pass
    return None

def construir_consolidado(df_meta, df_visma, df_ab=None):
    hoy       = date.today()
    ignorados = get_ignorados()

    # Identificar cesados y reingresos desde altas/bajas
    cesados_2026   = set()  # cesaron en 2026 → meta al 100%
    reingresos_2026= set()  # reingresaron en 2026 → sin meta este año

    if df_ab is not None and not df_ab.empty:
        # Cesados en 2026: baja registrada en 2026, Estado Inactivo
        # Incluimos independientemente del campo Vacaciones por si viene con valor distinto
        bajas = df_ab[
            (df_ab['Fecha_baja'].dt.year == 2026) &
            (df_ab['Estado'] == 'Inactivo')
        ]
        cesados_2026 = set(bajas['Legajo'].unique())

        # Reingresos: alta en 2026 Y tenían registro previo (Estado Activo ahora)
        altas_2026 = df_ab[
            (df_ab['Fecha_alta'].dt.year == 2026) &
            (df_ab['Estado'] == 'Activo')
        ]
        # Solo los que tienen más de un registro = reingresantes
        multi = df_ab.groupby('Legajo').size()
        reingresos_2026 = set(altas_2026[
            altas_2026['Legajo'].isin(multi[multi > 1].index)
        ]['Legajo'].unique())

    df = df_meta.copy()

    # Marcar cesados en 2026 desde altas/bajas en el flag es_cesado
    # (además de los que ya tienen es_cesado=True por area vacia en META)
    if cesados_2026:
        df['es_cesado'] = df.apply(
            lambda r: bool(r.get('es_cesado', False)) or str(r['Legajo']) in cesados_2026,
            axis=1
        )

    if not df_visma.empty:
        # Usar Visma como fuente de verdad para días programados 2026
        v2026 = df_visma[
            (df_visma['Fecha desde'].dt.year==2026) &
            (df_visma['Estado aus'].isin(['Aprobada','Pendiente']))
        ].copy()
        v2026['mes'] = v2026['Fecha desde'].dt.month
        pivot = v2026.pivot_table(index='Legajo', columns='mes',
                                   values='Cant dias', aggfunc='sum',
                                   fill_value=0).reset_index()
        for i in range(1,13):
            if i not in pivot.columns: pivot[i] = 0
        pivot.columns = ['Legajo'] + [MESES[i-1] for i in range(1,13)]
        pivot['Prog_visma'] = pivot[MESES].sum(axis=1)
        df = df.merge(pivot, on='Legajo', how='left')
        for m in MESES:
            if m not in df.columns: df[m] = 0.0
            else: df[m] = df[m].fillna(0)
        df['Prog_visma'] = df['Prog_visma'].fillna(0)
    else:
        # Fallback al META si no hay Visma
        col_prog_meta = next((c for c in ['Prog_meta','Programacion','Programación'] if c in df.columns), None)
        if col_prog_meta:
            df['Prog_visma'] = df[col_prog_meta].apply(safe_float)
        else:
            for m in MESES:
                if m not in df.columns: df[m] = 0.0
            df['Prog_visma'] = df[[m for m in MESES if m in df.columns]].apply(
                lambda col: col.apply(safe_float)).sum(axis=1)
        for m in MESES:
            if m not in df.columns: df[m] = 0.0

    col_meta = next((c for c in ['Meta2026'] if c in df.columns), None)

    estados,vencidos_r,fechas_l,dias_rest,pcts = [],[],[],[],[]
    for _, row in df.iterrows():
        leg       = str(row['Legajo'])
        com       = str(row.get('Comentario_ind','') or '')
        meta      = safe_float(row.get(col_meta) if col_meta else None)
        prog      = safe_float(row.get('Prog_visma'))
        pend      = safe_float(row.get('Pendientes'))
        es_ces    = bool(row.get('es_cesado', False)) or leg in cesados_2026
        es_reingreso = leg in reingresos_2026
        cargo_v   = str(row.get('Cargo','') or '')
        es_dir    = es_direccion(cargo_v)
        fl        = fecha_limite(com)
        dias_r    = (fl - hoy).days if fl else 999
        pct       = round(min(prog, meta)/meta*100,1) if meta>0 else 0
        # dias_x: cuántos días aún debe gozar según el comentario vs lo que gozó en Visma
        md_x      = re.search(r'DEBE GOZAR (\d+)', com.upper())
        if md_x:
            debia   = int(md_x.group(1))
            dias_x  = max(0, debia - prog)  # lo que falta según Visma real
        else:
            dias_x  = max(0, meta - prog)   # fallback: meta - lo programado

        if es_reingreso:
            # Reingresó en 2026 — sin meta este año, no genera alertas
            estado='SIN_SALDO'; venc=0; pct=0
        elif es_ces:
            # Cesó en 2026 con liquidación — meta al 100%
            estado='CUMPLIDO'; venc=0
        elif leg in ignorados:
            estado='IGNORADO'; venc=0
        elif es_dir:
            venc=0
            if meta>0 and prog>=meta:    estado='CUMPLIDO'
            elif meta==0 and pend==0:    estado='SIN_SALDO'
            else:                        estado='AL_DIA'
        else:
            venc=0
            if fl and fl<hoy:
                md = re.search(r'DEBE GOZAR (\d+)', com.upper())
                if md:
                    dd = int(md.group(1))
                    # Solo vencido si prog (de Visma) es menor a días que debía gozar
                    if prog < dd: venc = max(0, dd - prog)
            # VENCIDO: solo cuando realmente tiene días sin gozar después de la fecha límite
            if venc>0:                              estado='VENCIDO'
            elif fl and dias_r<=30 and dias_x>0:   estado='CRITICO'
            elif fl and dias_r<=90 and dias_x>0:   estado='EN_RIESGO'
            elif meta>0 and prog>=meta:             estado='CUMPLIDO'
            elif meta==0 and pend==0:               estado='SIN_SALDO'
            else:                                   estado='AL_DIA'

        estados.append(estado); vencidos_r.append(round(venc,1))
        fechas_l.append(str(fl) if fl else '')
        dias_rest.append(dias_r); pcts.append(pct)

    df = df.copy()
    df['Estado']        = estados
    df['Vencidos_real'] = vencidos_r
    df['Fecha límite']  = fechas_l
    df['Dias restantes']= dias_rest
    df['Pct avance']    = pcts

    # Recalcular Dias_x_prog desde Visma: max(0, meta - prog_visma)
    col_meta_calc = next((c for c in ['Meta2026'] if c in df.columns), None)
    if col_meta_calc:
        df['Dias_x_prog'] = df.apply(
            lambda r: max(0, int(safe_float(r.get(col_meta_calc)) - safe_float(r.get('Prog_visma'))))
            if str(r.get('Estado','')) not in ['CUMPLIDO','SIN_SALDO'] else 0,
            axis=1
        )
    return df

def filtrar_usuario(df, user_name, pa):
    # Vista: solo activos con area
    col_a = next((c for c in ['AREA','Area'] if c in df.columns), None)
    if col_a:
        df_act = df[df[col_a].notna() & (df[col_a].astype(str).str.strip()!='')].copy()
    else:
        df_act = df.copy()
    if user_name not in pa: return df_act
    info = pa[user_name]
    if info['role']=='Gerente' or not info.get('areas'): return df_act
    if col_a: return df_act[df_act[col_a].isin(info['areas'])].copy()
    return df_act

def emo(e):
    return {'VENCIDO':'🔴 Vencido','CRITICO':'🟠 Crítico','EN_RIESGO':'🟡 En riesgo',
            'CUMPLIDO':'🟢 Cumplido','AL_DIA':'🟡 Pendiente','SIN_SALDO':'⚪ Sin saldo',
            'IGNORADO':'🔕 Ignorado'}.get(str(e), str(e))

def to_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w: df.to_excel(w, index=False)
    return buf.getvalue()

# ── KPI card HTML ──────────────────────────────────────────────────────────────
def kpi(label, value, color=None):
    c = color or AZUL
    return f"""<div style='background:white;border:1.5px solid {BORDE};border-radius:12px;
    padding:14px 18px;text-align:center'>
    <div style='font-size:11px;color:#6b6860;margin-bottom:4px'>{label}</div>
    <div style='font-size:26px;font-weight:700;color:{c}'>{value}</div></div>"""

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

    mask  = ((df_visma['Fecha desde'].dt.year==anio) &
             (df_visma['Fecha desde'].dt.month==mes_num) &
             (df_visma['Legajo'].isin(legajos_ok)) &
             (df_visma['Estado aus'].isin(['Aprobada','Pendiente'])))
    mask2 = (df_visma['Fecha hasta'].notna() &
             (df_visma['Fecha hasta'].dt.year==anio) &
             (df_visma['Fecha hasta'].dt.month==mes_num) &
             (df_visma['Fecha desde'].dt.month<mes_num) &
             (df_visma['Legajo'].isin(legajos_ok)) &
             (df_visma['Estado aus'].isin(['Aprobada','Pendiente'])))
    vac_mes = pd.concat([df_visma[mask], df_visma[mask2]]).drop_duplicates()

    dias_p = {d:[] for d in range(1,32)}
    for _, r in vac_mes.iterrows():
        nombre = leg_nombre.get(str(r['Legajo']), str(r['Legajo']))
        fd = r['Fecha desde'].date() if hasattr(r['Fecha desde'],'date') else r['Fecha desde']
        fh = r['Fecha hasta'].date() if pd.notna(r['Fecha hasta']) and hasattr(r['Fecha hasta'],'date') else fd
        cur = fd
        while cur <= fh:
            if cur.year==anio and cur.month==mes_num:
                dias_p[cur.day].append(nombre.split()[0] if nombre else '')
            cur += timedelta(days=1)

    dias_sem = ['Lun','Mar','Miérc','Jue','Vie','Sáb','Dom']
    cols_h   = st.columns(7)
    for i,d in enumerate(dias_sem):
        cols_h[i].markdown(
            f"<div style='text-align:center;font-weight:600;font-size:12px;"
            f"color:{AZUL};padding:6px;background:{AZUL_L};border-radius:6px'>{d}</div>",
            unsafe_allow_html=True)
    for semana in calendar.monthcalendar(anio, mes_num):
        cols_w = st.columns(7)
        for i,dia in enumerate(semana):
            with cols_w[i]:
                if dia==0:
                    st.markdown("<div style='min-height:70px'></div>",unsafe_allow_html=True)
                else:
                    personas = dias_p.get(dia,[])
                    bg  = "#f0eff8" if i>=5 else "white"
                    brd = FUCSIA if personas else BORDE
                    html = (f"<div style='border:1.5px solid {brd};border-radius:8px;"
                            f"padding:5px 7px;min-height:70px;background:{bg}'>")
                    html += f"<div style='font-weight:700;color:{AZUL};font-size:13px;margin-bottom:3px'>{dia}</div>"
                    for p in personas[:3]:
                        html += (f"<div style='background:{AZUL_L};color:{AZUL};border-radius:4px;"
                                 f"padding:1px 5px;margin:2px 0;font-size:10px;overflow:hidden;"
                                 f"text-overflow:ellipsis;white-space:nowrap'>{p}</div>")
                    if len(personas)>3:
                        html += f"<div style='font-size:9px;color:{FUCSIA};font-weight:600'>+{len(personas)-3} más</div>"
                    html += "</div>"
                    st.markdown(html, unsafe_allow_html=True)

    total_p = len(set(n for ps in dias_p.values() for n in ps))
    st.caption(f"📊 {total_p} colaboradores — {vac_mes['Cant dias'].sum():.0f} días totales")
    if not vac_mes.empty:
        exp = vac_mes.copy()
        exp['Nombre'] = exp['Legajo'].map(leg_nombre)
        exp['Fecha desde'] = exp['Fecha desde'].dt.date
        exp['Fecha hasta'] = exp['Fecha hasta'].apply(lambda x: x.date() if pd.notna(x) else '')
        exp = exp[['Legajo','Nombre','Fecha desde','Fecha hasta','Cant dias','Estado aus','Periodo']]
        exp.columns = ['Legajo','Nombre','Desde','Hasta','Dias','Estado','Periodo']
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
        st.error("No se encontro el archivo de datos. Sube META_2026_-_Abril.xlsx al repositorio.")
        return

    df_ab   = cargar_altas_bajas()
    df_full = construir_consolidado(df_meta, df_visma, df_ab)
    df      = filtrar_usuario(df_full, user_name, pa)
    role    = pa.get(user_name, {}).get('role','RRHH')

    col_meta = next((c for c in ['Meta2026'] if c in df.columns), None)
    col_pend = 'Pendientes' if 'Pendientes' in df.columns else None
    col_dp   = 'Dias_x_prog' if 'Dias_x_prog' in df.columns else None
    col_area = next((c for c in ['AREA','Area'] if c in df.columns), None)
    col_cat  = next((c for c in ['Categoria','Categoría'] if c in df.columns), None)
    col_ger  = next((c for c in ['Gerencia'] if c in df.columns), None)
    col_jefe = 'Jefe' if 'Jefe' in df.columns else None
    col_nom  = next((c for c in ['Nombre','Apellidos y Nombres'] if c in df.columns), None)

    with st.sidebar:
        st.markdown(f"""<div style='text-align:center;padding:16px 0 8px'>
          <div style='font-size:22px;font-weight:800;color:white'>Apparka</div>
          <div style='font-size:11px;color:rgba(255,255,255,0.7);margin-top:2px'>
            Gestión de Vacaciones</div>
        </div><hr style='border-color:rgba(255,255,255,0.2);margin:0 0 8px'>""",
            unsafe_allow_html=True)
        st.markdown(f"**{user_name}**")
        st.caption(f"Rol: {role}")
        st.markdown("<hr style='border-color:rgba(255,255,255,0.2)'>",unsafe_allow_html=True)
        pagina = st.radio("", [
            "📊 Dashboard","👥 Colaboradores por Área","🔔 Centro de Alertas",
            "📅 Calendario","📋 Resumen Ejecutivo","📂 Historial de Vacaciones",
        ], label_visibility="collapsed")
        st.markdown("<hr style='border-color:rgba(255,255,255,0.2)'>",unsafe_allow_html=True)
        if st.button("Cerrar sesión"):
            st.session_state.authenticated = False; st.rerun()

    # ── DASHBOARD ──────────────────────────────────────────────────────────────
    if pagina == "📊 Dashboard":
        st.markdown("## Dashboard — Vacaciones")

        # KPIs scoped al usuario — si es RRHH ve empresa completa, si es Jefe/Admin ve su area
        is_rrhh = (role == 'RRHH')
        df_scope = df_full if is_rrhh else df

        col_meta_scope = next((c for c in ['Meta2026'] if c in df_scope.columns), None)
        meta_t   = col_sum(df_scope, col_meta_scope)

        # N colaboradores: para RRHH todos (incluyendo cesados), para otros solo activos
        if is_rrhh:
            n_colab = len(df_scope)
        else:
            n_colab = int((~df_scope['es_cesado']).sum()) if 'es_cesado' in df_scope.columns else len(df_scope)

        # Prog capeada: activos real + cesados al 100%
        df_act   = df_scope[~df_scope['es_cesado']] if 'es_cesado' in df_scope.columns else df_scope
        df_ces   = df_scope[df_scope['es_cesado']]  if 'es_cesado' in df_scope.columns else pd.DataFrame()
        prog_act = cap_sum(df_act, 'Prog_visma', col_meta_scope)
        prog_ces = col_sum(df_ces, col_meta_scope)
        prog_cap = prog_act + prog_ces
        pct      = round(prog_cap/meta_t*100,1) if meta_t>0 else 0
        # Vencidos: mismo criterio que Centro de Alertas — Estado == VENCIDO
        venc_n   = int((df['Estado']=='VENCIDO').sum()) if 'Estado' in df.columns else 0
        dp       = int(col_sum(df, col_dp))

        c1,c2,c3,c4,c5 = st.columns(5)
        c1.markdown(kpi("Meta 2026 (días)", fmt_num(meta_t)), unsafe_allow_html=True)
        c2.markdown(kpi("N° Colaboradores", fmt_num(n_colab)), unsafe_allow_html=True)
        c3.markdown(kpi("N° Colab. con vac. vencidas", venc_n, FUCSIA), unsafe_allow_html=True)
        c4.markdown(kpi("% Avance Meta 2026", f"{pct}%"), unsafe_allow_html=True)
        c5.markdown(kpi("Días por programar", fmt_num(dp), MORADO), unsafe_allow_html=True)

        st.markdown("---")
        col_l, col_r = st.columns([1.2,1])

        with col_l:
            st.markdown("### Top 10 áreas con mayor días pendientes por programar")
            if col_area and col_dp:
                top10 = (df[df[col_dp].apply(safe_float)>0]
                         .groupby(col_area)[col_dp]
                         .apply(lambda x: x.apply(safe_float).sum())
                         .sort_values(ascending=False)
                         .head(10).reset_index())
                top10.columns = ['Area','Dias']
                top10['Dias'] = top10['Dias'].astype(int)
                max_v = int(top10['Dias'].max()) if len(top10)>0 else 1
                for _, row_t in top10.iterrows():
                    pct_b = int(row_t['Dias']/max_v*100) if max_v>0 else 0
                    st.markdown(
                        f"<div style='margin-bottom:10px'>"
                        f"<div style='display:flex;justify-content:space-between;"
                        f"font-size:13px;margin-bottom:3px'>"
                        f"<span style='color:{AZUL};font-weight:500'>{row_t['Area']}</span>"
                        f"<span style='color:{FUCSIA};font-weight:700'>{row_t['Dias']:,} días</span></div>"
                        f"<div style='height:7px;background:{AZUL_L};border-radius:4px'>"
                        f"<div style='width:{pct_b}%;height:100%;background:{FUCSIA};"
                        f"border-radius:4px'></div></div></div>",
                        unsafe_allow_html=True)

        with col_r:
            # Nivel de agrupacion para el resumen segun rol:
            # RRHH/Gerente/SubGerente -> por Jefe
            # Jefe -> por Administrador (sus reportes directos)
            # Administrador -> resumen propio
            col_admin = 'Administrador' if 'Administrador' in df.columns else None
            if role == 'Jefe' and col_admin:
                grp_resumen    = col_admin
                titulo_resumen = "### Resumen por Administrador"
                lbl_col        = 'Administrador'
            elif role == 'Administrador':
                grp_resumen    = None
                titulo_resumen = "### Resumen de tu area"
                lbl_col        = None
            else:
                grp_resumen    = col_jefe
                titulo_resumen = "### Resumen por Jefe"
                lbl_col        = 'Jefe'

            st.markdown(titulo_resumen)
            if grp_resumen and grp_resumen in df.columns and 'Estado' in df.columns:
                rows_j = []
                for jv in sorted(df[grp_resumen].dropna().unique()):
                    gd   = df[df[grp_resumen]==jv]
                    mj   = col_sum(gd, col_meta)
                    pj   = cap_sum(gd, 'Prog_visma', col_meta)
                    pctj = round(pj/mj*100,1) if mj>0 else 0
                    vj   = int((gd['Vencidos_real']>0).sum()) if 'Vencidos_real' in gd.columns else 0
                    dpj  = int(col_sum(gd, col_dp))
                    rows_j.append({lbl_col: jv, 'HC': len(gd),
                                   '% Avance': f"{pctj}%", 'Vencidos': vj,
                                   'Dias x prog.': int(dpj)})
                if rows_j:
                    rj = pd.DataFrame(rows_j)
                    rj = rj.sort_values('Dias x prog.', ascending=False)
                    st.dataframe(rj, use_container_width=True, hide_index=True, height=360)
            elif role == 'Administrador' and 'Estado' in df.columns:
                mj   = col_sum(df, col_meta)
                pj   = cap_sum(df, 'Prog_visma', col_meta)
                pctj = round(pj/mj*100,1) if mj>0 else 0
                vj   = int((df['Vencidos_real']>0).sum()) if 'Vencidos_real' in df.columns else 0
                dpj  = int(col_sum(df, col_dp))
                rj   = pd.DataFrame([{'HC': len(df), '% Avance': f"{pctj}%",
                                       'Vencidos': vj, 'Dias x prog.': int(dpj)}])
                st.dataframe(rj, use_container_width=True, hide_index=True)

    # ── COLABORADORES POR AREA ─────────────────────────────────────────────────
    elif pagina == "👥 Colaboradores por Área":
        st.markdown("## Colaboradores por Área")
        c1,c2,c3,c4,c5 = st.columns(5)
        buscar = c1.text_input("🔍 Nombre o legajo")
        f_area = c2.selectbox("Área",['Todas']+sorted(df[col_area].dropna().unique().tolist()) if col_area else ['Todas'])
        f_cat  = c3.selectbox("Categoría",['Todas']+sorted(df[col_cat].dropna().unique().tolist()) if col_cat else ['Todas'])
        f_est  = c4.selectbox("Estado",['Todos','VENCIDO','CRITICO','EN_RIESGO','SIN_SALDO'])
        f_jefe = c5.selectbox("Jefe",['Todos']+sorted(df[col_jefe].dropna().unique().tolist()) if col_jefe else ['Todos'])

        # Solo colaboradores con area Y con dias pendientes por programar
        # Excluir cesados en 2026: su meta ya esta cumplida, Dias_x_prog=0
        df_f = df.copy()
        # Excluir por Estado CUMPLIDO que incluye cesados y los que completaron meta
        # Para la vista de "pendientes" solo interesan los que aun deben programar
        # Los cesados del archivo altas/bajas tienen es_cesado=True O Estado=CUMPLIDO via legajo
        if 'es_cesado' in df_f.columns:
            df_f = df_f[~df_f['es_cesado']].copy()
        # Excluir sin area
        if col_area:
            df_f = df_f[df_f[col_area].notna() & (df_f[col_area].astype(str).str.strip()!='') & (df_f[col_area].astype(str).str.lower()!='none')]
        if col_dp and col_meta:
            df_f = df_f[df_f[col_dp].apply(safe_float)>0]

        if buscar and col_nom:
            mk = df_f[col_nom].astype(str).str.upper().str.contains(buscar.upper(),na=False)
            mk |= df_f['Legajo'].astype(str).str.contains(buscar,na=False)
            df_f = df_f[mk]
        if f_area!='Todas' and col_area: df_f=df_f[df_f[col_area]==f_area]
        if f_cat !='Todas' and col_cat:  df_f=df_f[df_f[col_cat]==f_cat]
        if f_est !='Todos' and 'Estado' in df_f.columns: df_f=df_f[df_f['Estado']==f_est]
        if f_jefe!='Todos' and col_jefe: df_f=df_f[df_f[col_jefe]==f_jefe]

        st.caption(f"{len(df_f):,} colaboradores con días pendientes por programar")
        cols_t = [c for c in ['Legajo',col_nom,col_cat,col_area,col_jefe,'Administrador',
                               col_pend,'Truncos',col_meta,'Prog_visma','Pct avance',
                               col_dp,'Fecha límite','Estado']
                  if c and c in df_f.columns]
        show = df_f[cols_t].copy()
        ren  = {col_nom:'Nombre',col_cat:'Categoria',col_area:'Area',col_jefe:'Jefe',
                col_pend:'Pendientes',col_meta:'Meta 2026','Prog_visma':'Días programados',
                'Pct avance':'% Avance',col_dp:'Días x prog.','Fecha límite':'Fecha límite'}
        show = show.rename(columns={k:v for k,v in ren.items() if k and k in show.columns})
        if '% Avance' in show.columns:
            show['% Avance'] = show['% Avance'].apply(lambda x: f"{safe_float(x):.1f}%")
        if 'Estado' in show.columns:
            show['Estado'] = show['Estado'].apply(emo)
        # Redondear columnas numéricas a entero
        for col_int in ['Pendientes','Truncos','Días programados','Días x prog.','Meta 2026']:
            if col_int in show.columns:
                show[col_int] = show[col_int].apply(lambda x: int(safe_float(x)))

        st.dataframe(show, use_container_width=True, hide_index=True, height=500)
        st.download_button("⬇️ Descargar Excel", to_excel(show),
            file_name=f"colaboradores_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── ALERTAS ────────────────────────────────────────────────────────────────
    elif pagina == "🔔 Centro de Alertas":
        st.markdown("## Centro de Alertas")
        df_v = df[df['Estado']=='VENCIDO'].copy()   if 'Estado' in df.columns else df.head(0)
        df_c = df[df['Estado']=='CRITICO'].copy()   if 'Estado' in df.columns else df.head(0)
        df_r = df[df['Estado']=='EN_RIESGO'].copy() if 'Estado' in df.columns else df.head(0)

        c1,c2,c3 = st.columns(3)
        c1.metric("🔴 Vencidos",          len(df_v))
        c2.metric("🟠 Críticos ≤30 días", len(df_c))
        c3.metric("🟡 En riesgo ≤90 días",len(df_r))

        cols_a = [c for c in ['Legajo',col_nom,col_area,col_jefe,'Vencidos_real',
                               col_pend,col_dp,'Fecha límite','Dias restantes',
                               'Estado','Comentario_ind']
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
                if ci2.button("Ignorar") and leg_ig: ignorar(leg_ig); st.rerun()
                if ci3.button("Restaurar") and leg_ig: restaurar(leg_ig); st.rerun()
            st.download_button("⬇️ Descargar vencidos", to_excel(df_v),
                file_name=f"vencidos_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("Sin días vencidos en tu vista")

        st.markdown("### 🟠 Críticos — vencen en menos de 30 días")
        if not df_c.empty:
            show = df_c[cols_a].copy().sort_values('Dias restantes')
            show['Estado'] = show['Estado'].apply(emo)
            st.dataframe(show, use_container_width=True, hide_index=True)
        else:
            st.success("Sin críticos")

        st.markdown("### 🟡 En riesgo — vencen en 30–90 días")
        if not df_r.empty:
            show = df_r[cols_a].copy().sort_values('Dias restantes')
            show['Estado'] = show['Estado'].apply(emo)
            st.dataframe(show, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Descargar en riesgo", to_excel(df_r),
                file_name=f"riesgo_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("Sin colaboradores en riesgo")

    # ── CALENDARIO ─────────────────────────────────────────────────────────────
    elif pagina == "📅 Calendario":
        st.markdown("## Calendario de Vacaciones")
        c1,c2 = st.columns([1,4])
        with c1:
            mes_sel  = st.selectbox("Mes", MES_NAMES, index=date.today().month-1)
            anio_sel = st.selectbox("Año", [2025,2026,2027], index=1)
            f_jefe_c = st.selectbox("Jefe",['Todos']+sorted(df[col_jefe].dropna().unique().tolist()) if col_jefe else ['Todos'])
            f_area_c = st.selectbox("Área",['Todas']+sorted(df[col_area].dropna().unique().tolist()) if col_area else ['Todas'])
            f_leg_c  = st.text_input("Buscar legajo o nombre", placeholder="Ej: 1000097727")
        df_cal = df.copy()
        if f_jefe_c!='Todos' and col_jefe: df_cal=df_cal[df_cal[col_jefe]==f_jefe_c]
        if f_area_c!='Todas' and col_area: df_cal=df_cal[df_cal[col_area]==f_area_c]
        if f_leg_c:
            mk = df_cal['Legajo'].astype(str).str.contains(f_leg_c,na=False)
            if col_nom: mk |= df_cal[col_nom].astype(str).str.upper().str.contains(f_leg_c.upper(),na=False)
            df_cal = df_cal[mk]
        with c2:
            mes_num = MES_NAMES.index(mes_sel)+1
            st.markdown(f"### {mes_sel} {anio_sel} — {len(df_cal):,} colaboradores")
            render_calendario(df_visma, df_cal, mes_num, anio_sel)

    # ── RESUMEN EJECUTIVO ──────────────────────────────────────────────────────
    elif pagina == "📋 Resumen Ejecutivo":
        st.markdown("## Resumen Ejecutivo por Gerencia")
        grp = col_ger if col_ger else col_jefe
        if grp and grp in df.columns:
            CATS = ['OPERATIVOS','SUPERVISORES','BACK OFFICE']

            def fila_grupo(gd, label, nivel='gerencia'):
                meta = col_sum(gd, col_meta)
                prog = col_sum(gd, 'Prog_visma')
                pgc  = cap_sum(gd, 'Prog_visma', col_meta)
                pct2 = round(pgc/meta*100,1) if meta>0 else 0
                return {
                    'Nivel':            label,
                    'HC':               f"{len(gd):,}",
                    'Vencidos':         int((gd['Vencidos_real']>0).sum()) if 'Vencidos_real' in gd.columns else 0,
                    'Críticos':         int((gd['Estado']=='CRITICO').sum()) if 'Estado' in gd.columns else 0,
                    'En riesgo':        int((gd['Estado']=='EN_RIESGO').sum()) if 'Estado' in gd.columns else 0,
                    'Meta 2026':        f"{int(meta):,}",
                    'Días programados': f"{int(prog):,}",
                    '% Avance':         f"{pct2}%",
                    'Días x prog.':     f"{int(col_sum(gd, col_dp)):,}",
                }

            # Tabla expandible por gerencia con categorías
            all_rows_export = []
            for gv in sorted(df[grp].dropna().unique()):
                gd = df[df[grp]==gv]
                meta_g = col_sum(gd, col_meta)
                prog_g = col_sum(gd, 'Prog_visma')
                pgc_g  = cap_sum(gd, 'Prog_visma', col_meta)
                pct_g  = round(pgc_g/meta_g*100,1) if meta_g>0 else 0
                dp_g   = int(col_sum(gd, col_dp))
                venc_g = int((gd['Vencidos_real']>0).sum()) if 'Vencidos_real' in gd.columns else 0

                # Header de gerencia con colores Apparka
                col_a, col_b, col_c, col_d, col_e, col_f = st.columns([3,1,1.2,1.4,1,1.4])
                col_a.markdown(f"<div style='background:{AZUL};color:white;padding:6px 10px;"
                               f"border-radius:6px;font-weight:600;font-size:13px'>{gv}</div>",
                               unsafe_allow_html=True)
                col_b.markdown(f"<div style='text-align:right;font-size:12px'><b>HC</b><br>{len(gd):,}</div>",
                               unsafe_allow_html=True)
                col_c.markdown(f"<div style='text-align:right;font-size:12px'><b>Meta 2026</b><br>{int(meta_g):,}</div>",
                               unsafe_allow_html=True)
                col_d.markdown(f"<div style='text-align:right;font-size:12px'><b>Días prog.</b><br>{int(prog_g):,}</div>",
                               unsafe_allow_html=True)
                col_e.markdown(f"<div style='text-align:right;font-size:12px;color:{FUCSIA}'>"
                               f"<b>% Avance</b><br><b>{pct_g}%</b></div>",
                               unsafe_allow_html=True)
                col_f.markdown(f"<div style='text-align:right;font-size:12px;color:{MORADO}'>"
                               f"<b>Días x prog.</b><br><b>{dp_g:,}</b></div>",
                               unsafe_allow_html=True)

                # Categorías debajo
                if col_cat and col_cat in gd.columns:
                    for cat in CATS:
                        gc = gd[gd[col_cat].str.upper()==cat] if col_cat in gd.columns else pd.DataFrame()
                        if gc.empty: continue
                        meta_c = col_sum(gc, col_meta)
                        prog_c = col_sum(gc, 'Prog_visma')
                        pgc_c  = cap_sum(gc, 'Prog_visma', col_meta)
                        pct_c  = round(pgc_c/meta_c*100,1) if meta_c>0 else 0
                        dp_c   = int(col_sum(gc, col_dp))
                        ca,cb,cc,cd,ce,cf = st.columns([3,1,1.2,1.4,1,1.4])
                        cat_color = AZUL_L
                        ca.markdown(f"<div style='background:{cat_color};color:{AZUL};padding:4px 10px 4px 24px;"
                                    f"border-radius:4px;font-size:12px'>↳ {cat.title()}</div>",
                                    unsafe_allow_html=True)
                        cb.markdown(f"<div style='text-align:right;font-size:12px;color:#666'>{len(gc):,}</div>",
                                    unsafe_allow_html=True)
                        cc.markdown(f"<div style='text-align:right;font-size:12px;color:#666'>{int(meta_c):,}</div>",
                                    unsafe_allow_html=True)
                        cd.markdown(f"<div style='text-align:right;font-size:12px;color:#666'>{int(prog_c):,}</div>",
                                    unsafe_allow_html=True)
                        ce.markdown(f"<div style='text-align:right;font-size:12px;color:{FUCSIA}'>{pct_c}%</div>",
                                    unsafe_allow_html=True)
                        cf.markdown(f"<div style='text-align:right;font-size:12px;color:{MORADO}'>{dp_c:,}</div>",
                                    unsafe_allow_html=True)
                        all_rows_export.append({
                            'Gerencia': gv, 'Categoría': cat.title(),
                            'HC': len(gc), 'Meta 2026': int(meta_c),
                            'Días programados': int(prog_c),
                            '% Avance': f"{pct_c}%", 'Días x prog.': dp_c,
                            'Vencidos': int((gc['Vencidos_real']>0).sum()) if 'Vencidos_real' in gc.columns else 0,
                        })

                st.markdown("<hr style='margin:4px 0;border-color:#f0efe9'>", unsafe_allow_html=True)

            if all_rows_export:
                exp_df = pd.DataFrame(all_rows_export)
                st.download_button("⬇️ Descargar resumen completo", to_excel(exp_df),
                    file_name=f"resumen_categorias_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── HISTORIAL ──────────────────────────────────────────────────────────────
    elif pagina == "📂 Historial de Vacaciones":
        st.markdown("## Historial de Vacaciones de Colaboradores")
        if df_visma.empty:
            st.warning("Sube Vacaciones_-_Dias_solicitados__28_.xlsx al repositorio.")
            return
        leg_ok = set(df['Legajo'].astype(str).unique())
        hist   = df_visma[df_visma['Legajo'].isin(leg_ok)].copy()

        # Limpiar columnas
        hist['Fecha desde'] = hist['Fecha desde'].apply(lambda x: x.date() if pd.notna(x) else '')
        hist['Fecha hasta'] = hist['Fecha hasta'].apply(lambda x: x.date() if pd.notna(x) else '')
        hist['Cant dias']   = hist['Cant dias'].astype(int)

        # Unir nombre del colaborador
        if col_nom and 'Legajo' in df.columns:
            ln = df[['Legajo',col_nom]].drop_duplicates()
            ln['Legajo'] = ln['Legajo'].astype(str)
            hist = hist.merge(ln, on='Legajo', how='left')
            col_nombre_h = col_nom
        else:
            col_nombre_h = None

        c1,c2,c3,c4 = st.columns(4)
        bus  = c1.text_input("🔍 Legajo o nombre")
        anio = c2.selectbox("Año",['Todos']+sorted(
            [int(x) for x in hist['Fecha desde'].apply(
                lambda x: x.year if hasattr(x,'year') else 0).unique() if x>0],reverse=True))
        est  = c3.selectbox("Estado",['Todos']+sorted(hist['Estado aus'].dropna().unique().tolist()))
        per  = c4.selectbox("Periodo",['Todos']+sorted(hist['Periodo'].dropna().unique().tolist(),reverse=True))

        hf = hist.copy()
        if bus:
            mk = hf['Legajo'].astype(str).str.contains(bus,na=False)
            if col_nombre_h and col_nombre_h in hf.columns:
                mk |= hf[col_nombre_h].astype(str).str.upper().str.contains(bus.upper(),na=False)
            hf = hf[mk]
        if anio!='Todos':
            hf = hf[hf['Fecha desde'].apply(lambda x: x.year if hasattr(x,'year') else 0)==int(anio)]
        if est !='Todos': hf=hf[hf['Estado aus']==est]
        if per !='Todos': hf=hf[hf['Periodo']==per]

        # Seleccionar y renombrar columnas
        cols_h = ['Legajo']
        if col_nombre_h and col_nombre_h in hf.columns: cols_h.append(col_nombre_h)
        cols_h += ['Fecha desde','Fecha hasta','Cant dias','Periodo','Estado aus']
        cols_h  = [c for c in cols_h if c in hf.columns]
        hf_show = hf[cols_h].copy()
        hf_show = hf_show.rename(columns={
            col_nombre_h: 'Apellidos y Nombres',
            'Cant dias': 'Cantidad días',
            'Estado aus': 'Estado',
            'Fecha desde': 'Fecha desde',
            'Fecha hasta': 'Fecha hasta',
        })

        st.caption(f"{len(hf_show):,} registros — {int(hf['Cant dias'].sum()):,} días totales")
        st.dataframe(hf_show.sort_values('Fecha desde',ascending=False),
                     use_container_width=True, hide_index=True, height=500)
        st.download_button("⬇️ Descargar historial", to_excel(hf_show),
            file_name=f"historial_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == '__main__':
    main()
