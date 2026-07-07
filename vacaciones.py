"""
vacaciones.py - Gestión de Vacaciones Apparka
"""
import streamlit as st
import pandas as pd
import json, re, calendar, math
from datetime import date, datetime, timedelta
from dateutil.relativedelta import relativedelta
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
    import os
    # Intentar primero con hoja especifica, luego con la primera hoja disponible
    candidatos = [
        ('CONSOLIDADO_GENERADO.xlsx', None),
        ('META_2026_-_Abril.xlsx', 'Consolidado'),
        ('META_2026_-_Abril.xlsx', None),   # fallback: primera hoja
    ]
    for fn, sh in candidatos:
        if not os.path.exists(fn): continue
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
        # Detectar fila de header: buscar la que tiene 'Legajo' como primera celda con datos
        header_row = 2  # default
        for h in [2, 3, 1, 4]:
            try:
                test = pd.read_excel(archivo_visma, header=h, nrows=1)
                primera = str(test.columns[0]).strip().lower()
                if primera == 'legajo':
                    header_row = h
                    break
                # A veces el header esta una fila despues (cols unnamed)
                if 'legajo' in [str(c).strip().lower() for c in test.iloc[0].tolist()]:
                    header_row = h + 1
                    break
            except:
                pass

        v = pd.read_excel(archivo_visma, header=header_row)
        # Renombrar por posicion (tolerante a variaciones de nombre)
        cols_base = ['Legajo','_x','Apellidos y Nombre','Estado','Fecha desde',
                     'Fecha hasta','Cant dias','_2','Tipo dia','Periodo',
                     'Origen','Estado aus','Anticipo']
        if len(v.columns) >= len(cols_base):
            v.columns = cols_base + list(v.columns[len(cols_base):])
        v = v[v['Legajo'].notna() & (v['Legajo'].astype(str).str.strip()!='Legajo')].copy()
        v['Legajo']      = v['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
        v = v[v['Legajo'].str.match(r'^[0-9]+$')]
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

@st.cache_data(ttl=3600)
def cargar_area_sistema():
    """
    Lee Atributos_de_estructuras_por_colaborador.xlsx del sistema.
    Regla de area (en orden de prioridad):
      1. CENTROS DE COSTO: area de gestion interna (Back Office, admin, soporte)
      2. PCES: estacionamiento asignado (operativos de campo)
      3. AREA: fallback final si no tiene ninguno de los anteriores
    Retorna dict {'area': {legajo: area}, 'sede': {legajo: sede}}
    """
    import os
    nombres = [
        'Atributos_de_estructuras_por_colaborador.xlsx',
        'Atributos de estructuras por colaborador.xlsx',
    ]
    archivo = next((n for n in nombres if os.path.exists(n)), None)
    if not archivo:
        try:
            archivo = next(f for f in os.listdir('.') if f.startswith('Atributos') and f.endswith('.xlsx'))
        except StopIteration:
            return {}
    try:
        df = pd.read_excel(archivo, header=1)
        df.columns = ['Legajo','Nombre','Tipo_estructura','Atributo','Fecha_desde','Fecha_hasta']
        # Limpiar espacios en tipo estructura (el sistema exporta con espacios al final)
        df['Tipo_estructura'] = df['Tipo_estructura'].astype(str).str.strip()
        df['Atributo']        = df['Atributo'].astype(str).str.strip()
        df['Legajo']          = df['Legajo'].astype(str).str.replace('.0','',regex=False).str.strip()
        df = df[df['Legajo'].str.match(r'^[0-9]+$')]

        # Extraer cada tipo relevante
        ceco = (df[df['Tipo_estructura'] == 'CENTROS DE COSTO']
                [['Legajo','Atributo']].rename(columns={'Atributo':'CECO'})
                .drop_duplicates('Legajo'))
        pces = (df[df['Tipo_estructura'] == 'PCES']
                [['Legajo','Atributo']].rename(columns={'Atributo':'PCES'})
                .drop_duplicates('Legajo'))
        area = (df[df['Tipo_estructura'] == 'AREA']
                [['Legajo','Atributo']].rename(columns={'Atributo':'AREA_SYS'})
                .drop_duplicates('Legajo'))
        sede = (df[df['Tipo_estructura'] == 'SEDE']
                [['Legajo','Atributo']].rename(columns={'Atributo':'SEDE_SYS'})
                .drop_duplicates('Legajo'))

        base   = df[['Legajo']].drop_duplicates()
        result = (base
                  .merge(ceco, on='Legajo', how='left')
                  .merge(pces, on='Legajo', how='left')
                  .merge(area, on='Legajo', how='left')
                  .merge(sede, on='Legajo', how='left'))

        # Regla de prioridad: CECO > PCES > AREA
        # EXCEPCION: si CECO contiene 'VACAC' o 'RETEN' es un codigo temporal,
        # no es el area real -> ignorarlo y usar PCES o AREA en su lugar
        ceco_valido = result['CECO'].copy()
        mask_vacac  = ceco_valido.astype(str).str.upper().str.contains('VACAC|RETEN', na=False)
        ceco_valido[mask_vacac] = None  # descartar CECO invalido

        result['AREA_FINAL'] = (ceco_valido
                                .fillna(result['PCES'])
                                .fillna(result['AREA_SYS']))

        # Extraer PUESTO por legajo
        puesto = (df[df['Tipo_estructura'] == 'PUESTO']
                  [['Legajo','Atributo']].rename(columns={'Atributo':'PUESTO'})
                  .drop_duplicates('Legajo'))
        result = result.merge(puesto, on='Legajo', how='left')

        # Normalizar nombres de area con variaciones conocidas
        area_alias = {
            'SWISSOTEL': 'SUBTERRANEO SWISSOTEL',
        }
        result['AREA_FINAL'] = result['AREA_FINAL'].replace(area_alias)

        area_map   = result.dropna(subset=['AREA_FINAL']).set_index('Legajo')['AREA_FINAL'].to_dict()
        sede_map   = result.dropna(subset=['SEDE_SYS']).set_index('Legajo')['SEDE_SYS'].to_dict()
        puesto_map = result.dropna(subset=['PUESTO']).set_index('Legajo')['PUESTO'].to_dict()
        return {'area': area_map, 'sede': sede_map, 'puesto': puesto_map}
    except Exception as e:
        return {}


@st.cache_data(ttl=3600)
def cargar_grupos_area():
    """Lee grupos_area.json: {Jefe -> nombre_grupo}"""
    import os
    for fn in ['grupos_area.json']:
        if os.path.exists(fn):
            try:
                with open(fn, encoding='utf-8') as f:
                    return json.load(f)
            except:
                pass
    return {}

@st.cache_data(ttl=86400)
def cargar_jerarquia():
    """Carga acceso_persona.json para autenticacion y roles de usuario."""
    try:
        with open('acceso_persona.json', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

@st.cache_data(ttl=3600)
def cargar_tabla_jerarquia():
    """
    Carga jerarquia.csv con columnas: Gerente, Sub_Gerente, Jefe, Administrador.
    Retorna DataFrame con combinaciones unicas Administrador -> {Gerente, SubGerente, Jefe}.
    """
    import os
    nombres = ['jerarquia.csv', 'jerarquia.CSV']
    archivo = next((n for n in nombres if os.path.exists(n)), None)
    if not archivo:
        return pd.DataFrame()
    try:
        df = pd.read_csv(archivo)
        df.columns = [c.strip() for c in df.columns]
        # Normalizar nombres de columna (puede venir Sub_Gerente o SubGerente)
        rename = {}
        for c in df.columns:
            cl = c.lower().replace(' ','_').replace('-','_')
            if 'gerente' in cl and 'sub' not in cl:   rename[c] = 'Gerente'
            elif 'sub' in cl and 'gerente' in cl:      rename[c] = 'Sub_Gerente'
            elif 'jefe' in cl:                         rename[c] = 'Jefe'
            elif 'admin' in cl:                        rename[c] = 'Administrador'
        df = df.rename(columns=rename)
        # Limpiar valores
        for col in ['Gerente','Sub_Gerente','Jefe','Administrador']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace({'nan':'', 'None':'', 'NaN':''})
        return df.drop_duplicates()
    except Exception as e:
        return pd.DataFrame()

# ── Logica vacaciones ──────────────────────────────────────────────────────────
def fecha_limite(comentario):
    if not comentario or str(comentario).strip() in ['-','nan','None','']: return None
    m = re.search(r'ANTES DEL (\d{2}/\d{2}/\d{4})', str(comentario).upper())
    if m:
        try: return datetime.strptime(m.group(1),'%d/%m/%Y').date()
        except: pass
    return None

def construir_consolidado(df_meta, df_visma, df_ab=None, area_sistema=None, pa=None):
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

    # Actualizar AREA y SEDE desde el archivo de atributos del sistema (fuente mas fresca)
    if area_sistema:
        area_map = area_sistema.get('area', {})
        sede_map = area_sistema.get('sede', {})
        col_a = next((c for c in ['AREA','Area'] if c in df.columns), None)
        col_s = next((c for c in ['SEDE','Sede'] if c in df.columns), None)
        if area_map and col_a:
            df[col_a] = df['Legajo'].map(area_map).fillna(df[col_a])
        if sede_map and col_s:
            df[col_s] = df['Legajo'].map(sede_map).fillna(df[col_s])

    # Construir columnas de jerarquia desde acceso_persona.json y jerarquia.csv
    # Estrategia: area -> {Administrador, Jefe, Sub_Gerente, Gerente}
    import os
    col_a_adm = next((c for c in ['AREA','Area'] if c in df.columns), None)

    # Mapas area -> persona por nivel (construidos desde acceso_persona.json)
    area_to_admin   = {}  # area -> nombre Administrador
    area_to_jefe    = {}  # area -> nombre Jefe
    area_to_subger  = {}  # area -> nombre Sub_Gerente
    area_to_gerente = {}  # area -> nombre Gerente

    if pa:
        # Construir mapa area -> jerarquia desde acceso_persona.json
        for email, uinfo in pa.items():
            nombre = uinfo.get('nombre','')
            if not nombre: continue
            role_u = uinfo.get('role','')

            # Areas de acceso segun rol
            for ar in uinfo.get('areas',[]):
                if not ar or str(ar) in ('nan','None','NaN'): continue
                if role_u == 'Administrador':
                    area_to_admin[ar] = nombre
                elif role_u == 'Jefe':
                    area_to_jefe[ar] = nombre
                elif role_u in ('SubGerente', 'GerenteGeneral'):
                    area_to_subger[ar] = nombre
                elif role_u == 'Gerente':
                    area_to_gerente[ar] = nombre

            # admin_areas: areas que administra directamente aunque tenga otro rol de acceso
            # (ej: Oswaldo=RRHH administra ADM. RH; Fernando=GerenteGeneral administra ROL PRIVADO)
            for ar in uinfo.get('admin_areas', []):
                if not ar or str(ar) in ('nan','None','NaN'): continue
                if ar not in area_to_admin:
                    area_to_admin[ar] = nombre

        # Completar Jefe/SubGerente/Gerente desde jerarquia.csv usando Admin como puente
        df_jer = cargar_tabla_jerarquia()
        if not df_jer.empty:
            # Casos donde Jefe == Administrador: esa persona es su propio nivel
            # En ese caso NO propagar como Jefe (evitar que aparezca dos veces)
            jer_rows = df_jer[['Gerente','Sub_Gerente','Jefe','Administrador']].copy()
            # Mapa Admin -> {Jefe, SubGerente, Gerente} - usando primera aparicion
            jer_uniq = jer_rows.drop_duplicates('Administrador')
            jer_map  = jer_uniq.set_index('Administrador').to_dict('index')
            for ar, adm in area_to_admin.items():
                entry = jer_map.get(str(adm), {})
                jefe_csv  = entry.get('Jefe','')
                subg_csv  = entry.get('Sub_Gerente','')
                ger_csv   = entry.get('Gerente','')
                # Solo asignar Jefe si es distinto al Administrador
                if ar not in area_to_jefe and jefe_csv and jefe_csv != adm:
                    area_to_jefe[ar] = jefe_csv
                if ar not in area_to_subger and subg_csv and subg_csv != adm:
                    area_to_subger[ar] = subg_csv
                if ar not in area_to_gerente and ger_csv:
                    area_to_gerente[ar] = ger_csv

    # Asignar columnas al dataframe
    if col_a_adm:
        df['Administrador'] = df[col_a_adm].map(area_to_admin)
        df['Jefe']          = df[col_a_adm].map(area_to_jefe)
        df['Sub_Gerente']   = df[col_a_adm].map(area_to_subger)
        df['Gerente']       = df[col_a_adm].map(area_to_gerente)

        # Fallback Admin: Jefe -> SubGerente -> Gerente (en ese orden)
        mask = df['Administrador'].isna()
        df.loc[mask, 'Administrador'] = df.loc[mask, col_a_adm].map(area_to_jefe)
        mask = df['Administrador'].isna()
        df.loc[mask, 'Administrador'] = df.loc[mask, col_a_adm].map(area_to_subger)
        mask = df['Administrador'].isna()
        df.loc[mask, 'Administrador'] = df.loc[mask, col_a_adm].map(area_to_gerente)

        # Fallback Jefe: si no hay Jefe, usar el Administrador que ya se asigno
        mask_jefe = df['Jefe'].isna()
        df.loc[mask_jefe, 'Jefe'] = df.loc[mask_jefe, 'Administrador']

    else:
        for col in ['Administrador','Jefe','Sub_Gerente','Gerente']:
            if col not in df.columns:
                df[col] = None

    # Agregar columna Grupo desde grupos_area.json (Jefe -> nombre grupo)
    grupos_map = cargar_grupos_area()
    if grupos_map and 'Jefe' in df.columns:
        df['Grupo'] = df['Jefe'].map(grupos_map)
        # Para colaboradores sin Jefe (SubGerente directo), usar Sub_Gerente como clave
        mask_sin_grupo = df['Grupo'].isna() & df['Sub_Gerente'].notna()
        df.loc[mask_sin_grupo, 'Grupo'] = df.loc[mask_sin_grupo, 'Sub_Gerente'].map(grupos_map)

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

    # Pre-calcular registros Visma por legajo (todos los años aprobados/pendientes)
    # registros_visma     = {legajo: [(fecha, dias), ...]}  -- sin periodo
    # registros_visma_per = {legajo: {periodo_int: [(fecha, dias), ...]}} -- con periodo
    registros_visma     = {}
    registros_visma_per = {}
    if not df_visma.empty:
        v_todos = df_visma[
            df_visma['Estado aus'].isin(['Aprobada','Pendiente'])
        ].copy()
        for _, vrow in v_todos.iterrows():
            vleg    = str(vrow['Legajo']).replace('.0','').strip()
            vfech   = vrow['Fecha desde']
            vdias   = safe_float(vrow.get('Cant dias', 0))
            vper    = vrow.get('Periodo', None)
            vper_i  = int(vper) if pd.notna(vper) else None
            # Sin periodo
            if vleg not in registros_visma:
                registros_visma[vleg] = []
            registros_visma[vleg].append((vfech, vdias))
            # Con periodo
            if vper_i:
                if vleg not in registros_visma_per:
                    registros_visma_per[vleg] = {}
                if vper_i not in registros_visma_per[vleg]:
                    registros_visma_per[vleg][vper_i] = []
                registros_visma_per[vleg][vper_i].append((vfech, vdias))

    # Construir mapa legajo -> fecha de ingreso (la mas reciente activa)
    # Si fue cesado y reingreso, sus vacaciones se liquidaron y empiezan de cero
    fecha_ingreso_map = {}
    if df_ab is not None and not df_ab.empty and 'Fecha_alta' in df_ab.columns:
        ab_validos = df_ab[df_ab['Fecha_alta'].notna()].copy()
        # Normalizar legajo igual que en el resto del sistema
        ab_validos['Legajo'] = (ab_validos['Legajo'].astype(str)
                                .str.replace('.0','',regex=False).str.strip())
        ab_validos = ab_validos[ab_validos['Legajo'].str.match(r'^[0-9]+$')]
        ab_validos = ab_validos.sort_values('Fecha_alta', ascending=False)
        for leg_ab, grupo in ab_validos.groupby('Legajo'):
            primera = grupo.iloc[0]  # la mas reciente activa
            if pd.notna(primera['Fecha_alta']):
                fecha_ingreso_map[str(leg_ab)] = primera['Fecha_alta'].date()

    estados,vencidos_r,fechas_l,dias_rest,pcts = [],[],[],[],[]
    for _, row in df.iterrows():
        leg       = str(row['Legajo']).replace('.0','').strip()
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
            # ── VENCIMIENTO LEGAL (por periodo/aniversario) ───────────────────
            # Completamente independiente de la Meta de programacion Apparka
            # Fuente: fecha_ingreso (altas/bajas) + registros_visma_per (Visma por periodo)
            venc = 0
            fi = fecha_ingreso_map.get(leg)
            if fi:
                # Calcular vencimiento por AÑO LABORAL usando fechas reales
                # NO usar el campo Periodo de Visma (puede estar incorrecto)
                # Año laboral N: desde aniversario N hasta aniversario N+1 - 1 dia
                # Fecha limite para gozarlo: aniversario N+2 - 1 dia
                # Vencimiento legal: asignacion secuencial por fecha de ingreso
                # Cada registro cubre el periodo mas antiguo primero
                # Si un registro es mayor al saldo del periodo, el resto pasa al siguiente
                # SIN campo Periodo de Visma, SIN comentario META
                if fi:
                    regs_leg  = sorted(
                        [(f, d) for f, d in registros_visma.get(leg, []) if pd.notna(f)],
                        key=lambda x: x[0]
                    )
                    anios_max = relativedelta(hoy, fi).years + 2
                    # Periodos laborales
                    periodos_yl = []
                    for np_ in range(1, anios_max + 1):
                        periodos_yl.append({
                            'fl':     fi + relativedelta(years=np_+1) - relativedelta(days=1),
                            'antes':  0.0,
                            'despues':0.0
                        })
                    saldo_yl = [30.0] * len(periodos_yl)
                    # Asignar secuencialmente
                    for f_r, d_r in regs_leg:
                        _d_rest = float(d_r)
                        for idx_p, per_p in enumerate(periodos_yl):
                            if _d_rest <= 0: break
                            if saldo_yl[idx_p] <= 0: continue
                            aporte = min(_d_rest, saldo_yl[idx_p])
                            saldo_yl[idx_p] -= aporte
                            _d_rest         -= aporte
                            if f_r.date() <= per_p['fl']:
                                per_p['antes']   += aporte
                            else:
                                per_p['despues'] += aporte
                    # Detectar vencidos en el año actual
                    for per_p in periodos_yl:
                        if per_p['fl'].year != hoy.year: continue
                        if per_p['fl'] >= hoy:           continue
                        total_p = per_p['antes'] + per_p['despues']
                        if total_p < 15: continue  # part-time/trunco
                        if per_p['despues'] > 0:
                            venc = max(venc, int(per_p['despues']))

            # ── META DE PROGRAMACION APPARKA (independiente del vencimiento) ──
            # Fuente: comentario del META "DEBE GOZAR X DIAS ANTES DEL DD/MM/AAAA"
            # Controla estados CRITICO / EN_RIESGO para gestion operativa
            if venc > 0:
                estado = 'VENCIDO'
            elif fl and dias_r <= 30 and dias_x > 0:
                estado = 'CRITICO'
            elif fl and dias_r <= 90 and dias_x > 0:
                estado = 'EN_RIESGO'
            elif meta > 0 and prog >= meta:
                estado = 'CUMPLIDO'
            elif meta == 0 and pend == 0:
                estado = 'SIN_SALDO'
            else:
                estado = 'AL_DIA'

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

    # Calcular Dias_x_vencer: dias que aun debe gozar segun comentario vs Visma
    # Para CRITICO/VENCIDO: usa los dias del comentario "DEBE GOZAR X DIAS ANTES DEL..."
    # Esto es mas preciso que Dias_x_prog para las alertas
    def calc_dias_x_vencer(r):
        com  = str(r.get('Comentario_ind','') or '')
        leg  = str(r.get('Legajo',''))
        md   = re.search(r'DEBE GOZAR (\d+)', com.upper())
        if md:
            debia = int(md.group(1))
            fl    = fecha_limite(com)
            regs  = registros_visma.get(leg, [])
            if fl:
                if fl < hoy:
                    # Fecha ya paso: el numero del comentario ES la deuda
                    # (ya incorpora lo que habia gozado antes al exportar el META)
                    return debia
                else:
                    # Fecha futura: debia - dias gozados hasta hoy
                    prog_ref = sum(d for f, d in regs
                                   if pd.notna(f) and f.date() <= hoy)
                    return max(0, debia - prog_ref)
            else:
                prog_ref = safe_float(r.get('Prog_visma', 0))
                return max(0, debia - prog_ref)
        return safe_float(r.get('Dias_x_prog', 0))
    df['Dias_x_vencer'] = df.apply(calc_dias_x_vencer, axis=1)
    # Si el estado es VENCIDO, Dias_x_vencer = los dias vencidos (no la meta futura)
    if 'Estado' in df.columns and 'Vencidos_real' in df.columns:
        mask_vencido = df['Estado'] == 'VENCIDO'
        df.loc[mask_vencido, 'Dias_x_vencer'] = df.loc[mask_vencido, 'Vencidos_real']

    # Calcular vencimientos para anio siguiente (para selector de año en alertas)
    # Guardar en columna auxiliar 'Venc_anio_sig' para uso en main()
    anio_sig = hoy.year + 1
    def venc_anio_sig(row):
        leg = str(row['Legajo'])
        fi  = fecha_ingreso_map.get(leg)
        if not fi: return 0
        regs_todos = registros_visma.get(leg, [])
        n_check = relativedelta(date(anio_sig, 12, 31), fi).years
        for n in range(1, n_check + 1):
            fl_n = fi + relativedelta(years=n+1) - relativedelta(days=1)
            if fl_n.year != anio_sig: continue
            inicio_anio = fi + relativedelta(years=n)
            gozados_en_anio = sum(
                d for f, d in regs_todos
                if pd.notna(f) and inicio_anio <= f.date() <= fl_n
            )
            if gozados_en_anio == 0: continue
            if gozados_en_anio < 30:
                return 30 - gozados_en_anio
        return 0
    df['Venc_anio_sig'] = df.apply(venc_anio_sig, axis=1)
    return df

def filtrar_usuario(df, user_email, pa):
    # Vista: solo activos con area
    col_a = next((c for c in ['AREA','Area'] if c in df.columns), None)
    if col_a:
        df_act = df[df[col_a].notna() & (df[col_a].astype(str).str.strip()!='')].copy()
    else:
        df_act = df.copy()

    if user_email not in pa: return df_act
    info  = pa[user_email]
    role  = info.get('role', 'RRHH')
    areas = [a for a in info.get('areas', []) if a and str(a) not in ('nan','None','NaN')]

    # RRHH y Gerente ven todo
    if role in ('RRHH', 'Gerente', 'GerenteGeneral'):
        # Gerente General ve toda la empresa (igual que RRHH)
        # La restriccion de "solo reportes directos" aplica solo a las vistas especificas
        return df_act

    # Todos los demas (SubGerente, Jefe, Administrador): filtrar por areas del JSON
    if areas and col_a:
        return df_act[df_act[col_a].isin(areas)].copy()

    return df_act.iloc[0:0].copy()

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

    user_name  = st.session_state.user_name   # nombre legible, ej: "Mayra Huerta"
    user_email = st.session_state.get('user_email', '')  # email, ej: "mhuerta@apparka.pe"
    pa         = cargar_jerarquia()
    df_meta   = cargar_meta()
    df_visma  = cargar_visma()

    if df_meta.empty:
        st.error("No se encontro el archivo de datos. Sube META_2026_-_Abril.xlsx al repositorio.")
        return

    df_ab        = cargar_altas_bajas()
    area_sistema = cargar_area_sistema()
    df_full = construir_consolidado(df_meta, df_visma, df_ab, area_sistema, pa)
    df      = filtrar_usuario(df_full, user_email, pa)
    role    = pa.get(user_email, {}).get('role','RRHH')

    col_meta = next((c for c in ['Meta2026'] if c in df.columns), None)
    col_pend = 'Pendientes' if 'Pendientes' in df.columns else None
    col_dp   = 'Dias_x_prog' if 'Dias_x_prog' in df.columns else None
    col_area = next((c for c in ['AREA','Area'] if c in df.columns), None)
    col_cat  = next((c for c in ['Categoria','Categoría'] if c in df.columns), None)
    col_ger  = next((c for c in ['Gerencia','Gerente','Sub_Gerente'] if c in df.columns), None)
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
        # Buscar puesto real desde atributos del sistema usando legajo del JSON
        puesto_map  = area_sistema.get('puesto', {}) if area_sistema else {}
        puesto_user = ''
        if puesto_map:
            # Prioridad: legajo directo del acceso_persona.json
            legajo_user = str(pa.get(user_email, {}).get('legajo', ''))
            if legajo_user:
                puesto_user = puesto_map.get(legajo_user, '')
        if puesto_user:
            st.caption(puesto_user.title())
        else:
            roles_label = {'RRHH':'RRHH','Gerente':'Gerente','GerenteGeneral':'Gerente General',
                           'SubGerente':'Sub Gerente','Jefe':'Jefe de Área','Administrador':'Administrador'}
            st.caption(f"Rol: {roles_label.get(role, role)}")
        st.markdown("<hr style='border-color:rgba(255,255,255,0.2)'>",unsafe_allow_html=True)
        # Menu segun rol
        if role == 'GerenteGeneral':
            opciones_menu = [
                "📊 Dashboard",
                "📅 Calendario",
                "📋 Resumen Ejecutivo",
                "📂 Historial de Vacaciones",
            ]
        else:
            opciones_menu = [
                "📊 Dashboard","👥 Colaboradores por Área","🔔 Centro de Alertas",
                "📅 Calendario","📋 Resumen Ejecutivo","📂 Historial de Vacaciones",
            ]
        pagina = st.radio("", opciones_menu, label_visibility="collapsed")
        st.markdown("<hr style='border-color:rgba(255,255,255,0.2)'>",unsafe_allow_html=True)
        if st.button("Cerrar sesión"):
            st.session_state.authenticated = False; st.rerun()

    # ── DASHBOARD ──────────────────────────────────────────────────────────────
    if pagina == "📊 Dashboard":
        st.markdown("## Dashboard — Vacaciones")

        # KPIs scoped al usuario — si es RRHH ve empresa completa, si es Jefe/Admin ve su area
        is_rrhh = (role in ('RRHH', 'Gerente', 'GerenteGeneral'))  # ven empresa completa
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

        # ── Avance mensual de días gozados y programados ────────────────────────
        if not df_visma.empty:
            col_mes_sel, _ = st.columns([2,5])
            mes_avance = col_mes_sel.selectbox(
                "📅 Ver avance hasta el mes",
                options=MES_NAMES,
                index=date.today().month - 1,
                key='mes_avance_dash'
            )
            mes_num_av  = MES_NAMES.index(mes_avance) + 1
            dias_mes    = [31,28,31,30,31,30,31,31,30,31,30,31][mes_num_av-1]
            fecha_corte = date(2026, mes_num_av, dias_mes)
            # Misma base de Visma que los KPIs de arriba
            legs_scope  = set(df_scope['Legajo'].astype(str).unique())
            vis_scope   = df_visma[
                (df_visma['Legajo'].astype(str).isin(legs_scope)) &
                (df_visma['Estado aus'].isin(['Aprobada','Pendiente'])) &
                (df_visma['Fecha desde'].dt.year == 2026)
            ].copy()
            vis_scope['_d'] = vis_scope['Cant dias'].apply(safe_float)
            # Gozados hasta fin del mes (ya disfrutados o en curso)
            dias_gozados = int(vis_scope[vis_scope['Fecha desde'].dt.date <= fecha_corte]['_d'].sum())
            # Programados despues del mes (ya en Visma pero aun no disfrutados)
            dias_prog    = int(vis_scope[vis_scope['Fecha desde'].dt.date >  fecha_corte]['_d'].sum())
            # Sin programar = dias que faltan y NO estan en Visma = dp (mismo que KPI arriba)
            pct_goz  = round(dias_gozados/meta_t*100,1) if meta_t > 0 else 0
            pct_prog = round(dias_prog/meta_t*100,1)    if meta_t > 0 else 0
            cm1,cm2,cm3,cm4 = st.columns(4)
            cm1.markdown(kpi(f"✅ Gozados hasta {mes_avance}", fmt_num(dias_gozados)), unsafe_allow_html=True)
            cm2.markdown(kpi(f"📅 Programados resto 2026",    fmt_num(dias_prog)),    unsafe_allow_html=True)
            cm3.markdown(kpi(f"⚠️ Sin programar",             fmt_num(dp), FUCSIA),   unsafe_allow_html=True)
            cm4.markdown(kpi(f"📊 % Avance",                  f"{pct}%"),            unsafe_allow_html=True)
            st.caption(f"Gozados {pct_goz}% · Programados {pct_prog}% · Días sin programar {fmt_num(dp)} · Total avance {pct}%")

        st.markdown("---")

        # ── GerenteGeneral: tabla Personal a cargo a ancho completo ────────────
        if role == 'GerenteGeneral':
            st.markdown("### Personal a cargo")
            col_meta_gg   = next((c for c in ['Meta2026'] if c in df.columns), None)
            puesto_map_gg = area_sistema.get('puesto', {}) if area_sistema else {}
            legajo_gg     = str(pa.get(user_email, {}).get('legajo', ''))
            if col_meta_gg and puesto_map_gg:
                legs_dir = {leg: cargo for leg, cargo in puesto_map_gg.items()
                            if 'GERENTE' in cargo.upper() and leg != legajo_gg}
                rows_gg = []
                for legajo, cargo in sorted(legs_dir.items(), key=lambda x: x[1]):
                    fila = df[df['Legajo'].astype(str) == str(legajo)]
                    if fila.empty: continue
                    r      = fila.iloc[0]
                    col_n  = next((c for c in ['Nombre','Apellidos y Nombres'] if c in df.columns), None)
                    nombre_col = str(r.get(col_n,'')).title() if col_n else legajo
                    meta   = safe_float(r.get(col_meta_gg, 0))
                    prog   = min(safe_float(r.get('Prog_visma', 0)), meta)
                    pct    = round(prog/meta*100, 1) if meta > 0 else 0
                    dp_p   = max(0, int(meta - prog))
                    estado = str(r.get('Estado', ''))
                    rows_gg.append({
                        'Nombre':        nombre_col,
                        'Cargo':         cargo.title(),
                        'Meta (días)':   int(meta),
                        'Prog. (días)':  int(prog),
                        '% Avance':      f"{pct}%",
                        'Días x prog.':  dp_p,
                        'Estado':        emo(estado) if estado else ''
                    })
                if rows_gg:
                    rgg = pd.DataFrame(rows_gg).sort_values('Días x prog.', ascending=False)
                    st.dataframe(rgg, use_container_width=True, hide_index=True,
                                 height=min(50 + len(rgg)*38, 600))

        else:
            # ── Resto de roles: Top10 + Resumen ─────────────────────────────────
            col_l, col_r = st.columns([1.2,1])
            with col_l:
                st.markdown("### Top 10 áreas con mayor días pendientes por programar")
                if col_area and col_dp:
                    usar_grupo = (
                        role in ('RRHH', 'Gerente', 'SubGerente')
                        and 'Grupo' in df.columns
                        and df['Grupo'].notna().any()
                    )
                    col_top10 = 'Grupo' if usar_grupo else col_area
                    top10 = (df[df[col_dp].apply(safe_float)>0]
                             .groupby(col_top10)[col_dp]
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
                # Agrupacion del resumen segun rol:
                # RRHH/Gerente  -> por Jefe
                # SubGerente    -> por Jefe (sus reportes directos)
                # Jefe          -> por Administrador
                # Administrador -> resumen propio
                col_admin = 'Administrador' if 'Administrador' in df.columns else None
            # Buscar la primera columna jerarquica con datos reales en el df del usuario
            def mejor_col_resumen(candidatos):
                for c in candidatos:
                    if c and c in df.columns and df[c].notna().any():
                        return c
                return None

            if role == 'Administrador':
                grp_resumen    = None
                titulo_resumen = "### Resumen de tu area"
                lbl_col        = None
            elif role == 'GerenteGeneral':
                # GerenteGeneral: no muestra panel derecho en dashboard
                grp_resumen    = None
                titulo_resumen = ""
                lbl_col        = None
            elif role == 'Jefe':
                grp_resumen    = mejor_col_resumen(['Administrador', col_area])
                lbl_col        = grp_resumen
                titulo_resumen = "### Resumen por Administrador" if grp_resumen == 'Administrador' else "### Resumen por Area"
            else:
                # RRHH, Gerente, SubGerente: nivel inmediato inferior con datos
                grp_resumen = mejor_col_resumen(['Jefe', 'Administrador', col_area])
                if grp_resumen == 'Jefe':
                    titulo_resumen = "### Resumen por Jefe"
                    lbl_col        = 'Jefe'
                elif grp_resumen == 'Administrador':
                    titulo_resumen = "### Resumen por Administrador"
                    lbl_col        = 'Administrador'
                elif grp_resumen == col_area:
                    titulo_resumen = "### Resumen por Area"
                    lbl_col        = col_area
                else:
                    grp_resumen    = None
                    titulo_resumen = "### Resumen"
                    lbl_col        = None

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
        # Filtros - Gerencia usa Grupo (nombre agrupado) visible solo para niveles altos
        col_grp = 'Grupo' if 'Grupo' in df.columns else None
        if role in ('RRHH','Gerente','GerenteGeneral'):
            c1,c2,c3,c4,c5,c6 = st.columns(6)
            buscar = c1.text_input("🔍 Nombre o legajo")
            f_ger  = c2.selectbox("Gerencia", ['Todas']+sorted(df[col_grp].dropna().unique().tolist()) if col_grp else ['Todas'])
            f_area = c3.selectbox("Área",     ['Todas']+sorted(df[col_area].dropna().unique().tolist()) if col_area else ['Todas'])
            f_cat  = c4.selectbox("Categoría",['Todas']+sorted(df[col_cat].dropna().unique().tolist()) if col_cat else ['Todas'])
            f_est  = c5.selectbox("Estado",   ['Todos','VENCIDO','CRITICO','EN_RIESGO','SIN_SALDO'])
            f_jefe = c6.selectbox("Jefe",     ['Todos']+sorted(df[col_jefe].dropna().unique().tolist()) if col_jefe else ['Todos'])
        else:
            c1,c2,c3,c4,c5 = st.columns(5)
            buscar = c1.text_input("🔍 Nombre o legajo")
            f_ger  = 'Todas'
            f_area = c2.selectbox("Área",     ['Todas']+sorted(df[col_area].dropna().unique().tolist()) if col_area else ['Todas'])
            f_cat  = c3.selectbox("Categoría",['Todas']+sorted(df[col_cat].dropna().unique().tolist()) if col_cat else ['Todas'])
            f_est  = c4.selectbox("Estado",   ['Todos','VENCIDO','CRITICO','EN_RIESGO','SIN_SALDO'])
            f_jefe = c5.selectbox("Jefe",     ['Todos']+sorted(df[col_jefe].dropna().unique().tolist()) if col_jefe else ['Todos'])

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
        if f_ger !='Todas' and col_grp:  df_f=df_f[df_f[col_grp]==f_ger]
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

        # Selector de año para vencidos (solo RRHH puede ver otros años)
        anio_actual = date.today().year
        if role == 'RRHH':
            col_anio, _ = st.columns([1, 3])
            anio_venc = col_anio.selectbox(
                "📅 Ver vencimientos del año",
                options=list(range(anio_actual - 1, anio_actual + 2)),
                index=1,
                help="Filtra colaboradores cuyo periodo vence en este año"
            )
        else:
            anio_venc = anio_actual

        # Filtrar vencidos segun año seleccionado
        if anio_venc == anio_actual:
            # Año actual: usar Estado VENCIDO ya calculado
            df_v = df[df['Estado']=='VENCIDO'].copy() if 'Estado' in df.columns else df.head(0)
        elif anio_venc == anio_actual + 1:
            # Año siguiente: usar columna pre-calculada Venc_anio_sig
            df_v = df[df.get('Venc_anio_sig', pd.Series(0, index=df.index)) > 0].copy()                    if 'Venc_anio_sig' in df.columns else df.head(0)
            if not df_v.empty:
                df_v['Vencidos_real'] = df_v['Venc_anio_sig']
        else:
            # Año anterior: mostrar los que tuvieron vencidos (Estado VENCIDO es del año actual)
            df_v = df[df['Estado']=='VENCIDO'].copy() if 'Estado' in df.columns else df.head(0)

        df_c = df[df['Estado']=='CRITICO'].copy()   if 'Estado' in df.columns else df.head(0)
        df_r = df[df['Estado']=='EN_RIESGO'].copy() if 'Estado' in df.columns else df.head(0)

        c1,c2,c3 = st.columns(3)
        c1.metric(f"🔴 Vencidos {anio_venc}",   len(df_v))
        c2.metric("🟠 Críticos ≤30 días",        len(df_c))
        c3.metric("🟡 En riesgo ≤90 días",       len(df_r))

        # Columnas para alertas
        # VENCIDOS: no mostrar Dias_x_vencer (no aplica cuando ya esta vencido)
        cols_v = [c for c in ['Legajo',col_nom,col_area,col_jefe,'Vencidos_real',
                               col_dp,'Fecha límite','Estado','Comentario_ind']
                  if c and c in df_v.columns]
        # CRITICOS/EN_RIESGO: mostrar Dias_x_vencer (dias que faltan de la meta)
        cols_a = [c for c in ['Legajo',col_nom,col_area,col_jefe,'Vencidos_real',
                               col_dp,'Dias_x_vencer','Fecha límite',
                               'Estado','Comentario_ind']
                  if c and c in df.columns]
        rename_v = {
            col_dp:          'Días x programar',
            'Vencidos_real': 'Días vencidos',
        }
        rename_alertas = {
            col_dp:          'Días x programar',
            'Dias_x_vencer': 'Días x vencer',
            'Vencidos_real': 'Días vencidos',
        }

        st.markdown("---")
        st.markdown(f"### 🔴 Días vencidos {anio_venc} — riesgo de indemnización")
        if not df_v.empty:
            show = df_v[cols_v].copy().sort_values('Vencidos_real',ascending=False)
            show['Estado'] = show['Estado'].apply(emo) if 'Estado' in show.columns else show['Estado']
            show = show.rename(columns=rename_v)
            st.dataframe(show, use_container_width=True, hide_index=True)
            if role=='RRHH':
                st.markdown("**Marcar como ignorado** (ya gestionado por fuera):")
                ci1,ci2,ci3 = st.columns(3)
                leg_ig = ci1.text_input("Legajo",key="ig_in")
                if ci2.button("Ignorar") and leg_ig: ignorar(leg_ig); st.rerun()
                if ci3.button("Restaurar") and leg_ig: restaurar(leg_ig); st.rerun()
            st.download_button("⬇️ Descargar vencidos", to_excel(show),
                file_name=f"vencidos_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.success("Sin días vencidos en tu vista")

        # ── Ignorados: listado descargable ─────────────────────────────────────
        if role == 'RRHH':
            ignorados_set = get_ignorados()
            if ignorados_set:
                st.markdown("---")
                st.markdown("### 🚫 Ignorados — gestionados por fuera de RRHH")
                col_n_ig  = col_nom if col_nom else 'Nombre'
                cols_ig   = [c for c in ['Legajo', col_n_ig, col_area, col_jefe,
                                          'Vencidos_real', 'Fecha límite', 'Comentario_ind']
                             if c and c in df.columns]
                df_ig = df_full[df_full['Legajo'].astype(str).isin(ignorados_set)].copy()
                if not df_ig.empty and cols_ig:
                    show_ig = df_ig[[c for c in cols_ig if c in df_ig.columns]].copy()
                    st.dataframe(show_ig, use_container_width=True, hide_index=True)
                    st.download_button(
                        "⬇️ Descargar ignorados",
                        to_excel(show_ig),
                        file_name=f"ignorados_{date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info(f"{len(ignorados_set)} legajos ignorados (sin datos en el sistema)")
                    st.download_button(
                        "⬇️ Descargar legajos ignorados",
                        to_excel(pd.DataFrame({'Legajo': sorted(ignorados_set)})),
                        file_name=f"ignorados_{date.today()}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        st.markdown("### 🟠 Críticos — vencen en menos de 30 días")
        if not df_c.empty:
            show = df_c[cols_a].copy().sort_values('Fecha límite')
            show['Estado'] = show['Estado'].apply(emo)
            show = show.rename(columns=rename_alertas)
            st.dataframe(show, use_container_width=True, hide_index=True)
        else:
            st.success("Sin críticos")

        st.markdown("### 🟡 En riesgo — vencen en 30–90 días")
        if not df_r.empty:
            show = df_r[cols_a].copy().sort_values('Fecha límite')
            show['Estado'] = show['Estado'].apply(emo)
            show = show.rename(columns=rename_alertas)
            st.dataframe(show, use_container_width=True, hide_index=True)
            st.download_button("⬇️ Descargar en riesgo", to_excel(show),
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
            if role != 'GerenteGeneral':
                f_jefe_c = st.selectbox("Jefe",['Todos']+sorted(df[col_jefe].dropna().unique().tolist()) if col_jefe else ['Todos'])
                f_area_c = st.selectbox("Área",['Todas']+sorted(df[col_area].dropna().unique().tolist()) if col_area else ['Todas'])
            f_leg_c  = st.text_input("Buscar legajo o nombre", placeholder="Ej: 1000097727")

        # GerenteGeneral: filtrar solo sus directivos (puestos con GERENTE)
        if role == 'GerenteGeneral':
            puesto_map_cal = area_sistema.get('puesto', {}) if area_sistema else {}
            legajo_gg_cal  = str(pa.get(user_email, {}).get('legajo', ''))
            legs_dir_cal   = {str(leg) for leg, cargo in puesto_map_cal.items()
                              if 'GERENTE' in cargo.upper() and str(leg) != legajo_gg_cal}
            # Intentar desde df_full primero, sino construir desde legajos directivos
            df_cal = df_full[df_full['Legajo'].astype(str).isin(legs_dir_cal)].copy()
            # Si faltan directivos (no tienen area en META), agregarlos desde df_visma
            legs_en_cal = set(df_cal['Legajo'].astype(str).unique())
            legs_faltantes = legs_dir_cal - legs_en_cal
            if legs_faltantes and not df_visma.empty:
                vis_dir = df_visma[df_visma['Legajo'].astype(str).isin(legs_faltantes)][['Legajo','Apellidos y Nombre']].drop_duplicates('Legajo')
                for _, vr in vis_dir.iterrows():
                    fila_extra = pd.DataFrame([{
                        'Legajo': str(vr['Legajo']),
                        col_nom:  str(vr['Apellidos y Nombre']) if col_nom else '',
                    }])
                    df_cal = pd.concat([df_cal, fila_extra], ignore_index=True)
        else:
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
        # Para RRHH/Gerente: agrupar por Gerente o Sub_Gerente
        # Para SubGerente/Jefe/Admin: agrupar por Jefe (o Administrador si no hay Jefe)
        # Determinar columna de agrupacion segun rol
        def primera_col_con_datos(candidatos):
            for c in candidatos:
                if c and c in df.columns and df[c].notna().any():
                    return c
            return None

        if role == 'GerenteGeneral':
            # Ve resumen por Grupo (nombre agrupado de las areas)
            grp = primera_col_con_datos(['Grupo', 'Sub_Gerente', col_jefe])
        elif role in ('RRHH', 'Gerente'):
            grp = primera_col_con_datos(['Grupo', 'Jefe', 'Sub_Gerente', col_ger, col_jefe])
        elif role == 'SubGerente':
            # Usar Grupo para agrupar (mas legible que nombres de Jefe)
            grp = primera_col_con_datos(['Grupo', 'Jefe', 'Administrador', col_area, col_jefe])
        elif role == 'Jefe':
            grp = primera_col_con_datos(['Administrador', col_area, col_jefe])
        else:
            grp = primera_col_con_datos(['Administrador', col_area, col_jefe])
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
        # GerenteGeneral: solo reportes directos (SubGerentes + Jefe Auditoria)
        # Se filtra por cargos que contienen GERENTE o son jefes directos suyos
        if role == 'GerenteGeneral':
            # Solo reportes directos: colaboradores con cargo directivo
            # Filtrar por puesto en atributos del sistema
            puesto_map_hist = area_sistema.get('puesto', {}) if area_sistema else {}
            if puesto_map_hist:
                # Crear columna puesto temporal para filtrar
                df_hist_tmp = df.copy()
                df_hist_tmp['_puesto'] = df_hist_tmp['Legajo'].astype(str).map(puesto_map_hist).fillna('')
                mask_dir = df_hist_tmp['_puesto'].str.upper().str.contains(
                    'GERENTE|SUB.GERENTE|JEFE DE AUDITORIA|JEFE AUDITORIA|ASISTENTE DE GERENCIA', na=False
                )
                df_historial = df_hist_tmp[mask_dir]
            else:
                df_historial = df
            leg_ok = set(df_historial['Legajo'].astype(str).unique())
        else:
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
