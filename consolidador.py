"""
consolidador.py — Une los 5 archivos fuente y genera el consolidado de vacaciones
Uso: python consolidador.py
"""
import pandas as pd
import re
import json
from datetime import date, datetime
from pathlib import Path

# ── Configuración ──────────────────────────────────────────────────────────────
MES_ACTUAL = "ABRIL"   # Cambiar cada mes
ANIO       = 2026
MESES      = ['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO',
               'JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']
MESES_NUM  = {m:i+1 for i,m in enumerate(MESES)}

def cargar_jerarquia():
    """Carga los lookups de jerarquía desde los CSV generados"""
    with open('person_access.json', encoding='utf-8') as f:
        person_access = json.load(f)
    with open('area_to_admin.json', encoding='utf-8') as f:
        area_to_admin = json.load(f)
    with open('cargo_to_cat.json', encoding='utf-8') as f:
        cargo_to_cat = json.load(f)
    return person_access, area_to_admin, cargo_to_cat

def limpiar_legajo(s):
    return str(s).replace('.0','').strip() if pd.notna(s) else ''

def extraer_fecha_limite(comentario):
    """Extrae fecha límite de 'DEBE GOZAR X DÍAS ANTES DEL DD/MM/AAAA'"""
    if not comentario or str(comentario).strip() in ['-','nan','None','']:
        return None, 0
    match_f = re.search(r'ANTES DEL (\d{2}/\d{2}/\d{4})', str(comentario))
    match_d = re.search(r'DEBE GOZAR (\d+)', str(comentario))
    fecha   = datetime.strptime(match_f.group(1), '%d/%m/%Y').date() if match_f else None
    dias    = int(match_d.group(1)) if match_d else 0
    return fecha, dias

def calcular_estado(fecha_limite, dias_pendientes, programacion_total, meta):
    """Semáforo: VENCIDO / CRITICO / EN_RIESGO / AL_DIA / SIN_SALDO / CUMPLIDO"""
    hoy = date.today()
    if meta == 0 and dias_pendientes == 0:
        return 'SIN_SALDO'
    if meta > 0 and programacion_total >= meta:
        return 'CUMPLIDO'
    if fecha_limite:
        dias_rest = (fecha_limite - hoy).days
        if dias_rest < 0 and dias_pendientes > 0:
            return 'VENCIDO'
        if dias_rest <= 30:
            return 'CRITICO'
        if dias_rest <= 90:
            return 'EN_RIESGO'
    if dias_pendientes > 20:
        return 'EN_RIESGO'
    return 'AL_DIA'

def consolidar(
    path_vacaciones:  str,
    path_empleados:   str,
    path_atributos:   str,
    path_meta:        str,
    path_altas_bajas: str,
) -> pd.DataFrame:

    hoy = date.today()
    person_access, area_to_admin, cargo_to_cat = cargar_jerarquia()

    # ── 1. VACACIONES VISMA ─────────────────────────────────────────────────────
    print("Cargando vacaciones...")
    vac = pd.read_excel(path_vacaciones, header=2)
    vac.columns = ['Legajo','_','Nombre','Estado','Fecha_desde','Fecha_hasta',
                   'Cant_dias','Cant_dias_filtro','Tipo_dia','Periodo','Origen',
                   'Estado_ausencia','Anticipo']
    vac = vac[vac['Legajo'].notna() & (vac['Legajo'] != 'Legajo')].copy()
    vac['Legajo']      = vac['Legajo'].apply(limpiar_legajo)
    vac['Fecha_desde'] = pd.to_datetime(vac['Fecha_desde'], dayfirst=True, errors='coerce')
    vac['Cant_dias']   = pd.to_numeric(vac['Cant_dias'], errors='coerce').fillna(0)

    # Solo 2026, estados Aprobada + Pendiente
    vac2026 = vac[
        (vac['Fecha_desde'].dt.year == ANIO) &
        (vac['Estado_ausencia'].isin(['Aprobada','Pendiente']))
    ].copy()
    vac2026['mes_num'] = vac2026['Fecha_desde'].dt.month

    # Pivot: días por mes
    pivot = vac2026.pivot_table(
        index='Legajo', columns='mes_num', values='Cant_dias',
        aggfunc='sum', fill_value=0
    ).reset_index()
    for m in range(1,13):
        if m not in pivot.columns:
            pivot[m] = 0
    pivot.columns = ['Legajo'] + [MESES[m-1] for m in range(1,13)]
    pivot['Programacion'] = pivot[[MESES[m-1] for m in range(1,13)]].sum(axis=1)

    # ── 2. EMPLEADOS VISMA ──────────────────────────────────────────────────────
    print("Cargando empleados...")
    emp = pd.read_excel(path_empleados)
    emp['Legajo']       = emp['Legajo'].apply(limpiar_legajo)
    emp['Nombre_completo'] = (emp['Apellido'].fillna('') + ' ' + emp['Nombre'].fillna('')).str.strip()
    emp_activos = emp[emp['Est Emp'] == 'Activo'][['Legajo','Nombre_completo','Fec Ing','Nro Docu','Puesto','Correo Empresa']].copy()
    emp_activos.columns = ['Legajo','Nombre','Fec_ingreso','DNI','Cargo','Correo']

    # ── 3. ATRIBUTOS VISMA ──────────────────────────────────────────────────────
    print("Cargando atributos...")
    atr = pd.read_excel(path_atributos, header=0)
    atr.columns = ['Legajo','Nombre_atr','Tipo','Atributo','Fecha_desde_atr','Fecha_hasta_atr']
    atr = atr[atr['Legajo'] != 'Legajo'].copy()
    atr['Legajo'] = atr['Legajo'].apply(limpiar_legajo)
    atr['Fecha_desde_atr'] = pd.to_datetime(atr['Fecha_desde_atr'], dayfirst=True, errors='coerce')

    def get_ultimo_atributo(tipo):
        """Obtiene el atributo más reciente por tipo para cada legajo"""
        sub = atr[atr['Tipo'].str.upper() == tipo].copy()
        sub = sub.sort_values('Fecha_desde_atr', ascending=False)
        return sub.drop_duplicates('Legajo')[['Legajo','Atributo']].rename(columns={'Atributo': tipo})

    areas_df    = get_ultimo_atributo('AREA')
    sedes_df    = get_ultimo_atributo('SEDE')
    planilla_df = get_ultimo_atributo('PLANILLA')

    # ── 4. META ANUAL ───────────────────────────────────────────────────────────
    print("Cargando meta...")
    meta_df = pd.read_excel(path_meta, sheet_name='Consolidado')
    meta_df['Legajo'] = meta_df['Legajo'].apply(limpiar_legajo)
    meta_cols = ['Legajo','Vencidos','Pendientes','Truncos','Total',
                 'Meta SE','COMENTARIO PARA EVITAR INDEMNIZACION',
                 'COMENTARIOS PARA CUMPLIMIENTO META 2026']
    meta_cols_exist = [c for c in meta_cols if c in meta_df.columns]
    meta_clean = meta_df[meta_cols_exist].copy()
    meta_clean.columns = ['Legajo','Vencidos_excel','Pendientes','Truncos','Total',
                           'Meta2026','Comentario_ind','Comentario_meta'][:len(meta_cols_exist)]

    # ── 5. ALTAS Y BAJAS ───────────────────────────────────────────────────────
    print("Cargando altas y bajas...")
    ab = pd.read_excel(path_altas_bajas, header=0)
    ab.columns = ['Legajo','Apellidos','Nombres','Estado','Fecha_alta','Fecha_baja',
                  'Causa','Salario','Vacaciones','Indemnizacion','Real']
    ab = ab[ab['Legajo'] != 'Legajo'].copy()
    ab['Legajo']     = ab['Legajo'].apply(limpiar_legajo)
    ab['Fecha_baja'] = pd.to_datetime(ab['Fecha_baja'], dayfirst=True, errors='coerce')

    # Cesados en 2026 con vacaciones a liquidar
    cesados_2026 = ab[
        (ab['Fecha_baja'].dt.year == ANIO) &
        (ab['Estado'] == 'Inactivo') &
        (ab['Vacaciones'] == 'Si')
    ][['Legajo','Fecha_baja']].copy()
    cesados_2026['mes_cese'] = cesados_2026['Fecha_baja'].dt.month

    # ── CONSOLIDAR ─────────────────────────────────────────────────────────────
    print("Consolidando...")
    df = emp_activos.copy()

    # Merge vacaciones
    df = df.merge(pivot, on='Legajo', how='left')
    for mes in MESES:
        df[mes] = df[mes].fillna(0)
    df['Programacion'] = df['Programacion'].fillna(0)

    # Merge atributos
    df = df.merge(areas_df,    on='Legajo', how='left')
    df = df.merge(sedes_df,    on='Legajo', how='left')
    df = df.merge(planilla_df, on='Legajo', how='left')

    # Merge meta
    df = df.merge(meta_clean, on='Legajo', how='left')

    # Ajuste cesados: sus días pendientes van al mes de cese
    for _, row in cesados_2026.iterrows():
        leg = row['Legajo']
        mes_col = MESES[row['mes_cese']-1]
        mask = df['Legajo'] == leg
        if mask.any() and mes_col in df.columns:
            pendientes_val = df.loc[mask, 'Pendientes'].fillna(0).values[0]
            df.loc[mask, mes_col] = df.loc[mask, mes_col] + pendientes_val
            df.loc[mask, 'Programacion'] = df.loc[mask, 'Programacion'] + pendientes_val

    # Categoría por cargo
    df['Categoria'] = df['Cargo'].map(cargo_to_cat).fillna('BACK OFFICE')

    # Jerarquía por área
    df['Administrador'] = df['AREA'].map(area_to_admin)

    area_to_jefe      = {}
    area_to_subger    = {}
    area_to_gerente   = {}
    with open('areas.csv', encoding='utf-8') as f:
        import csv
        reader = csv.DictReader(f)
        for row in reader:
            adm = row['Administrador'].strip()
            area = row['Area'].strip()
            if adm in person_access:
                # look up chain in jerarquia.csv
                pass

    jer_df = pd.read_csv('jerarquia.csv')
    jer_df = jer_df.apply(lambda x: x.str.strip() if x.dtype=='object' else x)
    area_chain = {}
    for _, r in jer_df.iterrows():
        for _, area_row in pd.read_csv('areas.csv').iterrows():
            if area_row['Administrador'].strip() == str(r['Administrador']).strip():
                area_chain[area_row['Area'].strip()] = {
                    'gerente': r['Gerente'],
                    'sub_gerente': r['Sub_Gerente'],
                    'jefe': r['Jefe'],
                    'administrador': r['Administrador']
                }

    df['Gerente']     = df['AREA'].map({a:v['gerente'] for a,v in area_chain.items()})
    df['Sub_Gerente'] = df['AREA'].map({a:v['sub_gerente'] for a,v in area_chain.items()})
    df['Jefe']        = df['AREA'].map({a:v['jefe'] for a,v in area_chain.items()})

    # Calcular vencidos reales y estado
    def calc_vencidos_estado(row):
        comentario = row.get('Comentario_ind','')
        fecha_lim, dias_riesgo = extraer_fecha_limite(comentario)
        pendientes  = float(row.get('Pendientes',0) or 0)
        programado  = float(row.get('Programacion',0) or 0)
        meta        = float(row.get('Meta2026',0) or 0)

        # Vencidos reales: fecha ya pasó y tiene pendientes
        vencidos_real = 0
        if fecha_lim and fecha_lim < hoy and pendientes > 0:
            # Verificar si gozó los días en cuestión
            gozados = programado
            if gozados < dias_riesgo:
                vencidos_real = dias_riesgo - gozados

        estado = calcular_estado(fecha_lim, pendientes, programado, meta)
        pct    = round(programado/meta*100, 1) if meta > 0 else 0
        dias_x_prog = max(0, meta - programado)

        return pd.Series({
            'Vencidos_real': vencidos_real,
            'Estado': estado,
            'Pct_avance': min(pct, 999),
            'Dias_x_programar': dias_x_prog,
            'Fecha_limite': str(fecha_lim) if fecha_lim else ''
        })

    extras = df.apply(calc_vencidos_estado, axis=1)
    df = pd.concat([df, extras], axis=1)

    # Orden final de columnas
    cols_base = ['Legajo','Nombre','DNI','Cargo','Categoria',
                 'Fec_ingreso','AREA','SEDE','PLANILLA',
                 'Gerente','Sub_Gerente','Jefe','Administrador']
    cols_meta = ['Vencidos_real','Pendientes','Truncos','Total','Meta2026',
                 'Programacion','Pct_avance','Dias_x_programar','Estado',
                 'Fecha_limite','Comentario_ind','Comentario_meta']
    cols_meses = MESES

    df_final = df[[c for c in cols_base+cols_meta+cols_meses if c in df.columns]].copy()

    print(f"\n✓ Consolidado generado: {len(df_final)} colaboradores")
    print(f"  Vencidos reales: {(df_final['Vencidos_real']>0).sum()}")
    print(f"  En riesgo: {(df_final['Estado']=='EN_RIESGO').sum()}")
    print(f"  Cumplido: {(df_final['Estado']=='CUMPLIDO').sum()}")

    return df_final

if __name__ == '__main__':
    df = consolidar(
        path_vacaciones  = 'Vacaciones_-_Dias_solicitados__28_.xlsx',
        path_empleados   = 'Exportación_empleados__75_.xlsx',
        path_atributos   = 'Atributos_de_estructuras_por_colaborador__56_.xlsx',
        path_meta        = 'META_2026_-_Abril.xlsx',
        path_altas_bajas = 'Altas_y_bajas_de_colaboradores__27_.xlsx',
    )
    df.to_excel('CONSOLIDADO_GENERADO.xlsx', index=False)
    print("✓ CONSOLIDADO_GENERADO.xlsx guardado")
