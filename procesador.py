"""
procesador.py — Lógica de transformación Consolidado Citibank → Liquidaciones en detalle Rex+

Reglas según "Armado archivo de salida.docx":
  - Se lee una fila por trabajador/mes del Consolidado (una empresa por archivo).
  - Conceptos: encabezados desde la columna BK en adelante, traducidos con
    "Equivalencia columnas conceptos.xlsx". Montos en 0 se omiten, salvo los
    conceptos obligatorios (SIEMPRE_GENERAR).
  - Varias columnas que apuntan al mismo concepto Rex se suman (afp, apvi, ...),
    excepto isapre: si "Citización Salud Fonasa" > 0 se usa ese valor; si no,
    Obligatoria FO_6302 + Adicional FO_6304.
"""

import io
import re
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# ─────────────────────────────────────────────────────────────────
#  CONSTANTES / REGLAS DE NEGOCIO
# ─────────────────────────────────────────────────────────────────

COL_INICIO_CONCEPTOS = 62          # columna BK (0-based)

CONCEPTO_LICENCIA = 'licenciaDias'   # monto = DIAS LICENCIAS cuando > 0

SIEMPRE_GENERAR = ['sueldoBase', 'totalesEmpl', 'impuesto', 'cesEmpleado']

TIPOS_HABER_AFECTO = {'haber afecto', 'haber afecto especial', 'haber solo tributable'}
TIPOS_HABER_EXENTO = {'haber exento'}
TIPOS_DESCUENTO = {'descuento legal', 'descuento'}
TOLERANCIA = 1          # pesos, para cuadratura

# Isapre
COL_FONASA = 'Citización Salud Fonasa'
COLS_ISAPRE_NORMAL = [
    'Cotización Salud Obligatoria FO_6302 Descuento Legal',
    'Cotización Salud Adicional FO_6304 Descuento Legal',
]
COL_FONDO_SOLIDARIO = 'Cotización Seguro Cesantia Fondo Solidario FO_6308 Cálculo'

# Afecto
AFECTO_IMPONIBLE = {'afp', 'isapre', 'aporteAFPemp', 'mutual', 'aporteFAPPCEV', 'sis', 'cajaComp'}
AFECTO_CESANTIA = {'cesEmpleado', 'cesAporteCi', 'cesAporteSol'}

# Id de institución
INST_CCAF = {'cajaCred', 'cajaLeas', 'cajaComp'}
PREFIJO_APV = 'apv'   # apvi → 'apv' + Id institución de afp (ID AFP)
INST_AFP = {'afp', 'cesEmpleado', 'aporteAFPemp', 'cesAporteCi', 'cesAporteSol', 'sis'}

# Cotización de jubilación — valores fijos
COTIZ_FIJA = {'cesEmpleado': 0.6, 'aporteAFPemp': 0.1, 'aporteFAPPCEV': 0.9}

# Columnas de entrada requeridas (sin conceptos)
COLS_REQUERIDAS = [
    'Mes', 'ID EMPRESA', 'CCAF', 'APORTE CAJA', 'MUTUAL', '%MUTUAL', 'RUT',
    'ID AFP', 'PORC AFP', 'SIS', 'DIAS TRABAJADOS', 'DIAS VACACIONES',
    'DIAS LICENCIAS', 'ULT IMP', 'SUELDO ORIGINAL', 'IMPONIBLE',
    'BASE TRIBUTABLE', 'REBAJA LLSS', 'TOPE IMPONIBLE AFC TRABAJADOR',
]
COLS_ID_SALUD = ['ID SALUD', 'ID SDALUD']   # International viene con el typo

COLUMNAS_SALIDA = [
    'Fecha de proceso', 'Id empleado', 'Número de contrato', 'Id del concepto',
    'Monto del concepto', 'Afecto', 'Id de institución', 'Cotización de jubilación',
    'Días de licencias', 'Días trabajados', 'Fecha de aplicación', 'Empresa',
    'Total de rebajas por LLSS', 'Rentas no gravadas', 'Rebaja por zona extrema',
    'Jornada', 'Días de vacaciones', 'Monto Init', 'Fase', 'Parcial 7', 'Parcial 8',
]


# ─────────────────────────────────────────────────────────────────
#  AUXILIARES
# ─────────────────────────────────────────────────────────────────

def limpiar_texto(v):
    """Quita espacios, incluidos \\xa0. None/NaN → ''."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    return str(v).replace('\xa0', ' ').strip()


def num(v):
    """Convierte a float; None/NaN/texto inválido → 0."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return 0.0 if pd.isna(v) else float(v)
    s = limpiar_texto(v).replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0


def esta_vacio(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or limpiar_texto(v) == ''


def entero(v):
    return int(round(num(v)))


def normalizar_mes(v):
    """'2026-01', datetime o '01-2026' → '2026-01'."""
    if isinstance(v, (datetime, pd.Timestamp)):
        return v.strftime('%Y-%m')
    s = limpiar_texto(v)
    m = re.match(r'^(\d{4})[-/](\d{1,2})', s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.match(r'^(\d{1,2})[-/](\d{4})$', s)
    if m:
        return f"{m.group(2)}-{int(m.group(1)):02d}"
    return s


def normalizar_rut(v):
    s = limpiar_texto(v).replace('.', '').upper()
    if '-' in s:
        cuerpo, dv = s.rsplit('-', 1)
        return f"{cuerpo.lstrip('0')}-{dv}"
    return s


def rut_valido(rut):
    if '-' not in rut:
        return False
    cuerpo, dv = rut.split('-', 1)
    if not cuerpo.isdigit():
        return False
    total, factor = 0, 2
    for c in reversed(cuerpo):
        total += int(c) * factor
        factor = 2 if factor == 7 else factor + 1
    r = 11 - total % 11
    esperado = '0' if r == 11 else 'K' if r == 10 else str(r)
    return esperado == dv


# ─────────────────────────────────────────────────────────────────
#  CARGA
# ─────────────────────────────────────────────────────────────────

def cargar_equivalencias(archivo):
    """Devuelve (equiv, orden, tipo_concepto).
    equiv: {columna_entrada: id_concepto}
    orden: lista de id_concepto en el orden de la tabla (sin repetir)
    tipo_concepto: {id_concepto: tipo en minúsculas}
    """
    df = pd.read_excel(archivo, dtype=str)
    df.columns = [limpiar_texto(c) for c in df.columns]
    equiv, orden, tipo_concepto = {}, [], {}
    for _, r in df.iterrows():
        col = limpiar_texto(r.get('Columna de entrada'))
        idc = limpiar_texto(r.get('Id concepto Rex'))
        if not col or not idc:
            continue
        equiv[col] = idc
        tipo_concepto.setdefault(idc, limpiar_texto(r.get('Tipo')).lower())
        if idc not in orden:
            orden.append(idc)
    return equiv, orden, tipo_concepto


def cargar_consolidado(archivo):
    """Lee la primera hoja del Consolidado. Descarta columnas sin encabezado."""
    df = pd.read_excel(archivo, sheet_name=0)
    df.columns = [limpiar_texto(c) for c in df.columns]
    return df


# ─────────────────────────────────────────────────────────────────
#  PROCESO
# ─────────────────────────────────────────────────────────────────

def validar_estructura(df):
    """Errores bloqueantes de estructura. Lista vacía = OK."""
    errores = [f"Falta la columna '{c}'" for c in COLS_REQUERIDAS if c not in df.columns]
    if not any(c in df.columns for c in COLS_ID_SALUD):
        errores.append("Falta la columna 'ID SALUD'")
    if len(df.columns) <= COL_INICIO_CONCEPTOS:
        errores.append("El archivo no tiene columnas de conceptos desde la columna BK")
    if not errores:
        empresas = df['ID EMPRESA'].dropna().astype(str).str.strip().unique()
        if len(empresas) > 1:
            errores.append(f"El archivo trae más de una empresa ({', '.join(empresas)}). "
                           "Procese una empresa a la vez.")
    return errores


def procesar(df, equiv, orden, tipo_concepto, usa_fase=False, fase=None):
    """Genera las filas de salida y el log.
    Retorna (df_salida, df_log, resumen).
    """
    col_salud = next(c for c in COLS_ID_SALUD if c in df.columns)

    # Columnas de concepto (desde BK, con encabezado)
    cols_concepto = [c for c in df.columns[COL_INICIO_CONCEPTOS:]
                     if c and not c.startswith('Unnamed')]

    log = []
    sin_equiv = [c for c in cols_concepto if c not in equiv]
    for c in sin_equiv:
        log.append({'Tipo': 'ERROR', 'Mes': '', 'RUT': '', 'Detalle':
                    f"Columna sin equivalencia (se omite): {c}"})

    # concepto → columnas de entrada
    cols_por_concepto = {}
    for c in cols_concepto:
        if c in equiv:
            cols_por_concepto.setdefault(equiv[c], []).append(c)

    conceptos_orden = [c for c in orden if c in cols_por_concepto]
    for c in SIEMPRE_GENERAR:
        if c not in conceptos_orden:
            conceptos_orden.append(c)
    # licenciaDias: se genera cuando DIAS LICENCIAS > 0 (va primero, como dato)
    conceptos_orden.insert(0, CONCEPTO_LICENCIA)

    afectos = {c for c, t in tipo_concepto.items() if t in TIPOS_HABER_AFECTO}
    exentos = {c for c, t in tipo_concepto.items() if t in TIPOS_HABER_EXENTO}
    descuentos = {c for c, t in tipo_concepto.items()
                  if t in TIPOS_DESCUENTO and c != 'totalesEmpl'}

    filas = []
    cuadratura = []
    ruts_invalidos = set()
    for idx, r in df.iterrows():
        mes = normalizar_mes(r.get('Mes'))
        rut = normalizar_rut(r.get('RUT'))
        if not mes or not rut:
            continue
        if not rut_valido(rut) and rut not in ruts_invalidos:
            ruts_invalidos.add(rut)
            log.append({'Tipo': 'ADVERTENCIA', 'Mes': mes, 'RUT': rut,
                        'Detalle': 'RUT con dígito verificador inválido'})

        # ── Montos por concepto ──
        montos = {}
        for idc, cols in cols_por_concepto.items():
            if idc == 'isapre':
                fonasa = num(r.get(COL_FONASA)) if COL_FONASA in cols else 0
                if fonasa > 0:
                    montos[idc] = fonasa
                else:
                    montos[idc] = sum(num(r.get(c)) for c in cols if c in COLS_ISAPRE_NORMAL)
            else:
                montos[idc] = sum(num(r.get(c)) for c in cols)
            if montos[idc] < 0:
                log.append({'Tipo': 'ADVERTENCIA', 'Mes': mes, 'RUT': rut,
                            'Detalle': f"Monto negativo en {idc}: {montos[idc]:.0f}"})

        # Días de licencia como concepto (no suma en haberes ni descuentos)
        montos[CONCEPTO_LICENCIA] = max(num(r.get('DIAS LICENCIAS')), 0)

        suma_afectos = sum(v for k, v in montos.items() if k in afectos)
        suma_exentos = sum(v for k, v in montos.items() if k in exentos)

        # ── Cuadratura simple: haberes - descuentos (sin líquido) vs líquido ──
        suma_desc = sum(v for k, v in montos.items() if k in descuentos)
        liquido = montos.get('totalesEmpl', 0)
        calculado = suma_afectos + suma_exentos - suma_desc
        diferencia = calculado - liquido
        cuadratura.append({
            'Mes': mes, 'RUT': rut,
            'Total haberes': entero(suma_afectos + suma_exentos),
            'Total descuentos': entero(suma_desc),
            'Haberes - Descuentos': entero(calculado),
            'Sueldo líquido': entero(liquido),
            'Diferencia': entero(diferencia),
            'Estado': 'OK' if abs(diferencia) <= TOLERANCIA else 'DIFERENCIA',
        })

        imponible = entero(r.get('IMPONIBLE'))
        tope_afc = r.get('TOPE IMPONIBLE AFC TRABAJADOR')
        if not esta_vacio(tope_afc):
            afecto_ces = entero(tope_afc)
        elif num(r.get(COL_FONDO_SOLIDARIO)) > 0:
            afecto_ces = entero(num(r.get(COL_FONDO_SOLIDARIO)) / 0.008)
        else:
            afecto_ces = 0

        comunes = {
            'Fecha de proceso': mes,
            'Id empleado': rut,
            'Número de contrato': 1,
            'Días de licencias': entero(r.get('DIAS LICENCIAS')),
            'Días trabajados': entero(r.get('DIAS TRABAJADOS')),
            'Fecha de aplicación': mes,
            'Empresa': limpiar_texto(r.get('ID EMPRESA')),
            'Rebaja por zona extrema': 0,
            'Jornada': 'C',
            'Días de vacaciones': entero(r.get('DIAS VACACIONES')),
        }

        for idc in conceptos_orden:
            monto = montos.get(idc, 0)
            if monto == 0 and idc not in SIEMPRE_GENERAR:
                continue

            # Afecto
            if idc in AFECTO_IMPONIBLE:
                afecto = imponible
            elif idc == 'impuesto':
                afecto = entero(r.get('BASE TRIBUTABLE'))
            elif idc in AFECTO_CESANTIA:
                afecto = afecto_ces
            elif idc == 'totalesEmpl':
                afecto = entero(suma_afectos)
            else:
                afecto = 0

            # Id de institución
            if idc in INST_CCAF:
                inst = limpiar_texto(r.get('CCAF'))
            elif idc in INST_AFP:
                inst = limpiar_texto(r.get('ID AFP'))
            elif idc == 'apvi':
                inst = PREFIJO_APV + limpiar_texto(r.get('ID AFP'))
            elif idc == 'isapre':
                inst = limpiar_texto(r.get(col_salud))
            elif idc == 'impuesto':
                inst = 'Impuesto'
            elif idc == 'mutual':
                inst = limpiar_texto(r.get('MUTUAL'))
            elif idc == 'aporteFAPPCEV':
                inst = 'seguridadsocial'
            else:
                inst = ''

            # Cotización de jubilación
            if idc == 'afp':
                cotiz = num(r.get('PORC AFP'))
            elif idc == 'isapre':
                cotiz = entero(monto)
            elif idc in COTIZ_FIJA:
                cotiz = COTIZ_FIJA[idc]
            elif idc == 'mutual':
                cotiz = num(r.get('%MUTUAL'))
            elif idc == 'cajaComp':
                cotiz = num(r.get('APORTE CAJA'))
            elif idc == 'sis':
                cotiz = num(r.get('SIS'))
            elif idc == 'totalesEmpl':
                cotiz = entero(suma_afectos)
            else:
                cotiz = 0

            fila = dict(comunes)
            fila.update({
                'Id del concepto': idc,
                'Monto del concepto': entero(monto),
                'Afecto': afecto,
                'Id de institución': inst,
                'Cotización de jubilación': round(cotiz, 4),
                'Total de rebajas por LLSS': entero(r.get('REBAJA LLSS')) if idc == 'impuesto' else 0,
                'Rentas no gravadas': entero(suma_exentos) if idc == 'impuesto' else 0,
                'Monto Init': entero(r.get('SUELDO ORIGINAL')) if idc == 'sueldoBase' else 0,
                'Fase': fase if usa_fase else None,
                'Parcial 7': entero(r.get('ULT IMP')) if idc in ('mutual', 'sis') else 0,
                'Parcial 8': afecto_ces if idc in ('cesAporteSol', 'cesAporteCi') else 0,
            })
            filas.append(fila)

    columnas = [c for c in COLUMNAS_SALIDA if usa_fase or c != 'Fase']
    df_out = pd.DataFrame(filas, columns=COLUMNAS_SALIDA)[columnas]
    df_log = pd.DataFrame(log, columns=['Tipo', 'Mes', 'RUT', 'Detalle'])
    df_cuad = pd.DataFrame(cuadratura, columns=['Mes', 'RUT', 'Total haberes', 'Total descuentos',
                                                'Haberes - Descuentos', 'Sueldo líquido',
                                                'Diferencia', 'Estado'])

    resumen = {
        'empresa': limpiar_texto(df['ID EMPRESA'].dropna().iloc[0]) if df['ID EMPRESA'].notna().any() else '',
        'meses': sorted(df_out['Fecha de proceso'].unique().tolist()),
        'trabajadores': df_out['Id empleado'].nunique(),
        'liquidaciones': df_out[['Fecha de proceso', 'Id empleado']].drop_duplicates().shape[0],
        'filas': len(df_out),
        'por_mes': df_out.groupby('Fecha de proceso')['Id empleado'].nunique().to_dict(),
        'sin_equivalencia': sin_equiv,
        'cuadra_ok': int((df_cuad['Estado'] == 'OK').sum()),
        'cuadra_dif': int((df_cuad['Estado'] != 'OK').sum()),
    }
    return df_out, df_log, df_cuad, resumen


# ─────────────────────────────────────────────────────────────────
#  EXPORTACIÓN
# ─────────────────────────────────────────────────────────────────

COLS_TEXTO = {'Mes', 'RUT', 'Estado', 'Fecha de proceso', 'Id empleado', 'Id del concepto', 'Id de institución',
              'Fecha de aplicación', 'Empresa', 'Jornada'}


def a_excel(df, hoja='Liquidaciones'):
    """DataFrame → bytes .xlsx con encabezado formateado y números sin separador de miles."""
    return a_excel_hojas({hoja: df})


def a_excel_hojas(hojas):
    """{nombre_hoja: DataFrame} → bytes .xlsx."""
    wb = Workbook()
    wb.remove(wb.active)
    for hoja, df in hojas.items():
        _escribir_hoja(wb.create_sheet(hoja), df)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _escribir_hoja(ws, df):
    ws.append(list(df.columns))
    for c in ws[1]:
        c.font = Font(bold=True, color='FFFFFF')
        c.fill = PatternFill('solid', fgColor='1E5591')
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for fila in df.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in fila])
    for i, col in enumerate(df.columns, start=1):
        letra = ws.cell(row=1, column=i).column_letter
        ws.column_dimensions[letra].width = max(12, min(28, len(col) + 2))
        if col not in COLS_TEXTO:
            for cell in ws[letra][1:]:
                if isinstance(cell.value, float) and not cell.value.is_integer():
                    cell.number_format = '0.####'
                else:
                    cell.number_format = '0'
    ws.freeze_panes = 'A2'
    if 'Estado' in df.columns:
        rojo = PatternFill('solid', fgColor='F8D7DA')
        col = list(df.columns).index('Estado') + 1
        for fila in ws.iter_rows(min_row=2):
            if fila[col - 1].value == 'DIFERENCIA':
                for c in fila:
                    c.fill = rojo
