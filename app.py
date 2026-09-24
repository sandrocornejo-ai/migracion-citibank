"""
app.py — Migración Citibank → Rex+ (Liquidaciones en detalle)

Uso local:
    streamlit run app.py
"""

import os
from datetime import datetime

import streamlit as st

from procesador import (
    cargar_equivalencias, cargar_consolidado, validar_estructura, procesar, a_excel,
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EQUIV_DEFAULT = os.path.join(BASE_DIR, 'data', 'Equivalencia columnas conceptos.xlsx')

st.set_page_config(page_title='Rex+ | Migración Citibank', page_icon='🏦', layout='wide')

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.header-bar {
    background: linear-gradient(135deg, #0f2d5e 0%, #1a4a8a 100%);
    padding: 18px 32px; border-radius: 12px; margin-bottom: 28px;
}
.header-bar h1 { color: white !important; font-size: 1.6rem; font-weight: 700; margin: 0; }
.header-bar span { color: #7eb3ff; font-size: 0.95rem; }
.step { font-size: 1.05rem; font-weight: 700; color: #0f2d5e; margin: 18px 0 8px 0; }
</style>
<div class="header-bar">
  <h1>🏦 Migración Citibank → Rex+</h1>
  <span>Consolidado de Haberes y Descuentos → Liquidaciones en detalle</span>
</div>
""", unsafe_allow_html=True)

# ── Barra lateral: tabla de equivalencias ──
with st.sidebar:
    st.markdown('### Tabla de equivalencias')
    equiv_file = st.file_uploader('Reemplazar (opcional)', type=['xlsx'], key='equiv',
                                  help='Si no subes una, se usa la que viene con el programa.')
    fuente_equiv = equiv_file if equiv_file else EQUIV_DEFAULT
    try:
        equiv, orden, tipos = cargar_equivalencias(fuente_equiv)
        st.success(f"{len(equiv)} columnas → {len(orden)} conceptos Rex+")
        st.caption('Usando: ' + (equiv_file.name if equiv_file else 'tabla incluida'))
    except Exception as e:
        st.error(f'No se pudo leer la tabla de equivalencias: {e}')
        st.stop()

# ── Paso 1: archivo ──
st.markdown('<div class="step">1 · Archivo a procesar (una empresa)</div>', unsafe_allow_html=True)
archivo = st.file_uploader('Consolidado de la empresa (.xlsx)', type=['xlsx'], key='consolidado')

# ── Paso 2: fase ──
st.markdown('<div class="step">2 · ¿Va a usar fase?</div>', unsafe_allow_html=True)
col_a, col_b = st.columns([1, 2])
with col_a:
    usa_fase = st.radio('Usar fase', ['No', 'Sí'], horizontal=True, label_visibility='collapsed') == 'Sí'
fase = None
with col_b:
    if usa_fase:
        fase = st.number_input('Número de fase', min_value=1, step=1, value=1)
    else:
        st.caption('La columna Fase no se incluirá en el archivo de salida.')

if not archivo:
    st.info('Sube el Consolidado de una empresa para comenzar.')
    st.stop()

try:
    df = cargar_consolidado(archivo)
except Exception as e:
    st.error(f'No se pudo leer el archivo: {e}')
    st.stop()

errores = validar_estructura(df)
if errores:
    for err in errores:
        st.error(err)
    st.stop()

# ── Paso 3: procesar ──
st.markdown('<div class="step">3 · Procesar</div>', unsafe_allow_html=True)
if st.button('Generar archivo de salida', type='primary'):
    with st.spinner('Procesando…'):
        out, log, res = procesar(df, equiv, orden, tipos, usa_fase=usa_fase,
                                 fase=int(fase) if fase else None)
    st.session_state['resultado'] = (out, log, res, archivo.name)

if 'resultado' in st.session_state and st.session_state['resultado'][3] == archivo.name:
    out, log, res, _ = st.session_state['resultado']

    c1, c2, c3, c4 = st.columns(4)
    c1.metric('Empresa', res['empresa'])
    c2.metric('Trabajadores', res['trabajadores'])
    c3.metric('Liquidaciones', res['liquidaciones'])
    c4.metric('Filas de salida', f"{res['filas']:,}".replace(',', '.'))
    st.caption('Meses: ' + ', '.join(f"{m} ({n})" for m, n in res['por_mes'].items()))

    n_err = int((log['Tipo'] == 'ERROR').sum()) if len(log) else 0
    n_cua = int((log['Tipo'] == 'CUADRATURA').sum()) if len(log) else 0
    n_adv = int((log['Tipo'] == 'ADVERTENCIA').sum()) if len(log) else 0
    if n_err:
        st.error(f'{n_err} columnas sin equivalencia (no se migran). Revisa el log.')
    if n_cua:
        st.warning(f'{n_cua} diferencias de cuadratura contra los totales del archivo. Revisa el log.')
    if n_adv:
        st.warning(f'{n_adv} advertencias. Revisa el log.')
    if not (n_err or n_cua or n_adv):
        st.success('Sin errores ni diferencias de cuadratura.')

    sello = datetime.now().strftime('%Y%m%d_%H%M%S')
    d1, d2 = st.columns(2)
    d1.download_button('⬇️ Descargar archivo de salida', a_excel(out),
                       file_name=f"migracion_citibank_{res['empresa']}_{sello}.xlsx",
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                       type='primary', use_container_width=True)
    if len(log):
        d2.download_button('⬇️ Descargar log de validaciones', a_excel(log, 'Log'),
                           file_name=f"log_citibank_{res['empresa']}_{sello}.xlsx",
                           mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                           use_container_width=True)

    tab1, tab2 = st.tabs(['Vista previa salida', 'Log'])
    with tab1:
        st.dataframe(out.head(500), use_container_width=True, hide_index=True)
    with tab2:
        st.dataframe(log, use_container_width=True, hide_index=True)
