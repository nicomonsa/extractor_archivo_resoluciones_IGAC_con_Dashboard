#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
================================================================================
 GENERADOR DE DASHBOARD + BATCH
================================================================================

QUE HACE
    1. Busca TODAS las carpetas de procesamiento en la carpeta de Salida
    2. Genera un dashboard HTML para CADA procesamiento
    3. Crea un índice (index.html) con todos los procesamientos disponibles
    4. Si se especifica un procesamiento específico, genera solo ese

COMO SE USA
    # Para TODOS los procesamientos (BATCH)
    python dashboard_batch.py

    # Para UN procesamiento específico
    python dashboard_batch.py "ruta/a/procesamiento_20260107_1430"

    # Para usar rutas personalizadas
    python dashboard_batch.py --entrada "ruta/salida" --salida "ruta/dashboard"

NOTA
    Este script NO modifica el parser. Solo lee los Excel generados por él.

CORRECCIONES APLICADAS
    1. Las llaves del CSS del índice estaban sencillas y .format() las tomaba
       como campos a reemplazar, lo que rompía la generación con KeyError.
       Ahora están duplicadas, igual que en la plantilla del dashboard.
    2. Las alertas sin predio asociado mostraban "nan". Ahora muestran "-".
    3. Las alertas se ordenan por nivel antes de recortar, para que arriba
       queden las que exigen acción y no las informativas.
================================================================================
"""

import os
import sys
import glob
import argparse
import pandas as pd
from datetime import datetime
from pathlib import Path

# =============================================================================
# CONFIGURACIÓN - AJUSTA ESTAS RUTAS POR DEFECTO
# =============================================================================

RUTA_SALIDA = r"D:\ESCRITORIO POR D\SEC_PLANEACION\SEPTIEMBRE\CATASTRO\ACTUALIZACION\RESOLUCIONES\SCRIPTS\Salida"
RUTA_DASHBOARD = r"D:\ESCRITORIO POR D\SEC_PLANEACION\SEPTIEMBRE\CATASTRO\ACTUALIZACION\RESOLUCIONES\SCRIPTS\Dashboard"

# =============================================================================
# PLANTILLA HTML DEL DASHBOARD
# =============================================================================

HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard - Resoluciones Catastrales</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/dataTables.bootstrap5.min.css">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {{ background: #f0f2f5; font-family: 'Segoe UI', sans-serif; }}
        .header {{ background: linear-gradient(135deg, #0a1628, #1a3a5c); color: white; padding: 2rem 0; margin-bottom: 2rem; }}
        .stat-card {{ background: white; border-radius: 12px; padding: 1.5rem; box-shadow: 0 2px 8px rgba(0,0,0,0.08); height: 100%; transition: transform 0.2s; }}
        .stat-card:hover {{ transform: translateY(-3px); }}
        .stat-card .number {{ font-size: 2.2rem; font-weight: 700; }}
        .stat-card .label {{ color: #6c757d; font-size: 0.8rem; text-transform: uppercase; letter-spacing: 0.5px; }}
        .stat-card .icon {{ font-size: 1.8rem; opacity: 0.4; margin-bottom: 0.3rem; }}
        .stat-card.primary .number {{ color: #0d6efd; }}
        .stat-card.success .number {{ color: #198754; }}
        .stat-card.danger .number {{ color: #dc3545; }}
        .stat-card.warning .number {{ color: #fd7e14; }}
        .stat-card.info .number {{ color: #0dcaf0; }}
        .section-title {{ margin: 2.5rem 0 1.5rem 0; padding-bottom: 0.5rem; border-bottom: 3px solid #0d6efd; }}
        .section-title h2 {{ font-weight: 300; }}
        .table-container {{ background: white; border-radius: 12px; padding: 1.5rem; box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-bottom: 2rem; }}
        .chart-container {{ background: white; border-radius: 12px; padding: 1.5rem; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}
        .chart-container canvas {{ max-height: 340px; }}
        .footer {{ margin-top: 3rem; padding: 1.5rem 0; color: #6c757d; text-align: center; border-top: 1px solid #dee2e6; }}
        .badge-custom {{ padding: 0.4rem 0.8rem; border-radius: 20px; }}
        .info-badge {{ font-size: 0.8rem; opacity: 0.8; }}
    </style>
</head>
<body>
    <div class="header">
        <div class="container">
            <div class="row align-items-center">
                <div class="col-md-8">
                    <h1>🏛️ Dashboard de Resoluciones Catastrales</h1>
                    <small>Convenio IGAC - Municipio de Mosquera</small>
                </div>
                <div class="col-md-4 text-end">
                    <small>📅 {fecha}</small><br>
                    <small>📁 {procesamiento}</small>
                </div>
            </div>
        </div>
    </div>

    <div class="container">
        <!-- ESTADÍSTICAS -->
        <div class="row g-4">{stats}</div>

        <!-- GRÁFICOS -->
        <div class="row g-4 mt-4">
            <div class="col-md-6">
                <div class="chart-container">
                    <h6 class="mb-1">📊 {titulo_barras}</h6>
                    <div class="text-muted small mb-2">{nota_barras}</div>
                    <canvas id="chartMovimientos"></canvas>
                </div>
            </div>
            <div class="col-md-6">
                <div class="chart-container">
                    <h6 class="mb-1">💰 {titulo_dona}</h6>
                    <div class="text-muted small mb-2">Agrupado por clase, no por resolución</div>
                    <canvas id="chartAvaluo"></canvas>
                </div>
            </div>
        </div>

        <!-- TABLA RESOLUCIONES -->
        <div class="section-title"><h2>📋 Resoluciones Procesadas</h2></div>
        <div class="table-container">
            <table id="tablaResoluciones" class="table table-hover table-striped">
                <thead>
                    <tr>
                        <th>Resolución</th>
                        <th>Radicación</th>
                        <th>Clase Mutación</th>
                        <th>Predios</th>
                        <th>✅ Inscritos</th>
                        <th>❌ Cancelados</th>
                        <th>💰 Avaluo</th>
                    </tr>
                </thead>
                <tbody>{tabla_res}</tbody>
            </table>
        </div>

        <!-- TABLA PREDIOS -->
        <div class="section-title"><h2>🏠 Predios Consolidados</h2></div>
        <div class="table-container">
            <div class="row mb-3 align-items-end">
                <div class="col-md-7">
                    <label class="form-label fw-bold mb-1">🔎 Consultar predio</label>
                    <input type="text" id="buscarPredio" class="form-control"
                           placeholder="Escriba el número predial, la resolución o parte del movimiento">
                    <div class="form-text">La búsqueda recorre los {total_predios_tabla} predios, no solo la página visible.</div>
                </div>
                <div class="col-md-5 text-md-end mt-3 mt-md-0">
                    <span class="badge bg-secondary" id="contadorPredios"></span>
                </div>
            </div>
            <table id="tablaPredios" class="table table-hover table-striped">
                <thead>
                    <tr>
                        <th>NPN</th>
                        <th>Resolución</th>
                        <th>Situación</th>
                        <th>Movimiento</th>
                        <th>Área (m²)</th>
                        <th>Avaluo ($)</th>
                        <th>Var. Área</th>
                        <th>Var. Avaluo</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
        </div>

        <!-- ALERTAS -->
        <div class="section-title"><h2>⚠️ Alertas e Inconsistencias</h2></div>
        <div class="row g-4">{alertas}</div>

        <div class="footer">
            <p>Generado automáticamente el {fecha}</p>
            <small class="text-muted">Dashboard generado desde: {procesamiento}</small>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
    <script src="https://cdn.datatables.net/1.13.4/js/dataTables.bootstrap5.min.js"></script>
    <script>
        {predios_js}

        // Cada bloque va aislado: si uno falla, los demas siguen funcionando.
        $(document).ready(function() {{
            try {{
                $('#tablaResoluciones').DataTable({{
                    pageLength: 25, order: [[0, 'desc']],
                    lengthMenu: [[25, 50, 100, -1], [25, 50, 100, "Todas"]],
                    language: {{ search: 'Buscar resolucion:',
                                lengthMenu: 'Ver _MENU_ resoluciones',
                                info: 'Mostrando _START_ a _END_ de _TOTAL_ resoluciones',
                                paginate: {{ first: 'Primera', last: 'Ultima',
                                            next: 'Siguiente', previous: 'Anterior' }} }}
                }});
            }} catch (e) {{ console.error('Tabla de resoluciones:', e); }}

            try {{
                var pesos = function(v) {{
                    if (v === null || v === undefined || v === 0) return '-';
                    var a = Math.abs(v), s, u = '';
                    if (a >= 1e12) {{ s = (v/1e12).toFixed(2); u = ' billones'; }}
                    else if (a >= 1e6) {{ s = Math.round(v/1e6).toString(); u = ' millones'; }}
                    else {{ s = Math.round(v).toString(); }}
                    return '$' + s.replace('.', ',').replace(/\\\B(?=(\\d{{3}})+(?!\\d))/g, '.') + u;
                }};
                var metros = function(v) {{
                    if (v === null || v === undefined) return '-';
                    return Math.round(v).toString().replace(/\\B(?=(\\d{{3}})+(?!\\d))/g, '.');
                }};
                var colorSit = function(s) {{
                    var l = (s || '').toLowerCase();
                    if (l.indexOf('inscripcion') >= 0 && l.indexOf('cancelacion') >= 0) return 'warning';
                    if (l.indexOf('inscripcion') >= 0) return 'success';
                    if (l.indexOf('cancelacion') >= 0) return 'danger';
                    return 'secondary';
                }};
                var tp = $('#tablaPredios').DataTable({{
                    data: PREDIOS,
                    deferRender: true,
                    pageLength: 25,
                    order: [[5, 'desc']],
                    lengthMenu: [[25, 50, 100, 500], [25, 50, 100, 500]],
                    dom: 'lfrtip',
                    columns: [
                        {{ data: 0, render: function(d) {{ return '<code>' + d + '</code>'; }} }},
                        {{ data: 1, render: function(d) {{ return D_RES[d]; }} }},
                        {{ data: 2, render: function(d) {{
                            var s = D_SIT[d];
                            return '<span class="badge bg-' + colorSit(s) + '">' + s + '</span>'; }} }},
                        {{ data: 3, render: function(d) {{ return D_MOV[d]; }} }},
                        {{ data: 4, render: function(d, t) {{ return t === 'display' ? metros(d) : d; }} }},
                        {{ data: 5, render: function(d, t) {{ return t === 'display' ? pesos(d) : d; }} }},
                        {{ data: 6, render: function(d, t) {{ return t === 'display' ? metros(d) : d; }} }},
                        {{ data: 7, render: function(d, t) {{ return t === 'display' ? pesos(d) : d; }} }}
                    ]
                }});
                $('#buscarPredio').on('keyup', function() {{ tp.search(this.value).draw(); }});
                var actualizar = function() {{
                    var info = tp.page.info();
                    $('#contadorPredios').text(
                        info.recordsDisplay.toLocaleString('es-CO') + ' de ' +
                        info.recordsTotal.toLocaleString('es-CO') + ' predios');
                }};
                tp.on('draw', actualizar); actualizar();
            }} catch (e) {{ console.error('Tabla de predios:', e); }}

            try {{
                if (typeof Chart === 'undefined') {{
                    throw new Error('Chart.js no cargo. Revisar la conexion a internet.');
                }}
                {graficos_js}
            }} catch (e) {{
                console.error('Graficos:', e);
                document.querySelectorAll('.chart-container').forEach(function(c) {{
                    c.innerHTML += '<div class="alert alert-warning mt-2 small">' +
                        'No fue posible dibujar el grafico. ' + e.message + '</div>';
                }});
            }}
        }});
    </script>
</body>
</html>'''

# =============================================================================
# PLANTILLA HTML DEL ÍNDICE
# Las llaves del CSS van DUPLICADAS porque esta plantilla pasa por .format().
# =============================================================================

INDEX_TEMPLATE = '''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Índice - Dashboards Resoluciones</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body {{ background: #f0f2f5; font-family: 'Segoe UI', sans-serif; }}
        .header {{ background: linear-gradient(135deg, #0a1628, #1a3a5c); color: white; padding: 2rem 0; margin-bottom: 2rem; }}
        .card-dashboard {{
            background: white; border-radius: 12px; padding: 1.5rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
            transition: transform 0.2s, box-shadow 0.2s;
            cursor: pointer;
            height: 100%;
        }}
        .card-dashboard:hover {{
            transform: translateY(-5px);
            box-shadow: 0 4px 16px rgba(0,0,0,0.15);
        }}
        .card-dashboard .fecha {{ font-size: 0.85rem; color: #6c757d; }}
        .card-dashboard .badge {{ font-size: 0.8rem; }}
        .footer {{ margin-top: 3rem; padding: 1.5rem 0; color: #6c757d; text-align: center; border-top: 1px solid #dee2e6; }}
        .stats-resumen {{
            background: white; border-radius: 12px; padding: 1rem 1.5rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08); margin-bottom: 2rem;
        }}
        .stats-resumen .numero {{ font-size: 1.5rem; font-weight: 700; }}
    </style>
</head>
<body>
    <div class="header">
        <div class="container">
            <h1>📊 Dashboards de Resoluciones Catastrales</h1>
            <small>Todos los procesamientos disponibles</small>
        </div>
    </div>
    <div class="container">
        <div class="stats-resumen">
            <div class="row text-center">
                <div class="col-md-4">
                    <div class="numero">{total}</div>
                    <small class="text-muted">Dashboards disponibles</small>
                </div>
                <div class="col-md-4">
                    <div class="numero">{ultimo}</div>
                    <small class="text-muted">Último procesamiento</small>
                </div>
                <div class="col-md-4">
                    <div class="numero">{fecha_actual}</div>
                    <small class="text-muted">Actualizado</small>
                </div>
            </div>
        </div>
        <div class="row g-4">
{tarjetas}
        </div>
        <div class="footer">
            <p>Generado automáticamente el {fecha_actual}</p>
            <small class="text-muted">Sistema de monitoreo de resoluciones catastrales</small>
        </div>
    </div>
</body>
</html>'''

# =============================================================================
# FUNCIONES DE UTILIDAD
# =============================================================================

def miles(v, dec=0):
    """Separador de miles con punto y decimal con coma, como en Colombia."""
    s = f'{v:,.{dec}f}'
    return s.replace(',', '@').replace('.', ',').replace('@', '.')


def fmt_num(v):
    """Formatea pesos con la escala larga que se usa en Colombia.

    Aqui un billon son un millon de millones (1e12), NO mil millones como en
    la escala corta anglosajona. Por eso no se usan las abreviaturas B ni K:
    1e9 se muestra como miles de millones expresados en millones.
    """
    if v is None or pd.isna(v):
        return '-'
    if isinstance(v, (int, float)):
        a = abs(v)
        if a >= 1e12:
            return f'${miles(v/1e12, 2)} billones'
        if a >= 1e6:
            return f'${miles(v/1e6)} millones'
        return f'${miles(v)}'
    return str(v)


def fmt_area(v):
    if v is None or pd.isna(v):
        return '-'
    return miles(v)


def texto(v, defecto='-'):
    """Devuelve el valor como texto, o el defecto si viene vacío o es NaN.
    Evita que aparezcan 'nan' en el HTML."""
    if v is None:
        return defecto
    try:
        if pd.isna(v):
            return defecto
    except (TypeError, ValueError):
        pass
    s = str(v).strip()
    return s if s and s.lower() != 'nan' else defecto


def cargar_excel(ruta_excel):
    """Carga todas las hojas del Excel en un diccionario de DataFrames."""
    try:
        excel = pd.ExcelFile(ruta_excel)
        hojas = {}
        for hoja in excel.sheet_names:
            df = pd.read_excel(excel, sheet_name=hoja)
            hojas[hoja] = df
        return hojas
    except Exception as e:
        print(f"  ⚠️ Error al cargar {ruta_excel}: {e}")
        return None


# =============================================================================
# FUNCIÓN PRINCIPAL: GENERAR DASHBOARD
# =============================================================================

def generar_dashboard(carpeta_procesamiento, carpeta_destino):
    """
    Genera un dashboard HTML a partir de una carpeta de procesamiento.

    Args:
        carpeta_procesamiento (str): Ruta a la carpeta del procesamiento
        carpeta_destino (str): Ruta donde guardar el dashboard

    Returns:
        str: Ruta al archivo HTML generado, o None si falla
    """

    ruta_excel = os.path.join(carpeta_procesamiento, "resoluciones.xlsx")
    if not os.path.exists(ruta_excel):
        print(f"  ⚠️ No se encontró resoluciones.xlsx")
        return None

    print(f"  📊 Procesando: {os.path.basename(carpeta_procesamiento)}")

    # Cargar datos
    hojas = cargar_excel(ruta_excel)
    if not hojas:
        return None

    df_res = hojas.get('01_resoluciones_resumen', pd.DataFrame())
    df_cons = hojas.get('05_consolidado', pd.DataFrame())
    df_alert = hojas.get('06_alertas', pd.DataFrame())

    # ===== MÉTRICAS CLAVE =====
    total_res = len(df_res)
    total_predios = df_cons['npn'].nunique() if not df_cons.empty else 0
    total_transacciones = len(df_cons)

    if not df_cons.empty and 'situacion' in df_cons.columns:
        sit_txt = df_cons['situacion'].fillna('').astype(str)
        inscritos = df_cons[sit_txt.str.contains('inscripcion', case=False, na=False)]
        cancelados = df_cons[sit_txt.str.contains('cancelacion', case=False, na=False)]
        total_inscritos = inscritos['npn'].nunique()
        total_cancelados = cancelados['npn'].nunique()
        solo_ins = int((sit_txt == 'Solo inscripcion').sum())
        solo_can = int((sit_txt == 'Solo cancelacion').sum())
        ambos = int((sit_txt == 'Cancelacion e inscripcion').sum())
    else:
        total_inscritos = total_cancelados = solo_ins = solo_can = ambos = 0

    total_alertas = len(df_alert)
    alertas_criticas = 0
    if not df_alert.empty and 'nivel' in df_alert.columns:
        alertas_criticas = len(df_alert[df_alert['nivel'].astype(str).str.startswith(('1.', '2.'))])

    # ===== TARJETAS DE ESTADÍSTICAS =====
    clase_alerta = 'danger' if alertas_criticas > 0 else ('warning' if total_alertas > 0 else 'success')
    stats = f'''
        <div class="col-md-3">
            <div class="stat-card primary">
                <div class="icon">📄</div>
                <div class="number">{total_res}</div>
                <div class="label">Resoluciones</div>
            </div>
        </div>
                <div class="col-md-3">
            <div class="stat-card info">
                <div class="icon">🔁</div>
                <div class="number">{miles(total_transacciones)}</div>
                <div class="label">Total transacciones procesadas</div>
                <small class="text-muted">actos sobre predios</small>
            </div>
        </div>
        <div class="col-md-3">
            <div class="stat-card primary">
                <div class="icon">🏠</div>
                <div class="number">{total_predios}</div>
                <div class="label">Predios Totales</div>
            </div>
        </div>
        <div class="col-md-3">
            <div class="stat-card success">
                <div class="icon">✅</div>
                <div class="number">{total_inscritos}</div>
                <div class="label">Predios Inscritos</div>
                <small class="text-muted">({solo_ins} nuevos)</small>
            </div>
        </div>
        <div class="col-md-3">
            <div class="stat-card danger">
                <div class="icon">❌</div>
                <div class="number">{total_cancelados}</div>
                <div class="label">Predios Cancelados</div>
                <small class="text-muted">({solo_can} retirados)</small>
            </div>
        </div>
        <div class="col-md-3">
            <div class="stat-card warning">
                <div class="icon">🔄</div>
                <div class="number">{ambos}</div>
                <div class="label">Cambios (C + I)</div>
                <small class="text-muted">cancelación e inscripción</small>
            </div>
        </div>
        <div class="col-md-3">
            <div class="stat-card {clase_alerta}">
                <div class="icon">⚠️</div>
                <div class="number">{total_alertas}</div>
                <div class="label">Alertas ({alertas_criticas} críticas)</div>
            </div>
        </div>
    '''

    # ===== TABLA DE RESOLUCIONES =====
    filas_res = []
    if df_res.empty:
        filas_res.append('<tr><td colspan="7" class="text-center">No hay datos</td></tr>')
    else:
        for _, row in df_res.iterrows():
            avaluo = row.get('avaluo_inscrito', 0)
            if pd.isna(avaluo) or not avaluo:
                avaluo = row.get('avaluo_cancelado', 0)
            filas_res.append(f'''
            <tr>
                <td><strong>{texto(row.get('resolucion'))}</strong></td>
                <td>{texto(row.get('radicacion_nro'))}</td>
                <td><span class="badge bg-info">{texto(row.get('clase_mutacion_desc'))[:40]}</span></td>
                <td>{texto(row.get('predios'), '0')}</td>
                <td><span class="badge bg-success">{texto(row.get('predios_inscritos'), '0')}</span></td>
                <td><span class="badge bg-danger">{texto(row.get('predios_cancelados'), '0')}</span></td>
                <td>{fmt_num(avaluo)}</td>
            </tr>
            ''')
    tabla_res = ''.join(filas_res)

    # ===== TABLA DE PREDIOS =====
    # Se embeben TODOS los predios como arreglo JS compacto. Los textos que se
    # repiten (resolucion, situacion, movimiento) se guardan una sola vez en
    # diccionarios y en cada fila va solo su indice: reduce el peso a la
    # tercera parte frente a escribir el HTML fila por fila.
    import json as _json

    if df_cons.empty:
        predios_js = "const PREDIOS=[];const D_RES=[];const D_SIT=[];const D_MOV=[];"
        total_predios_tabla = 0
    else:
        d_res, d_sit, d_mov = {}, {}, {}

        def idx(dic, val):
            v = texto(val, '')
            if v not in dic:
                dic[v] = len(dic)
            return dic[v]

        filas = []
        for r in df_cons.itertuples(index=False):
            g = lambda c, d=None: getattr(r, c, d)
            def num(c):
                v = getattr(r, c, None)
                try:
                    return 0 if v is None or pd.isna(v) else round(float(v), 1)
                except (TypeError, ValueError):
                    return 0
            filas.append([
                texto(g('npn'), ''),
                idx(d_res, g('resolucion')),
                idx(d_sit, g('situacion')),
                idx(d_mov, g('movimiento_observado')),
                num('area_m2_inscripcion'),
                num('avaluo_pesos_inscripcion'),
                num('variacion_area'),
                num('variacion_avaluo'),
            ])
        inv = lambda d: [k for k, _ in sorted(d.items(), key=lambda kv: kv[1])]
        predios_js = (
            "const PREDIOS=" + _json.dumps(filas, ensure_ascii=False, separators=(',', ':')) + ";"
            + "const D_RES=" + _json.dumps(inv(d_res), ensure_ascii=False) + ";"
            + "const D_SIT=" + _json.dumps(inv(d_sit), ensure_ascii=False) + ";"
            + "const D_MOV=" + _json.dumps(inv(d_mov), ensure_ascii=False) + ";"
        )
        total_predios_tabla = len(filas)

    # ===== ALERTAS =====
    # Se ordenan por nivel para que las que exigen accion queden arriba y no
    # se pierdan entre las informativas, que suelen ser mayoria.
    alertas_html = []
    if df_alert.empty:
        alertas_html.append('<div class="col-12"><div class="alert alert-success">✅ No se encontraron inconsistencias</div></div>')
    else:
        df_ord = df_alert.copy()
        if 'nivel' in df_ord.columns:
            df_ord = df_ord.sort_values('nivel', kind='stable')
        for _, row in df_ord.head(10).iterrows():
            nivel = texto(row.get('nivel'), 'Informativo')
            if nivel.startswith(('1.', '2.')):
                clase = 'danger'
            elif nivel.startswith('3.'):
                clase = 'warning'
            else:
                clase = 'info'

            alertas_html.append(f'''
            <div class="col-md-6">
                <div class="stat-card" style="border-left: 4px solid var(--bs-{clase});">
                    <span class="badge bg-{clase} mb-2">{nivel}</span>
                    <p class="mb-1"><strong>{texto(row.get('tipo'), '')}</strong></p>
                    <p class="mb-1 text-muted small">{texto(row.get('detalle'), '')}</p>
                    <p class="mb-0 text-muted small">NPN: {texto(row.get('npn'))}</p>
                </div>
            </div>
            ''')
        if len(df_alert) > 10:
            alertas_html.append(f'<div class="col-12"><p class="text-muted small">Mostrando 10 de {len(df_alert)} alertas. El detalle completo está en 06_alertas.csv</p></div>')
    alertas = ''.join(alertas_html)

    # ===== GRÁFICOS =====
    # Un grafico de barras con 85 resoluciones es ilegible. Por eso:
    #   - El primero muestra solo las 12 resoluciones de mayor volumen.
    #   - El segundo agrupa por clase de mutacion, que siempre son pocas
    #     categorias sin importar cuantas resoluciones haya.
    TOPE = 12
    if not df_res.empty and 'resolucion' in df_res.columns:
        d = df_res.copy()
        for c in ('predios', 'predios_inscritos', 'predios_cancelados',
                  'avaluo_inscrito', 'avaluo_cancelado'):
            if c not in d.columns:
                d[c] = 0
            d[c] = pd.to_numeric(d[c], errors='coerce').fillna(0)

        top = d.sort_values('predios', ascending=False).head(TOPE)
        res_list = [str(v) for v in top['resolucion']]
        insc_list = [float(v) for v in top['predios_inscritos']]
        can_list = [float(v) for v in top['predios_cancelados']]

        titulo_barras = (f'Top {len(top)} resoluciones por volumen'
                         if len(d) > TOPE else 'Movimientos por resolución')
        nota_barras = (f'De {len(d)} resoluciones se muestran las {len(top)} de mayor volumen. '
                       f'El detalle completo está en la tabla de abajo.'
                       if len(d) > TOPE else '')

        # avaluo agrupado por clase de mutacion
        col_clase = 'clase_mutacion_desc' if 'clase_mutacion_desc' in d.columns else None
        if col_clase:
            d['_av'] = d['avaluo_inscrito'].where(d['avaluo_inscrito'] > 0, d['avaluo_cancelado'])
            g = (d.groupby(col_clase)['_av'].sum().sort_values(ascending=False))
            clase_lbl = [str(k)[:45] for k in g.index]
            clase_val = [round(float(v) / 1e6, 1) for v in g.values]
            titulo_dona = 'Avalúo por clase de mutación (millones de pesos)'
        else:
            clase_lbl, clase_val = [], []
            titulo_dona = 'Avalúo'

        graficos_js = f'''
        new Chart(document.getElementById('chartMovimientos'), {{
            type: 'bar',
            data: {{
                labels: {res_list},
                datasets: [
                    {{ label: 'Inscritos', data: {insc_list}, backgroundColor: 'rgba(25,135,84,0.75)' }},
                    {{ label: 'Cancelados', data: {can_list}, backgroundColor: 'rgba(220,53,69,0.75)' }}
                ]
            }},
            options: {{ responsive: true, maintainAspectRatio: false, indexAxis: 'y',
                plugins: {{ legend: {{ position: 'top' }} }},
                scales: {{ x: {{ beginAtZero: true }} }} }}
        }});
        new Chart(document.getElementById('chartAvaluo'), {{
            type: 'doughnut',
            data: {{
                labels: {clase_lbl},
                datasets: [{{ data: {clase_val}, backgroundColor: ['#0d6efd','#198754','#fd7e14','#6f42c1','#d63384','#20c997','#ffc107','#dc3545'] }}]
            }},
            options: {{ responsive: true, maintainAspectRatio: false,
                plugins: {{ legend: {{ position: 'bottom' }},
                    tooltip: {{ callbacks: {{ label: function(c) {{
                        return c.label + ': $' + c.parsed.toLocaleString('es-CO') + ' millones'; }} }} }} }} }}
        }});
        '''
    else:
        graficos_js = '// Sin datos para gráficos'
        titulo_barras, titulo_dona, nota_barras = 'Movimientos', 'Avalúo', ''

    # ===== GENERAR HTML =====
    fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
    nombre_procesamiento = os.path.basename(carpeta_procesamiento.rstrip("\\/"))

    html = HTML_TEMPLATE.format(
        fecha=fecha,
        procesamiento=nombre_procesamiento,
        stats=stats,
        tabla_res=tabla_res,
        predios_js=predios_js,
        total_predios_tabla=miles(total_predios_tabla),
        alertas=alertas,
        graficos_js=graficos_js,
        titulo_barras=titulo_barras,
        titulo_dona=titulo_dona,
        nota_barras=nota_barras
    )

    # Guardar
    os.makedirs(carpeta_destino, exist_ok=True)
    nombre_html = f"dashboard_{nombre_procesamiento}.html"
    ruta_html = os.path.join(carpeta_destino, nombre_html)

    with open(ruta_html, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"  ✅ Dashboard generado: {nombre_html}")
    return ruta_html


# =============================================================================
# GENERAR ÍNDICE (BATCH)
# =============================================================================

def generar_indice(carpeta_dashboard):
    """Genera un índice HTML con todos los dashboards disponibles."""

    # Buscar todos los archivos HTML de dashboard
    patron = os.path.join(carpeta_dashboard, "dashboard_*.html")
    archivos = glob.glob(patron)

    if not archivos:
        print("⚠️ No se encontraron dashboards para indexar")
        return None

    # Extraer información de cada dashboard
    dashboards = []
    for archivo in archivos:
        nombre_base = os.path.basename(archivo)
        nombre = nombre_base.replace('dashboard_', '').replace('.html', '')

        # Intentar extraer fecha del nombre
        fecha_mostrar = nombre
        try:
            fecha_str = nombre.split('_', 1)[1] if '_' in nombre else nombre
            fecha_obj = datetime.strptime(fecha_str, "%Y%m%d_%H%M")
            fecha_mostrar = fecha_obj.strftime("%d/%m/%Y %H:%M")
        except (ValueError, IndexError):
            pass

        dashboards.append({
            'nombre': nombre,
            'archivo': nombre_base,
            'fecha': fecha_mostrar
        })

    # Ordenar por nombre (más reciente primero)
    dashboards.sort(key=lambda x: x['nombre'], reverse=True)

    tarjetas = []
    for d in dashboards:
        tarjetas.append(f'''
            <div class="col-md-4">
                <div class="card-dashboard" onclick="window.location.href='{d['archivo']}'">
                    <h5 class="card-title">📁 {d['nombre']}</h5>
                    <p class="fecha">🕐 {d['fecha']}</p>
                    <span class="badge bg-primary">Ver dashboard →</span>
                </div>
            </div>
        ''')

    html_index = INDEX_TEMPLATE.format(
        total=len(dashboards),
        ultimo=dashboards[0]['nombre'] if dashboards else '-',
        fecha_actual=datetime.now().strftime("%d/%m/%Y %H:%M"),
        tarjetas=''.join(tarjetas)
    )

    # Guardar índice
    ruta_index = os.path.join(carpeta_dashboard, "index.html")
    with open(ruta_index, 'w', encoding='utf-8') as f:
        f.write(html_index)

    print(f"\n✅ Índice generado: {ruta_index}")
    return ruta_index


# =============================================================================
# FUNCIÓN PRINCIPAL
# =============================================================================

def main():
    # Configurar argumentos de línea de comandos
    parser = argparse.ArgumentParser(
        description='Generador de Dashboard + Batch para Resoluciones Catastrales',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
EJEMPLOS:
  # Generar dashboards para TODOS los procesamientos (BATCH)
  python dashboard_batch.py

  # Generar dashboard para UN procesamiento específico
  python dashboard_batch.py "Salida/procesamiento_20260107_1430"

  # Usar rutas personalizadas
  python dashboard_batch.py --entrada "mi_ruta_salida" --salida "mi_ruta_dashboard"
        '''
    )

    parser.add_argument(
        'procesamiento_especifico',
        nargs='?',
        help='Ruta a un procesamiento específico (opcional)'
    )

    parser.add_argument(
        '--entrada',
        dest='ruta_salida',
        default=RUTA_SALIDA,
        help='Ruta a la carpeta de Salida del parser'
    )

    parser.add_argument(
        '--salida',
        dest='ruta_dashboard',
        default=RUTA_DASHBOARD,
        help='Ruta donde guardar los dashboards'
    )

    args = parser.parse_args()

    print("\n" + "=" * 62)
    print("📊 GENERADOR DE DASHBOARD + BATCH")
    print("=" * 62)
    print(f"📁 Dashboards en: {args.ruta_dashboard}")
    print(f"📁 Buscando en:   {args.ruta_salida}")
    print("=" * 62)

    # Caso 1: Procesamiento específico
    if args.procesamiento_especifico:
        print(f"\n🎯 Procesando específico: {args.procesamiento_especifico}")
        ruta_html = generar_dashboard(args.procesamiento_especifico, args.ruta_dashboard)
        if ruta_html:
            print(f"\n✅ Dashboard generado: {ruta_html}")
            generar_indice(args.ruta_dashboard)
        return

    # Caso 2: Batch - Procesar TODOS los procesamientos
    print("\n🔍 Buscando todos los procesamientos...")

    if not os.path.isdir(args.ruta_salida):
        print(f"❌ La carpeta de Salida no existe:\n   {args.ruta_salida}")
        print("   Revisar la constante RUTA_SALIDA al inicio del script.")
        sys.exit(1)

    # Buscar carpetas de procesamiento (coincide con el nombre que genera el parser)
    patron = os.path.join(args.ruta_salida, "procesamiento_*")
    carpetas = sorted([c for c in glob.glob(patron) if os.path.isdir(c)])

    if not carpetas:
        print(f"❌ No se encontraron procesamientos en: {args.ruta_salida}")
        print("   Buscando carpetas con formato: procesamiento_*")
        sys.exit(1)

    print(f"🔍 Encontrados {len(carpetas)} procesamientos\n")

    # Procesar cada carpeta
    generados = 0
    for carpeta in carpetas:
        print(f"📁 {os.path.basename(carpeta)}")
        ruta_html = generar_dashboard(carpeta, args.ruta_dashboard)
        if ruta_html:
            generados += 1
        print()

    # Generar índice
    if generados > 0:
        print("-" * 40)
        print("📄 Generando índice...")
        ruta_index = generar_indice(args.ruta_dashboard)

        print("\n" + "=" * 62)
        print(f"✅ PROCESO COMPLETADO")
        print(f"   Dashboards generados: {generados} de {len(carpetas)}")
        print(f"   Índice: {ruta_index}")
        print("=" * 62)
    else:
        print("\n❌ No se generó ningún dashboard")
        sys.exit(1)


if __name__ == "__main__":
    main()
