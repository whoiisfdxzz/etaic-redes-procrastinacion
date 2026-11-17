# -*- coding: utf-8 -*-
"""Análisis definitivo ETAIC: redes sociales, procrastinación y nota media.

Reproduce las cifras del artículo IEEE y sus comprobaciones posteriores.

REGLAS DE LIMPIEZA (importante):
- Una celda vacía se conserva como valor ausente. Nunca se rellena con cero.
- Un duplicado exacto se elimina conservando su primera aparición.
- Una respuesta imposible o incoherente provoca la eliminación de TODA la fila.
  Esto incluye las cuatro notas imposibles y las horas de redes > pantalla total.
- Los casos completos se seleccionan DESPUÉS, solo para las variables requeridas.

Dependencias: pip install artifact-tool numpy scipy matplotlib
Uso, con el Excel en la misma carpeta:
    python script_etaic_definitivo.py
O indicando su ubicación:
    python script_etaic_definitivo.py --excel "ruta/al/ETAIC_Dataset_VidaUniversitaria2026.xlsx"
La salida se guarda en ETAIC_resultados_definitivos/.
El Excel original nunca se modifica.
"""

from pathlib import Path
import argparse
import csv
import json

import numpy as np
import matplotlib.pyplot as plt
from scipy.stats import spearmanr
from artifact_tool import Blob, SpreadsheetFile

CARPETA = Path(__file__).resolve().parent
argumentos = argparse.ArgumentParser(description='Reproduce el análisis del artículo IEEE de ETAIC.')
argumentos.add_argument('--excel', type=Path, help='Archivo XLSX original (opcional si está junto al script).')
argumentos.add_argument('--salida', type=Path, help='Directorio de salida (por defecto: ETAIC_resultados_definitivos).')
args = argumentos.parse_args()
if args.excel:
    EXCEL = args.excel.expanduser().resolve()
else:
    candidatos = sorted(CARPETA.glob('ETAIC_Dataset_VidaUniversitaria2026*.xlsx'))
    if not candidatos:
        raise FileNotFoundError('Pon el Excel ETAIC_Dataset_VidaUniversitaria2026*.xlsx junto al script, o indica --excel.')
    EXCEL = candidatos[0]
if not EXCEL.is_file():
    raise FileNotFoundError(f'No se encuentra el Excel: {EXCEL}')
SALIDA = (args.salida.expanduser().resolve() if args.salida else CARPETA / 'ETAIC_resultados_definitivos')
SALIDA.mkdir(parents=True, exist_ok=True)
print(f'Excel analizado: {EXCEL}')
print(f'Directorio de resultados: {SALIDA}')

NUMERICAS = {
    'dig_redes_h': (0, 24), 'dig_pantalla_h': (0, 24),
    'nota_media': (0, 10), 'est_procrastinacion': (1, 5),
    'dig_multitarea': (1, 5), 'curso': (1, 4),
}
CATEGORICAS = {
    'dig_notificaciones': {'Ninguna', 'Algunas', 'Todas'},
    'dig_movil_clase': {'Nunca', 'A veces', 'Casi siempre'},
}
PRINCIPALES = ('dig_redes_h', 'est_procrastinacion', 'nota_media')


def esta_vacio(valor):
    return valor is None or (isinstance(valor, str) and not valor.strip())


def num(valor):
    if esta_vacio(valor) or isinstance(valor, bool):
        return None
    try:
        resultado = float(valor)
        return resultado if np.isfinite(resultado) else None
    except (ValueError, TypeError):
        return None


def escribir_csv(nombre, encabezados, registros):
    with (SALIDA / nombre).open('w', newline='', encoding='utf-8-sig') as archivo:
        escritor = csv.writer(archivo)
        escritor.writerow(encabezados)
        escritor.writerows(registros)


# 1. Importar exactamente la hoja de 1.200 respuestas y 92 columnas.
libro = SpreadsheetFile.import_xlsx(Blob.load(str(EXCEL)))
matriz = libro.worksheets.get_item('Datos').get_range('A1:CN1201').values
columnas = [str(valor).strip() for valor in matriz[0]]
requeridas = set(NUMERICAS) | set(CATEGORICAS) | {'id', 'titulacion'}
if not requeridas.issubset(columnas):
    raise ValueError(f'Columnas que faltan: {requeridas - set(columnas)}')
originales = [
    {'fila_excel': pos, **dict(zip(columnas, celdas))}
    for pos, celdas in enumerate(matriz[1:], start=2)
    if any(not esta_vacio(celda) for celda in celdas)
]

# 2. Excluir seis duplicados exactos, conservando la primera copia.
vistos, sin_duplicados, auditoria = set(), [], []
for fila in originales:
    clave = tuple(fila[col] for col in columnas)
    if clave in vistos:
        auditoria.append((fila['fila_excel'], fila['id'], 'Duplicado exacto'))
    else:
        vistos.add(clave)
        sin_duplicados.append(fila)

# 3. Corregir 15 etiquetas de titulación; descartar FILAS ENTERAS si
#    hay valores imposibles o uso de redes > tiempo total de pantalla.
depurados, invalidos, normalizados = [], [], 0
for fila in sin_duplicados:
    tit_original = fila['titulacion']
    tit = str(tit_original).strip().upper()
    if tit != tit_original:
        normalizados += 1
    fila['titulacion'] = tit

    motivos = []
    if tit not in {'INF', 'ADET'}:
        motivos.append('Titulación no válida')
    for nombre, (minimo, maximo) in NUMERICAS.items():
        valor = fila[nombre]
        n = num(valor)
        if not esta_vacio(valor) and (n is None or not minimo <= n <= maximo):
            motivos.append(f'{nombre}={valor} fuera de rango')
    for nombre, opciones in CATEGORICAS.items():
        if not esta_vacio(fila[nombre]) and fila[nombre] not in opciones:
            motivos.append(f'{nombre}={fila[nombre]} categoría no válida')
    redes, pantalla = num(fila['dig_redes_h']), num(fila['dig_pantalla_h'])
    if redes is not None and pantalla is not None and redes > pantalla:
        motivos.append('Horas en redes superiores a horas totales de pantalla')

    if motivos:
        invalidos.append(fila)
        auditoria.append((fila['fila_excel'], fila['id'], '; '.join(motivos)))
    else:
        depurados.append(fila)

# 4. Casos completos en las tres variables centrales; vacíos no eliminan
#    la fila de la base general, solo de la comparación que los necesita.
base = [fila for fila in depurados if all(num(fila[c]) is not None for c in PRINCIPALES)]

assert len(originales) == 1200, f'No hay 1.200 registros: {len(originales)}'
assert len(sin_duplicados) == 1194, 'Revisar eliminación de duplicados'
assert len(invalidos) == 5, f'Revisar filas inválidas: {len(invalidos)}'
assert len(depurados) == 1189, f'Revisar limpieza: {len(depurados)}'
assert len(base) == 1073, f'Revisar muestra analítica: {len(base)}'
assert sum(1 for f in base if f['titulacion'] == 'ADET') == 491
assert sum(1 for f in base if f['titulacion'] == 'INF') == 582
assert sorted([num(f['nota_media']) for f in invalidos if num(f['nota_media']) is not None and num(f['nota_media']) > 10]) == [47.0, 71.0, 84.0, 88.0]

# 5. TABLA I: cruce por titulación. Las frecuencias son n (%) dentro
#    de cada titulación en la misma muestra de n=1.073.
subgrupos = {t: [fila for fila in base if fila['titulacion'] == t] for t in ('ADET', 'INF')}
condiciones = [
    ('Redes: menos de 2 h', lambda f: num(f['dig_redes_h']) < 2),
    ('Redes: de 2 a <4 h', lambda f: 2 <= num(f['dig_redes_h']) < 4),
    ('Redes: 4 h o más', lambda f: num(f['dig_redes_h']) >= 4),
    ('Procrastinación: nivel 1–2', lambda f: num(f['est_procrastinacion']) in (1, 2)),
    ('Procrastinación: nivel 3', lambda f: num(f['est_procrastinacion']) == 3),
    ('Procrastinación: nivel 4–5', lambda f: num(f['est_procrastinacion']) in (4, 5)),
    ('Calificación: inferior a 5', lambda f: num(f['nota_media']) < 5),
    ('Calificación: 5 o superior', lambda f: num(f['nota_media']) >= 5),
]
filas_tabla = []
for nombre, condicion in condiciones:
    fila = [nombre]
    for titulo in ('ADET', 'INF'):
        total = len(subgrupos[titulo])
        cuenta = sum(condicion(f) for f in subgrupos[titulo])
        fila.append(f'{cuenta} ({100 * cuenta / total:.1f} %)')
    filas_tabla.append(fila)
encabezado = ['Variable / categoría', f'ADET (n={len(subgrupos["ADET"])})', f'INF (n={len(subgrupos["INF"])})']
escribir_csv('tabla_I_por_titulacion.csv', encabezado, filas_tabla)

# 6. Dos primeras figuras con barras claras y el tamaño de muestra sobre cada barra.
def barras(nombre_archivo, etiquetas, conjuntos, eje_x, eje_y, minimo, maximo):
    medias = [float(np.mean(valores)) if valores else np.nan for valores in conjuntos]
    fig, ax = plt.subplots(figsize=(8.2, 5.0))
    posiciones = np.arange(len(etiquetas))
    rects = ax.bar(posiciones, medias, width=0.62)
    ax.set_xticks(posiciones, etiquetas)
    ax.set_ylim(minimo, maximo)
    ax.set_xlabel(eje_x)
    ax.set_ylabel(eje_y)
    ax.grid(axis='y', alpha=0.17)
    ax.set_axisbelow(True)
    for barra, media, valores in zip(rects, medias, conjuntos):
        ax.text(barra.get_x() + barra.get_width()/2, media + 0.08,
                f'{media:.2f} (n={len(valores)})', ha='center', va='bottom', fontsize=10)
    fig.tight_layout()
    fig.savefig(SALIDA / nombre_archivo, dpi=220, bbox_inches='tight')
    plt.close(fig)
    return medias, [len(g) for g in conjuntos]

# Fig. 1. Promedio de procrastinación por tramos de horas en redes.
redes_grupos = [
    [num(f['est_procrastinacion']) for f in base if num(f['dig_redes_h']) < 2],
    [num(f['est_procrastinacion']) for f in base if 2 <= num(f['dig_redes_h']) < 4],
    [num(f['est_procrastinacion']) for f in base if num(f['dig_redes_h']) >= 4],
]
medias_redes, n_redes = barras(
    'figura_1_redes_procrastinacion.png',
    ['Menos de 2 h', 'De 2 a <4 h', '4 h o más'], redes_grupos,
    'Horas diarias en redes sociales', 'Procrastinación media (escala 1–5)', 0, 5
)

# Fig. 2. Calificación media por nivel de procrastinación, también barras.
notas_grupos = [
    [num(f['nota_media']) for f in base if num(f['est_procrastinacion']) == nivel]
    for nivel in range(1, 6)
]
medias_notas, n_notas = barras(
    'figura_2_procrastinacion_notas.png',
    ['1', '2', '3', '4', '5'], notas_grupos,
    'Nivel de procrastinación (1–5)', 'Nota media académica (escala 0–10)', 0, 10
)

# 7. Correlaciones de Spearman sobre las MISMAS 1.073 observaciones.
pares = [
    ('Horas en redes y procrastinación', 'dig_redes_h', 'est_procrastinacion'),
    ('Procrastinación y calificaciones', 'est_procrastinacion', 'nota_media'),
    ('Horas en redes y calificaciones', 'dig_redes_h', 'nota_media'),
]
correlaciones = []
for etiqueta, primera, segunda in pares:
    resultado = spearmanr(
        [num(f[primera]) for f in base], [num(f[segunda]) for f in base]
    )
    correlaciones.append([etiqueta, len(base), round(float(resultado.statistic), 5), float(resultado.pvalue)])
escribir_csv('correlaciones_spearman.csv', ['Relación', 'n', 'rho', 'p'], correlaciones)
for fila, esperado in zip(correlaciones, [0.00537, -0.46142, -0.05365]):
    assert np.isclose(fila[2], esperado, atol=1e-5), f'Coeficiente distinto del artículo: {fila}'

# Fig. 3. Gráfico de puntos: mantiene visibles incluso las correlaciones
# casi nulas (las barras anteriores ocultaban rho = 0,005).
# Diseño compacto para encajar en una columna de la plantilla IEEE.
labels_corr = [
    'Redes y\nprocrastinación',
    'Procrastinación\ny nota media',
    'Redes y\nnota media',
]
valores_corr = [float(fila[2]) for fila in correlaciones]
fig, ax = plt.subplots(figsize=(4.3, 2.65))
fig.subplots_adjust(left=0.36, right=0.96, bottom=0.23, top=0.94)
posiciones = np.arange(len(labels_corr))
ax.axvline(0, color='0.30', linewidth=1.0, zorder=1)
for y, valor in enumerate(valores_corr):
    ax.plot([0, valor], [y, y], color='0.50', linewidth=2.3,
            solid_capstyle='round', zorder=2)
    ax.scatter([valor], [y], s=72, marker='o',
               facecolor='#1d648a', edgecolor='white', linewidth=1.15,
               zorder=5)
    # La cifra se coloca en una columna fija a la derecha para facilitar
    # la lectura con independencia de la longitud del segmento.
    ax.text(0.98, y, f'{valor:+.3f}'.replace('.', ','),
            transform=ax.get_yaxis_transform(), ha='right', va='center',
            fontsize=9.4, fontweight='bold',
            bbox={'facecolor':'white', 'edgecolor':'none', 'pad':1.2},
            zorder=7)
ax.set_yticks(posiciones, labels_corr)
ax.invert_yaxis()
ax.set_xlim(-1, 1)
ax.set_xticks([-1, -0.5, 0, 0.5, 1])
ax.set_xticklabels(['−1', '−0,5', '0', '+0,5', '+1'])
ax.set_xlabel('Correlación de Spearman (ρ)', fontsize=9)
ax.tick_params(axis='y', labelsize=8.4, length=0, pad=6)
ax.tick_params(axis='x', labelsize=8, length=3)
ax.grid(axis='x', color='0.86', linewidth=0.6)
ax.set_axisbelow(True)
for spine in ('top', 'right', 'left'):
    ax.spines[spine].set_visible(False)
fig.savefig(SALIDA / 'figura_3_correlaciones_spearman.png', dpi=240,
            bbox_inches='tight', facecolor='white')
plt.close(fig)
escribir_csv('auditoria_filas_eliminadas.csv', ['Fila Excel', 'ID', 'Motivo'], auditoria)

# 8. Informe de salida breve para redactar resultados SIN interpretaciones.
resumen = {
    'originales': len(originales),
    'duplicados': len(originales) - len(sin_duplicados),
    'filas_invalidas': len(invalidos),
    'depurados': len(depurados),
    'base_tres_variables': len(base),
    'titulaciones_normalizadas': normalizados,
    'base_ADET': len(subgrupos['ADET']),
    'base_INF': len(subgrupos['INF']),
    'medias_figura_1': medias_redes,
    'n_figura_1': n_redes,
    'medias_figura_2': medias_notas,
    'n_figura_2': n_notas,
    'correlaciones': correlaciones,
}
(SALIDA / 'resumen_resultados.json').write_text(
    json.dumps(resumen, ensure_ascii=False, indent=2), encoding='utf-8'
)

# 9. Preparar también el apartado IV en texto (sin interpretación causal).
#    Las cifras proceden siempre de la misma ejecución, no se transcriben a mano.
def es(numero, decimales=2):
    return f'{numero:.{decimales}f}'.replace('.', ',')

lineas_tabla = ['| Variable / categoría | ADET (n=' + str(len(subgrupos['ADET'])) + ') | INF (n=' + str(len(subgrupos['INF'])) + ') |',
                '|---|---:|---:|']
for fila in filas_tabla:
    lineas_tabla.append('| ' + ' | '.join(c.replace('.', ',') for c in fila) + ' |')

informe = f"""# IV. RESULTADOS

La Tabla I presenta la distribución del uso diario de redes sociales, el nivel de procrastinación y las calificaciones, agrupados por titulación. El análisis incluye {len(base)} observaciones completas, de las cuales {len(subgrupos['ADET'])} corresponden a ADET y {len(subgrupos['INF'])} a Ingeniería Informática (INF).

**TABLA I. DISTRIBUCIÓN DE LAS VARIABLES PRINCIPALES POR TITULACIÓN**

{chr(10).join(lineas_tabla)}

*Nota.* Se muestran frecuencias absolutas n y porcentajes calculados dentro de cada titulación. Las categorías de procrastinación 1–2, 3 y 4–5 son agrupaciones descriptivas. Datos simulados de la asignatura.

Se calcularon correlaciones de Spearman en las {len(base)} observaciones completas: horas en redes y procrastinación (ρ = {es(correlaciones[0][2], 3)}), procrastinación y nota media (ρ = {es(correlaciones[1][2], 3)}) y horas en redes y nota media (ρ = {es(correlaciones[2][2], 3)}).

La Fig. 1 presenta el promedio de procrastinación de los estudiantes distribuidos en tres intervalos de uso diario de redes sociales.

![Gráfica de procrastinación media según horas de redes sociales](figura_1_redes_procrastinacion.png)

*Fig. 1. Procrastinación media según las horas diarias de uso de redes sociales. Datos simulados. Las etiquetas indican media y número de registros.*

La Fig. 2 presenta la nota media académica correspondiente a cada uno de los cinco niveles de procrastinación.

![Gráfica de nota media según procrastinación](figura_2_procrastinacion_notas.png)

*Fig. 2. Calificación media según el nivel de procrastinación académica. Datos simulados. Las etiquetas indican media y número de registros.*

La Fig. 3 recoge los tres coeficientes de Spearman calculados con las mismas {len(base)} respuestas completas. El coeficiente entre horas en redes y procrastinación fue de {es(correlaciones[0][2], 3)}, entre procrastinación y nota media fue de {es(correlaciones[1][2], 3)}, y entre horas en redes y nota media fue de {es(correlaciones[2][2], 3)}.

![Correlaciones entre las tres variables principales](figura_3_correlaciones_spearman.png)

*Fig. 3. Coeficientes de correlación de Spearman entre las variables principales. Datos simulados (n = {len(base)}).*
"""
(SALIDA / 'apartado_IV_resultados.md').write_text(informe, encoding='utf-8')

print('\n=== AUDITORÍA ===')
print(f'Originales: {len(originales)} | Duplicados: {len(originales) - len(sin_duplicados)}')
print(f'Filas inválidas: {len(invalidos)} | Depuradas: {len(depurados)} | Completas: {len(base)}')
print('\n=== TABLA I ===')
print(' | '.join(encabezado))
for fila in filas_tabla:
    print(' | '.join(fila))
print('\n=== SPEARMAN ===')
for fila in correlaciones:
    print(fila)
print('\n=== GRÁFICOS ===')
for nombre in ('figura_1_redes_procrastinacion.png', 'figura_2_procrastinacion_notas.png', 'figura_3_correlaciones_spearman.png'):
    print(SALIDA / nombre)

# 10. COMPROBACIONES ADICIONALES, REALIZADAS DESPUÉS DEL ANÁLISIS PRINCIPAL.
# Estas correlaciones no se utilizaron para elegir las hipótesis principales.
# Diferentes n por eliminación por pares válidos. No hay ajuste por multiplicidad.
from scipy.stats import pearsonr

complementarias = []

def registrar(tipo, izquierda, derecha, registros, metodo='spearman'):
    pares_validos = [(num(f[izquierda]), num(f[derecha])) for f in registros]
    pares_validos = [(x, y) for x, y in pares_validos if x is not None and y is not None]
    x, y = zip(*pares_validos)
    resultado = pearsonr(x, y) if metodo == 'pearson' else spearmanr(x, y)
    complementarias.append((tipo, izquierda, derecha, metodo,
                            len(pares_validos), float(resultado.statistic), float(resultado.pvalue)))

# Sensibilidad: solo las dos relaciones que incluyen HORAS EN REDES.
for y in ('est_procrastinacion', 'nota_media'):
    registrar('sensibilidad, muestra común', 'dig_redes_h', y, base, 'pearson')
    registrar('sensibilidad, pares disponibles', 'dig_redes_h', y, depurados)

# Exploración motivada por el resultado inicial: no análisis confirmatorio.
for x, y in (
    ('dig_multitarea', 'est_procrastinacion'),
    ('dig_multitarea', 'nota_media'),
    ('dig_pantalla_h', 'est_procrastinacion'),
    ('dig_pantalla_h', 'nota_media'),
):
    registrar('exploratorio, pares disponibles', x, y, depurados)

escribir_csv('sensibilidad_y_exploracion.csv',
             ['Tipo', 'Variable A', 'Variable B', 'Coeficiente', 'n', 'r/rho', 'p'],
             [(t, x, y, met, n, f'{r:.6f}', f'{p:.10g}')
              for t, x, y, met, n, r, p in complementarias])
print('\n=== COMPROBACIONES POSTERIORES NO CONFIRMATORIAS ===')
for fila in complementarias:
    print(fila)


# 11. Tabla de control que distingue la limpieza general de las muestras.
recuento = [
    ('Excel original', len(originales)),
    ('Tras eliminar duplicados exactos', len(sin_duplicados)),
    ('Filas completas excluidas por valores imposibles', len(invalidos)),
    ('Registros depurados, incluidos los que tienen vacíos', len(depurados)),
    ('Casos completos para tres correlaciones principales', len(base)),
    ('ADET dentro de casos completos', len(subgrupos['ADET'])),
    ('INF dentro de casos completos', len(subgrupos['INF'])),
]
escribir_csv('control_muestras.csv', ['Etapa', 'Registros'], recuento)
print('\nCOMPROBACIÓN FINAL CORRECTA: 1.200 → 1.194 → 1.189 → 1.073')
