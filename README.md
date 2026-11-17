# Análisis de redes sociales procrastinación y nota media

Código para reproducir un análisis correlacional de la «Encuesta sobre vida universitaria 2026». La base contiene datos simulados con fines docentes. Los resultados no describen estudiantes reales.

La depuración conserva una copia de cada fila exacta, normaliza las etiquetas de titulación y elimina por completo las filas con valores inválidos o incoherencias. Las casillas vacías siguen siendo ausentes. La secuencia comprobada es 1.200 → 1.194 → 1.189 → 1.073; los casos principales se reparten entre 491 registros de ADET y 582 de INF.

## Ejecutar el análisis

Se necesita Python y el Excel original, que debe obtenerse en Moodle con el acceso correspondiente. El repositorio no redistribuye ese archivo.

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python script_etaic_definitivo.py --excel /ruta/ETAIC_Dataset_VidaUniversitaria2026.xlsx --salida resultados
```

## Archivos y procedencia

- `original/script_etaic_definitivo_recibido.py` conserva el archivo recibido sin cambios.
- `script_etaic_definitivo.py` adapta la lectura del Excel a openpyxl, añade la comprobación de Pearson para procrastinación y nota, y exporta descriptivos y disponibilidad. Conserva el criterio de limpieza y las correlaciones principales del archivo recibido.
- `resultados/` contiene las salidas de la ejecución comprobada, incluidas las tres figuras y los tamaños de cada análisis.
- `CITATION.cff` permite identificar esta versión del código.

Los análisis exploratorios utilizan pares válidos, por lo que su tamaño cambia. Sus valores p no tienen ajuste por comparaciones múltiples. La asociación observada no establece causalidad y no permite generalizar a una población real.

ChatGPT intervino en la adaptación, documentación y comprobación del código. La redacción del artículo y su adecuación a las normas de uso de IA deben evaluarse por separado.
