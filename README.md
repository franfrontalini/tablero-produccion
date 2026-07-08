# Tablero de Gestión de Producción

Dashboard de planificación y control de producción para planta metalmecánica
(herramientas de fresado/pesca para pozos: Junk Mill, Taper Mill, etc.).

Es un **único archivo HTML autocontenido** ([`index.html`](index.html)) con los datos
embebidos como JSON. No tiene backend ni build step en tiempo de ejecución: se puede
abrir con doble-click, servir por HTTP o publicar en GitHub Pages.

## Vistas

1. **Planificación** — OF pendientes, carga vs. capacidad por centro y proceso en un
   horizonte de N semanas, con filtros y gauges de carga (umbrales 85% / 100%).
2. **En Proceso** — OF en proceso / vencidas / en riesgo (≤7 días), semáforo de
   utilización de la semana actual y tendencia de las últimas 12 semanas vs. objetivo 85%.
3. **Mes Cerrado** — cumplimiento de entrega vs. 95%, utilización vs. 85%, heatmap por
   centro × mes y producción terminada por familia.

## Cómo actualizar los datos

Los datos viven en un Excel ([`data/`](data/)) y se inyectan en `index.html` con
[`build.py`](build.py). Para actualizar el tablero cuando cambia el Excel:

```bash
# 1. Preparar el entorno (una sola vez)
python3 -m venv .buildenv
.buildenv/bin/pip install -r requirements.txt

# 2. Reemplazar el Excel en data/ (mismo nombre y estructura de hojas) y regenerar
.buildenv/bin/python build.py

# 3. Commitear el index.html actualizado
git add index.html data/*.xlsx && git commit -m "Actualizar datos" && git push
```

`build.py` reemplaza *in-place* solo el bloque `<script id="data-store">` del HTML;
no toca el resto del archivo.

### Opciones de `build.py`

| Flag | Descripción |
|------|-------------|
| `--xlsx RUTA` | Usar otra planilla (default: `data/Prueba - Tablero de Gestión.xlsx`). |
| `--today YYYY-MM-DD` | Fijar la fecha de referencia (default: hoy del sistema). |
| `--check` | No escribe: compara con el JSON embebido actual y reporta diferencias. |

La **fecha de referencia** (hoy, semana y mes actual, meses cerrados) se **deriva del
calendario** de la planilla, así que se actualiza sola al cambiar los datos o el `--today`.

## Cómo se arman los datos

`build.py` lee las hojas del Excel y produce el JSON que consume el dashboard:

| Bloque JSON | Hoja de origen |
|-------------|----------------|
| `config` | `01_Conf` |
| `meta` (fecha de referencia) | derivado de `02_Calendario` |
| `centros` | `03_Recursos` |
| `productos` | `04_Productos` |
| `ordenesFabricacion` | `06_OF` (+ campos calculados: `aTiempo`, `diasAtraso`, `diasParaVencer`) |
| `cargaOF` | `07_Modelo_Carga_OF` |
| `capacidadSemanal` | `08_Carga_Capacidad` (dedup por año/semana/centro) |
| `indicadores*` | **recalculados** desde `capacidadSemanal` |

> Los indicadores semanales/mensuales/por centro se **recalculan** desde la capacidad
> (no se leen de las hojas `09A/09B/09C`, que traen inconsistencias de mayúsculas y filas
> basura). Los indicadores son autoconsistentes con la capacidad que muestra el tablero.

## Publicación (GitHub Pages)

El sitio se sirve directamente desde `index.html` en la raíz de la rama `main`.
Cualquier `push` con un `index.html` actualizado se publica automáticamente.

## Stack

HTML/CSS/JS sin framework · [Chart.js](https://www.chartjs.org/) vía CDN ·
Google Fonts (Oswald / Inter / IBM Plex Mono).
