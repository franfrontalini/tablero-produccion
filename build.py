#!/usr/bin/env python3
"""
build.py — Regenera el JSON embebido del dashboard a partir del Excel de origen.

Lee "Prueba - Tablero de Gestión.xlsx", arma el objeto de datos que consume
el dashboard y lo inyecta in-place dentro de <script id="data-store"> en index.html.

Uso:
    python build.py                          # usa data/*.xlsx e index.html por defecto
    python build.py --xlsx ruta/al.xlsx      # otra planilla (misma estructura)
    python build.py --today 2026-07-08       # fija la fecha de referencia (default: hoy)
    python build.py --check                  # no escribe; sólo compara y reporta diffs

La "fecha de referencia" (today / semana / mes actual / meses cerrados) se DERIVA
del calendario de la planilla, así que actualizando el Excel se actualiza sola.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import re
import sys
import unicodedata
from collections import OrderedDict
from pathlib import Path

try:
    import openpyxl
except ImportError:
    sys.exit("Falta openpyxl. Instalá con:  pip install -r requirements.txt")

ROOT = Path(__file__).resolve().parent
DEFAULT_XLSX = ROOT / "data" / "Prueba - Tablero de Gestión.xlsx"
DEFAULT_HTML = ROOT / "index.html"

MONTH_ORDER = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]
MONTH_IDX = {m: i for i, m in enumerate(MONTH_ORDER)}


# ----------------------------------------------------------------------------- helpers
def cap_mes(v):
    """'enero' / 'ENERO' -> 'Enero'. Deja intactos otros textos."""
    if v is None:
        return None
    s = str(v).strip()
    return s[:1].upper() + s[1:].lower() if s else s


def num(v):
    """float entero (9.0) -> int 9; deja floats no enteros y no-números como están."""
    if isinstance(v, bool):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def as_bool(v):
    """'SI' / 'Sí' / 'Si' -> True ; 'No' -> False."""
    if isinstance(v, bool):
        return v
    if v is None:
        return None
    s = unicodedata.normalize("NFKD", str(v)).encode("ascii", "ignore").decode().strip().lower()
    if s in ("si", "s", "true", "1"):
        return True
    if s in ("no", "n", "false", "0"):
        return False
    return None


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return None


def date_str(v):
    d = to_date(v)
    return d.isoformat() if d else None


def r2(v):
    return round(float(v), 2) if v is not None else 0.0


def r4(v):
    return round(float(v), 4) if v is not None else 0.0


def load_sheet(wb, name):
    """Devuelve (headers, list[dict]) de una hoja, ignorando filas totalmente vacías."""
    ws = wb[name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    headers = [h for h in rows[0]]
    out = []
    for r in rows[1:]:
        if all(c is None for c in r):
            continue
        out.append({h: (r[i] if i < len(r) else None) for i, h in enumerate(headers) if h is not None})
    return headers, out


def estado_util(util):
    """Umbrales de estado para indicadores recalculados por centro."""
    if not util:
        return "Sin carga"
    if util >= 1:
        return "Saturado"
    if util >= 0.85:
        return "Al límite"
    return "Disponible"


# ----------------------------------------------------------------------------- build
def build_data(xlsx_path: Path, today: dt.date) -> "OrderedDict":
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)

    # --- 02_Calendario: mapeo fecha -> (semana, mes) y derivación de la referencia
    _, cal = load_sheet(wb, "02_Calendario")
    fecha_to_mes = {}
    for row in cal:
        d = to_date(row.get("Fecha"))
        if d:
            fecha_to_mes[d] = cap_mes(row.get("Mes"))

    cal_row = next((r for r in cal if to_date(r.get("Fecha")) == today), None)
    if cal_row is not None:
        current_week = num(cal_row.get("Semana"))
        current_month = cap_mes(cal_row.get("Mes"))
    else:  # fecha fuera del calendario: usar ISO week y nombre de mes por índice
        current_week = today.isocalendar()[1]
        current_month = MONTH_ORDER[today.month - 1]
    closed_months = MONTH_ORDER[: MONTH_IDX[current_month]]

    meta = OrderedDict(
        today=today.isoformat(),
        currentWeek=current_week,
        currentMonth=current_month,
        closedMonths=closed_months,
        monthOrder=MONTH_ORDER,
    )

    # --- 01_Conf: parámetros
    _, conf = load_sheet(wb, "01_Conf")
    config = OrderedDict()
    for row in conf:
        k = row.get("Parámetro")
        if k is not None:
            config[str(k)] = num(row.get("Valor"))

    # --- 03_Recursos: centros de trabajo
    _, rec = load_sheet(wb, "03_Recursos")
    centros = [
        OrderedDict(
            id=row.get("ID_CT"),
            nombre=row.get("Centro de Trabajo"),
            proceso=row.get("Proceso"),
            tipo=row.get("Tipo"),
            planificable=as_bool(row.get("Planificable")),
            cantidadRecursos=num(row.get("Cantidad recursos")),
            hsDiaNominal=num(row.get("Hs/día nominal por recurso")),
            turnos=num(row.get("Turnos base")),
            valorHora=row.get("Valor hora absorción USD"),
            activo=as_bool(row.get("Activo")),
            observaciones=row.get("Observaciones"),
        )
        for row in rec
        if row.get("ID_CT")
    ]

    # --- 04_Productos
    _, prod = load_sheet(wb, "04_Productos")
    productos = [
        OrderedDict(
            sku=row.get("SKU"),
            descripcion=row.get("Descripción"),
            familia=row.get("Familia"),
            linea=row.get("Línea"),
            criticidad=row.get("Criticidad"),
        )
        for row in prod
        if row.get("SKU")
    ]

    # --- 06_OF: órdenes de fabricación + campos calculados
    _, ofs = load_sheet(wb, "06_OF")
    ordenes = []
    for row in ofs:
        if not row.get("OF"):
            continue
        f_req = to_date(row.get("Fecha requerida"))
        f_fin = to_date(row.get("Fecha fin real"))
        estado = row.get("Estado")
        terminada = estado == "Terminada" and f_fin is not None
        ordenes.append(
            OrderedDict(
                of=row.get("OF"),
                fechaCreacion=date_str(row.get("Fecha creación")),
                fechaRequerida=date_str(row.get("Fecha requerida")),
                semanaRequerida=num(row.get("Semana requerida")),
                sku=row.get("SKU"),
                descripcion=row.get("Descripción"),
                familia=row.get("Familia"),
                cantidad=num(row.get("Cantidad")),
                prioridad=row.get("Prioridad"),
                estado=estado,
                clienteInterno=row.get("Cliente interno"),
                responsable=row.get("Responsable"),
                fechaInicioReal=date_str(row.get("Fecha inicio real")),
                fechaFinReal=date_str(row.get("Fecha fin real")),
                aTiempo=(f_fin <= f_req) if terminada else None,
                diasAtraso=(f_fin - f_req).days if terminada else None,
                diasParaVencer=(f_req - today).days if f_req else None,
                # Opcional: si el Excel trae "Horas reales" por OF, habilita el contralor
                # de costo real vs. estándar en la vista Mes Cerrado.
                horasReales=num(row.get("Horas reales")) if row.get("Horas reales") is not None else None,
            )
        )

    # --- 07_Modelo_Carga_OF: carga por OF x operación x centro
    _, mod = load_sheet(wb, "07_Modelo_Carga_OF")
    carga_of = [
        OrderedDict(
            of=row.get("OF"),
            fechaRequerida=date_str(row.get("Fecha requerida")),
            anio=num(row.get("Año")),
            mes=cap_mes(row.get("Mes")),
            semana=num(row.get("Semana")),
            sku=row.get("SKU"),
            cantidad=num(row.get("Cantidad")),
            operacion=num(row.get("Operacion N°")),
            centro=row.get("Centro de Trabajo"),
            proceso=row.get("Proceso"),
            hsRequeridas=r2(row.get("Hs requeridas")),
            estado=row.get("Estado OF"),
            prioridad=row.get("Prioridad"),
            planificable=as_bool(row.get("Planificable")),
        )
        for row in mod
        if row.get("OF")
    ]

    # --- 08_Carga_Capacidad: capacidad semanal por centro (dedup (año,semana,centro))
    _, cap = load_sheet(wb, "08_Carga_Capacidad")
    seen = set()
    capacidad = []
    for row in cap:
        centro = row.get("Centro de Trabajo")
        anio = num(row.get("Año"))
        semana = num(row.get("Semana"))
        if centro is None:
            continue
        key = (anio, semana, centro)
        if key in seen:
            continue
        seen.add(key)
        capacidad.append(
            OrderedDict(
                anio=anio,
                mes=cap_mes(row.get("Mes")),
                semana=semana,
                centro=centro,
                proceso=row.get("Proceso"),
                hsRequeridas=r2(row.get("Hs requeridas")),
                hsPlanificables=num(row.get("Hs planificables")),
                utilizacion=r4(row.get("Utilización")),
                estado=row.get("Estado capacidad"),
            )
        )

    # --- Indicadores RECALCULADOS desde capacidad (no se confía en 09A/B/C) ---
    def agg(rows):
        req = round(sum(r["hsRequeridas"] for r in rows), 2)
        plan = int(sum(r["hsPlanificables"] for r in rows))
        util = round(req / plan, 4) if plan else 0.0
        return req, plan, util

    # Semanales: agrupado por (año, semana, mes) -> reproduce el split en borde de mes
    sem_groups = OrderedDict()
    for r in capacidad:
        sem_groups.setdefault((r["anio"], r["semana"], r["mes"]), []).append(r)
    ind_semanales = []
    for (anio, semana, mes), rows in sem_groups.items():
        req, plan, util = agg(rows)
        top = max(rows, key=lambda r: r["hsRequeridas"])
        ind_semanales.append(
            OrderedDict(
                anio=anio,
                semana=semana,
                mes=mes,
                hsRequeridas=req,
                hsPlanificables=plan,
                utilizacionProm=util,
                centrosSaturados=sum(1 for r in rows if r["utilizacion"] >= 1),
                centroMasCargado=top["centro"],
            )
        )
    ind_semanales.sort(key=lambda x: (x["anio"], x["semana"], MONTH_IDX[x["mes"]]))

    # Mensuales: agrupado por (año, mes)
    mes_groups = OrderedDict()
    for r in capacidad:
        mes_groups.setdefault((r["anio"], r["mes"]), []).append(r)
    ind_mensuales = []
    for (anio, mes), rows in mes_groups.items():
        req, plan, util = agg(rows)
        # semanas del mes con al menos un centro saturado
        semanas_sat = sum(
            1 for s in ind_semanales
            if s["anio"] == anio and s["mes"] == mes and s["centrosSaturados"] > 0
        )
        # centro con más horas requeridas en el mes
        por_centro = OrderedDict()
        for r in rows:
            por_centro[r["centro"]] = por_centro.get(r["centro"], 0) + r["hsRequeridas"]
        centro_critico = max(por_centro, key=por_centro.get) if por_centro else None
        ind_mensuales.append(
            OrderedDict(
                anio=anio,
                mes=mes,
                hsRequeridas=req,
                hsPlanificables=plan,
                utilizacionProm=util,
                semanasSaturadas=semanas_sat,
                centroCritico=centro_critico,
            )
        )
    ind_mensuales.sort(key=lambda x: (x["anio"], MONTH_IDX[x["mes"]]))

    # Por centro: agrupado por (año, mes, centro)
    ct_groups = OrderedDict()
    for r in capacidad:
        ct_groups.setdefault((r["anio"], r["mes"], r["centro"]), []).append(r)
    ind_centros = []
    for (anio, mes, centro), rows in ct_groups.items():
        req, plan, util = agg(rows)
        ind_centros.append(
            OrderedDict(
                anio=anio,
                mes=mes,
                centro=centro,
                proceso=rows[0]["proceso"],
                hsRequeridas=req,
                hsPlanificables=plan,
                utilizacion=util,
                estado=estado_util(util),
            )
        )

    return OrderedDict(
        meta=meta,
        config=config,
        centros=centros,
        productos=productos,
        ordenesFabricacion=ordenes,
        cargaOF=carga_of,
        capacidadSemanal=capacidad,
        indicadoresSemanales=ind_semanales,
        indicadoresMensuales=ind_mensuales,
        indicadoresCentros=ind_centros,
    )


# ----------------------------------------------------------------------------- inject
DATA_RE = re.compile(
    r'(<script id="data-store" type="application/json">)(.*?)(</script>)', re.S
)


def inject(html: str, data: "OrderedDict") -> str:
    payload = json.dumps(data, ensure_ascii=False, separators=(", ", ": "))
    if not DATA_RE.search(html):
        sys.exit("No se encontró el <script id=\"data-store\"> en el HTML.")
    return DATA_RE.sub(lambda m: m.group(1) + payload + m.group(3), html, count=1)


def current_json(html: str):
    m = DATA_RE.search(html)
    return json.loads(m.group(2)) if m else None


# ----------------------------------------------------------------------------- cli
def main():
    ap = argparse.ArgumentParser(description="Regenera el JSON del dashboard desde el Excel.")
    ap.add_argument("--xlsx", type=Path, default=DEFAULT_XLSX, help="Planilla de origen (.xlsx)")
    ap.add_argument("--html", type=Path, default=DEFAULT_HTML, help="index.html a actualizar")
    ap.add_argument("--today", default=None, help="Fecha de referencia YYYY-MM-DD (default: hoy)")
    ap.add_argument("--check", action="store_true", help="No escribe; compara con el JSON actual")
    args = ap.parse_args()

    if not args.xlsx.exists():
        sys.exit(f"No existe la planilla: {args.xlsx}")
    today = dt.date.fromisoformat(args.today) if args.today else dt.date.today()

    data = build_data(args.xlsx, today)
    html = args.html.read_text(encoding="utf-8")

    if args.check:
        old = current_json(html)
        new = json.loads(json.dumps(data))
        if old == new:
            print("✓ Sin cambios respecto al JSON embebido actual.")
            return
        for key in new:
            if old is None or old.get(key) != new.get(key):
                a = old.get(key) if old else None
                print(f"Δ {key}: "
                      f"{len(a) if isinstance(a, list) else 'n/a'} -> "
                      f"{len(new[key]) if isinstance(new[key], list) else new[key]}")
        return

    args.html.write_text(inject(html, data), encoding="utf-8")
    m = data["meta"]
    print(f"✓ index.html actualizado — ref {m['today']} (sem {m['currentWeek']}, {m['currentMonth']}), "
          f"{len(data['ordenesFabricacion'])} OF, {len(data['capacidadSemanal'])} filas de capacidad.")


if __name__ == "__main__":
    main()
