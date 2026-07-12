"""Tools de introspeccion visual (Bloque C): shapes, charts y slicers.

Origen: los dashboards reales llevan la logica a la vista en objetos flotantes
(gauges de dona, tarjetas hechas con rectangulos, segmentadores) que
analyze_workbook/read_range no ven. Estas tools los inventarian sin VBA ad hoc.

Reglas COM criticas (MANUAL.md §5):
- UNA propiedad fragil por try/except: una que falle no debe perder las demas (§5.4).
- NUNCA iterar SlicerCacheItems (ni leer su .Count) en un slicer OLAP/Data Model:
  puede colgar >120s (§5.2). Aqui los items solo se tocan si include_items=True
  Y olap=False; si la lectura de .OLAP falla, se asume OLAP (conservador).
- GroupItems se itera por indice 1..Count (For Each es fragil en late binding).
"""

import logging

from ..utils.excel_utils import (
    MSO_SHAPE_CHART,
    MSO_SHAPE_GROUP,
    chart_type_name,
    get_active_workbook,
    get_sheet,
    shape_type_name,
)

logger = logging.getLogger(__name__)

MAX_GROUP_DEPTH = 5
MAX_SLICER_ITEMS = 50


def register(mcp, session, run):
    @mcp.tool()
    def list_shapes(sheet: str) -> dict:
        """Inventario de shapes de una hoja: nombre, tipo, posicion y tamano,
        recursando dentro de grupos (msoGroup). Ve lo que read_range no ve:
        graficos, segmentadores, rectangulos de dashboards, imagenes."""
        return run(_list_shapes, session, sheet)

    @mcp.tool()
    def list_charts(sheet: str) -> dict:
        """Inventario de graficos embebidos de una hoja: tipo (con nombre
        legible), tabla dinamica origen si es PivotChart, y formula de cada
        serie. Incluye graficos anidados dentro de grupos (group_path)."""
        return run(_list_charts, session, sheet)

    @mcp.tool()
    def list_slicers(include_items: bool = False) -> dict:
        """Inventario de segmentadores (SlicerCaches) del workbook activo:
        origen, si es OLAP, que tablas dinamicas controla y donde esta cada
        slicer visible. Los ITEMS solo se listan con include_items=True y
        SOLO para caches no-OLAP: iterar items de un slicer del Data Model
        cuelga la sesion COM (limite conocido del proveedor MSOLAP)."""
        return run(_list_slicers, session, include_items)


# =============================================================================
# Helpers
# =============================================================================


def _prop(obj, name, default=None):
    """Lee UNA propiedad COM fragil; si falla devuelve default sin romper el resto."""
    try:
        return getattr(obj, name)
    except Exception:  # noqa: BLE001 - com_error o attributo inexistente
        return default


def _num(value):
    """float redondeado a 1 decimal o None."""
    try:
        return round(float(value), 1)
    except (TypeError, ValueError):
        return None


# =============================================================================
# list_shapes
# =============================================================================


def _shape_info(shape, depth: int) -> dict:
    type_code = _prop(shape, "Type")
    try:
        type_code = int(type_code) if type_code is not None else None
    except (TypeError, ValueError):
        type_code = None

    visible = _prop(shape, "Visible")
    info = {
        "name": _prop(shape, "Name"),
        "type": shape_type_name(type_code) if type_code is not None else None,
        "type_code": type_code,
        "left": _num(_prop(shape, "Left")),
        "top": _num(_prop(shape, "Top")),
        "width": _num(_prop(shape, "Width")),
        "height": _num(_prop(shape, "Height")),
        "visible": bool(visible) if visible is not None else None,
        "children": None,
    }
    if type_code == MSO_SHAPE_GROUP and depth < MAX_GROUP_DEPTH:
        children = []
        group_items = _prop(shape, "GroupItems")
        if group_items is not None:
            count = _prop(group_items, "Count", 0) or 0
            for i in range(1, int(count) + 1):
                try:
                    children.append(_shape_info(group_items.Item(i), depth + 1))
                except Exception as e:  # noqa: BLE001
                    children.append({"name": None, "error": str(e)})
        info["children"] = children
    return info


def _count_tree(shapes: list) -> int:
    total = 0
    for s in shapes:
        total += 1
        if s.get("children"):
            total += _count_tree(s["children"])
    return total


def _list_shapes(session, sheet: str) -> dict:
    wb = get_active_workbook(session)
    ws = get_sheet(wb, sheet)
    shapes = []
    shapes_col = ws.Shapes  # cachear: cada acceso a ws.Shapes es un round-trip COM
    count = int(shapes_col.Count)
    for i in range(1, count + 1):
        try:
            shapes.append(_shape_info(shapes_col.Item(i), depth=0))
        except Exception as e:  # noqa: BLE001
            shapes.append({"name": None, "error": str(e)})
    return {
        "sheet": str(ws.Name),
        "top_level_count": count,
        "total_count": _count_tree(shapes),
        "shapes": shapes,
    }


# =============================================================================
# list_charts
# =============================================================================


def _walk_chart_shapes(container, path: str, depth: int, found: list) -> None:
    """Recorre un contenedor de shapes acumulando (shape, group_path) de charts."""
    count = _prop(container, "Count", 0) or 0
    for i in range(1, int(count) + 1):
        try:
            shape = container.Item(i)
        except Exception:  # noqa: BLE001
            continue
        type_code = _prop(shape, "Type")
        try:
            type_code = int(type_code) if type_code is not None else None
        except (TypeError, ValueError):
            type_code = None
        if type_code == MSO_SHAPE_CHART:
            found.append((shape, path or None))
        elif type_code == MSO_SHAPE_GROUP and depth < MAX_GROUP_DEPTH:
            name = _prop(shape, "Name") or "?"
            sub_path = f"{path}/{name}" if path else str(name)
            group_items = _prop(shape, "GroupItems")
            if group_items is not None:
                _walk_chart_shapes(group_items, sub_path, depth + 1, found)


def _chart_info(shape, group_path) -> dict:
    info = {
        "name": _prop(shape, "Name"),
        "group_path": group_path,
        "left": _num(_prop(shape, "Left")),
        "top": _num(_prop(shape, "Top")),
        "width": _num(_prop(shape, "Width")),
        "height": _num(_prop(shape, "Height")),
        "chart_type": None,
        "chart_type_name": None,
        "title": None,
        "is_pivot_chart": False,
        "pivot_source": None,
        "series_count": 0,
        "series": [],
    }
    chart = _prop(shape, "Chart")
    if chart is None:
        info["error"] = "Shape.Chart inaccesible"
        return info

    raw_type = _prop(chart, "ChartType")
    if raw_type is not None:
        try:
            info["chart_type"] = int(raw_type)
            info["chart_type_name"] = chart_type_name(raw_type)
        except (TypeError, ValueError):
            info["chart_type_name"] = "mixed/unavailable"

    # Titulo: HasTitle y ChartTitle.Text por separado (cada uno puede fallar)
    if _prop(chart, "HasTitle"):
        title_obj = _prop(chart, "ChartTitle")
        if title_obj is not None:
            info["title"] = _prop(title_obj, "Text")

    # PivotChart: el acceso encadenado LANZA en charts normales -> un try completo
    try:
        pt = chart.PivotLayout.PivotTable
        info["is_pivot_chart"] = True
        info["pivot_source"] = {
            "pivot_table": _prop(pt, "Name"),
            "sheet": _prop(_prop(pt, "Parent"), "Name"),
        }
    except Exception:  # noqa: BLE001 - no es PivotChart (caso normal)
        pass

    # SeriesCollection es METODO en el binding gen_py (early) y propiedad-como-
    # coleccion en dynamic: intentar la llamada primero.
    try:
        series_col = chart.SeriesCollection()
    except Exception:  # noqa: BLE001
        series_col = _prop(chart, "SeriesCollection")
    if series_col is not None:
        n = _prop(series_col, "Count", 0) or 0
        info["series_count"] = int(n)
        for i in range(1, int(n) + 1):
            try:
                ser = series_col.Item(i)
            except Exception:  # noqa: BLE001
                info["series"].append({"index": i, "name": None, "formula": None})
                continue
            info["series"].append({
                "index": i,
                "name": _prop(ser, "Name"),
                "formula": _prop(ser, "Formula"),
            })
    return info


def _list_charts(session, sheet: str) -> dict:
    wb = get_active_workbook(session)
    ws = get_sheet(wb, sheet)

    found: list = []
    _walk_chart_shapes(ws.Shapes, "", 0, found)
    charts = [_chart_info(shape, path) for shape, path in found]
    seen = {c["name"] for c in charts if c["name"]}

    # Pasada secundaria SOLO si hay indicios de omision: ChartObjects() es un
    # subconjunto (charts top-level) del arbol de Shapes ya recorrido, asi que
    # solo puede aportar algo si la lectura de Shape.Type fallo en la caminata.
    # Comparar conteos cuesta 1 round-trip; la pasada completa cuesta N.
    try:
        chart_objects = ws.ChartObjects()
        n = int(_prop(chart_objects, "Count", 0) or 0)
        if n > len(charts):
            for i in range(1, n + 1):
                try:
                    co = chart_objects.Item(i)
                except Exception:  # noqa: BLE001
                    continue
                name = _prop(co, "Name")
                if name and name in seen:
                    continue
                info = _chart_info(co, None)
                info["name"] = name
                charts.append(info)
    except Exception as e:  # noqa: BLE001
        logger.debug("ChartObjects() fallo en %s: %s", sheet, e)

    return {"sheet": str(ws.Name), "chart_count": len(charts), "charts": charts}


# =============================================================================
# list_slicers
# =============================================================================


def _slicer_positions(sc) -> list:
    slicers = []
    col = _prop(sc, "Slicers")
    if col is None:
        return slicers
    n = _prop(col, "Count", 0) or 0
    for i in range(1, int(n) + 1):
        try:
            sl = col.Item(i)
        except Exception:  # noqa: BLE001
            continue
        sheet_name = None
        shape = _prop(sl, "Shape")
        if shape is not None:
            parent = _prop(shape, "Parent")
            if parent is not None:
                sheet_name = _prop(parent, "Name")
        slicers.append({
            "name": _prop(sl, "Name"),
            "caption": _prop(sl, "Caption"),
            "sheet": sheet_name,
            "left": _num(_prop(sl, "Left")),
            "top": _num(_prop(sl, "Top")),
            "width": _num(_prop(sl, "Width")),
            "height": _num(_prop(sl, "Height")),
        })
    return slicers


def _slicer_items(sc) -> dict:
    """SOLO llamar con cache no-OLAP verificado (ver docstring del modulo)."""
    items = {"selected": [], "total": 0, "truncated": False}
    try:
        # La coleccion se llama SlicerItems (SlicerCacheItems NO existe en el
        # modelo de objetos; en caches OLAP los items viven en
        # SlicerCacheLevels(n).SlicerItems, que aqui jamas tocamos).
        try:
            cache_items = sc.SlicerItems
        except AttributeError:
            import win32com.client.dynamic as _dyn

            cache_items = _dyn.Dispatch(sc._oleobj_).SlicerItems
        total = int(cache_items.Count)
        items["total"] = total
        limit = min(total, MAX_SLICER_ITEMS)
        items["truncated"] = total > MAX_SLICER_ITEMS
        for i in range(1, limit + 1):
            try:
                it = cache_items.Item(i)
            except Exception:  # noqa: BLE001
                continue
            if _prop(it, "Selected"):
                caption = _prop(it, "Caption")
                if caption is None:
                    caption = _prop(it, "Name")
                if caption is not None:
                    items["selected"].append(str(caption))
    except Exception as e:  # noqa: BLE001
        items["error"] = str(e)
    return items


def _list_slicers(session, include_items: bool = False) -> dict:
    wb = get_active_workbook(session)
    caches = []
    col = wb.SlicerCaches
    count = int(_prop(col, "Count", 0) or 0)
    for i in range(1, count + 1):
        try:
            sc = col.Item(i)
        except Exception as e:  # noqa: BLE001
            caches.append({"name": None, "error": str(e)})
            continue

        olap_raw = _prop(sc, "OLAP")
        # Si no pudimos leer OLAP, asumimos True: jamas tocar items en la duda
        olap = True if olap_raw is None else bool(olap_raw)

        pivot_tables = []
        try:
            for pt in sc.PivotTables:  # seguro y rapido (MANUAL §5.2)
                pt_name = _prop(pt, "Name")
                pt_sheet = _prop(_prop(pt, "Parent"), "Name")
                pivot_tables.append(f"{pt_name}@{pt_sheet}")
        except Exception as e:  # noqa: BLE001
            pivot_tables = [f"error: {e}"]

        entry = {
            "name": _prop(sc, "Name"),
            "source_name": _prop(sc, "SourceName"),
            "olap": olap,
            "olap_unverified": olap_raw is None,
            "pivot_tables": pivot_tables,
            "slicers": _slicer_positions(sc),
            "items": None,
            "items_note": None,
        }
        if include_items:
            if olap:
                entry["items_note"] = (
                    "omitidos: cache OLAP/Data Model (iterar SlicerCacheItems "
                    "cuelga la sesion COM; limite conocido MANUAL.md §5.2)"
                )
            else:
                entry["items"] = _slicer_items(sc)
        caches.append(entry)

    return {"slicer_cache_count": count, "slicer_caches": caches}
