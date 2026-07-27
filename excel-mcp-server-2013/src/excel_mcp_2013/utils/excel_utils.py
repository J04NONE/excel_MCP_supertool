"""Helpers COM compartidos por los modulos de tools."""

import logging
from datetime import date, datetime, time as dtime

logger = logging.getLogger(__name__)

# --- Constantes Excel (xl*) ---
XL_CALCULATION_MANUAL = -4135
XL_CALCULATION_AUTOMATIC = -4105
XL_CELL_TYPE_FORMULAS = -4123

# FileFormat para SaveAs
XL_FILE_FORMATS = {
    ".xlsx": 51,  # xlOpenXMLWorkbook
    ".xlsm": 52,  # xlOpenXMLWorkbookMacroEnabled
    ".xlsb": 50,  # xlExcel12
    ".xls": 56,   # xlExcel8
}

# PivotField orientations
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
XL_PAGE_FIELD = 3
XL_DATA_FIELD = 4

# Funciones de agregacion para pivots
XL_AGG_FUNCTIONS = {
    "sum": -4157,
    "count": -4112,
    "average": -4106,
    "max": -4136,
    "min": -4139,
    "product": -4149,
    "count_numbers": -4113,
    "stdev": -4155,
    "var": -4164,
}

# msoAutomationSecurity
MSO_AUTOMATION_SECURITY_LOW = 1
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

# VBA component types (VBIDE.vbext_ComponentType)
VBA_COMPONENT_TYPES = {
    1: "standard_module",
    2: "class_module",
    3: "userform",
    100: "document",  # ThisWorkbook / hojas
}

# Timeout para operaciones COM largas (refresh PQ/pivots/modelo, document_workbook)
LONG_OP_TIMEOUT = 600

# MsoShapeType (introspeccion visual, Bloque C)
MSO_SHAPE_TYPES = {
    -2: "mixed",
    1: "auto_shape",
    2: "callout",
    3: "chart",
    4: "comment",
    5: "freeform",
    6: "group",
    7: "embedded_ole_object",
    8: "form_control",
    9: "line",
    10: "linked_ole_object",
    11: "linked_picture",
    12: "ole_control",
    13: "picture",
    14: "placeholder",
    15: "text_effect",
    16: "media",
    17: "text_box",
    19: "table",
    21: "diagram",
    24: "smart_art",
    25: "slicer",
    28: "graphic",
    30: "3d_model",
}
MSO_SHAPE_GROUP = 6  # msoGroup -> recursar GroupItems
MSO_SHAPE_CHART = 3  # msoChart

# XlChartType numero -> nombre legible (los mas comunes; fallback "unknown(<n>)")
# OJO: -4111 es xlCombination (grupos de tipos mezclados; los "gauges" de
# dashboards leen asi), y xlDoughnut es -4120 (verificado contra Excel real:
# asignar -4111 como si fuera dona lanza E_INVALIDARG).
XL_CHART_TYPE_NAMES = {
    -4169: "xlXYScatter",
    -4151: "xlRadar",
    -4120: "xlDoughnut",
    -4111: "xlCombination",
    -4102: "xl3DPie",
    -4101: "xl3DLine",
    -4100: "xl3DColumn",
    -4098: "xl3DArea",
    1: "xlArea",
    4: "xlLine",
    5: "xlPie",
    15: "xlBubble",
    51: "xlColumnClustered",
    52: "xlColumnStacked",
    53: "xlColumnStacked100",
    54: "xl3DColumnClustered",
    55: "xl3DColumnStacked",
    56: "xl3DColumnStacked100",
    57: "xlBarClustered",
    58: "xlBarStacked",
    59: "xlBarStacked100",
    60: "xl3DBarClustered",
    63: "xlLineStacked",
    65: "xlLineMarkers",
    68: "xlPieOfPie",
    69: "xlPieExploded",
    71: "xlBarOfPie",
    72: "xlXYScatterSmooth",
    74: "xlXYScatterLines",
    76: "xlAreaStacked",
    77: "xlAreaStacked100",
    80: "xlDoughnutExploded",
    83: "xlSurface",
    92: "xlCylinderColClustered",
    93: "xlCylinderColStacked",
    94: "xlCylinderColStacked100",
}


def shape_type_name(code) -> str:
    """Nombre legible de un MsoShapeType (fallback unknown(<n>))."""
    try:
        return MSO_SHAPE_TYPES.get(int(code), f"unknown({int(code)})")
    except (TypeError, ValueError):
        return f"unknown({code!r})"


def chart_type_name(code) -> str:
    """Nombre legible de un XlChartType (fallback unknown(<n>))."""
    try:
        return XL_CHART_TYPE_NAMES.get(int(code), f"unknown({int(code)})")
    except (TypeError, ValueError):
        return f"unknown({code!r})"


# Celdas en ERROR: COM las entrega como VT_ERROR, que pywin32 convierte a un int
# con el scode 0x800A0000 | codigo (verificado: #N/A -> -2146826246 = 0x800A07FA).
# Sin traducir, un #N/A saldria como el numero -2146826246 y contaminaria lecturas
# y exports. Los valores numericos reales SIEMPRE llegan como float, nunca como int,
# asi que la deteccion no tiene falsos positivos en la practica.
XL_CVERR_NAMES = {
    2000: "#NULL!",
    2007: "#DIV/0!",
    2015: "#VALUE!",
    2023: "#REF!",
    2029: "#NAME?",
    2036: "#NUM!",
    2042: "#N/A",
    2043: "#GETTING_DATA",  # celda de cubo con la fuente OLAP inalcanzable
}
_CVERR_HI = 0x800A
_CVERR_MIN, _CVERR_MAX = 2000, 2100


def excel_error_name(value):
    """Nombre del error de Excel si `value` es un VT_ERROR, si no None.

    Fallback "Error <codigo>" para codigos del rango que no esten mapeados.
    """
    if not isinstance(value, int) or isinstance(value, bool):
        return None
    unsigned = value & 0xFFFFFFFF
    if (unsigned >> 16) != _CVERR_HI:
        return None
    code = unsigned & 0xFFFF
    if not (_CVERR_MIN <= code <= _CVERR_MAX):
        return None
    return XL_CVERR_NAMES.get(code, f"Error {code}")


def to_jsonable(value):
    """Convierte un valor COM (pywintypes datetime, error de celda, etc.) a
    tipo JSON-safe."""
    if isinstance(value, (datetime, date, dtime)):
        return value.isoformat()
    err = excel_error_name(value)
    if err is not None:
        return err
    if value is None or isinstance(value, (bool, int, float, str)):
        return value
    return str(value)


def matrix_to_jsonable(values):
    """Convierte el resultado de Range.Value/.Formula a matriz 2D JSON-safe."""
    if values is None:
        return []
    if not isinstance(values, tuple):  # celda unica
        return [[to_jsonable(values)]]
    return [[to_jsonable(cell) for cell in row] for row in values]


def normalize_matrix(values: list):
    """Normaliza una lista (posiblemente 1D) a tuple-of-tuples rectangular.

    Returns (matrix, n_rows, n_cols).
    """
    if not isinstance(values[0], list):
        values = [values]
    n_rows = len(values)
    n_cols = max(len(row) for row in values)
    matrix = tuple(
        tuple(row[i] if i < len(row) else None for i in range(n_cols)) for row in values
    )
    return matrix, n_rows, n_cols


def target_range(ws, range_addr: str, n_rows: int, n_cols: int):
    """Rango destino de n_rows x n_cols anclado en la esquina de range_addr.

    NO usar Range(...).Resize(r, c) con late binding: pywin32 lo resuelve como
    Range.Item(r, c) y desplaza el destino (bug real detectado en pruebas).
    """
    start = ws.Range(range_addr)
    r1, c1 = start.Row, start.Column
    return ws.Range(ws.Cells(r1, c1), ws.Cells(r1 + n_rows - 1, c1 + n_cols - 1))


def get_active_workbook(session):
    """Workbook activo o error claro si no hay ninguno."""
    wb = session.get_application().ActiveWorkbook
    if not wb:
        raise RuntimeError("No hay workbook activo. Usa open_workbook primero.")
    return wb


def get_sheet(wb, sheet: str):
    """Hoja por nombre, con error claro que lista las hojas disponibles."""
    try:
        return wb.Worksheets(sheet)
    except Exception:
        names = [ws.Name for ws in wb.Worksheets]
        raise RuntimeError(f"Hoja '{sheet}' no encontrada. Hojas: {names}")
