"""Tools de rangos: valores, formulas, formato y layout."""

import logging
from typing import Optional

from ..utils.excel_utils import (
    get_active_workbook,
    get_sheet,
    matrix_to_jsonable,
    normalize_matrix,
    target_range,
)
from .bulk import MAX_INLINE_CELLS

logger = logging.getLogger(__name__)

# Alineacion (xlHAlign / xlVAlign)
H_ALIGN = {"left": -4131, "center": -4108, "right": -4152, "justify": -4130}
V_ALIGN = {"top": -4160, "center": -4108, "bottom": -4107}

# Bordes: (LineStyle, Weight). xlContinuous=1; thin=2, medium=-4138, thick=4.
# "none" usa xlLineStyleNone=-4142 (limpia todas las aristas).
BORDER_STYLES = {"thin": (1, 2), "medium": (1, -4138), "thick": (1, 4), "none": None}
XL_LINESTYLE_NONE = -4142

# Claves admitidas en una spec de formato (apply_format y apply_format_batch)
FORMAT_KEYS = frozenset({
    "bold", "italic", "font_size", "font_name", "font_color_rgb", "fill_color_rgb",
    "number_format", "h_align", "v_align", "wrap_text", "border",
    "column_width", "row_height", "merge",
})


def register(mcp, session, run):
    @mcp.tool()
    def read_range(sheet: str, range_addr: str) -> list:
        """Leer VALORES de un rango (ej: sheet='Hoja1', range_addr='A1:C10').
        Maximo 50.000 celdas inline; para rangos mayores usa
        export_sheet(sheet, dest, range_addr=...)."""
        return run(_read_range, session, sheet, range_addr)

    @mcp.tool()
    def write_range(sheet: str, range_addr: str, values: list) -> str:
        """Escribir valores en un rango. values = lista de filas (lista de listas)."""
        return run(_write_range, session, sheet, range_addr, values)

    @mcp.tool()
    def read_formulas(sheet: str, range_addr: str, local: bool = False) -> list:
        """Leer FORMULAS de un rango sin evaluar. Celdas sin formula devuelven su
        valor. local=True devuelve formulas en el idioma de la UI (ej: =SUMA)."""
        return run(_read_formulas, session, sheet, range_addr, local)

    @mcp.tool()
    def write_formulas(sheet: str, range_addr: str, formulas: list, local: bool = False) -> str:
        """Escribir formulas en un rango (matriz 2D de strings '=...').
        local=True interpreta formulas en el idioma de la UI (ej: =SUMA)."""
        return run(_write_formulas, session, sheet, range_addr, formulas, local)

    @mcp.tool()
    def apply_format(
        sheet: str,
        range_addr: str,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        font_size: Optional[int] = None,
        font_name: Optional[str] = None,
        font_color_rgb: Optional[str] = None,
        fill_color_rgb: Optional[str] = None,
        number_format: Optional[str] = None,
        h_align: Optional[str] = None,
        v_align: Optional[str] = None,
        wrap_text: Optional[bool] = None,
        border: Optional[str] = None,
        column_width: Optional[float] = None,
        row_height: Optional[float] = None,
        merge: Optional[bool] = None,
    ) -> str:
        """Aplicar formato a un rango (una llamada COM por propiedad, no por celda).

        Colores hex 'RRGGBB'. number_format ej: '#,##0.00', 'dd/mm/yyyy', '0.0%'.
        h_align: left|center|right|justify. v_align: top|center|bottom.
        border: thin|medium|thick|none (todas las aristas del rango).
        column_width/row_height afectan las columnas/filas completas del rango.
        Para estilizar muchas zonas de golpe usa apply_format_batch."""
        spec = {k: v for k, v in {
            "bold": bold, "italic": italic, "font_size": font_size,
            "font_name": font_name, "font_color_rgb": font_color_rgb,
            "fill_color_rgb": fill_color_rgb, "number_format": number_format,
            "h_align": h_align, "v_align": v_align, "wrap_text": wrap_text,
            "border": border, "column_width": column_width,
            "row_height": row_height, "merge": merge,
        }.items() if v is not None}
        return run(_apply_format, session, sheet, range_addr, spec)

    @mcp.tool()
    def apply_format_batch(sheet: str, formats: list) -> dict:
        """Aplicar N especificaciones de formato en UNA llamada (todo el estilo de
        un informe en un solo viaje).

        formats = [{"range": "A1:F1", "bold": true, "fill_color_rgb": "D9E1F2",
                    "number_format": "#,##0.00", ...}] — mismas claves que
        apply_format. Un item puede traer "sheet" propia (informes multi-hoja).
        TODO el batch se valida ANTES de aplicar nada: una spec invalida no deja
        el formato a medias."""
        return run(_apply_format_batch, session, sheet, formats)

    @mcp.tool()
    def set_freeze_panes(sheet: str, at: Optional[str] = None) -> dict:
        """Congelar paneles. at='C5' congela las filas 1-4 y las columnas A-B
        (todo lo arriba/izquierda de esa celda). at=None descongela."""
        return run(_set_freeze_panes, session, sheet, at)

    @mcp.tool()
    def auto_fit_columns(sheet: str, range_addr: Optional[str] = None) -> str:
        """Auto-ajustar ancho de columnas (todas si no se indica rango)."""
        return run(_auto_fit_columns, session, sheet, range_addr)


def _get_sheet(session, sheet: str):
    wb = get_active_workbook(session)
    return wb.Sheets(sheet)


def _read_range(session, sheet: str, range_addr: str) -> list:
    ws = _get_sheet(session, sheet)
    rng = ws.Range(range_addr)
    n_cells = int(rng.Rows.Count) * int(rng.Columns.Count)
    if n_cells > MAX_INLINE_CELLS:
        raise ValueError(
            f"Rango {range_addr} tiene {n_cells} celdas (> {MAX_INLINE_CELLS}). "
            "Usa export_sheet(sheet, dest, range_addr=...) para volcarlo a archivo."
        )
    return matrix_to_jsonable(rng.Value)


def _write_range(session, sheet: str, range_addr: str, values: list) -> str:
    if not values:
        return "OK: 0 filas escritas (values vacio)"
    ws = _get_sheet(session, sheet)
    matrix, n_rows, n_cols = normalize_matrix(values)
    target_range(ws, range_addr, n_rows, n_cols).Value = matrix
    logger.info("Escritas %sx%s celdas en %s!%s", n_rows, n_cols, sheet, range_addr)
    return f"OK: {n_rows} filas x {n_cols} columnas escritas en {range_addr}"


def _read_formulas(session, sheet: str, range_addr: str, local: bool) -> list:
    ws = _get_sheet(session, sheet)
    rng = ws.Range(range_addr)
    formulas = rng.FormulaLocal if local else rng.Formula
    return matrix_to_jsonable(formulas)


def _write_formulas(session, sheet: str, range_addr: str, formulas: list, local: bool) -> str:
    if not formulas:
        return "OK: 0 formulas escritas (lista vacia)"
    ws = _get_sheet(session, sheet)
    matrix, n_rows, n_cols = normalize_matrix(formulas)
    target = target_range(ws, range_addr, n_rows, n_cols)
    if local:
        target.FormulaLocal = matrix
    else:
        target.Formula = matrix
    logger.info("Escritas %sx%s formulas en %s!%s", n_rows, n_cols, sheet, range_addr)
    return f"OK: {n_rows} filas x {n_cols} columnas de formulas escritas en {range_addr}"


def _hex_to_ole_color(rgb_hex: str) -> int:
    """'RRGGBB' -> entero BGR que espera COM (Font.Color / Interior.Color)."""
    rgb_hex = rgb_hex.lstrip("#")
    r, g, b = int(rgb_hex[0:2], 16), int(rgb_hex[2:4], 16), int(rgb_hex[4:6], 16)
    return (b << 16) | (g << 8) | r


def _validate_spec(spec: dict, label: str) -> None:
    """Valida claves y valores de una spec de formato. Lanza ValueError accionable.

    Es una funcion PURA (sin COM): el batch la corre sobre TODOS los items antes
    de tocar la hoja, para no dejar el formato a medias."""
    desconocidas = set(spec) - FORMAT_KEYS - {"range", "sheet"}
    if desconocidas:
        raise ValueError(
            f"{label}: claves desconocidas {sorted(desconocidas)}. "
            f"Validas: {sorted(FORMAT_KEYS)}"
        )
    if "h_align" in spec and spec["h_align"] not in H_ALIGN:
        raise ValueError(f"{label}: h_align '{spec['h_align']}' invalido "
                         f"(usa {sorted(H_ALIGN)})")
    if "v_align" in spec and spec["v_align"] not in V_ALIGN:
        raise ValueError(f"{label}: v_align '{spec['v_align']}' invalido "
                         f"(usa {sorted(V_ALIGN)})")
    if "border" in spec and spec["border"] not in BORDER_STYLES:
        raise ValueError(f"{label}: border '{spec['border']}' invalido "
                         f"(usa {sorted(BORDER_STYLES)})")
    for key in ("font_color_rgb", "fill_color_rgb"):
        if key in spec:
            try:
                _hex_to_ole_color(str(spec[key]))
            except (ValueError, IndexError):
                raise ValueError(f"{label}: {key} '{spec[key]}' no es hex 'RRGGBB'")


def _apply_spec(rng, spec: dict) -> list:
    """Aplica una spec (ya validada) a un Range. Devuelve las props aplicadas."""
    applied = []
    if "bold" in spec:
        rng.Font.Bold = bool(spec["bold"])
        applied.append("bold")
    if "italic" in spec:
        rng.Font.Italic = bool(spec["italic"])
        applied.append("italic")
    if "font_size" in spec:
        rng.Font.Size = spec["font_size"]
        applied.append("font_size")
    if "font_name" in spec:
        rng.Font.Name = str(spec["font_name"])
        applied.append("font_name")
    if "font_color_rgb" in spec:
        rng.Font.Color = _hex_to_ole_color(str(spec["font_color_rgb"]))
        applied.append("font_color")
    if "fill_color_rgb" in spec:
        rng.Interior.Color = _hex_to_ole_color(str(spec["fill_color_rgb"]))
        applied.append("fill_color")
    if "number_format" in spec:
        rng.NumberFormat = str(spec["number_format"])
        applied.append("number_format")
    if "h_align" in spec:
        rng.HorizontalAlignment = H_ALIGN[spec["h_align"]]
        applied.append("h_align")
    if "v_align" in spec:
        rng.VerticalAlignment = V_ALIGN[spec["v_align"]]
        applied.append("v_align")
    if "wrap_text" in spec:
        rng.WrapText = bool(spec["wrap_text"])
        applied.append("wrap_text")
    if "border" in spec:
        estilo = BORDER_STYLES[spec["border"]]
        if estilo is None:
            rng.Borders.LineStyle = XL_LINESTYLE_NONE
        else:
            rng.Borders.LineStyle, rng.Borders.Weight = estilo
        applied.append("border")
    if "column_width" in spec:
        rng.ColumnWidth = spec["column_width"]  # columnas completas del rango
        applied.append("column_width")
    if "row_height" in spec:
        rng.RowHeight = spec["row_height"]
        applied.append("row_height")
    if spec.get("merge"):
        rng.Merge()
        applied.append("merge")
    return applied


def _apply_format(session, sheet: str, range_addr: str, spec: dict) -> str:
    _validate_spec(spec, f"formato {range_addr}")
    ws = _get_sheet(session, sheet)
    applied = _apply_spec(ws.Range(range_addr), spec)
    return f"OK: formato aplicado a {range_addr}: {', '.join(applied) or 'nada'}"


def _apply_format_batch(session, sheet: str, formats: list) -> dict:
    if not formats:
        return {"applied": 0, "failed": 0, "details": []}
    wb = get_active_workbook(session)
    # Pre-vuelo COMPLETO sin COM: si un item esta mal, no se aplica NADA.
    for i, spec in enumerate(formats):
        if not isinstance(spec, dict) or not spec.get("range"):
            raise ValueError(f"formats[{i}]: cada item es un dict con clave 'range'")
        _validate_spec(spec, f"formats[{i}] ({spec['range']})")
    applied, details = 0, []
    for i, spec in enumerate(formats):
        target_sheet = spec.get("sheet") or sheet
        try:
            ws = get_sheet(wb, target_sheet)
            props = _apply_spec(ws.Range(spec["range"]), spec)
            applied += 1
            details.append({"range": f"{target_sheet}!{spec['range']}", "applied": props})
        except Exception as e:
            # Ya validado: un fallo aqui es COM puro (ej. rango malformado).
            logger.warning("apply_format_batch item %d fallo: %s", i, e)
            details.append({"range": f"{target_sheet}!{spec.get('range')}",
                            "error": str(e)[:200]})
    return {"applied": applied, "failed": len(formats) - applied, "details": details}


def _set_freeze_panes(session, sheet: str, at) -> dict:
    wb = get_active_workbook(session)
    ws = get_sheet(wb, sheet)
    # FreezePanes es propiedad de la VENTANA activa: activar la hoja primero.
    # Verificado en instancia oculta: funciona sin Select usando SplitRow/Col.
    ws.Activate()
    win = ws.Application.ActiveWindow
    win.FreezePanes = False
    if at is None:
        win.SplitRow = 0
        win.SplitColumn = 0
        return {"sheet": str(ws.Name), "frozen": False}
    celda = ws.Range(at)  # resuelve 'C5' -> Row/Column (no parsear A1 a mano)
    win.SplitRow = int(celda.Row) - 1
    win.SplitColumn = int(celda.Column) - 1
    win.FreezePanes = True
    return {"sheet": str(ws.Name), "frozen": True, "at": str(celda.Address),
            "frozen_rows": int(celda.Row) - 1, "frozen_cols": int(celda.Column) - 1}


def _auto_fit_columns(session, sheet: str, range_addr: Optional[str]) -> str:
    ws = _get_sheet(session, sheet)
    if range_addr:
        ws.Range(range_addr).Columns.AutoFit()
    else:
        ws.UsedRange.Columns.AutoFit()
    return f"OK: columnas auto-ajustadas en {sheet}"
