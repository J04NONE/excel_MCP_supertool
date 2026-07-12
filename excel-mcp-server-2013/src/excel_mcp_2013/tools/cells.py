"""Tools de rangos: valores, formulas y formato."""

import logging
from typing import Optional

from ..utils.excel_utils import (
    get_active_workbook,
    matrix_to_jsonable,
    normalize_matrix,
    target_range,
)

logger = logging.getLogger(__name__)


def register(mcp, session, run):
    @mcp.tool()
    def read_range(sheet: str, range_addr: str) -> list:
        """Leer VALORES de un rango (ej: sheet='Hoja1', range_addr='A1:C10')."""
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
        font_color_rgb: Optional[str] = None,
        fill_color_rgb: Optional[str] = None,
        number_format: Optional[str] = None,
    ) -> str:
        """Aplicar formato a un rango. Colores en hex 'RRGGBB'.
        number_format ej: '#,##0.00', 'dd/mm/yyyy', '0.0%'."""
        return run(
            _apply_format, session, sheet, range_addr,
            bold, italic, font_size, font_color_rgb, fill_color_rgb, number_format,
        )

    @mcp.tool()
    def auto_fit_columns(sheet: str, range_addr: Optional[str] = None) -> str:
        """Auto-ajustar ancho de columnas (todas si no se indica rango)."""
        return run(_auto_fit_columns, session, sheet, range_addr)


def _get_sheet(session, sheet: str):
    wb = get_active_workbook(session)
    return wb.Sheets(sheet)


def _read_range(session, sheet: str, range_addr: str) -> list:
    ws = _get_sheet(session, sheet)
    return matrix_to_jsonable(ws.Range(range_addr).Value)


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


def _apply_format(
    session, sheet, range_addr, bold, italic, font_size,
    font_color_rgb, fill_color_rgb, number_format,
) -> str:
    ws = _get_sheet(session, sheet)
    rng = ws.Range(range_addr)
    applied = []
    if bold is not None:
        rng.Font.Bold = bold
        applied.append("bold")
    if italic is not None:
        rng.Font.Italic = italic
        applied.append("italic")
    if font_size is not None:
        rng.Font.Size = font_size
        applied.append("font_size")
    if font_color_rgb:
        rng.Font.Color = _hex_to_ole_color(font_color_rgb)
        applied.append("font_color")
    if fill_color_rgb:
        rng.Interior.Color = _hex_to_ole_color(fill_color_rgb)
        applied.append("fill_color")
    if number_format:
        rng.NumberFormat = number_format
        applied.append("number_format")
    return f"OK: formato aplicado a {range_addr}: {', '.join(applied) or 'nada'}"


def _auto_fit_columns(session, sheet: str, range_addr: Optional[str]) -> str:
    ws = _get_sheet(session, sheet)
    if range_addr:
        ws.Range(range_addr).Columns.AutoFit()
    else:
        ws.UsedRange.Columns.AutoFit()
    return f"OK: columnas auto-ajustadas en {sheet}"
