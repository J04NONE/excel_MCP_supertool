"""Tools de lectura masiva: tablas completas y export de hojas/rangos a archivo.

Fast path COM: UNA llamada .Value por bloque (nunca celda a celda). El tamano
se verifica ANTES de materializar datos (Rows.Count/Columns.Count son baratos).
"""

import csv
import json
import logging
import os
from typing import Optional

from ..utils.excel_utils import get_active_workbook, get_sheet, matrix_to_jsonable

logger = logging.getLogger(__name__)

MAX_INLINE_CELLS = 50_000
MAX_EXPORT_CELLS = 5_000_000
SAMPLE_ROWS_DEFAULT = 5


def register(mcp, session, run):
    @mcp.tool()
    def read_table(table_name: str, dest: Optional[str] = None) -> dict:
        """Leer una tabla (ListObject) completa por nombre (case-insensitive).

        Sin dest: inline (max 50.000 celdas) -> {table, sheet, headers, rows,
        row_count, col_count}; 'rows' es la MATRIZ de datos.
        Con dest (.csv/.tsv/.json): escribe el archivo (headers en la primera
        fila) -> {file, rows, cols, headers, sample}; 'rows'/'cols' son CONTEOS."""
        return run(_read_table, session, table_name, dest)

    @mcp.tool()
    def export_sheet(sheet: str, dest: str, sample_rows: int = 5,
                     range_addr: Optional[str] = None) -> dict:
        """Exportar una hoja (UsedRange) o un rango (range_addr='B5:X9000') a
        .csv/.tsv/.json en UNA llamada COM. Techo: 5M celdas.

        OJO: UsedRange puede NO empezar en A1 — usa la clave 'range' de la
        respuesta para mapear offsets. Celdas combinadas: el valor queda solo
        en la celda superior-izquierda (el resto sale null)."""
        return run(_export_sheet, session, sheet, dest, sample_rows, range_addr)


def _norm_cell(v):
    """Excel guarda TODO numero como double: un entero llega como 1.0. Al exportar
    eso ensucia (y peor: un codigo PT '5060094' saldria '5060094.0' y rompe cruces).
    Se colapsan los floats de valor entero a int. Los 8-9 digitos de un codigo PT
    caben exactos en double (< 2^53), no hay perdida de precision."""
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def _norm_matrix(matrix: list) -> list:
    return [[_norm_cell(c) for c in row] for row in matrix]


def _find_table(wb, table_name: str):
    """(worksheet, ListObject) por nombre case-insensitive, o error con listado."""
    disponibles = []
    for ws in wb.Worksheets:
        for lo in ws.ListObjects:
            disponibles.append(f"{lo.Name} (hoja {ws.Name})")
            if str(lo.Name).lower() == table_name.lower():
                return ws, lo
    listado = ", ".join(disponibles) or "ninguna"
    raise ValueError(f"Tabla '{table_name}' no encontrada. Disponibles: {listado}")


def _headers_of(lo, n_cols: int) -> list:
    try:
        if lo.ShowHeaders:
            hdr = matrix_to_jsonable(lo.HeaderRowRange.Value)
            if hdr:
                return [str(h) if h is not None else f"col{i + 1}"
                        for i, h in enumerate(hdr[0])]
    except Exception:
        logger.debug("HeaderRowRange inaccesible; headers sinteticos")
    return [f"col{i + 1}" for i in range(n_cols)]


def _ext_of(dest: str) -> str:
    ext = os.path.splitext(dest)[1].lower()
    if ext not in (".csv", ".tsv", ".json"):
        raise ValueError(f"Extension no soportada: '{ext}' (usa .csv, .tsv o .json)")
    return ext


def _write_file(dest: str, matrix: list) -> str:
    """Escribe matriz 2D JSON-safe a dest segun extension. Devuelve ruta absoluta."""
    dest = os.path.abspath(dest)
    ext = _ext_of(dest)
    if ext == ".json":
        with open(dest, "w", encoding="utf-8") as f:
            json.dump(matrix, f, ensure_ascii=False)
    else:
        delim = "," if ext == ".csv" else "\t"
        # utf-8-sig: BOM para que Excel reabra el archivo con acentos correctos
        with open(dest, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=delim, quoting=csv.QUOTE_MINIMAL)
            for row in matrix:
                w.writerow(["" if c is None else c for c in row])
    return dest


def _read_table(session, table_name: str, dest) -> dict:
    wb = get_active_workbook(session)
    ws, lo = _find_table(wb, table_name)
    body = lo.DataBodyRange  # None si la tabla no tiene filas de datos
    if body is None:
        n_rows, n_cols = 0, int(lo.ListColumns.Count)
    else:
        n_rows, n_cols = int(body.Rows.Count), int(body.Columns.Count)
    headers = _headers_of(lo, n_cols)
    if not dest and n_rows * n_cols > MAX_INLINE_CELLS:
        raise ValueError(
            f"Tabla '{lo.Name}' tiene {n_rows * n_cols} celdas "
            f"(> {MAX_INLINE_CELLS}): pasa dest='ruta.csv|.tsv|.json'."
        )
    rows = _norm_matrix(matrix_to_jsonable(body.Value)) if body is not None else []
    if dest:
        path = _write_file(dest, [headers] + rows)
        return {"file": path, "rows": n_rows, "cols": n_cols, "headers": headers,
                "sample": rows[:SAMPLE_ROWS_DEFAULT]}
    return {"table": str(lo.Name), "sheet": str(ws.Name), "headers": headers,
            "rows": rows, "row_count": n_rows, "col_count": n_cols}


def _export_sheet(session, sheet: str, dest: str, sample_rows: int,
                  range_addr) -> dict:
    wb = get_active_workbook(session)
    ws = get_sheet(wb, sheet)
    ur = ws.Range(range_addr) if range_addr else ws.UsedRange
    n_rows, n_cols = int(ur.Rows.Count), int(ur.Columns.Count)
    if n_rows * n_cols > MAX_EXPORT_CELLS:
        raise ValueError(
            f"{n_rows * n_cols} celdas (> {MAX_EXPORT_CELLS}): exporta por partes "
            "con range_addr o divide la hoja."
        )
    values = _norm_matrix(matrix_to_jsonable(ur.Value))
    path = _write_file(dest, values)
    return {"file": path, "rows": n_rows, "cols": n_cols,
            "range": str(ur.Address), "sample": values[:max(0, int(sample_rows))]}
