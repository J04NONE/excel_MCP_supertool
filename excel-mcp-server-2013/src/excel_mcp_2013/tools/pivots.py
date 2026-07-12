"""Tools de tablas dinamicas: inventario, creacion nativa y refresco."""

import logging
from typing import Optional

from ..utils.excel_utils import (
    LONG_OP_TIMEOUT,
    XL_AGG_FUNCTIONS,
    XL_COLUMN_FIELD,
    XL_DATA_FIELD,
    XL_PAGE_FIELD,
    XL_ROW_FIELD,
    get_active_workbook,
)

logger = logging.getLogger(__name__)


def register(mcp, session, run):
    @mcp.tool()
    def list_pivot_tables() -> list:
        """Inventario de tablas dinamicas del workbook activo: hoja, origen de
        datos y campos por zona (filas, columnas, filtros, valores)."""
        return run(_list_pivot_tables, session)

    @mcp.tool()
    def create_pivot_table(
        source_sheet: str,
        source_range: str,
        dest_sheet: str,
        dest_cell: str,
        rows: Optional[list] = None,
        columns: Optional[list] = None,
        filters: Optional[list] = None,
        values: Optional[list] = None,
        name: Optional[str] = None,
    ) -> dict:
        """Crear una tabla dinamica NATIVA. values = [{"field": str,
        "agg": "sum|count|average|max|min"}]. rows/columns/filters = nombres de campo."""
        return run(
            _create_pivot_table, session, source_sheet, source_range,
            dest_sheet, dest_cell, rows, columns, filters, values, name,
        )

    @mcp.tool()
    def refresh_pivot_tables(sheet: Optional[str] = None) -> dict:
        """Refrescar todas las tablas dinamicas (de una hoja o de todo el libro)."""
        return run(_refresh_pivot_tables, session, sheet, timeout=LONG_OP_TIMEOUT)


def _pivot_fields_by_orientation(pt) -> dict:
    zones = {"rows": [], "columns": [], "filters": [], "values": []}
    try:
        zones["rows"] = [str(f.Name) for f in pt.RowFields]
    except Exception:
        pass
    try:
        zones["columns"] = [str(f.Name) for f in pt.ColumnFields]
    except Exception:
        pass
    try:
        zones["filters"] = [str(f.Name) for f in pt.PageFields]
    except Exception:
        pass
    try:
        zones["values"] = [str(f.Name) for f in pt.DataFields]
    except Exception:
        pass
    return zones


def _list_pivot_tables(session) -> list:
    wb = get_active_workbook(session)
    pivots = []
    for ws in wb.Worksheets:
        try:
            sheet_pivots = ws.PivotTables()
        except Exception:
            continue
        for pt in sheet_pivots:
            info = {"name": str(pt.Name), "sheet": str(ws.Name)}
            try:
                info["location"] = str(pt.TableRange2.Address).replace("$", "")
            except Exception:
                pass
            try:
                info["source"] = str(pt.PivotCache().SourceData)
            except Exception:
                info["source"] = None
            info["fields"] = _pivot_fields_by_orientation(pt)
            pivots.append(info)
    return pivots


def _create_pivot_table(
    session, source_sheet, source_range, dest_sheet, dest_cell,
    rows, columns, filters, values, name,
) -> dict:
    wb = get_active_workbook(session)
    app = session.get_application()

    source = f"'{source_sheet}'!{source_range}"
    dest_ws = wb.Sheets(dest_sheet)
    dest = dest_ws.Range(dest_cell)

    cache = wb.PivotCaches().Create(SourceType=1, SourceData=source)  # xlDatabase
    pt_name = name or f"TablaDinamica_{wb.PivotCaches().Count}"
    pt = cache.CreatePivotTable(TableDestination=dest, TableName=pt_name)

    for field in rows or []:
        pt.PivotFields(field).Orientation = XL_ROW_FIELD
    for field in columns or []:
        pt.PivotFields(field).Orientation = XL_COLUMN_FIELD
    for field in filters or []:
        pt.PivotFields(field).Orientation = XL_PAGE_FIELD
    for spec in values or []:
        field = spec["field"]
        agg = str(spec.get("agg", "sum")).lower()
        func = XL_AGG_FUNCTIONS.get(agg)
        if func is None:
            raise ValueError(f"Agregacion '{agg}' invalida. Usa: {sorted(XL_AGG_FUNCTIONS)}")
        df = pt.AddDataField(pt.PivotFields(field))
        df.Function = func

    logger.info("Tabla dinamica '%s' creada en %s!%s", pt_name, dest_sheet, dest_cell)
    return {
        "name": pt_name,
        "sheet": dest_sheet,
        "location": str(pt.TableRange2.Address).replace("$", ""),
        "fields": _pivot_fields_by_orientation(pt),
    }


def _refresh_pivot_tables(session, sheet: Optional[str]) -> dict:
    wb = get_active_workbook(session)
    refreshed, failed = [], []
    sheets = [wb.Sheets(sheet)] if sheet else list(wb.Worksheets)
    for ws in sheets:
        try:
            sheet_pivots = ws.PivotTables()
        except Exception:
            continue
        for pt in sheet_pivots:
            try:
                pt.RefreshTable()
                refreshed.append(f"{ws.Name}!{pt.Name}")
            except Exception as e:
                failed.append({"pivot": f"{ws.Name}!{pt.Name}", "error": str(e)})
    return {"refreshed": refreshed, "failed": failed}
