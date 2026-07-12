"""Tools de ciclo de vida y radiografia de workbooks."""

import logging
import os
from typing import Optional

from ..utils.excel_utils import (
    XL_CELL_TYPE_FORMULAS,
    XL_FILE_FORMATS,
    get_active_workbook,
)

logger = logging.getLogger(__name__)


def register(mcp, session, run):
    @mcp.tool()
    def save_workbook(path: Optional[str] = None) -> dict:
        """Guardar el workbook activo. Con path hace SaveAs (formato por extension)."""
        return run(_save_workbook, session, path)

    @mcp.tool()
    def close_workbook(save_changes: bool = False) -> dict:
        """Cerrar el workbook activo (por defecto SIN guardar cambios)."""
        return run(_close_workbook, session, save_changes)

    @mcp.tool()
    def create_sheet(name: str, after: Optional[str] = None) -> dict:
        """Crear una hoja nueva en el workbook activo."""
        return run(_create_sheet, session, name, after)

    @mcp.tool()
    def delete_sheet(name: str) -> dict:
        """Eliminar una hoja del workbook activo (irreversible dentro del libro)."""
        return run(_delete_sheet, session, name)

    @mcp.tool()
    def analyze_workbook() -> dict:
        """Radiografia completa del workbook activo: hojas con dimensiones y
        cantidad de formulas, tablas, tablas dinamicas, nombres definidos,
        conexiones de datos y si tiene proyecto VBA. Es el punto de partida
        para desenmaranar un libro complejo."""
        return run(_analyze_workbook, session)


def _save_workbook(session, path: Optional[str]) -> dict:
    wb = get_active_workbook(session)
    if path is None:
        wb.Save()
        return {"saved": True, "path": str(wb.FullName)}
    path = os.path.abspath(path)
    ext = os.path.splitext(path)[1].lower()
    file_format = XL_FILE_FORMATS.get(ext)
    if file_format is None:
        raise ValueError(f"Extension no soportada: {ext} (usa .xlsx, .xlsm, .xlsb, .xls)")
    wb.SaveAs(path, FileFormat=file_format)
    logger.info("SaveAs %s (formato %s)", path, file_format)
    return {"saved": True, "path": str(wb.FullName), "format": ext}


def _close_workbook(session, save_changes: bool) -> dict:
    wb = get_active_workbook(session)
    name, full_name = str(wb.Name), str(wb.FullName)
    wb.Close(SaveChanges=save_changes)
    session._workbooks.pop(full_name, None)
    return {"closed": name, "saved": save_changes}


def _create_sheet(session, name: str, after: Optional[str]) -> dict:
    wb = get_active_workbook(session)
    if after:
        ws = wb.Sheets.Add(After=wb.Sheets(after))
    else:
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = name
    return {"name": ws.Name, "index": int(ws.Index)}


def _delete_sheet(session, name: str) -> dict:
    wb = get_active_workbook(session)
    wb.Sheets(name).Delete()
    return {"deleted": name}


def _analyze_workbook(session) -> dict:
    wb = get_active_workbook(session)
    app = session.get_application()

    sheets = []
    total_formulas = 0
    for ws in wb.Worksheets:
        info = {
            "name": str(ws.Name),
            "index": int(ws.Index),
            "visible": int(ws.Visible) == -1,  # xlSheetVisible
        }
        info["formula_cells"] = 0
        try:
            ur = ws.UsedRange
            info["used_range"] = str(ur.Address).replace("$", "")
            info["rows"] = int(ur.Rows.Count)
            info["columns"] = int(ur.Columns.Count)
            try:
                info["formula_cells"] = int(ur.SpecialCells(XL_CELL_TYPE_FORMULAS).Count)
            except Exception:
                pass  # SpecialCells lanza error si la hoja no tiene formulas
        except Exception:
            info["used_range"] = None
        total_formulas += info["formula_cells"]
        try:
            info["tables"] = [str(t.Name) for t in ws.ListObjects]
        except Exception:
            info["tables"] = []
        try:
            info["pivot_tables"] = [str(pt.Name) for pt in ws.PivotTables()]
        except Exception:
            info["pivot_tables"] = []
        sheets.append(info)

    result = {
        "name": str(wb.Name),
        "path": str(wb.FullName),
        "sheet_count": len(sheets),
        "sheets": sheets,
        "total_formula_cells": total_formulas,
    }

    try:
        result["has_vba_project"] = bool(wb.HasVBProject)
    except Exception:
        result["has_vba_project"] = None

    try:
        names = []
        for n in wb.Names:
            try:
                names.append({"name": str(n.Name), "refers_to": str(n.RefersTo)})
            except Exception:
                continue
        result["defined_names"] = names[:100]
        result["defined_name_count"] = len(names)
    except Exception:
        result["defined_names"] = []

    try:
        conns = []
        for c in wb.Connections:
            conn = {"name": str(c.Name)}
            try:
                conn["description"] = str(c.Description)
            except Exception:
                pass
            conns.append(conn)
        result["connections"] = conns
    except Exception:
        result["connections"] = []

    try:
        result["power_queries"] = [str(q.Name) for q in wb.Queries]
    except Exception:
        result["power_queries"] = None  # coleccion Queries no existe en Excel 2013

    return result
