"""Tools de Power Pivot / Data Model: DAX, tablas del modelo, medidas (Fase 5).

La via de acceso es Workbook.Model (introducido en Excel 2013) y su conexion
ADO in-process: Model.DataModelConnection.ModelConnection.ADOConnection.
Las consultas DAX (EVALUATE ...) y DMV ($SYSTEM.MDSCHEMA_*) van por ADODB.

Restricciones (plan maestro C2/C5):
- Power Pivot solo existe en SKU Professional Plus / Stand-alone / 365 ProPlus.
- En x86 el modelo esta limitado a 4 GB: validate_environment ya lo advierte.
"""

import logging

from ..utils.excel_utils import LONG_OP_TIMEOUT, get_active_workbook, to_jsonable

logger = logging.getLogger(__name__)

XL_CMD_EXCEL = 7  # xlCmdExcel: origen "WORKSHEET;<ruta>"


def register(mcp, session, run):
    @mcp.tool()
    def list_data_model_tables() -> list:
        """Listar las tablas del Data Model (Power Pivot) del workbook activo."""
        return run(_list_data_model_tables, session)

    @mcp.tool()
    def evaluate_dax_query(dax: str, max_rows: int = 1000) -> dict:
        """Ejecutar una consulta DAX contra el Data Model (ej: "EVALUATE 'Tabla'"
        o "EVALUATE SUMMARIZECOLUMNS(...)"). Tambien acepta DMV
        ($SYSTEM.MDSCHEMA_MEASURES, etc.). Devuelve columnas y filas."""
        return run(_evaluate_dax_query, session, dax, max_rows)

    @mcp.tool()
    def get_data_model_measures() -> dict:
        """Medidas del Data Model con su expresion DAX. Devuelve
        {measures, diagnostic}: diagnostic distingue 'sin medidas' (modelo
        presente, medidas implicitas) de 'sin modelo'."""
        return run(_get_data_model_measures, session)

    @mcp.tool()
    def add_table_to_data_model(sheet: str, range_addr: str, table_name: str) -> dict:
        """Agregar un rango como tabla (ListObject) y cargarla al Data Model.
        Pieza base para ELT: cargar datos crudos al modelo y agregar con DAX."""
        return run(_add_table_to_data_model, session, sheet, range_addr, table_name)

    @mcp.tool()
    def refresh_data_model() -> dict:
        """Refrescar el Data Model completo (todas sus conexiones)."""
        return run(_refresh_data_model, session, timeout=LONG_OP_TIMEOUT)


def _get_model(wb):
    try:
        model = wb.Model
        _ = model.Name  # fuerza el binding
        return model
    except Exception as e:
        raise RuntimeError(
            "Este workbook no tiene Data Model accesible. En Excel 2013 requiere el "
            "add-in Power Pivot activo (solo SKU Professional Plus / Stand-alone / 365 ProPlus)."
        ) from e


def _get_ado_connection(model):
    try:
        return model.DataModelConnection.ModelConnection.ADOConnection
    except Exception as e:
        raise RuntimeError(
            "No se pudo abrir la conexion ADO al modelo. El modelo puede estar vacio: "
            "agrega una tabla primero (add_table_to_data_model)."
        ) from e


def _list_data_model_tables(session) -> list:
    wb = get_active_workbook(session)
    model = _get_model(wb)
    tables = []
    for t in model.ModelTables:
        info = {"name": str(t.Name)}
        try:
            info["record_count"] = int(t.RecordCount)
        except Exception:
            info["record_count"] = None
        try:
            info["source"] = str(t.SourceName)
        except Exception:
            pass
        tables.append(info)
    return tables


_MSOLAP_HELP = (
    "El proveedor OLE DB MSOLAP no esta registrado para este proceso. Es un quirk "
    "conocido de Office Click-to-Run (dev host 2016+): el motor DAX del Data Model "
    "no es visible por ADO hasta que se abre el modelo en la UI de Excel al menos una "
    "vez. En el TARGET real (Excel 2013 Professional Plus, instalacion MSI) MSOLAP "
    "esta registrado a nivel maquina y esta consulta funciona directamente. "
    "Alternativa portable: usa create_pivot_table (agrega contra el mismo Data Model "
    "sin depender de MSOLAP externo)."
)


def _evaluate_dax_query(session, dax: str, max_rows: int) -> dict:
    """DAX via la conexion ADO in-process del Data Model.

    NO se usa puente VBA (Application.Run desde el thread STA hace deadlock con el
    message pump de Excel y deja el proceso colgado). Si ADO no puede por MSOLAP
    no registrado, se falla rapido con un error accionable en vez de colgar.
    """
    wb = get_active_workbook(session)
    model = _get_model(wb)
    conn = _get_ado_connection(model)
    try:
        rs = conn.Execute(dax)
    except Exception as e:
        msg = str(e)
        if "registrada" in msg or "registered" in msg or "MSOLAP" in msg or "Provider" in msg:
            raise RuntimeError(f"No se pudo ejecutar DAX. {_MSOLAP_HELP}\n\nDetalle: {msg}") from e
        raise RuntimeError(f"Error ejecutando DAX: {msg}") from e

    if isinstance(rs, tuple):  # dynamic dispatch puede devolver (recordset, affected)
        rs = rs[0]
    try:
        columns = [str(rs.Fields(i).Name) for i in range(rs.Fields.Count)]
        rows = []
        truncated = False
        if not rs.EOF:
            data = rs.GetRows(max_rows)  # tuple de COLUMNAS
            n_rows = len(data[0]) if data else 0
            rows = [
                [to_jsonable(data[c][r]) for c in range(len(columns))]
                for r in range(n_rows)
            ]
            truncated = not rs.EOF
    finally:
        try:
            rs.Close()
        except Exception:
            pass
    logger.info("DAX via ADO: %s filas x %s columnas", len(rows), len(columns))
    return {"columns": columns, "rows": rows, "row_count": len(rows), "truncated": truncated}


def _get_data_model_measures(session) -> dict:
    wb = get_active_workbook(session)
    diagnostic = {"model_present": False, "model_tables": 0}
    try:
        model = _get_model(wb)
        diagnostic["model_present"] = True
        try:
            diagnostic["model_tables"] = int(model.ModelTables.Count)
        except Exception:
            pass
    except Exception as e:
        diagnostic["note"] = f"Sin Data Model accesible: {e}"
        return {"measures": [], "diagnostic": diagnostic}

    measures = []
    try:
        # Via COM (Excel 2016+): ModelMeasures expone nombre + formula directamente
        for m in model.ModelMeasures:
            measures.append(
                {
                    "name": str(m.Name),
                    "expression": str(m.Formula),
                    "table": str(m.AssociatedTable.Name),
                }
            )
    except Exception:
        logger.debug("ModelMeasures no disponible (host 2013); fallback DMV")
        try:
            # Fallback DMV (funciona tambien en 2013): medidas visibles del cubo
            result = _evaluate_dax_query(
                session,
                "SELECT [MEASURE_NAME], [EXPRESSION], [MEASURE_CAPTION] "
                "FROM $SYSTEM.MDSCHEMA_MEASURES WHERE [MEASURE_IS_VISIBLE]",
                1000,
            )
            measures = [
                {"name": r[0], "expression": r[1], "caption": r[2]}
                for r in result["rows"]
            ]
        except Exception as e:
            diagnostic["note"] = f"ModelMeasures y DMV fallaron: {e}"

    if not measures and diagnostic["model_tables"] > 0 and "note" not in diagnostic:
        diagnostic["note"] = (
            f"Modelo presente con {diagnostic['model_tables']} tablas y 0 medidas "
            "explicitas: probablemente pivots con agregaciones implicitas."
        )
    return {"measures": measures, "diagnostic": diagnostic}


def _add_table_to_data_model(session, sheet: str, range_addr: str, table_name: str) -> dict:
    wb = get_active_workbook(session)
    ws = wb.Sheets(sheet)

    if not str(wb.Path):
        raise RuntimeError(
            "El workbook debe estar guardado en disco antes de cargar tablas al modelo "
            "(la conexion WORKSHEET usa la ruta del archivo). Usa save_workbook primero."
        )

    # 1. Crear el ListObject si el rango aun no es tabla
    existing = None
    for lo in ws.ListObjects:
        if str(lo.Name).lower() == table_name.lower():
            existing = lo
            break
    if existing is None:
        lo = ws.ListObjects.Add(1, ws.Range(range_addr), None, 1)  # xlSrcRange, xlYes
        lo.Name = table_name
        created_table = True
    else:
        lo = existing
        created_table = False

    # 2. Cargar al Data Model via conexion WORKSHEET con CreateModelConnection=True
    conn_name = f"ModelConnection_{table_name}"
    wb.Connections.Add2(
        conn_name,
        f"Tabla {table_name} en el Data Model",
        f"WORKSHEET;{wb.FullName}",
        f"{ws.Name}!{table_name}",
        XL_CMD_EXCEL,
        True,   # CreateModelConnection
        False,  # ImportRelationships
    )
    logger.info("Tabla %s cargada al Data Model (conexion %s)", table_name, conn_name)

    model_tables = _list_data_model_tables(session)
    return {
        "table": table_name,
        "created_list_object": created_table,
        "connection": conn_name,
        "model_tables": model_tables,
    }


def _refresh_data_model(session) -> dict:
    wb = get_active_workbook(session)
    model = _get_model(wb)
    model.Refresh()
    return {"refreshed": True, "tables": _list_data_model_tables(session)}
