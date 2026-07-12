"""Kit ELT: medidas DAX persistentes, queries Power Query, dashboards CUBE y
macro de refresh. Cierra el ciclo Extract (PQ) -> Load (Data Model) ->
Transform (DAX) -> Salida (pivots/CUBEVALUE) usando SOLO Excel.

Compatibilidad 2013: las medidas y queries se CREAN en el host 2021 (las APIs
ModelMeasures/Queries son 2016+) pero VIAJAN dentro del archivo y funcionan en
2013. Las formulas CUBE existen desde Excel 2007. El piloto ida-vuelta valida esto.
"""

import logging
from typing import Optional

from ..utils.excel_utils import get_active_workbook
from ..utils.m_constraints import get_alternative, validate_m_expression
from .power_pivot import _get_model, _list_data_model_tables
from .vba import _inject_vba_code

logger = logging.getLogger(__name__)

MASHUP_PROVIDER = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="

REFRESH_MACRO_TEMPLATE = '''
Public Sub {name}()
    ' Refresca todo en orden: conexiones/PQ -> Data Model -> tablas dinamicas.
    ' Generada por excel-mcp-server-2013 (setup_refresh_macro).
    Dim cn As Object, ws As Worksheet, pt As PivotTable
    Application.ScreenUpdating = False
    On Error Resume Next

    ' 1. Conexiones (Power Query) en modo SINCRONO
    For Each cn In ThisWorkbook.Connections
        cn.OLEDBConnection.BackgroundQuery = False
        cn.Refresh
    Next cn

    ' 2. Data Model (si existe)
    ThisWorkbook.Model.Refresh

    ' 3. Tablas dinamicas
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    On Error GoTo 0
    Application.ScreenUpdating = True
    MsgBox "Actualizacion completa: " & Format(Now, "yyyy-mm-dd hh:nn"), vbInformation
End Sub
'''


def register(mcp, session, run):
    @mcp.tool()
    def add_data_model_measure(
        table_name: str, measure_name: str, dax: str, number_format: str = "general"
    ) -> dict:
        """Crear una medida DAX persistente en el Data Model (ej: dax=
        "SUM(Ventas[Cajas])"). number_format: general|decimal|whole|percentage|currency.
        La medida viaja en el archivo y funciona en Excel 2013."""
        return run(_add_data_model_measure, session, table_name, measure_name, dax, number_format)

    @mcp.tool()
    def add_power_query(
        query_name: str,
        m_code: str,
        load_to: str = "connection_only",
        target_sheet: Optional[str] = None,
    ) -> dict:
        """Crear una query de Power Query (valida el M contra Excel 2013 primero).
        load_to: 'connection_only' | 'sheet' (vuelca a target_sheet) | 'data_model'."""
        return run(_add_power_query, session, query_name, m_code, load_to, target_sheet)

    @mcp.tool()
    def write_cube_formulas(
        sheet: str,
        start_cell: str,
        title: str,
        rows: list,
        values: list,
        model_name: str = "ThisWorkbookDataModel",
    ) -> dict:
        """Escribir un mini-dashboard con formulas CUBE (nativas desde Excel 2007,
        funcionan en 2013). rows=[{"caption":"Norte","member":"[Ventas].[Filial].&[Norte]"}],
        values=[{"caption":"Total Cajas","measure":"[Measures].[Total Cajas]"}]."""
        return run(_write_cube_formulas, session, sheet, start_cell, title, rows, values, model_name)

    @mcp.tool()
    def setup_refresh_macro(macro_name: str = "ActualizarTodo") -> dict:
        """Inyectar la macro de actualizacion total (PQ -> modelo -> pivots, en orden
        y sincrono). Solo la INYECTA: ejecutala tu con un boton/F5, o via
        execute_vba_macro con el libro abierto con enable_macros=True."""
        return run(_setup_refresh_macro, session, macro_name)


def _add_data_model_measure(
    session, table_name: str, measure_name: str, dax: str, number_format: str
) -> dict:
    wb = get_active_workbook(session)
    model = _get_model(wb)
    try:
        formats = {
            "general": model.ModelFormatGeneral,
            "decimal": model.ModelFormatDecimalNumber,
            "whole": model.ModelFormatWholeNumber,
            "percentage": model.ModelFormatPercentageNumber,
            "currency": model.ModelFormatCurrency,
        }
        fmt = formats.get(number_format.lower())
        if fmt is None:
            raise ValueError(f"number_format invalido: {number_format} (usa {sorted(formats)})")
        table = model.ModelTables(table_name)
        measure = model.ModelMeasures.Add(measure_name, table, dax, fmt)
    except ValueError:
        raise
    except Exception as e:
        raise RuntimeError(
            "No se pudo crear la medida. La API ModelMeasures requiere host Excel 2016+ "
            "(en 2013 las medidas se crean desde la ventana de Power Pivot). "
            f"Detalle: {e}"
        ) from e
    logger.info("Medida '%s' creada en tabla '%s': %s", measure_name, table_name, dax)
    return {
        "measure": str(measure.Name),
        "table": table_name,
        "dax": dax,
        "cube_reference": f"[Measures].[{measure_name}]",
    }


def _add_power_query(
    session, query_name: str, m_code: str, load_to: str, target_sheet: Optional[str]
) -> dict:
    if load_to not in ("connection_only", "sheet", "data_model"):
        raise ValueError("load_to debe ser: connection_only | sheet | data_model")

    # 1. Guardian M 2013: nunca crear una query que reviente en el trabajo
    check = validate_m_expression(m_code)
    if not check["valid"]:
        alts = {fn: get_alternative(fn) for fn in check["blocked_found"]}
        raise ValueError(
            f"El codigo M usa funciones NO disponibles en Excel 2013: "
            f"{check['blocked_found']}. Alternativas: {alts}"
        )

    wb = get_active_workbook(session)

    # 2. Crear la query (API 2016+; viaja en el archivo hacia 2013)
    try:
        wb.Queries.Add(query_name, m_code)
    except Exception as e:
        raise RuntimeError(
            "Workbook.Queries.Add no disponible (host 2013 no expone la API; crear "
            f"la query en el host 2021). Detalle: {e}"
        ) from e

    result = {"query": query_name, "load_to": load_to, "m_valid_2013": True}

    # 3. Carga
    if load_to == "sheet":
        if not target_sheet:
            raise ValueError("load_to='sheet' requiere target_sheet")
        ws = None
        for s in wb.Worksheets:
            if str(s.Name).lower() == target_sheet.lower():
                ws = s
                break
        if ws is None:
            ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
            ws.Name = target_sheet
        # QueryTables.Add: la API clasica funciona cross-process (ListObjects.Add
        # con fuentes externas NO — falla igual que xlSrcModel, probado)
        qt = ws.QueryTables.Add(
            Connection=f"{MASHUP_PROVIDER}{query_name}",
            Destination=ws.Range("A1"),
            Sql=f"SELECT * FROM [{query_name}]",
        )
        qt.RefreshStyle = 1  # xlInsertDeleteCells
        qt.AdjustColumnWidth = True
        qt.Refresh(BackgroundQuery=False)
        result["sheet"] = str(ws.Name)
        result["rows_loaded"] = int(ws.UsedRange.Rows.Count)
    elif load_to == "data_model":
        wb.Connections.Add2(
            f"Query - {query_name}",
            f"Conexion a la query '{query_name}' del libro",
            f"{MASHUP_PROVIDER}{query_name}",
            f'"{query_name}"',
            6,  # xlCmdTableCollection
            True,   # CreateModelConnection
            False,  # ImportRelationships
        )
        result["model_tables"] = _list_data_model_tables(session)

    logger.info("Power Query '%s' creada (load_to=%s)", query_name, load_to)
    return result


def _write_cube_formulas(
    session, sheet: str, start_cell: str, title: str, rows: list, values: list, model_name: str
) -> dict:
    wb = get_active_workbook(session)
    ws = None
    for s in wb.Worksheets:
        if str(s.Name).lower() == sheet.lower():
            ws = s
            break
    if ws is None:
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheet

    anchor = ws.Range(start_cell)
    r0, c0 = int(anchor.Row), int(anchor.Column)

    # Titulo
    ws.Cells(r0, c0).Value = title
    ws.Cells(r0, c0).Font.Bold = True

    # Encabezados de medidas (fila r0+1): CUBEMEMBER de cada medida
    for j, v in enumerate(values):
        cell = ws.Cells(r0 + 1, c0 + 1 + j)
        cell.Formula = f'=CUBEMEMBER("{model_name}","{v["measure"]}","{v["caption"]}")'
        cell.Font.Bold = True

    # Filas: CUBEMEMBER del miembro + CUBEVALUE cruzando member x medida
    for i, row in enumerate(rows):
        member_cell = ws.Cells(r0 + 2 + i, c0)
        member_cell.Formula = (
            f'=CUBEMEMBER("{model_name}","{row["member"]}","{row["caption"]}")'
        )
        for j, _v in enumerate(values):
            value_cell = ws.Cells(r0 + 2 + i, c0 + 1 + j)
            member_ref = member_cell.Address  # $A$5
            header_ref = ws.Cells(r0 + 1, c0 + 1 + j).Address
            value_cell.Formula = f'=CUBEVALUE("{model_name}",{member_ref},{header_ref})'

    # Las formulas CUBE se resuelven de forma ASINCRONA (#GETTING_DATA hasta que
    # el motor responde). Con Calculation=Manual (default de la sesion) las
    # consultas nunca arrancan: pasar a automatico, esperar, y restaurar.
    app = session.get_application()
    try:
        prev_calc = app.Calculation
        app.Calculation = -4105  # xlCalculationAutomatic
        app.CalculateFull()
        app.CalculateUntilAsyncQueriesDone()
        app.Calculation = prev_calc
    except Exception as e:
        logger.warning("Recalculo de formulas CUBE fallo: %s", e)

    ws.UsedRange.Columns.AutoFit()
    n_cells = len(rows) * len(values) + len(rows) + len(values)
    logger.info("Dashboard CUBE escrito en %s!%s (%s formulas)", sheet, start_cell, n_cells)
    return {
        "sheet": str(ws.Name),
        "start_cell": start_cell,
        "cube_formulas_written": n_cells,
        "note": "Las formulas CUBE se recalculan al refrescar el Data Model (compatibles Excel 2013).",
    }


def _setup_refresh_macro(session, macro_name: str) -> dict:
    code = REFRESH_MACRO_TEMPLATE.format(name=macro_name)
    result = _inject_vba_code(session, "MCP_Refresh", code, replace=True)
    return {
        **result,
        "macro": macro_name,
        "how_to_run": (
            f"En el trabajo: Alt+F8 -> {macro_name}, o asigna la macro a un boton. "
            f"Desde el MCP: execute_vba_macro('{macro_name}') con el libro abierto "
            "con enable_macros=True. Guarda como .xlsm."
        ),
    }
