"""Tools de Power Query: listar, extraer M, refrescar y validar compatibilidad 2013.

Notas de compatibilidad (plan maestro C1/C3/C4):
- Excel 2013: PQ es un ADD-IN (Microsoft.Data.Mashup.ExcelAddin); la coleccion
  Workbook.Queries NO existe (se introdujo en 2016). El M vive en las conexiones.
- Refresh SIEMPRE sincrono: BackgroundQuery=False antes de .Refresh().
- El M generado debe validarse contra utils.m_constraints (motor v2.62 legacy).
"""

import logging
from typing import Optional

from ..utils.excel_utils import LONG_OP_TIMEOUT, get_active_workbook
from ..utils.m_constraints import (
    BLOCKED_FUNCTIONS,
    get_alternative,
    validate_m_expression,
)

logger = logging.getLogger(__name__)


def register(mcp, session, run):
    @mcp.tool()
    def list_power_queries() -> dict:
        """Listar queries de Power Query y conexiones de datos del workbook activo."""
        return run(_list_power_queries, session)

    @mcp.tool()
    def get_power_query_m(query_name: str) -> dict:
        """Extraer el codigo M de una query de Power Query.
        (Requiere Excel 2016+ en el host; en 2013 el M no es accesible via COM)."""
        return run(_get_power_query_m, session, query_name)

    @mcp.tool()
    def refresh_power_query(connection_name: Optional[str] = None) -> dict:
        """Refrescar conexiones de datos de forma SINCRONA (BackgroundQuery=False).
        Sin connection_name refresca todas."""
        return run(_refresh_power_query, session, connection_name, timeout=LONG_OP_TIMEOUT)

    @mcp.tool()
    def validate_m_code(m_code: str) -> dict:
        """Validar que un codigo M sea compatible con Power Query de Excel 2013
        (v2.62 legacy). Devuelve funciones bloqueadas y sus alternativas."""
        result = validate_m_expression(m_code)
        result["alternatives"] = {
            fn: get_alternative(fn) for fn in result["blocked_found"]
        }
        return result

    @mcp.tool()
    def m_function_compatible(function_name: str) -> dict:
        """Verificar si UNA funcion M es compatible con Excel 2013."""
        blocked = BLOCKED_FUNCTIONS.get(function_name)
        if blocked:
            return {
                "function": function_name,
                "compatible": False,
                "introduced": blocked.get("introduced"),
                "alternative": blocked.get("alternative"),
                "example_legacy": blocked.get("example_legacy"),
            }
        return {
            "function": function_name,
            "compatible": True,
            "note": "No esta en la lista de bloqueadas; verificar con 'type shared' en el 2013 real si es critica.",
        }


def _list_power_queries(session) -> dict:
    wb = get_active_workbook(session)
    result = {"queries": [], "connections": [], "queries_api_available": True}

    try:
        for q in wb.Queries:
            result["queries"].append({"name": str(q.Name), "description": str(q.Description or "")})
    except Exception:
        # Excel 2013: la coleccion Queries no existe
        result["queries_api_available"] = False

    try:
        for c in wb.Connections:
            conn = {"name": str(c.Name), "type": int(c.Type)}
            try:
                conn["description"] = str(c.Description or "")
            except Exception:
                pass
            try:
                conn["refresh_date"] = str(c.OLEDBConnection.RefreshDate)
            except Exception:
                pass
            result["connections"].append(conn)
    except Exception as e:
        logger.warning("No se pudieron listar conexiones: %s", e)

    return result


def _get_power_query_m(session, query_name: str) -> dict:
    wb = get_active_workbook(session)
    try:
        queries = wb.Queries
    except Exception as e:
        raise RuntimeError(
            "La coleccion Workbook.Queries no esta disponible (Excel 2013 no la expone via COM). "
            "En 2013 el M solo puede verse desde el editor de Power Query."
        ) from e
    q = queries(query_name)
    return {"name": str(q.Name), "m_code": str(q.Formula)}


def _refresh_power_query(session, connection_name: Optional[str]) -> dict:
    wb = get_active_workbook(session)
    refreshed, failed = [], []
    for c in wb.Connections:
        name = str(c.Name)
        if connection_name and name != connection_name:
            continue
        try:
            # C4: forzar refresh SINCRONO o el tool devuelve antes de terminar
            try:
                c.OLEDBConnection.BackgroundQuery = False
            except Exception:
                pass  # conexiones no-OLEDB no tienen BackgroundQuery
            c.Refresh()
            refreshed.append(name)
            logger.info("Conexion refrescada: %s", name)
        except Exception as e:
            failed.append({"name": name, "error": str(e)})
            logger.warning("Fallo refrescando %s: %s", name, e)
    if connection_name and not refreshed and not failed:
        raise ValueError(f"Conexion '{connection_name}' no encontrada")
    return {"refreshed": refreshed, "failed": failed}
