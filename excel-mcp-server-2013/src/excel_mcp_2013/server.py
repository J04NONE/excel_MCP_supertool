"""FastMCP server: automatizacion de Excel via COM.

- Excel arranca LAZY: en el primer tool call que lo necesite (ping no lo necesita).
- Toda operacion COM pasa por ExcelWriteGuard (thread STA dedicado).
- El shutdown (lifespan) cierra Excel y mata el proceso si sobrevive.
- Tools organizadas en modulos: workbook, cells, vba, power_query, pivots, discovery.
"""

import logging
from contextlib import asynccontextmanager
from datetime import datetime
from typing import Optional

import anyio
import pywintypes
from fastmcp import FastMCP

from .com_guard import ExcelWriteGuard, STAWedgedError
from .orphan_guard import OrphanProcessGuard
from .session import SessionManager
from .utils.com_errors import ExcelCrashedError, is_dead_server, translate_com_error

logger = logging.getLogger(__name__)

session = SessionManager()
guard = ExcelWriteGuard()

# Timeout dedicado (corto) para abrir libros: un open sano tarda pocos segundos;
# uno que se pasa de esto casi siempre es un dialogo modal. Fallar en ~45s en vez
# del default de 120s. La deteccion de dialogo (has_modal_dialog) hace fail-fast
# a las llamadas SIGUIENTES sin esperar nada.
OPEN_WORKBOOK_TIMEOUT = 45


def _shutdown_sync() -> None:
    """Cierra Excel en el STA thread y apaga el guard.

    Si el STA quedo colgado (ej: dialogo modal de Excel), el Quit cooperativo
    no puede correr: se mata el proceso por PID directamente (psutil no
    necesita COM ni thread STA).
    """
    try:
        if session.get_application() is not None or session.get_pid() is not None:
            guard.execute(session.close)
    except Exception:
        logger.exception("Cierre cooperativo fallo; matando Excel por PID")
        pid = session.get_pid()
        if pid:
            OrphanProcessGuard.kill_process(pid)
    finally:
        guard.shutdown()


@asynccontextmanager
async def _lifespan(app):
    logger.info("Servidor iniciado. Excel arrancara en el primer tool call que lo use.")
    try:
        yield
    finally:
        logger.info("Shutting down server...")
        await anyio.to_thread.run_sync(_shutdown_sync)
        logger.info("Shutdown completo")


mcp = FastMCP(
    "excel-mcp-server-2013",
    instructions=(
        "MCP server for Excel automation via COM. "
        "Dev: Office LTSC 2021, Target: Excel 2013 (Power Query 2.62 legacy). "
        "Flujo tipico: open_workbook -> analyze_workbook -> read_formulas/get_vba_code/"
        "list_pivot_tables para desarmar; write_formulas/create_pivot_table/inject_vba_code "
        "para armar; save_workbook para persistir. El codigo M generado debe validarse "
        "con validate_m_code (Excel 2013 usa el motor M v2.62 legacy)."
    ),
    version="1.5.0",
    lifespan=_lifespan,
)


def run_with_excel(func, *args, timeout: Optional[float] = None, **kwargs):
    """Ejecuta func en el STA thread, arrancando Excel antes si hace falta.

    El start() DEBE ocurrir en el mismo thread STA que ejecutara el resto de
    las operaciones COM (afinidad de thread de COM).

    `timeout` (keyword-only, segundos) sobreescribe el default de 120s del
    guard para esta llamada; nunca llega a `func`.

    Recovery (Fase 7):
    - Pre-flight: si EXCEL.EXE murio entre llamadas (is_alive por PID), se
      limpia y reinicia la sesion y la operacion PROCEDE (nada se habia
      ejecutado aun, es seguro para lecturas y escrituras).
    - Mid-flight: si Excel muere DURANTE la operacion (com_error de servidor
      muerto), se limpia la sesion y se lanza ExcelCrashedError. NUNCA se
      reintenta: no se pueden reabrir los workbooks fielmente (flags no
      guardados) ni garantizar idempotencia de escrituras.
    - TimeoutError NO dispara recovery: Excel vivo pero ocupado no es crash.
    """

    def runner():
        if session.get_application() is None:
            logger.info("Starting Excel session (lazy init)...")
            session.start()
        elif not session.is_alive():
            logger.warning(
                "EXCEL.EXE murio entre llamadas (PID %s); reiniciando sesion",
                session.get_pid(),
            )
            session.reset_after_crash()
            session.start()
        try:
            result = func(*args, **kwargs)
        except pywintypes.com_error as e:
            if is_dead_server(e):
                session.reset_after_crash()
                raise ExcelCrashedError(
                    "Excel murio durante la operacion. La sesion fue reiniciada; "
                    "los workbooks abiertos (y cambios no guardados) se perdieron. "
                    "Reabre con open_workbook."
                ) from e
            raise
        # Post-flight: si una tool trago internamente el com_error de servidor
        # muerto (handlers amplios, ej. _collect_documentation), el resultado
        # es parcial/basura. Detectarlo aqui (psutil, <1ms) y fallar honesto.
        if not session.is_alive():
            session.reset_after_crash()
            raise ExcelCrashedError(
                "Excel murio durante la operacion (un handler interno enmascaro "
                "el error y el resultado seria parcial). Sesion reiniciada; "
                "reabre con open_workbook."
            )
        return result

    # Fail-fast: si el STA esta ocupado en una tarea Y Excel muestra un dialogo
    # modal, el STA esta bloqueado (el dialogo congela el hilo de Excel). No tiene
    # sentido encolar y esperar el timeout completo: avisar ya. La deteccion es
    # una senal POSITIVA (ventana de dialogo), asi que una tarea larga legitima
    # —sin dialogo— sigue encolando normal.
    info = guard.inflight_info()
    if info is not None and session.has_modal_dialog():
        raise STAWedgedError(
            f"STA bloqueado: la operacion en curso ('{info['desc']}') lleva "
            f"{info['elapsed']:.0f}s y Excel muestra un dialogo modal. "
            "Corre close_excel para recuperar la sesion."
        )

    try:
        return guard.execute(runner, timeout=timeout)
    except pywintypes.com_error as e:
        # Errores COM que no son crash: traducirlos a algo accionable
        raise RuntimeError(translate_com_error(e)) from e


# =============================================================================
# Tools de sesion (las demas viven en tools/*.py)
# =============================================================================


@mcp.tool()
def ping() -> str:
    """Verificar que el servidor responde (no requiere Excel)."""
    return f'{{"status":"ok","timestamp":"{datetime.now().isoformat()}"}}'


@mcp.tool()
def get_session_info() -> dict:
    """Obtener informacion de la sesion de Excel (la inicia si no existe)."""
    return run_with_excel(_get_session_info_impl)


def _get_session_info_impl() -> dict:
    app = session.get_application()
    return {
        "initialized": app is not None,
        "pid": session.get_pid(),
        "version": str(app.Version) if app else None,
        "visible": bool(app.Visible) if app else None,
        "workbook_count": int(app.Workbooks.Count) if app else 0,
    }


@mcp.tool()
def open_workbook(
    path: str,
    read_only: bool = False,
    password: Optional[str] = None,
    enable_macros: bool = False,
    sanitize_names: str = "auto",
) -> dict:
    """Abrir un workbook (.xlsx/.xlsm/.xlsb). Por seguridad las macros NO se
    habilitan al abrir salvo enable_macros=True (necesario para execute_vba_macro).

    sanitize_names (auto|always|never): en 'auto' abre una copia temporal sin
    los nombres definidos ROTOS cuando detecta riesgo (OOXML con vinculos
    externos + nombres #REF!), evitando el cuelgue del dialogo 'Conflicto de
    nombres'. El archivo original nunca se modifica; la respuesta incluye
    'sanitized' cuando ocurrio."""
    return run_with_excel(
        _open_workbook_impl, path, read_only, password, enable_macros, sanitize_names,
        timeout=OPEN_WORKBOOK_TIMEOUT,
    )


def _open_workbook_impl(
    path: str,
    read_only: bool,
    password: Optional[str],
    enable_macros: bool,
    sanitize_names: str,
) -> dict:
    wb = session.open_workbook(path, read_only, password, enable_macros, sanitize_names)
    sheets = [{"name": ws.Name, "type": "worksheet"} for ws in wb.Worksheets]
    info = session.last_open_info()
    result = {
        # Ruta ORIGINAL solicitada (si hubo saneo, wb.FullName seria la copia temporal).
        "path": info.get("path", str(wb.FullName)),
        "sheets": sheets,
        "sheet_count": len(sheets),
        "read_only": bool(wb.ReadOnly),
        "macros_enabled": enable_macros,
    }
    if info.get("sanitized"):
        result["sanitized"] = {
            "definednames_removed": info.get("definednames_removed"),
            "external_links": info.get("external_links"),
            "note": "Se abrio una COPIA saneada para evitar el cuelgue por "
                    "'Conflicto de nombres'. El archivo original NO fue modificado. "
                    "Los valores cacheados se conservan; para refrescar formulas "
                    "externas, abrir los libros de origen.",
        }
    elif info.get("sanitize_error"):
        result["sanitize_warning"] = (
            f"Archivo de riesgo pero el saneo fallo ({info['sanitize_error']}); "
            "se abrio el original (puede colgar por 'Conflicto de nombres')."
        )
    return result


@mcp.tool()
def list_sheets() -> list:
    """Listar las hojas del workbook activo."""
    return run_with_excel(_list_sheets_impl)


def _list_sheets_impl() -> list:
    wb = session.get_application().ActiveWorkbook
    if not wb:
        return []
    return [{"name": ws.Name, "type": "worksheet"} for ws in wb.Worksheets]


@mcp.tool()
def close_excel() -> dict:
    """Cerrar la sesion de Excel (workbooks sin guardar se descartan).
    Es la via de escape si una operacion quedo colgada: si el STA no responde
    en 30s, mata el proceso por PID (psutil, sin COM)."""
    if session.get_application() is None and session.get_pid() is None:
        return {"closed": False, "reason": "Excel no estaba iniciado"}
    # Si hay un dialogo modal, el STA esta bloqueado y el Quit cooperativo
    # colgaria: matar por PID de una (sin esperar 30s). psutil no necesita COM.
    if session.has_modal_dialog():
        pid = session.get_pid()
        if pid:
            OrphanProcessGuard.kill_process(pid)
        logger.warning("close_excel: dialogo modal detectado; Excel PID %s matado directo", pid)
        return {"closed": True, "method": "kill", "pid": pid, "reason": "dialogo modal"}
    try:
        guard.execute(session.close, timeout=30)
        return {"closed": True, "method": "cooperative"}
    except TimeoutError:
        # STA bloqueado por una tarea colgada (sin dialogo visible): el Quit
        # cooperativo no puede correr. Matar por PID es seguro desde este thread
        # (solo psutil). No tocamos _app aqui (afinidad COM): la proxima llamada
        # lo limpia via is_alive() -> reset_after_crash() dentro del STA.
        pid = session.get_pid()
        if pid:
            OrphanProcessGuard.kill_process(pid)
        logger.warning("close_excel: STA bloqueado; Excel PID %s matado directamente", pid)
        return {"closed": True, "method": "kill", "pid": pid}


# =============================================================================
# Registro de modulos de tools
# =============================================================================

from .tools import (  # noqa: E402
    bulk,
    cells,
    discovery,
    documentation,
    elt,
    introspect,
    pivots,
    power_pivot,
    power_query,
    semantics,
    shapes,
    vba,
    workbook,
)

workbook.register(mcp, session, run_with_excel)
bulk.register(mcp, session, run_with_excel)
cells.register(mcp, session, run_with_excel)
vba.register(mcp, session, run_with_excel)
power_query.register(mcp, session, run_with_excel)
power_pivot.register(mcp, session, run_with_excel)
pivots.register(mcp, session, run_with_excel)
discovery.register(mcp, session, run_with_excel)
introspect.register(mcp, session, run_with_excel)
semantics.register(mcp, session, run_with_excel)
documentation.register(mcp, session, run_with_excel)
elt.register(mcp, session, run_with_excel)
shapes.register(mcp, session, run_with_excel)
