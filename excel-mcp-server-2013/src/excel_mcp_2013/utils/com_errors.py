"""Traduccion de pywintypes.com_error a mensajes accionables.

Excel reporta casi todos sus errores como DISP_E_EXCEPTION (0x80020009) con el
codigo real anidado en excepinfo[5] (scode). Por eso extract_scodes devuelve
tanto el hresult exterior como el scode interior, y todas las tablas se
consultan contra ambos.

Modulo puro Python (sin dependencia de Excel vivo): testeable offline con
com_error construidos a mano.
"""

from __future__ import annotations

import logging

logger = logging.getLogger(__name__)


def _i32(value: int) -> int:
    """Normaliza un HRESULT a entero de 32 bits con signo (como los reporta pywin32)."""
    value &= 0xFFFFFFFF
    return value - 0x100000000 if value >= 0x80000000 else value


# --- HRESULTs de "proceso Excel muerto" -> disparan recovery ---
RPC_S_SERVER_UNAVAILABLE = _i32(0x800706BA)  # el proceso ya no existe
RPC_S_CALL_FAILED = _i32(0x800706BE)         # murio a mitad de la llamada
RPC_E_DISCONNECTED = _i32(0x80010108)        # objeto desconectado de su servidor
CO_E_SERVER_EXEC_FAILURE = _i32(0x80080005)  # no pudo arrancar/re-lanzar

DEAD_SERVER_HRESULTS = frozenset({
    RPC_S_SERVER_UNAVAILABLE,
    RPC_S_CALL_FAILED,
    RPC_E_DISCONNECTED,
    CO_E_SERVER_EXEC_FAILURE,
})

# --- HRESULTs de "Excel vivo pero ocupado" -> NUNCA recovery ---
RPC_E_CALL_REJECTED = _i32(0x80010001)        # dialogo modal / celda en edicion
RPC_E_SERVERCALL_RETRYLATER = _i32(0x8001010A)  # ocupado procesando

BUSY_HRESULTS = frozenset({
    RPC_E_CALL_REJECTED,
    RPC_E_SERVERCALL_RETRYLATER,
})

# --- Otros con mensaje dedicado ---
DISP_E_EXCEPTION = _i32(0x80020009)   # envoltorio generico de Excel
DISP_E_BADINDEX = _i32(0x8002000B)    # item inexistente en coleccion (Worksheets("X"))
VBA_E_APPLICATION = _i32(0x800A03EC)  # "1004": Excel rechazo la operacion
E_FAIL = _i32(0x80004005)


class ExcelCrashedError(RuntimeError):
    """El proceso de Excel murio durante una operacion; la sesion fue reiniciada."""


def extract_scodes(e) -> list:
    """Devuelve [hresult, scode] de un pywintypes.com_error (los que existan).

    com_error.args = (hresult, strerror, excepinfo, argerror);
    excepinfo = (wCode, source, description, helpFile, helpContext, scode) o None.
    """
    codes = []
    args = getattr(e, "args", ())
    if len(args) >= 1 and isinstance(args[0], int):
        codes.append(_i32(args[0]))
    if len(args) >= 3 and args[2] is not None:
        try:
            scode = args[2][5]
            if isinstance(scode, int):
                codes.append(_i32(scode))
        except (IndexError, TypeError):
            pass
    return codes


def extract_description(e) -> str:
    """Mejor descripcion disponible: excepinfo[2] (mensaje de Excel) o strerror."""
    args = getattr(e, "args", ())
    if len(args) >= 3 and args[2] is not None:
        try:
            desc = args[2][2]
            if desc:
                return str(desc).strip()
        except (IndexError, TypeError):
            pass
    if len(args) >= 2 and args[1]:
        return str(args[1]).strip()
    return repr(e)


def is_dead_server(e) -> bool:
    """True si el com_error indica que el proceso servidor (Excel) murio."""
    return any(code in DEAD_SERVER_HRESULTS for code in extract_scodes(e))


def _hex(code: int) -> str:
    return f"0x{code & 0xFFFFFFFF:08X}"


def translate_com_error(e) -> str:
    """Mensaje accionable en espanol para un pywintypes.com_error."""
    codes = extract_scodes(e)
    desc = extract_description(e)
    codes_hex = "/".join(_hex(c) for c in codes) or "sin codigo"

    if any(c in DEAD_SERVER_HRESULTS for c in codes):
        return (
            f"El proceso de Excel murio o se desconecto ({codes_hex}). "
            "La sesion fue reiniciada; reabre el archivo con open_workbook."
        )
    if any(c in BUSY_HRESULTS for c in codes):
        return (
            f"Excel esta ocupado ({codes_hex}): dialogo abierto o celda en edicion. "
            "Reintenta en unos segundos; si persiste, usa close_excel para reiniciar."
        )
    if VBA_E_APPLICATION in codes:
        return (
            f"Excel rechazo la operacion ({_hex(VBA_E_APPLICATION)}): '{desc}'. "
            "Causas tipicas: nombre de hoja/rango invalido, hoja protegida, "
            "u operacion no valida en el estado actual del libro."
        )
    if DISP_E_BADINDEX in codes:
        return (
            f"Elemento no encontrado ({_hex(DISP_E_BADINDEX)}): '{desc}'. "
            "El nombre/indice pedido no existe en la coleccion (hoja, tabla "
            "dinamica, modulo, etc.). Verifica el nombre exacto con la tool "
            "de listado correspondiente."
        )
    if E_FAIL in codes:
        return f"Error COM generico ({_hex(E_FAIL)}): '{desc}'."
    return f"Error COM ({codes_hex}): '{desc}'."
