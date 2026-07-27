# -*- coding: utf-8 -*-
"""Tools de introspeccion: buscar en celdas, inventariar/limpiar nombres
definidos y mapear de donde sale la data (conexiones + vinculos).

Regla dura de este modulo: NINGUNA operacion sondea la red. Las fuentes de
intranet (SSAS/OLAP, UNC) se MAPEAN y se explican, nunca se contactan — un
sondeo bloquearia el unico thread STA (mismo cuelgue que ya se blindo en
open_workbook).
"""

import logging
import os
import re
from typing import Optional

from ..utils.excel_utils import get_active_workbook, get_sheet, matrix_to_jsonable, to_jsonable
from ..utils.workbook_sanitize import _BROKEN_RE

logger = logging.getLogger(__name__)

MAX_SEARCH_CELLS = 2_000_000  # techo por hoja para el motor regex (lee la matriz)
MATCH_CAP = 500
NAMES_SAMPLE = 300
VALUE_TRUNC = 200

# Constantes Excel
XL_VALUES = -4163
XL_FORMULAS = -4123
XL_WHOLE = 1
XL_PART = 2
XL_BY_ROWS = 1
XL_EXCEL_LINKS = 1

# XlConnectionType (valores documentados; OJO: NO coinciden con el orden
# "intuitivo" ODBC=1/OLEDB=2 que se suele asumir).
XL_CONNECTION_TYPES = {
    1: "OLE DB",
    2: "ODBC",
    3: "XML Map",
    4: "Texto",
    5: "Web",
    6: "Data Feed",
    7: "Data Model",
    8: "Rango de hoja",
    9: "Sin origen",
}

# Nombres definidos built-in de Excel. En el XML llevan el prefijo `_xlnm.`,
# pero por COM llegan SIN el (p.ej. 'Datos!Print_Area'): hay que reconocer las
# dos formas. Borrarlos rompe areas de impresion/autofiltros (y Excel suele
# lanzar 1004 al intentarlo).
_BUILTIN_BASE_NAMES = {
    "print_area", "print_titles", "_filterdatabase", "filterdatabase",
    "criteria", "extract", "database", "consolidate_area", "sheet_title",
    "data_form", "auto_open", "auto_close", "auto_activate", "auto_deactivate",
}


def register(mcp, session, run):
    @mcp.tool()
    def search_workbook(query: str, where: str = "both", regex: bool = False,
                        match_case: bool = False, whole_cell: bool = False,
                        sheet: Optional[str] = None, max_results: int = 500) -> dict:
        """Buscar texto en las CELDAS del libro activo (el 'grep' de Excel).

        where: values|formulas|both. regex=False usa el motor nativo de Excel
        (rapido, sin techo de tamano); regex=True lee la hoja y aplica un patron
        Python (techo 2M celdas/hoja). sheet=None recorre todas las hojas,
        incluidas las ocultas.

        OJO idioma: buscar en formulas con regex=False usa el texto LOCALIZADO
        (en un Excel en espanol la funcion es 'SUMA(', no 'SUM('). Con
        regex=True se busca sobre .Formula, que SIEMPRE esta en ingles. El
        campo 'formula' de la respuesta va en ingles en ambos casos.

        Busca SOLO en celdas. NO busca en nombres de hoja, shapes/charts,
        comentarios ni codigo VBA (para nombres definidos usa
        list_defined_names; para VBA, get_vba_code)."""
        return run(_search_workbook, session, query, where, regex, match_case,
                   whole_cell, sheet, max_results)

    @mcp.tool()
    def list_defined_names(broken_only: bool = False) -> dict:
        """Inventario de nombres definidos: conteos del universo completo
        (total/broken/hidden/builtin, scope libro vs hoja) + muestra de hasta
        300 nombres. 'broken' = el RefersTo tiene #REF! o apunta a un origen
        externo [n] no disponible.

        OJO: si open_workbook saneo el archivo (respuesta con bloque
        'sanitized'), los nombres rotos YA fueron removidos de la copia de
        trabajo y aqui saldran ~0 rotos. Es lo esperado."""
        return run(_list_defined_names, session, broken_only)

    @mcp.tool()
    def clean_defined_names(broken_only: bool = True, include_builtin: bool = False,
                            save: bool = False) -> dict:
        """ESCRITURA: borrar nombres definidos del libro activo.

        Por defecto borra solo los ROTOS (#REF! / [n]) y PRESERVA los built-in
        de Excel (areas de impresion, _FilterDatabase) aunque esten rotos.
        save=False por defecto: el borrado queda en memoria hasta que guardes.

        Si el libro se abrio como copia saneada, aqui no habra rotos que borrar
        (ya los quito el saneo) y para persistir hay que usar
        save_workbook(<ruta original>)."""
        return run(_clean_defined_names, session, broken_only, include_builtin, save)

    @mcp.tool()
    def list_connections() -> dict:
        """Mapear de donde sale la data: conexiones del libro (OLE DB, ODBC,
        Power Query, cubos OLAP/SSAS, Data Model, texto) + vinculos a otros
        libros de Excel.

        Nunca se conecta ni sondea la red: 'reachable' distingue rutas locales
        verificadas de fuentes de servidor/UNC ('no_verificable'). Las cadenas
        de conexion se devuelven parseadas y SIN credenciales."""
        return run(_list_connections, session)


# =============================================================================
# Componente 1: search_workbook
# =============================================================================


def _short(value):
    """Valor JSON-safe y acotado para mostrar en un match."""
    v = to_jsonable(value)
    # Excel entrega todo numero como double: un codigo '5060094' llegaria como
    # 5060094.0 y confundiria al leer los resultados.
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, str) and len(v) > VALUE_TRUNC:
        return v[:VALUE_TRUNC] + "..."
    return v


def _as_text(value) -> str:
    """Texto sobre el que corre el regex. Los numeros llegan como double: sin
    esto, buscar '^5$' no matchearia un 5 (seria '5.0') y '^5060094$' tampoco
    encontraria un codigo PT — justo lo contrario de lo que ve el usuario."""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


class _Matches:
    """Acumulador con dedupe por (hoja, celda) y tope de resultados."""

    def __init__(self, cap: int):
        self.cap = cap
        self.items = {}
        self.truncated = False

    def add(self, sheet: str, addr: str, value, formula, matched_in: str) -> bool:
        """Registra un match. Devuelve False si ya no cabe (cap alcanzado)."""
        key = (sheet, addr)
        entry = self.items.get(key)
        if entry is not None:
            if matched_in not in entry["matched_in"]:
                entry["matched_in"] = "value+formula"
            return True
        if len(self.items) >= self.cap:
            self.truncated = True
            return False
        self.items[key] = {
            "sheet": sheet, "cell": addr, "value": _short(value),
            "formula": _short(formula), "matched_in": matched_in,
        }
        return True


def _find_pass(ws, query: str, look_in: int, whole_cell: bool, match_case: bool,
               label: str, acc: _Matches) -> bool:
    """Una pasada Find/FindNext sobre una hoja. False si se lleno el cap."""
    sheet_name = str(ws.Name)
    try:
        cell = ws.Cells.Find(
            What=query, LookIn=look_in, LookAt=(XL_WHOLE if whole_cell else XL_PART),
            SearchOrder=XL_BY_ROWS, MatchCase=match_case,
        )
    except Exception as e:
        # Hojas de grafico/protegidas pueden rechazar Find: no aborta la busqueda.
        logger.debug("Find fallo en '%s' (%s): %s", sheet_name, label, e)
        return True
    first = None
    while cell is not None:
        addr = str(cell.Address)
        if first is None:
            first = addr
        elif addr == first:
            break  # dio la vuelta completa a la hoja
        try:
            value, formula = cell.Value, cell.Formula
        except Exception:
            value, formula = None, None
        if not acc.add(sheet_name, addr, value, formula, label):
            return False
        try:
            cell = ws.Cells.FindNext(cell)
        except Exception as e:
            logger.debug("FindNext fallo en '%s': %s", sheet_name, e)
            break
    return True


def _regex_pass(ws, pattern, where: str, acc: _Matches) -> bool:
    """Escaneo regex de una hoja leyendo la matriz completa. False si lleno el cap."""
    sheet_name = str(ws.Name)
    ur = ws.UsedRange
    n_rows, n_cols = int(ur.Rows.Count), int(ur.Columns.Count)
    total = n_rows * n_cols
    if total > MAX_SEARCH_CELLS:
        raise ValueError(
            f"La hoja '{sheet_name}' tiene {total} celdas (> {MAX_SEARCH_CELLS}) "
            "para busqueda regex: acota con sheet=... o usa regex=False "
            "(motor nativo de Excel, sin techo de tamano)."
        )
    want_values = where in ("values", "both")
    want_formulas = where in ("formulas", "both")
    values = matrix_to_jsonable(ur.Value) if want_values else None
    formulas = matrix_to_jsonable(ur.Formula) if want_formulas else None
    # UsedRange NO empieza necesariamente en A1: la direccion se resuelve con
    # Cells(fila, col) — 1 llamada COM por MATCH (acotado por el cap), no por celda.
    row0, col0 = int(ur.Row), int(ur.Column)
    for i in range(n_rows):
        for j in range(n_cols):
            val = values[i][j] if values else None
            fml = formulas[i][j] if formulas else None
            hit_v = val is not None and pattern.search(_as_text(val))
            hit_f = fml is not None and pattern.search(_as_text(fml))
            if not (hit_v or hit_f):
                continue
            matched_in = "value+formula" if (hit_v and hit_f) else (
                "value" if hit_v else "formula")
            addr = str(ws.Cells(row0 + i, col0 + j).Address)
            if not acc.add(sheet_name, addr, val, fml, matched_in):
                return False
    return True


def _search_workbook(session, query: str, where: str, regex: bool, match_case: bool,
                     whole_cell: bool, sheet, max_results: int) -> dict:
    if not query:
        raise ValueError("query vacio: indica el texto (o patron) a buscar.")
    where = (where or "both").lower()
    if where not in ("values", "formulas", "both"):
        raise ValueError(f"where='{where}' invalido: usa 'values', 'formulas' o 'both'.")
    cap = max(1, min(int(max_results), MATCH_CAP))
    wb = get_active_workbook(session)
    sheets = [get_sheet(wb, sheet)] if sheet else list(wb.Worksheets)
    acc = _Matches(cap)

    if regex:
        try:
            pattern = re.compile(query, 0 if match_case else re.IGNORECASE)
        except re.error as e:
            raise ValueError(f"Patron regex invalido: {e}")
        for ws in sheets:
            if not _regex_pass(ws, pattern, where, acc):
                break
    else:
        passes = []
        if where in ("values", "both"):
            passes.append((XL_VALUES, "value"))
        if where in ("formulas", "both"):
            passes.append((XL_FORMULAS, "formula"))
        for ws in sheets:
            lleno = False
            for look_in, label in passes:
                if not _find_pass(ws, query, look_in, whole_cell, match_case, label, acc):
                    lleno = True
                    break
            if lleno:
                break

    matches = list(acc.items.values())
    return {"query": query, "count": len(matches), "truncated": acc.truncated,
            "engine": "regex" if regex else "excel_find",
            "sheets_searched": len(sheets), "matches": matches}


# =============================================================================
# Componentes 2 y 3: nombres definidos
# =============================================================================


def _base_name(raw: str) -> str:
    """'Datos!Print_Area' -> 'Print_Area'; '_xlnm.Print_Area' -> 'Print_Area'."""
    base = raw.split("!")[-1] if "!" in raw else raw
    if base.lower().startswith("_xlnm."):
        base = base[len("_xlnm."):]
    return base


def _is_builtin(raw: str) -> bool:
    if raw.lower().startswith("_xlnm.") or ".xlnm." in raw.lower():
        return True
    return _base_name(raw).lower() in _BUILTIN_BASE_NAMES


def _name_scope(raw: str, hojas: set) -> str:
    """'workbook' o 'sheet:<hoja>' a partir del nombre COM ('Hoja!Nombre').

    Se resuelve por STRING, no por `nm.Parent`: sobre un libro con miles de
    nombres, preguntar el Parent (y sondearlo con una excepcion) costaba
    ~2 llamadas COM por nombre y multiplicaba el tiempo por 3.
    """
    if "!" not in raw:
        return "workbook"
    prefijo, _, _ = raw.partition("!")
    prefijo = prefijo.strip("'").replace("''", "'")
    return f"sheet:{prefijo}" if prefijo in hojas else "workbook"


def _collect_names(wb) -> list:
    """Todos los nombres del libro, deduplicados por (scope, nombre base).

    wb.Names YA incluye los de scope hoja (verificado: aparecen como
    'Hoja!Nombre' con Parent Worksheet). Se recorren igual los ws.Names por si
    algun host no los expusiera, con dedupe.
    """
    seen = set()
    out = []
    hojas = set()
    for ws in wb.Worksheets:
        try:
            hojas.add(str(ws.Name))
        except Exception:
            pass

    def emit(nm):
        try:
            raw = str(nm.Name)
        except Exception:
            return
        scope = _name_scope(raw, hojas)
        key = (scope, _base_name(raw).lower())
        if key in seen:
            return
        seen.add(key)
        try:
            refers_to = str(nm.RefersTo)
            legible = True
        except Exception:
            refers_to, legible = "", False
        try:
            visible = bool(nm.Visible)
        except Exception:
            visible = True
        out.append({
            "name": _base_name(raw),
            "raw": raw,
            "scope": scope,
            "refers_to": refers_to,
            "visible": visible,
            # Un RefersTo ilegible se cuenta como roto: no hay forma de que
            # sirva y es justo lo que interesa limpiar.
            "broken": (not legible) or bool(_BROKEN_RE.search(refers_to)),
            "builtin": _is_builtin(raw),
        })

    for nm in wb.Names:
        emit(nm)
    for ws in wb.Worksheets:
        try:
            for nm in ws.Names:
                emit(nm)
        except Exception as e:
            logger.debug("ws.Names inaccesible en '%s': %s", ws.Name, e)
    return out


def _public_name(entry: dict) -> dict:
    return {k: v for k, v in entry.items() if k != "raw"}


def _list_defined_names(session, broken_only: bool) -> dict:
    wb = get_active_workbook(session)
    todos = _collect_names(wb)
    muestra = [n for n in todos if n["broken"]] if broken_only else todos
    return {
        # Los conteos son SIEMPRE del universo completo, no de la muestra.
        "total": len(todos),
        "broken": sum(1 for n in todos if n["broken"]),
        "hidden": sum(1 for n in todos if not n["visible"]),
        "builtin": sum(1 for n in todos if n["builtin"]),
        "workbook_scoped": sum(1 for n in todos if n["scope"] == "workbook"),
        "sheet_scoped": sum(1 for n in todos if n["scope"] != "workbook"),
        "broken_only": broken_only,
        "names": [_public_name(n) for n in muestra[:NAMES_SAMPLE]],
        "truncated": len(muestra) > NAMES_SAMPLE,
    }


def _resolve_name(wb, entry: dict):
    """Reobtiene el objeto Name (la coleccion se reindexa al borrar)."""
    try:
        return wb.Names(entry["raw"])
    except Exception:
        scope = entry["scope"]
        if scope.startswith("sheet:"):
            return wb.Worksheets(scope[len("sheet:"):]).Names(entry["name"])
        raise


def _clean_defined_names(session, broken_only: bool, include_builtin: bool,
                         save: bool) -> dict:
    wb = get_active_workbook(session)
    if bool(wb.ReadOnly):
        raise RuntimeError(
            "El libro esta en solo lectura; reabrelo con "
            "open_workbook(read_only=False) para limpiar nombres."
        )
    # Se captura la LISTA primero (strings, no objetos COM): borrar mientras se
    # itera reindexa la coleccion y se saltan elementos.
    objetivos = [
        n for n in _collect_names(wb)
        if (n["broken"] or not broken_only) and (include_builtin or not n["builtin"])
    ]
    removed, failed = 0, []
    for entry in objetivos:
        try:
            _resolve_name(wb, entry).Delete()
            removed += 1
        except Exception as e:
            # Tipico en built-ins: Excel lanza 1004 al borrar un Print_Area.
            logger.debug("No se pudo borrar el nombre '%s': %s", entry["raw"], e)
            failed.append(entry["name"])
    saved = False
    if save and removed:
        full_name = str(wb.FullName)
        if session.is_temp_copy(full_name):
            raise RuntimeError(
                f"Se borraron {removed} nombres, pero el libro abierto es la COPIA "
                "saneada temporal: un guardado aqui se perderia. Persiste con "
                "save_workbook('<ruta original>')."
            )
        wb.Save()
        saved = True
    return {
        "removed": removed,
        "failed": len(failed),
        "failed_names": failed[:20],
        "remaining": int(wb.Names.Count),
        "broken_only": broken_only,
        "include_builtin": include_builtin,
        "saved": saved,
    }


# =============================================================================
# Componente 4: list_connections
# =============================================================================

_CRED_KEYS = {"password", "pwd", "user id", "uid"}


def _parse_conn_string(s: str) -> dict:
    """Cadena de conexion -> dict {clave_minuscula: valor}, respetando comillas.

    Los tokens sin '=' (el prefijo 'OLEDB;'/'ODBC;') se ignoran. NO filtra
    credenciales: eso lo hace el caller (ver _safe_conn_fields).
    """
    out = {}
    if not s:
        return out
    partes, token, quote = [], [], None
    for ch in str(s):
        if quote:
            token.append(ch)
            if ch == quote:
                quote = None
        elif ch in ('"', "'"):
            quote = ch
            token.append(ch)
        elif ch == ";":
            partes.append("".join(token))
            token = []
        else:
            token.append(ch)
    partes.append("".join(token))
    for p in partes:
        if "=" not in p:
            continue
        k, v = p.split("=", 1)
        k, v = k.strip().lower(), v.strip()
        if len(v) >= 2 and v[0] == v[-1] and v[0] in ('"', "'"):
            v = v[1:-1]
        if k:
            out[k] = v
    return out


def _safe_conn_fields(parsed: dict) -> dict:
    """Campos utiles SIN credenciales (la cadena cruda no se devuelve nunca)."""
    return {
        "provider": parsed.get("provider", ""),
        "data_source": parsed.get("data source", "") or parsed.get("dsn", ""),
        "initial_catalog": parsed.get("initial catalog", "") or parsed.get("database", ""),
        "location": parsed.get("location", ""),
        "integrated_security": "integrated security" in parsed,
        "had_credentials": bool(parsed.get("password") or parsed.get("pwd")),
    }


def _reachability(target: str) -> str:
    """Alcanzabilidad SIN tocar la red (un os.path.exists sobre UNC bloquea)."""
    if not target:
        return "n/a"
    t = str(target).strip().strip('"').strip("'")
    if not t or t.startswith("$"):  # $Embedded$ = modelo interno del libro
        return "n/a"
    low = t.lower()
    if low.startswith(("http://", "https://", "ftp://")):
        return "no_verificable (url)"
    if t.startswith("\\\\") or t.startswith("//"):
        return "no_verificable (red/UNC)"
    if re.match(r"^[A-Za-z]:[\\/]", t):
        return "local:existe" if os.path.exists(t) else "local:no_existe"
    if "\\" not in t and "/" not in t:
        return "no_verificable (servidor)"
    return "no_verificable (ruta relativa)"


def _classify(conn_type, provider: str, cmd_type) -> str:
    prov = (provider or "").lower()
    if prov.startswith("msolap"):
        return "OLAP/Cubo (SSAS)"
    if prov.startswith("microsoft.mashup"):
        return "Power Query"
    try:
        return XL_CONNECTION_TYPES.get(int(conn_type), f"tipo {int(conn_type)}")
    except (TypeError, ValueError):
        return "desconocido"


_CMD_TYPES = {1: "cube", 2: "sql", 3: "table", 4: "default", 5: "list"}


def _connection_info(conn) -> dict:
    info = {"name": "", "description": "", "type": "desconocido"}
    for attr, key in (("Name", "name"), ("Description", "description")):
        try:
            info[key] = str(getattr(conn, attr) or "")
        except Exception:
            pass
    try:
        conn_type = int(conn.Type)
    except Exception:
        conn_type = None
    # Una conexion WORKSHEET/MODEL/TEXT no expone OLEDBConnection: todo va en
    # try/except por tipo, nunca se asume la forma.
    oledb = None
    for attr in ("OLEDBConnection", "ODBCConnection"):
        try:
            candidate = getattr(conn, attr)
            if candidate is not None:
                oledb = candidate
                break
        except Exception:
            continue
    parsed, extra = {}, {}
    if oledb is not None:
        try:
            parsed = _parse_conn_string(str(oledb.Connection))
        except Exception as e:
            logger.debug("Cadena de conexion ilegible en '%s': %s", info["name"], e)
        for attr, key, conv in (
            ("CommandText", "command_text", str),
            ("SourceConnectionFile", "odc_file", str),
            ("RefreshOnFileOpen", "refresh_on_open", bool),
        ):
            try:
                raw = getattr(oledb, attr)
                extra[key] = conv(raw) if raw not in (None, "") else None
            except Exception:
                extra[key] = None
        try:
            extra["command_type"] = _CMD_TYPES.get(int(oledb.CommandType), "default")
        except Exception:
            extra["command_type"] = None
    campos = _safe_conn_fields(parsed)
    info["type"] = _classify(conn_type, campos["provider"], extra.get("command_type"))
    info.update(campos)
    info.update(extra)
    if extra.get("command_text") and len(extra["command_text"]) > VALUE_TRUNC:
        info["command_text"] = extra["command_text"][:VALUE_TRUNC] + "..."
    # Alcanzabilidad del origen y del .odc (ambos sin tocar la red)
    info["reachable"] = _reachability(campos["data_source"])
    if extra.get("odc_file"):
        info["odc_reachable"] = _reachability(extra["odc_file"])
    return info


def _list_connections(session) -> dict:
    wb = get_active_workbook(session)
    conexiones = []
    try:
        count = int(wb.Connections.Count)
    except Exception as e:
        logger.debug("wb.Connections inaccesible: %s", e)
        count = 0
    for i in range(1, count + 1):
        try:
            conexiones.append(_connection_info(wb.Connections(i)))
        except Exception as e:
            logger.warning("Conexion %d ilegible: %s", i, e)
            conexiones.append({"name": f"<conexion {i} ilegible>", "type": "error",
                               "error": str(e)[:VALUE_TRUNC]})
    # LinkSources devuelve None (no lista vacia) cuando no hay vinculos.
    vinculos = []
    try:
        fuentes = wb.LinkSources(XL_EXCEL_LINKS)
    except Exception as e:
        logger.debug("LinkSources fallo: %s", e)
        fuentes = None
    if fuentes:
        if isinstance(fuentes, str):
            fuentes = (fuentes,)
        for src in fuentes:
            ruta = to_jsonable(src)
            vinculos.append({"source": ruta, "reachable": _reachability(str(ruta))})
    return {"count": len(conexiones), "connections": conexiones,
            "excel_links": vinculos, "excel_link_count": len(vinculos)}
