"""Tools de VBA: listar modulos, extraer codigo, ejecutar macros, inyectar codigo.

Restriccion C11 del plan maestro:
- EJECUTAR macros (Application.Run) NO requiere configuracion especial.
- LEER/ESCRIBIR codigo (wb.VBProject) requiere activar "Confiar en el acceso al
  modelo de objetos de proyectos de VBA" en el Centro de Confianza.
"""

import logging
from typing import Optional

from ..utils.excel_utils import (
    LONG_OP_TIMEOUT,
    VBA_COMPONENT_TYPES,
    get_active_workbook,
    to_jsonable,
)

logger = logging.getLogger(__name__)

TRUST_HELP = (
    "Acceso al proyecto VBA denegado. Activa: Archivo > Opciones > Centro de confianza > "
    "Configuracion del Centro de confianza > Configuracion de macros > "
    "'Confiar en el acceso al modelo de objetos de proyectos de VBA'."
)


def register(mcp, session, run):
    @mcp.tool()
    def list_vba_modules() -> list:
        """Listar los modulos VBA del workbook activo (nombre, tipo, lineas).
        Requiere 'Trust access to the VBA project object model'."""
        return run(_list_vba_modules, session)

    @mcp.tool()
    def get_vba_code(module_name: str) -> dict:
        """Extraer el codigo fuente completo de un modulo VBA.
        Requiere 'Trust access to the VBA project object model'."""
        return run(_get_vba_code, session, module_name)

    @mcp.tool()
    def execute_vba_macro(
        macro_name: str, arguments: Optional[list] = None, timeout_s: int = 120
    ) -> dict:
        """Ejecutar una macro VBA del workbook activo via Application.Run.
        OJO: el workbook debe haberse abierto con enable_macros=True.
        timeout_s: limite en segundos para macros largas (cap 600); el default
        de 120s aplica tambien a macros inyectadas."""
        timeout_s = max(1, min(int(timeout_s), LONG_OP_TIMEOUT))
        return run(_execute_vba_macro, session, macro_name, arguments, timeout=timeout_s)

    @mcp.tool()
    def inject_vba_code(module_name: str, code: str, replace: bool = False) -> dict:
        """Inyectar codigo VBA en un modulo (lo crea si no existe).
        replace=True borra el contenido previo del modulo.
        Requiere 'Trust access to the VBA project object model'."""
        return run(_inject_vba_code, session, module_name, code, replace)

    @mcp.tool()
    def analyze_vba_project() -> dict:
        """Analisis estatico del proyecto VBA completo: procedimientos por modulo,
        call graph (quien llama a quien), hojas y rangos que toca cada macro,
        y manejadores de eventos (Workbook_Open, Worksheet_Change...).
        Requiere 'Trust access to the VBA project object model'."""
        return run(_analyze_vba_project, session)


def _get_vbproject(wb):
    try:
        vbp = wb.VBProject
        _ = vbp.VBComponents.Count  # fuerza el chequeo de acceso
        return vbp
    except Exception as e:
        logger.warning("VBProject inaccesible: %s", e)
        raise PermissionError(TRUST_HELP) from e


def _list_vba_modules(session) -> list:
    wb = get_active_workbook(session)
    vbp = _get_vbproject(wb)
    modules = []
    for comp in vbp.VBComponents:
        ctype = int(comp.Type)
        modules.append(
            {
                "name": str(comp.Name),
                "type": VBA_COMPONENT_TYPES.get(ctype, f"unknown({ctype})"),
                "lines": int(comp.CodeModule.CountOfLines),
            }
        )
    return modules


def _get_vba_code(session, module_name: str) -> dict:
    wb = get_active_workbook(session)
    vbp = _get_vbproject(wb)
    comp = vbp.VBComponents(module_name)
    cm = comp.CodeModule
    total = int(cm.CountOfLines)
    code = cm.Lines(1, total) if total > 0 else ""
    return {
        "module": str(comp.Name),
        "type": VBA_COMPONENT_TYPES.get(int(comp.Type), "unknown"),
        "lines": total,
        "code": code,
    }


def _execute_vba_macro(session, macro_name: str, arguments: Optional[list]) -> dict:
    app = session.get_application()
    wb = get_active_workbook(session)
    # Calificar con el nombre del libro evita ambiguedad con varios libros abiertos
    qualified = macro_name if "!" in macro_name else f"'{wb.Name}'!{macro_name}"
    args = arguments or []
    logger.info("Ejecutando macro %s con %s args", qualified, len(args))
    result = app.Run(qualified, *args)
    return {"macro": qualified, "result": to_jsonable(result)}


import re

PROC_RE = re.compile(
    r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property\s+\w+)\s+(\w+)",
    re.IGNORECASE | re.MULTILINE,
)
SHEETS_TOUCHED_RE = re.compile(r"(?:Worksheets|Sheets)\(\s*\"([^\"]+)\"", re.IGNORECASE)
RANGES_TOUCHED_RE = re.compile(r"Range\(\s*\"([^\"]+)\"", re.IGNORECASE)


def _strip_vba_comments(code: str) -> str:
    """Quita comentarios ' hasta fin de linea (aprox: ignora ' dentro de strings)."""
    lines = []
    for line in code.splitlines():
        in_str = False
        for i, ch in enumerate(line):
            if ch == '"':
                in_str = not in_str
            elif ch == "'" and not in_str:
                line = line[:i]
                break
        lines.append(line)
    return "\n".join(lines)


def _analyze_vba_project(session) -> dict:
    wb = get_active_workbook(session)
    vbp = _get_vbproject(wb)

    # Pasada 1: extraer codigo y procedimientos de cada modulo
    modules = []
    all_procs = {}  # nombre_proc -> "Modulo.Proc"
    for comp in vbp.VBComponents:
        cm = comp.CodeModule
        total = int(cm.CountOfLines)
        code = cm.Lines(1, total) if total > 0 else ""
        clean = _strip_vba_comments(code)
        procs = []
        for m in PROC_RE.finditer(clean):
            kind, name = m.group(1).split()[0].capitalize(), m.group(2)
            procs.append({"name": name, "kind": kind, "start": m.start()})
            all_procs[name.lower()] = f"{comp.Name}.{name}"
        modules.append(
            {"component": comp, "name": str(comp.Name), "type": int(comp.Type),
             "code": clean, "procs": procs, "lines": total}
        )

    # Pasada 2: por procedimiento, cuerpo = desde su inicio hasta el siguiente proc
    result_modules = []
    call_graph = []
    for mod in modules:
        mod_out = {
            "module": mod["name"],
            "type": VBA_COMPONENT_TYPES.get(mod["type"], "unknown"),
            "lines": mod["lines"],
            "procedures": [],
        }
        procs = mod["procs"]
        for i, p in enumerate(procs):
            end = procs[i + 1]["start"] if i + 1 < len(procs) else len(mod["code"])
            body = mod["code"][p["start"]:end]
            calls = sorted(
                {
                    all_procs[t.lower()]
                    for t in re.findall(r"\b(\w+)\b", body)
                    if t.lower() in all_procs
                    and all_procs[t.lower()] != f"{mod['name']}.{p['name']}"
                }
            )
            is_event = bool(re.match(r"(Workbook_|Worksheet_|Auto_)", p["name"], re.IGNORECASE))
            mod_out["procedures"].append(
                {
                    "name": p["name"],
                    "kind": p["kind"],
                    "is_event_handler": is_event,
                    "calls": calls,
                    "sheets_touched": sorted(set(SHEETS_TOUCHED_RE.findall(body))),
                    "ranges_touched": sorted(set(RANGES_TOUCHED_RE.findall(body)))[:20],
                }
            )
            for callee in calls:
                call_graph.append({"caller": f"{mod['name']}.{p['name']}", "callee": callee})
        if mod_out["procedures"] or mod["lines"] > 0:
            result_modules.append(mod_out)

    return {
        "modules": result_modules,
        "call_graph": call_graph,
        "total_procedures": sum(len(m["procedures"]) for m in result_modules),
    }


def _inject_vba_code(session, module_name: str, code: str, replace: bool) -> dict:
    wb = get_active_workbook(session)
    vbp = _get_vbproject(wb)
    comp = None
    for c in vbp.VBComponents:
        if str(c.Name).lower() == module_name.lower():
            comp = c
            break
    created = False
    if comp is None:
        comp = vbp.VBComponents.Add(1)  # vbext_ct_StdModule
        comp.Name = module_name
        created = True
    cm = comp.CodeModule
    if replace and int(cm.CountOfLines) > 0:
        cm.DeleteLines(1, int(cm.CountOfLines))
    cm.AddFromString(code)
    logger.info("Codigo inyectado en modulo %s (created=%s, replace=%s)", module_name, created, replace)
    return {"module": str(comp.Name), "created": created, "lines": int(cm.CountOfLines)}
