"""Tools de comprension semantica: entender QUE calcula un libro y POR QUE.

- profile_formulas: las formulas arrastradas colapsan a patrones unicos en R1C1
  (114.898 formulas -> ~80 patrones). Leer 80 resumenes en vez de 114k celdas.
- trace_cell: precedentes/dependientes/nombres de una celda, con valores.
- check_2013_compatibility: guardian del viaje 2021 -> 2013.
- map_dependencies: grafo hoja->hoja + clasificacion entrada/calculo/salida.
"""

import logging
import re
from typing import Optional

from ..utils.compat_2013 import FUNC_RES, MODERN_FUNCTIONS, SPILL_REF_RE
from ..utils.excel_utils import get_active_workbook, to_jsonable
from ..utils.m_constraints import validate_m_expression

logger = logging.getLogger(__name__)

# Referencias con hoja: 'Mi Hoja'!A1:B2  o  Hoja1!A1
SHEET_REF_RE = re.compile(
    r"(?:'([^']+)'|(?<![\w!])([A-Za-z_\w.]+))!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)"
)
# Referencias sin hoja: A1, $B$2:C3 — no precedidas de ! ni parte de un token mayor,
# no seguidas de ( (seria funcion tipo LOG10( )
BARE_REF_RE = re.compile(
    r"(?<![\w!:$\]])(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)(?![\w(#])"
)
# Tokens candidatos a nombre definido
NAME_TOKEN_RE = re.compile(r"(?<![\w.'!])([A-Za-z_][\w.]{2,})(?!\s*\(|[\w!])")


def _col_letter(n: int) -> str:
    """1 -> A, 27 -> AA."""
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def register(mcp, session, run):
    @mcp.tool()
    def profile_formulas(sheet: str, max_patterns: int = 200) -> dict:
        """Huella de formulas de una hoja: agrupa las formulas por patron unico
        R1C1 (una formula arrastrada = 1 patron). Es LA forma de entender una
        hoja con miles de formulas sin leerlas todas."""
        return run(_profile_formulas, session, sheet, max_patterns)

    @mcp.tool()
    def trace_cell(sheet: str, cell: str, max_precedents: int = 25) -> dict:
        """Explicar una celda: su formula, que celdas/rangos/nombres la alimentan
        (con valores actuales) y que celdas dependen de ella."""
        return run(_trace_cell, session, sheet, cell, max_precedents)

    @mcp.tool()
    def check_2013_compatibility() -> dict:
        """Guardian pre-entrega: detecta funciones post-2013 (XLOOKUP, TEXTJOIN,
        LET...), referencias de derrame (A1#), nombres _xlfn.* rotos y M
        incompatible. Ejecutar SIEMPRE antes de entregar un libro a Excel 2013."""
        return run(_check_2013_compatibility, session)

    @mcp.tool()
    def map_dependencies() -> dict:
        """Grafo hoja->hoja del workbook activo (quien alimenta a quien, via
        formulas y pivots) + clasificacion entrada/calculo/salida/estatica."""
        return run(_map_dependencies, session)


# =============================================================================
# profile_formulas
# =============================================================================


def _iter_formula_cells(ur):
    """Itera (fila_abs, col_abs, formula_r1c1) de las celdas con formula."""
    base_row, base_col = int(ur.Row), int(ur.Column)
    data = ur.FormulaR1C1
    if data is None:
        return
    if not isinstance(data, tuple):
        data = ((data,),)
    for r, row in enumerate(data):
        for c, val in enumerate(row):
            if isinstance(val, str) and val.startswith("="):
                yield base_row + r, base_col + c, val


def _profile_formulas(session, sheet: str, max_patterns: int) -> dict:
    wb = get_active_workbook(session)
    ws = wb.Sheets(sheet)
    ur = ws.UsedRange

    patterns: dict = {}
    total = 0
    for row, col, f in _iter_formula_cells(ur):
        total += 1
        p = patterns.get(f)
        if p is None:
            patterns[f] = {
                "count": 1,
                "min_row": row, "max_row": row,
                "min_col": col, "max_col": col,
                "first": (row, col),
            }
        else:
            p["count"] += 1
            p["min_row"] = min(p["min_row"], row)
            p["max_row"] = max(p["max_row"], row)
            p["min_col"] = min(p["min_col"], col)
            p["max_col"] = max(p["max_col"], col)

    ordered = sorted(patterns.items(), key=lambda kv: -kv[1]["count"])
    result_patterns = []
    for f_r1c1, p in ordered[:max_patterns]:
        first_addr = f"{_col_letter(p['first'][1])}{p['first'][0]}"
        covers = (
            f"{_col_letter(p['min_col'])}{p['min_row']}:"
            f"{_col_letter(p['max_col'])}{p['max_row']}"
        )
        cols = (
            _col_letter(p["min_col"])
            if p["min_col"] == p["max_col"]
            else f"{_col_letter(p['min_col'])}-{_col_letter(p['max_col'])}"
        )
        try:
            example_a1 = str(ws.Range(first_addr).Formula)
        except Exception:
            example_a1 = None
        result_patterns.append(
            {
                "formula_r1c1": f_r1c1,
                "example_cell": first_addr,
                "example_a1": example_a1,
                "count": p["count"],
                "covers": covers,
                "columns": cols,
            }
        )

    return {
        "sheet": str(ws.Name),
        "total_formula_cells": total,
        "unique_patterns": len(patterns),
        "patterns": result_patterns,
        "truncated": len(patterns) > max_patterns,
    }


# =============================================================================
# trace_cell
# =============================================================================


def _read_small_range(ws_or_wb_range) -> object:
    """Valores de un rango si es <=3x3; si no, solo dimensiones."""
    rng = ws_or_wb_range
    n_rows, n_cols = int(rng.Rows.Count), int(rng.Columns.Count)
    if n_rows <= 3 and n_cols <= 3:
        v = rng.Value
        if not isinstance(v, tuple):
            return to_jsonable(v)
        return [[to_jsonable(c) for c in row] for row in v]
    return {"rows": n_rows, "cols": n_cols, "corner": to_jsonable(rng.Cells(1, 1).Value)}


def _trace_cell(session, sheet: str, cell: str, max_precedents: int) -> dict:
    wb = get_active_workbook(session)
    ws = wb.Sheets(sheet)
    rng = ws.Range(cell)
    formula = str(rng.Formula) if rng.Formula is not None else ""

    result = {
        "cell": f"{ws.Name}!{cell}",
        "formula": formula,
        "formula_local": str(rng.FormulaLocal) if rng.FormulaLocal else "",
        "value": to_jsonable(rng.Value),
        "number_format": str(rng.NumberFormat),
        "has_formula": formula.startswith("="),
    }
    if not formula.startswith("="):
        result["note"] = "La celda no tiene formula: es un valor constante."
        return result

    sheet_names = {str(s.Name) for s in wb.Worksheets}

    # --- Precedentes con hoja explicita ---
    precedents = []
    seen = set()
    stripped = SHEET_REF_RE.sub(" ", formula)  # lo que queda para refs sin hoja
    for m in SHEET_REF_RE.finditer(formula):
        target_sheet = m.group(1) or m.group(2)
        addr = m.group(3)
        key = (target_sheet, addr)
        if key in seen or target_sheet not in sheet_names:
            continue
        seen.add(key)
        entry = {"ref": f"'{target_sheet}'!{addr}", "sheet": target_sheet}
        try:
            entry["values"] = _read_small_range(wb.Sheets(target_sheet).Range(addr))
        except Exception as e:
            entry["error"] = str(e)
        precedents.append(entry)
        if len(precedents) >= max_precedents:
            break

    # --- Precedentes en la misma hoja ---
    if len(precedents) < max_precedents:
        for m in BARE_REF_RE.finditer(stripped):
            addr = m.group(1)
            key = (str(ws.Name), addr.replace("$", ""))
            if key in seen:
                continue
            seen.add(key)
            entry = {"ref": addr, "sheet": str(ws.Name)}
            try:
                entry["values"] = _read_small_range(ws.Range(addr))
            except Exception:
                continue  # falso positivo del parseo
            precedents.append(entry)
            if len(precedents) >= max_precedents:
                break
    result["precedents"] = precedents

    # --- Nombres definidos usados ---
    names_used = []
    try:
        wb_names = {str(n.Name).split("!")[-1]: n for n in wb.Names}
        for tok in set(NAME_TOKEN_RE.findall(stripped)):
            n = wb_names.get(tok)
            if n is not None:
                names_used.append({"name": tok, "refers_to": str(n.RefersTo)})
    except Exception:
        pass
    result["defined_names_used"] = names_used

    # --- Dependientes (limitacion COM: solo misma hoja) ---
    try:
        deps = rng.DirectDependents
        result["dependents"] = str(deps.Address).replace("$", "")
    except Exception:
        result["dependents"] = None
    result["dependents_note"] = (
        "DirectDependents solo ve la misma hoja (limitacion COM). Para dependencias "
        "entre hojas usa map_dependencies."
    )
    return result


# =============================================================================
# check_2013_compatibility
# =============================================================================


def _check_2013_compatibility(session) -> dict:
    wb = get_active_workbook(session)
    findings = []
    spill_refs = []

    for ws in wb.Worksheets:
        try:
            data = ws.UsedRange.Formula
        except Exception:
            continue
        if data is None:
            continue
        if not isinstance(data, tuple):
            data = ((data,),)
        base_row, base_col = int(ws.UsedRange.Row), int(ws.UsedRange.Column)
        per_func: dict = {}
        for r, row in enumerate(data):
            for c, val in enumerate(row):
                if not (isinstance(val, str) and val.startswith("=")):
                    continue
                for fn, rx in FUNC_RES.items():
                    if rx.search(val):
                        info = per_func.setdefault(
                            fn,
                            {
                                "count": 0,
                                "example_cell": f"{_col_letter(base_col + c)}{base_row + r}",
                                "example_formula": val[:120],
                            },
                        )
                        info["count"] += 1
                if SPILL_REF_RE.search(val):
                    spill_refs.append(
                        {
                            "sheet": str(ws.Name),
                            "cell": f"{_col_letter(base_col + c)}{base_row + r}",
                            "formula": val[:120],
                        }
                    )
        for fn, info in per_func.items():
            findings.append(
                {
                    "function": fn,
                    "introduced": MODERN_FUNCTIONS[fn],
                    "sheet": str(ws.Name),
                    **info,
                }
            )

    # Nombres _xlfn.*: el formato OOXML registra con ese prefijo TODA funcion
    # posterior a Excel 2003. Solo son BLOQUEANTES si la funcion tampoco existe
    # en 2013 (p.ej. _xlfn.XLOOKUP); _xlfn.SUMIFS o _xlfn.CUBEVALUE (era 2007)
    # son residuales inofensivos para un target 2013.
    broken_names = []
    try:
        for n in wb.Names:
            name = str(n.Name)
            try:
                refers = str(n.RefersTo)
            except Exception:
                refers = ""
            if name.startswith(("_xlfn.", "_xleta.")) or "#NAME?" in refers:
                fn = name.split(".", 1)[-1].upper()
                blocking = fn in MODERN_FUNCTIONS
                broken_names.append(
                    {"name": name, "refers_to": refers, "blocking_for_2013": blocking}
                )
    except Exception:
        pass
    blocking_names = [n for n in broken_names if n["blocking_for_2013"]]

    pq_issues = []
    try:
        for q in wb.Queries:
            check = validate_m_expression(str(q.Formula))
            if not check["valid"]:
                pq_issues.append({"query": str(q.Name), "blocked": check["blocked_found"]})
    except Exception:
        pass  # coleccion Queries no disponible

    residual_count = len(broken_names) - len(blocking_names)
    compatible = not findings and not spill_refs and not pq_issues and not blocking_names
    return {
        "compatible": compatible,
        "modern_functions_found": findings,
        "spill_references": spill_refs[:50],
        "broken_names": broken_names,
        "power_query_issues": pq_issues,
        "verdict": (
            "El libro es seguro para Excel 2013."
            if compatible
            else "El libro NO es seguro para Excel 2013: corrige los hallazgos antes de entregar."
        )
        + (
            f" ({residual_count} nombres _xlfn.* residuales inofensivos: funciones era "
            "2007-2013 que OOXML registra asi; no afectan)."
            if residual_count
            else ""
        ),
    }


# =============================================================================
# map_dependencies
# =============================================================================


def _sheets_referenced_in(text: str, sheet_names: set) -> set:
    found = set()
    for m in re.finditer(r"'([^']+)'!|(?<![\w!'])([\w.]+)!", text):
        name = m.group(1) or m.group(2)
        if name in sheet_names:
            found.add(name)
    return found


def _map_dependencies(session) -> dict:
    wb = get_active_workbook(session)
    sheet_names = {str(s.Name) for s in wb.Worksheets}
    edges = []
    formula_counts = {}
    pivot_sheets = set()
    hidden = []

    for ws in wb.Worksheets:
        name = str(ws.Name)
        if int(ws.Visible) != -1:
            hidden.append(name)

        # formulas -> aristas
        refs = set()
        n_formulas = 0
        try:
            data = ws.UsedRange.Formula
            if data is not None:
                if not isinstance(data, tuple):
                    data = ((data,),)
                for row in data:
                    for val in row:
                        if isinstance(val, str) and val.startswith("="):
                            n_formulas += 1
                            refs |= _sheets_referenced_in(val, sheet_names)
        except Exception:
            pass
        formula_counts[name] = n_formulas
        for src in refs:
            if src != name:
                edges.append({"from": src, "to": name, "via": "formulas"})

        # pivots -> aristas
        try:
            for pt in ws.PivotTables():
                pivot_sheets.add(name)
                try:
                    source = str(pt.PivotCache().SourceData)
                except Exception:
                    continue
                src_sheet = source.split("!")[0].strip("'")
                if src_sheet in sheet_names and src_sheet != name:
                    edges.append({"from": src_sheet, "to": name, "via": f"pivot:{pt.Name}"})
        except Exception:
            pass

    # dedup
    unique_edges = [dict(t) for t in {tuple(sorted(e.items())) for e in edges}]
    feeds_others = {e["from"] for e in unique_edges}

    classification = {}
    for name in sheet_names:
        n = formula_counts.get(name, 0)
        if name in pivot_sheets and n < 100:
            classification[name] = "salida (pivots)"
        elif n == 0 and name in feeds_others:
            classification[name] = "entrada (datos)"
        elif n > 0:
            classification[name] = "calculo"
        else:
            classification[name] = "estatica"

    return {
        "edges": sorted(unique_edges, key=lambda e: (e["from"], e["to"])),
        "classification": classification,
        "formula_counts": formula_counts,
        "hidden_sheets": hidden,
    }
