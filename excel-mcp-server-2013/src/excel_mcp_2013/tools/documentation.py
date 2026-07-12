"""document_workbook: el expediente tecnico completo de una herramienta.

Orquesta las tools de analisis en UNA pasada STA y genera un markdown con la
documentacion que nadie tiene tiempo de escribir a mano.
"""

import logging
import os
from typing import Optional

from ..utils.excel_utils import LONG_OP_TIMEOUT
from .pivots import _list_pivot_tables
from .power_query import _list_power_queries
from .semantics import _check_2013_compatibility, _map_dependencies, _profile_formulas
from .vba import _analyze_vba_project
from .workbook import _analyze_workbook

logger = logging.getLogger(__name__)


def register(mcp, session, run):
    @mcp.tool()
    def document_workbook(
        output_path: Optional[str] = None, max_patterns_per_sheet: int = 40
    ) -> dict:
        """Generar la documentacion tecnica completa (markdown) del workbook
        activo: hojas, dependencias, patrones de formulas, pivots, Power Query,
        VBA y compatibilidad con Excel 2013. El entregable 'expediente de la
        herramienta'."""
        data = run(_collect_documentation, session, max_patterns_per_sheet, timeout=LONG_OP_TIMEOUT)
        md = _render_markdown(data)
        if output_path is None:
            folder = os.path.dirname(data["analysis"]["path"]) or "."
            stem = os.path.splitext(data["analysis"]["name"])[0]
            output_path = os.path.join(folder, f"{stem}_DOCUMENTACION.md")
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(md)
        logger.info("Documentacion generada: %s (%s chars)", output_path, len(md))
        return {
            "path": output_path,
            "sheet_count": data["analysis"]["sheet_count"],
            "warnings": data["warnings"],
            "compatible_2013": data["compat"]["compatible"],
        }


def _collect_documentation(session, max_patterns_per_sheet: int) -> dict:
    """Corre TODO el analisis en una sola pasada (ya estamos en el STA thread)."""
    warnings = []
    analysis = _analyze_workbook(session)

    profiles = {}
    for s in analysis["sheets"]:
        if s.get("formula_cells", 0) > 0:
            try:
                profiles[s["name"]] = _profile_formulas(
                    session, s["name"], max_patterns_per_sheet
                )
            except Exception as e:
                warnings.append(f"profile_formulas fallo en '{s['name']}': {e}")

    try:
        dependencies = _map_dependencies(session)
    except Exception as e:
        dependencies = None
        warnings.append(f"map_dependencies fallo: {e}")

    try:
        pivots = _list_pivot_tables(session)
    except Exception as e:
        pivots = []
        warnings.append(f"list_pivot_tables fallo: {e}")

    try:
        power_queries = _list_power_queries(session)
    except Exception as e:
        power_queries = None
        warnings.append(f"list_power_queries fallo: {e}")

    try:
        vba = _analyze_vba_project(session)
    except PermissionError:
        vba = None
        warnings.append("VBA no analizado: falta trust del VBA object model.")
    except Exception as e:
        vba = None
        warnings.append(f"analyze_vba_project fallo: {e}")

    compat = _check_2013_compatibility(session)

    return {
        "analysis": analysis,
        "profiles": profiles,
        "dependencies": dependencies,
        "pivots": pivots,
        "power_queries": power_queries,
        "vba": vba,
        "compat": compat,
        "warnings": warnings,
    }


def _render_markdown(d: dict) -> str:
    a = d["analysis"]
    out = [f"# Documentación técnica — {a['name']}", ""]
    out.append(f"Ruta: `{a['path']}`  ")
    out.append(
        f"Hojas: {a['sheet_count']} · Celdas con fórmula: {a['total_formula_cells']:,} · "
        f"VBA: {'sí' if a.get('has_vba_project') else 'no'} · "
        f"Nombres definidos: {a.get('defined_name_count', 0)}"
    )
    out.append("")

    # --- Hojas ---
    out.append("## 1. Hojas")
    out.append("")
    out.append("| Hoja | Rango usado | Filas | Fórmulas | Pivots | Tablas | Visible |")
    out.append("|---|---|---:|---:|---|---|---|")
    for s in a["sheets"]:
        out.append(
            f"| {s['name']} | {s.get('used_range') or '-'} | {s.get('rows', 0):,} "
            f"| {s['formula_cells']:,} | {', '.join(s['pivot_tables']) or '-'} "
            f"| {', '.join(s['tables']) or '-'} | {'sí' if s['visible'] else '**OCULTA**'} |"
        )
    out.append("")

    # --- Dependencias ---
    out.append("## 2. Dependencias entre hojas")
    out.append("")
    dep = d["dependencies"]
    if dep:
        for e in dep["edges"]:
            out.append(f"- `{e['from']}` → `{e['to']}` _(vía {e['via']})_")
        out.append("")
        out.append("Clasificación: " + " · ".join(
            f"**{k}**: {v}" for k, v in sorted(dep["classification"].items())
        ))
        if dep["hidden_sheets"]:
            out.append("")
            out.append(f"⚠ Hojas ocultas con posible lógica: {', '.join(dep['hidden_sheets'])}")
    else:
        out.append("_No disponible._")
    out.append("")

    # --- Formulas ---
    out.append("## 3. Patrones de fórmulas por hoja")
    out.append("")
    for sheet, prof in d["profiles"].items():
        out.append(
            f"### {sheet} — {prof['total_formula_cells']:,} fórmulas, "
            f"{prof['unique_patterns']} patrones únicos"
        )
        out.append("")
        out.append("| Celdas | Cubre | Fórmula (ejemplo) |")
        out.append("|---:|---|---|")
        for p in prof["patterns"][:15]:
            example = (p["example_a1"] or p["formula_r1c1"]).replace("|", "\\|")
            if len(example) > 110:
                example = example[:110] + "…"
            out.append(f"| {p['count']:,} | {p['covers']} | `{example}` |")
        if prof["truncated"] or prof["unique_patterns"] > 15:
            out.append("")
            out.append(f"_({prof['unique_patterns'] - 15} patrones más — usar `profile_formulas`)_")
        out.append("")

    # --- Pivots ---
    out.append("## 4. Tablas dinámicas")
    out.append("")
    if d["pivots"]:
        for pt in d["pivots"]:
            f = pt["fields"]
            out.append(
                f"- **{pt['name']}** en `{pt['sheet']}` ({pt.get('location', '?')}) — "
                f"origen `{pt.get('source') or '?'}` · filas: {', '.join(f['rows']) or '-'} · "
                f"columnas: {', '.join(f['columns']) or '-'} · filtros: {', '.join(f['filters']) or '-'} · "
                f"valores: {', '.join(f['values']) or '-'}"
            )
    else:
        out.append("_Sin tablas dinámicas._")
    out.append("")

    # --- Power Query ---
    out.append("## 5. Power Query / Conexiones")
    out.append("")
    pq = d["power_queries"]
    if pq and (pq["queries"] or pq["connections"]):
        for q in pq["queries"]:
            out.append(f"- Query: **{q['name']}** {q.get('description', '')}")
        for c in pq["connections"]:
            out.append(f"- Conexión: **{c['name']}** (tipo {c.get('type')})")
    else:
        out.append("_Sin queries ni conexiones de datos._")
    out.append("")

    # --- VBA ---
    out.append("## 6. Proyecto VBA")
    out.append("")
    vba = d["vba"]
    if vba and vba["total_procedures"]:
        for mod in vba["modules"]:
            if not mod["procedures"]:
                continue
            out.append(f"### {mod['module']} ({mod['type']}, {mod['lines']} líneas)")
            out.append("")
            for p in mod["procedures"]:
                extras = []
                if p["is_event_handler"]:
                    extras.append("**evento**")
                if p["calls"]:
                    extras.append(f"llama a: {', '.join(p['calls'])}")
                if p["sheets_touched"]:
                    extras.append(f"hojas: {', '.join(p['sheets_touched'])}")
                if p["ranges_touched"]:
                    extras.append(f"rangos: {', '.join(p['ranges_touched'][:6])}")
                out.append(f"- `{p['kind']} {p['name']}` " + (" · ".join(extras) if extras else ""))
            out.append("")
        if vba["call_graph"]:
            out.append("Call graph: " + " · ".join(
                f"`{e['caller']}`→`{e['callee']}`" for e in vba["call_graph"][:20]
            ))
            out.append("")
    elif vba is None:
        out.append("_No analizado (falta trust del VBA object model o no hay proyecto)._")
    else:
        out.append("_Proyecto VBA sin procedimientos._")
    out.append("")

    # --- Compatibilidad ---
    out.append("## 7. Compatibilidad con Excel 2013")
    out.append("")
    compat = d["compat"]
    out.append(("✅ " if compat["compatible"] else "❌ ") + compat["verdict"])
    out.append("")
    for f in compat["modern_functions_found"]:
        out.append(
            f"- `{f['function']}` (Excel {f['introduced']}) × {f['count']} en "
            f"`{f['sheet']}` (ej. {f['example_cell']})"
        )
    for s in compat["spill_references"][:10]:
        out.append(f"- Referencia de derrame `{s['formula'][:60]}` en `{s['sheet']}!{s['cell']}`")
    for n in compat["broken_names"]:
        out.append(f"- Nombre roto: `{n['name']}` → `{n['refers_to']}`")
    for q in compat["power_query_issues"]:
        out.append(f"- Query `{q['query']}` usa M no disponible en 2013: {', '.join(q['blocked'])}")
    out.append("")

    # --- Advertencias ---
    if d["warnings"]:
        out.append("## 8. Advertencias del análisis")
        out.append("")
        for w in d["warnings"]:
            out.append(f"- {w}")
        out.append("")

    out.append("---")
    out.append("_Generado automáticamente por excel-mcp-server-2013 (`document_workbook`)._")
    return "\n".join(out)
