"""Validacion de las tools visuales contra el dashboard real ya radiografiado
a mano (VBA inyectado, 2026-07-08): 'Seguimiento  PT-CUOTA  Julio 8.xlsb'.

Datos conocidos del archivo (ground truth, CORREGIDO por la primera corrida de
estas tools — el analisis manual con VBA de 2026-07-08 tenia 3 errores):
- Hoja DashBoard: 26 shapes top-level (34 contando anidados); 3 slicers;
  7 charts en total, 1 de ellos anidado (Chart 25 dentro de Group 24)
- Gauges reales: exactamente 2 (Chart 25 y Chart 34), ChartType=-4111
  xlCombination con 2 series; Chart 21 tambien es -4111 pero con 3 series (combo)
- 14 SlicerCaches; TODOS tienen slicer visible en alguna hoja (no hay huerfanos:
  los que no estan en DashBoard viven en Detalle / NO NETO / Resumen IND);
  los 3 de DashBoard controlan 'Tabla dinamica1'@Detalle

Solo lectura, macros OFF, sin refrescar nada. NO guarda.

Uso:  .venv/Scripts/python.exe test_dashboard_real.py
"""

import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")

import psutil

from src.excel_mcp_2013.com_guard import ExcelWriteGuard
from src.excel_mcp_2013.session import SessionManager
from src.excel_mcp_2013.tools.shapes import _list_charts, _list_shapes, _list_slicers

XLSB = r"C:\Users\muril\Desktop\MCP Excel\Seguimiento  PT-CUOTA  Julio 8.xlsb"

FAILED = []


def check(name: str, cond: bool, extra: str = "") -> None:
    tag = "[OK]" if cond else "[ERR]"
    print(f"{tag} {name}" + (f" -> {extra}" if extra else ""))
    if not cond:
        FAILED.append(name)


def snapshot_excel_pids() -> set:
    return {p.pid for p in psutil.process_iter(["name"]) if p.info["name"] == "EXCEL.EXE"}


def main() -> int:
    if not os.path.exists(XLSB):
        print(f"[ERR] archivo no encontrado: {XLSB}")
        return 1

    pids_before = snapshot_excel_pids()
    session = SessionManager()
    guard = ExcelWriteGuard()

    def run(func, *args, timeout=None):
        def runner():
            if session.get_application() is None:
                session.start()
            return func(*args)
        return guard.execute(runner, timeout=timeout)

    try:
        print("=== Abriendo dashboard real (read_only, macros OFF) ===")
        run(session.open_workbook, XLSB, True, None, False, timeout=180)
        print("[OK] abierto")

        print("=== list_shapes(DashBoard) ===")
        rs = run(_list_shapes, session, "DashBoard", timeout=180)
        check("26 shapes top-level", rs["top_level_count"] == 26, str(rs["top_level_count"]))
        check("total > 26 (anidados en grupos)", rs["total_count"] > 26, str(rs["total_count"]))
        by_type = {}
        def count_types(shapes):
            for s in shapes:
                by_type[s.get("type")] = by_type.get(s.get("type"), 0) + 1
                if s.get("children"):
                    count_types(s["children"])
        count_types(rs["shapes"])
        check("3 slicers visibles", by_type.get("slicer") == 3, str(by_type))
        check("grupos detectados", by_type.get("group", 0) >= 4, str(by_type.get("group")))

        print("=== list_charts(DashBoard) ===")
        rc = run(_list_charts, session, "DashBoard", timeout=180)
        check("7 charts (incluido el anidado)", rc["chart_count"] == 7, str(rc["chart_count"]))
        nested = [c for c in rc["charts"] if c["group_path"]]
        check("Chart 25 anidado en Group 24",
              len(nested) == 1 and nested[0]["name"] == "Chart 25"
              and nested[0]["group_path"] == "Group 24",
              str([(c["name"], c["group_path"]) for c in nested]))
        gauges = [c for c in rc["charts"]
                  if c["chart_type"] == -4111 and c["series_count"] == 2]
        check("2 gauges xlCombination con 2 series (Chart 25 y 34)",
              {g["name"] for g in gauges} == {"Chart 25", "Chart 34"},
              str([(g["name"], g["series_count"], g["chart_type_name"]) for g in gauges]))
        pivot_charts = [c for c in rc["charts"] if c["is_pivot_chart"]]
        check("PivotCharts apuntan a TD CUBO",
              pivot_charts and all(
                  (c["pivot_source"] or {}).get("sheet") == "TD CUBO" for c in pivot_charts
              ),
              str([(c["name"], c["pivot_source"]) for c in pivot_charts])[:120])

        print("=== list_slicers() — cronometrado, NO debe colgarse con OLAP ===")
        t0 = time.time()
        rsl = run(_list_slicers, session, False, timeout=120)
        elapsed = time.time() - t0
        check("14 slicer caches", rsl["slicer_cache_count"] == 14, str(rsl["slicer_cache_count"]))
        check("sin cuelgue OLAP (<30s)", elapsed < 30, f"{elapsed:.1f}s")
        olap_caches = [c for c in rsl["slicer_caches"] if c["olap"]]
        check("caches OLAP detectados", len(olap_caches) >= 5, str(len(olap_caches)))
        visibles = [c for c in rsl["slicer_caches"] if c["slicers"]]
        check("los 14 caches tienen slicer en alguna hoja (no hay huerfanos)",
              len(visibles) == 14, str(len(visibles)))
        en_dash = [
            c for c in rsl["slicer_caches"]
            if any(s["sheet"] == "DashBoard" for s in c["slicers"])
        ]
        check("3 caches con slicer EN DashBoard", len(en_dash) == 3,
              str([c["name"] for c in en_dash]))
        # CORRECCION al analisis manual de 2026-07-08: los 3 slicers del
        # DashBoard pertenecen a los caches no-OLAP Marca2/Grupo1/Tipo_Producto1
        # y controlan las 6 pivots de TD CUBO que alimentan los graficos del
        # dashboard (SI es interactivo). El analisis VBA consulto caches
        # OLAP de nombre parecido (Marca1/Grupo/Tipo_Producto -> Detalle).
        td_cubo_ok = en_dash and all(
            len(c["pivot_tables"]) >= 6
            and all("@TD CUBO" in pt for pt in c["pivot_tables"])
            and c["olap"] is False
            for c in en_dash
        )
        check("los 3 de DashBoard controlan las 6 pivots de TD CUBO (no-OLAP)",
              td_cubo_ok, str([c["pivot_tables"] for c in en_dash])[:150])
        otras_hojas = {
            s["sheet"] for c in rsl["slicer_caches"] for s in c["slicers"]
        } - {"DashBoard"}
        check("el resto de slicers vive en Detalle/NO NETO/Resumen IND",
              otras_hojas == {"Detalle", "NO NETO", "Resumen IND"}, str(otras_hojas))

        # include_items=True: los OLAP deben quedar omitidos con nota, sin colgarse
        t0 = time.time()
        rsl2 = run(_list_slicers, session, True, timeout=120)
        elapsed2 = time.time() - t0
        olap2 = [c for c in rsl2["slicer_caches"] if c["olap"]]
        check("include_items=True omite items OLAP con nota",
              all(c["items"] is None and c["items_note"] for c in olap2),
              (olap2[0]["items_note"] or "")[:60] if olap2 else "sin caches OLAP")
        check("include_items=True tampoco cuelga (<30s)", elapsed2 < 30, f"{elapsed2:.1f}s")

    finally:
        print("=== Limpieza (cerrar SIN guardar) ===")
        guard.execute(session.close)
        guard.shutdown()
        time.sleep(1)
        leftover = snapshot_excel_pids() - pids_before
        check("0 procesos EXCEL.EXE huerfanos", not leftover, str(leftover or "ninguno"))

    print()
    if FAILED:
        print(f"[ERR] {len(FAILED)} fallos: {FAILED}")
        return 1
    print("[OK] Tools visuales validadas contra el dashboard real")
    return 0


if __name__ == "__main__":
    sys.exit(main())
