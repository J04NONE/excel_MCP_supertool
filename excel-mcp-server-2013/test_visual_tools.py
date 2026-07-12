"""Prueba de las tools de introspeccion visual (Bloque C) con un libro sintetico.

Construye en memoria (sin guardar): 2 rectangulos + textbox + grupo, un chart
normal, un chart dentro de un grupo, un PivotChart sobre una pivot nativa y un
slicer no-OLAP. Valida list_shapes / list_charts / list_slicers contra eso.

Uso:  .venv/Scripts/python.exe test_visual_tools.py
"""

import sys
import time

sys.stdout.reconfigure(encoding="utf-8")

import psutil

from src.excel_mcp_2013.com_guard import ExcelWriteGuard
from src.excel_mcp_2013.session import SessionManager
from src.excel_mcp_2013.tools.shapes import _list_charts, _list_shapes, _list_slicers

FAILED = []

MSO_SHAPE_RECTANGLE = 1  # msoShapeRectangle (AddShape)
MSO_TEXT_ORIENTATION_HORIZONTAL = 1
XL_COLUMN_CLUSTERED = 51
XL_DOUGHNUT = -4120  # el -4111 es xlCombination (error clasico de tablas de enums)


def check(name: str, cond: bool, extra: str = "") -> None:
    tag = "[OK]" if cond else "[ERR]"
    print(f"{tag} {name}" + (f" -> {extra}" if extra else ""))
    if not cond:
        FAILED.append(name)


def snapshot_excel_pids() -> set:
    return {p.pid for p in psutil.process_iter(["name"]) if p.info["name"] == "EXCEL.EXE"}


def build_workbook(session):
    """Arma el libro sintetico completo. Corre en el STA."""
    app = session.get_application()
    wb = app.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Lienzo"

    # Datos para charts y pivot
    data = [("Cat", "Val"), ("Norte", 10), ("Sur", 20), ("Centro", 15), ("Norte", 5)]
    for r, row in enumerate(data, start=1):
        ws.Cells(r, 1).Value = row[0]
        ws.Cells(r, 2).Value = row[1]

    # --- Shapes basicos ---
    r1 = ws.Shapes.AddShape(MSO_SHAPE_RECTANGLE, 300, 10, 60, 30)
    r1.Name = "RectA"
    r2 = ws.Shapes.AddShape(MSO_SHAPE_RECTANGLE, 370, 10, 60, 30)
    r2.Name = "RectB"
    tb = ws.Shapes.AddTextbox(MSO_TEXT_ORIENTATION_HORIZONTAL, 300, 50, 130, 25)
    tb.Name = "Titulo"
    grp = ws.Shapes.Range(["RectA", "RectB"]).Group()
    grp.Name = "GrupoRects"

    # --- Chart normal (suelto) ---
    co = ws.ChartObjects().Add(300, 90, 220, 130)
    co.Name = "ChartSuelto"
    co.Chart.SetSourceData(ws.Range("A1:B5"))
    co.Chart.ChartType = XL_COLUMN_CLUSTERED

    # --- Chart dentro de un grupo (el caso del gauge real) ---
    co2 = ws.ChartObjects().Add(300, 240, 180, 110)
    co2.Name = "ChartAnidado"
    co2.Chart.SetSourceData(ws.Range("A1:B5"))
    co2.Chart.ChartType = XL_DOUGHNUT
    r3 = ws.Shapes.AddShape(MSO_SHAPE_RECTANGLE, 490, 240, 40, 110)
    r3.Name = "Marco"
    grp2 = ws.Shapes.Range(["ChartAnidado", "Marco"]).Group()
    grp2.Name = "GrupoGauge"

    # --- Pivot nativa + PivotChart + slicer no-OLAP ---
    ws2 = wb.Worksheets.Add()
    ws2.Name = "TD"
    cache = wb.PivotCaches().Create(SourceType=1, SourceData=f"'{ws.Name}'!A1:B5")  # xlDatabase
    pt = cache.CreatePivotTable(TableDestination=f"'{ws2.Name}'!R1C1", TableName="PivotTest")
    pt.PivotFields("Cat").Orientation = 1  # xlRowField
    df = pt.AddDataField(pt.PivotFields("Val"))
    df.Function = -4157  # xlSum

    co3 = ws2.ChartObjects().Add(250, 10, 220, 130)
    co3.Name = "ChartPivot"
    co3.Chart.SetSourceData(pt.TableRange1)
    co3.Chart.ChartType = XL_COLUMN_CLUSTERED

    slicer_cache = wb.SlicerCaches.Add2(pt, "Cat")
    # OJO: pasar None en los opcionales de Slicers.Add marshalea VT_NULL ->
    # E_INVALIDARG. Omitir opcionales y asignar propiedades despues.
    sl = slicer_cache.Slicers.Add(ws2)
    sl.Name = "SlicerCat"
    sl.Caption = "Categoria"
    sl.Top = 160
    sl.Left = 250
    sl.Width = 100
    sl.Height = 100
    return True


def main() -> int:
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
        print("=== Construyendo libro sintetico ===")
        run(build_workbook, session, timeout=180)
        print("[OK] libro construido")

        print("=== C1: list_shapes ===")
        rs = run(_list_shapes, session, "Lienzo")
        # Top-level: Titulo, GrupoRects, ChartSuelto, GrupoGauge = 4
        check("top_level_count == 4", rs["top_level_count"] == 4, str(rs["top_level_count"]))
        # Total: 4 top + 2 en GrupoRects + 2 en GrupoGauge = 8
        check("total_count == 8", rs["total_count"] == 8, str(rs["total_count"]))
        by_name = {s["name"]: s for s in rs["shapes"]}
        grp = by_name.get("GrupoRects")
        check("GrupoRects es group con 2 hijos",
              grp is not None and grp["type"] == "group"
              and grp["children"] is not None and len(grp["children"]) == 2,
              str([c["name"] for c in (grp or {}).get("children") or []]))
        child_names = {c["name"] for c in grp["children"]} if grp and grp["children"] else set()
        check("hijos RectA/RectB", child_names == {"RectA", "RectB"}, str(child_names))
        check("tipos legibles", by_name.get("Titulo", {}).get("type") == "text_box"
              and by_name.get("ChartSuelto", {}).get("type") == "chart")
        check("posiciones numericas", isinstance(by_name.get("Titulo", {}).get("left"), float))

        print("=== C2: list_charts ===")
        rc = run(_list_charts, session, "Lienzo")
        check("chart_count == 2 en Lienzo", rc["chart_count"] == 2, str(rc["chart_count"]))
        charts = {c["name"]: c for c in rc["charts"]}
        suelto = charts.get("ChartSuelto", {})
        anidado = charts.get("ChartAnidado", {})
        check("ChartSuelto sin group_path", suelto.get("group_path") is None)
        check("ChartAnidado con group_path", anidado.get("group_path") == "GrupoGauge",
              str(anidado.get("group_path")))
        check("chart_type_name correctos",
              suelto.get("chart_type_name") == "xlColumnClustered"
              and anidado.get("chart_type") == -4120
              and anidado.get("chart_type_name") == "xlDoughnut",
              f"{suelto.get('chart_type_name')} / {anidado.get('chart_type_name')}")
        check("ChartSuelto no es pivot y CONSERVA ChartType (regresion §5.4)",
              suelto.get("is_pivot_chart") is False and suelto.get("chart_type") == 51)
        check("series con formula", suelto.get("series") and suelto["series"][0].get("formula"),
              str((suelto.get("series") or [{}])[0].get("formula"))[:60])

        rc2 = run(_list_charts, session, "TD")
        cp = {c["name"]: c for c in rc2["charts"]}.get("ChartPivot", {})
        check("ChartPivot es PivotChart", cp.get("is_pivot_chart") is True)
        check("pivot_source correcto",
              (cp.get("pivot_source") or {}).get("pivot_table") == "PivotTest"
              and (cp.get("pivot_source") or {}).get("sheet") == "TD",
              str(cp.get("pivot_source")))

        print("=== C3: list_slicers ===")
        t0 = time.time()
        rsl = run(_list_slicers, session, False)
        elapsed = time.time() - t0
        check("1 slicer cache", rsl["slicer_cache_count"] == 1, str(rsl["slicer_cache_count"]))
        cache = rsl["slicer_caches"][0]
        check("olap == False", cache["olap"] is False)
        check("pivot_tables == ['PivotTest@TD']", cache["pivot_tables"] == ["PivotTest@TD"],
              str(cache["pivot_tables"]))
        check("slicer visible con hoja y posicion",
              cache["slicers"] and cache["slicers"][0]["sheet"] == "TD"
              and isinstance(cache["slicers"][0]["left"], float),
              str(cache["slicers"]))
        check("default NO trae items", cache["items"] is None and cache["items_note"] is None)

        rsl2 = run(_list_slicers, session, True)
        cache2 = rsl2["slicer_caches"][0]
        check("include_items=True trae items no-OLAP",
              cache2["items"] is not None and cache2["items"]["total"] == 3
              and set(cache2["items"]["selected"]) == {"Norte", "Sur", "Centro"},
              str(cache2["items"]))
        check("list_slicers rapido (<5s)", elapsed < 5, f"{elapsed:.2f}s")

    finally:
        print("=== Limpieza ===")
        try:
            wb = session.get_application()
            if wb is not None:
                guard.execute(lambda: [b.Close(SaveChanges=False) for b in list(session.get_application().Workbooks)])
        except Exception as e:  # noqa: BLE001
            print(f"[WARN] cierre de workbooks: {e}")
        guard.execute(session.close)
        guard.shutdown()
        time.sleep(1)
        leftover = snapshot_excel_pids() - pids_before
        check("0 procesos EXCEL.EXE huerfanos", not leftover, str(leftover or "ninguno"))

    print()
    if FAILED:
        print(f"[ERR] {len(FAILED)} fallos: {FAILED}")
        return 1
    print("[OK] Tools visuales verificadas contra libro sintetico")
    return 0


if __name__ == "__main__":
    sys.exit(main())
