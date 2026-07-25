# -*- coding: utf-8 -*-
"""Suite del paquete lectura masiva + recalculate (spec 2026-07-24).

Excel real via COM (patron test_hardening). Corre TODO en el thread principal
(un solo hilo = STA implicito valido para estas pruebas).

Uso:  .venv/Scripts/python.exe test_bulk_tools.py
"""
import os
import sys
import tempfile
import time

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

import psutil
import pythoncom

FAILED = []


def check(name, cond, extra=""):
    tag = "[OK]" if cond else "[ERR]"
    print(f"{tag} {name}" + (f" -> {extra}" if extra else ""))
    if not cond:
        FAILED.append(name)


def snapshot_excel_pids():
    return {p.pid for p in psutil.process_iter(["name"]) if p.info["name"] == "EXCEL.EXE"}


def build_fixture(sm):
    """Libro con: tabla 100x10 (Datos), formulas (Formulas), datos desde C4 (Offset),
    tabla vacia (Datos!L1:N1). Devuelve el workbook."""
    app = sm.get_application()
    wb = app.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Datos"
    ws.Range("A1:J1").Value = [tuple(f"H{i}" for i in range(1, 11))]
    filas = tuple(tuple(r * 10 + c for c in range(10)) for r in range(100))
    ws.Range("A2:J101").Value = filas
    # TRAMPA COM conocida del repo: NO pasar None posicional (VT_NULL -> E_INVALIDARG).
    lo = ws.ListObjects.Add(SourceType=1, Source=ws.Range("A1:J101"),
                            XlListObjectHasHeaders=1)  # xlSrcRange, xlYes
    lo.Name = "TablaDatos"
    # tabla vacia (solo encabezados)
    ws.Range("L1:N1").Value = [("X", "Y", "Z")]
    lo2 = ws.ListObjects.Add(SourceType=1, Source=ws.Range("L1:N1"),
                             XlListObjectHasHeaders=1)
    lo2.Name = "TablaVacia"
    ws2 = wb.Worksheets.Add()
    ws2.Name = "Formulas"
    ws2.Range("A1").Value = 2
    ws2.Range("A2").Value = 3
    ws2.Range("A3").Formula = "=A1+A2"
    ws3 = wb.Worksheets.Add()
    ws3.Name = "Offset"
    ws3.Range("C4:E6").Value = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
    return wb


def main():
    pids_before = snapshot_excel_pids()
    pythoncom.CoInitialize()
    tmpdir = tempfile.mkdtemp(prefix="bulk_test_")
    try:
        from excel_mcp_2013.session import SessionManager
        sm = SessionManager(visible=False)
        sm.start()
        wb = build_fixture(sm)
        check("fixture creada", wb.Worksheets.Count == 3, f"{wb.Worksheets.Count} hojas")

        # === recalculate ===
        from excel_mcp_2013.tools.workbook import _recalculate
        app = sm.get_application()
        app.Calculation = -4135  # xlCalculationManual (Workbooks.Add no pasa por open)
        ws2 = wb.Worksheets("Formulas")
        val_inicial = ws2.Range("A3").Value
        check("A3 = 5 antes", val_inicial == 5, str(val_inicial))
        ws2.Range("A1").Value = 10  # en Manual, A3 queda desactualizada (5)
        stale = ws2.Range("A3").Value
        check("A3 desactualizada en Manual", stale == 5, str(stale))
        r = _recalculate(sm, False, None, False)
        check("recalculate dirty devuelve dict", r.get("calculated") is True, str(r))
        check("A3 recalculada = 13", ws2.Range("A3").Value == 13,
              str(ws2.Range("A3").Value))
        check("mode dirty", r.get("mode") == "dirty", str(r.get("mode")))
        check("state done", r.get("calculation_state") == "done", str(r))
        r2 = _recalculate(sm, True, None, False)
        check("mode full", r2.get("mode") == "full", str(r2.get("mode")))
        ws2.Range("A2").Value = 20
        r3 = _recalculate(sm, False, "Formulas", False)
        check("mode sheet", r3.get("mode") == "sheet:Formulas", str(r3.get("mode")))
        check("A3 = 30 tras sheet calc", ws2.Range("A3").Value == 30,
              str(ws2.Range("A3").Value))
        r4 = _recalculate(sm, False, None, True)  # wait_async sin CUBE: no debe romper
        check("wait_async no rompe sin CUBE", r4.get("mode") == "async_cube", str(r4))
        check("Calculation sigue Manual tras wait_async",
              int(app.Calculation) == -4135, str(app.Calculation))
        try:
            _recalculate(sm, False, "NoExiste", False)
            check("sheet inexistente lanza error", False)
        except Exception as e:
            check("sheet inexistente lanza error", "NoExiste" in str(e) or "Hoja" in str(e),
                  str(e)[:60])

        # === read_table ===
        from excel_mcp_2013.tools.bulk import (
            MAX_INLINE_CELLS,
            _export_sheet,
            _read_table,
        )
        t = _read_table(sm, "tabladatos", None)  # case-insensitive a proposito
        check("read_table headers", t["headers"] == [f"H{i}" for i in range(1, 11)],
              str(t["headers"][:3]))
        check("read_table 100 filas", t["row_count"] == 100 and len(t["rows"]) == 100)
        check("read_table col_count 10", t["col_count"] == 10)
        check("read_table celda [0][0]=0 y [99][9]=999",
              t["rows"][0][0] == 0 and t["rows"][99][9] == 999,
              f"{t['rows'][0][0]}, {t['rows'][99][9]}")
        # tabla vacia
        tv = _read_table(sm, "TablaVacia", None)
        check("tabla vacia rows=[]", tv["rows"] == [] and tv["row_count"] == 0, str(tv))
        # tabla inexistente -> error con listado
        try:
            _read_table(sm, "NoExiste", None)
            check("tabla inexistente lanza error", False)
        except Exception as e:
            check("tabla inexistente lanza error", "TablaDatos" in str(e), str(e)[:80])
        # dest .csv y .json
        dest_csv = os.path.join(tmpdir, "tabla.csv")
        tc = _read_table(sm, "TablaDatos", dest_csv)
        check("dest csv escrito", os.path.exists(dest_csv))
        check("dest csv conteos", tc["rows"] == 100 and tc["cols"] == 10, str(tc))
        check("dest csv sample 5", len(tc["sample"]) == 5)
        with open(dest_csv, encoding="utf-8-sig") as f:
            lineas = f.read().splitlines()
        check("csv: encabezado + 100 filas", len(lineas) == 101, str(len(lineas)))
        check("csv primera linea headers", lineas[0].startswith("H1,H2"), lineas[0][:20])
        dest_json = os.path.join(tmpdir, "tabla.json")
        tj = _read_table(sm, "TablaDatos", dest_json)
        import json as _json
        data = _json.load(open(dest_json, encoding="utf-8"))
        check("json: 101 filas (headers+100)", len(data) == 101, str(len(data)))
        check("json celda [1][0] == 0", data[1][0] == 0, str(data[1][0]))
        # cap inline: tabla 501x100 = 50.100 > 50.000
        ws_cap = wb.Worksheets.Add()
        ws_cap.Name = "Cap"
        ws_cap.Range("A1:CV1").Value = [tuple(f"C{i}" for i in range(1, 101))]
        bloque = tuple(tuple(1 for _ in range(100)) for _ in range(501))
        ws_cap.Range("A2:CV502").Value = bloque
        lo_cap = ws_cap.ListObjects.Add(SourceType=1, Source=ws_cap.Range("A1:CV502"),
                                        XlListObjectHasHeaders=1)
        lo_cap.Name = "TablaGrande"
        try:
            _read_table(sm, "TablaGrande", None)
            check("cap inline lanza error", False)
        except Exception as e:
            check("cap inline lanza error", str(MAX_INLINE_CELLS) in str(e) or "dest" in str(e),
                  str(e)[:80])
        tg = _read_table(sm, "TablaGrande", os.path.join(tmpdir, "grande.csv"))
        check("cap con dest OK", tg["rows"] == 501, str(tg["rows"]))

        # === export_sheet ===
        dest_off = os.path.join(tmpdir, "offset.csv")
        eo = _export_sheet(sm, "Offset", dest_off, 5, None)
        check("export offset range $C$4:$E$6", eo["range"] == "$C$4:$E$6", eo["range"])
        check("export offset 3x3", eo["rows"] == 3 and eo["cols"] == 3, str(eo))
        with open(dest_off, encoding="utf-8-sig") as f:
            check("offset csv primera linea 1,2,3", f.readline().strip() == "1,2,3")
        dest_tsv = os.path.join(tmpdir, "offset.tsv")
        _export_sheet(sm, "Offset", dest_tsv, 5, None)
        with open(dest_tsv, encoding="utf-8-sig") as f:
            check("tsv usa tab", "\t" in f.readline())
        dest_rng = os.path.join(tmpdir, "rango.csv")
        er = _export_sheet(sm, "Datos", dest_rng, 5, "B5:D20")
        check("range_addr respetado", er["range"] == "$B$5:$D$20", er["range"])
        check("range_addr 16x3", er["rows"] == 16 and er["cols"] == 3, str(er))
        dest_json2 = os.path.join(tmpdir, "offset.json")
        ej = _export_sheet(sm, "Offset", dest_json2, 2, None)
        check("export json sample 2", len(ej["sample"]) == 2)
        try:
            _export_sheet(sm, "Offset", os.path.join(tmpdir, "malo.xlsx"), 5, None)
            check("extension mala lanza error", False)
        except Exception as e:
            check("extension mala lanza error", "xlsx" in str(e).lower()
                  or "Extension" in str(e), str(e)[:60])

        # === read_range cap ===
        from excel_mcp_2013.tools.cells import _read_range
        chico = _read_range(sm, "Offset", "C4:E6")
        check("read_range chico sigue devolviendo list", isinstance(chico, list)
              and chico[0][0] == 1, str(type(chico)))
        try:
            _read_range(sm, "Datos", "A1:ZZ5000")  # 702*5000 = 3.5M celdas
            check("read_range cap lanza error", False)
        except Exception as e:
            check("read_range cap lanza error", "export_sheet" in str(e), str(e)[:80])

        # === get_data_model_measures honesto ===
        from excel_mcp_2013.tools.power_pivot import _get_data_model_measures
        gm = _get_data_model_measures(sm)
        check("measures devuelve dict", isinstance(gm, dict), str(type(gm)))
        check("measures lista vacia en fixture", gm.get("measures") == [], str(gm)[:80])
        diag = gm.get("diagnostic", {})
        check("diagnostic presente", "model_present" in diag, str(diag)[:80])
        if diag.get("model_present"):
            check("model_tables entero", isinstance(diag.get("model_tables"), int),
                  str(diag))

        sm.close()
    finally:
        pythoncom.CoUninitialize()
        for pid in snapshot_excel_pids() - pids_before:
            try:
                psutil.Process(pid).kill()
            except Exception:
                pass
    print()
    if FAILED:
        print(f"[FALLARON {len(FAILED)}]: {FAILED}")
        return 1
    print("[OK] Suite bulk completa")
    return 0


if __name__ == "__main__":
    sys.exit(main())
