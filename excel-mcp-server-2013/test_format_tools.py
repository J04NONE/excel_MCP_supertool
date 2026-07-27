# -*- coding: utf-8 -*-
"""Suite del paquete de autoria de estilos v1.6.0 (spec 2026-07-26).

Excel real via COM (patron test_hardening: thread principal = STA implicito).

Uso:  .venv/Scripts/python.exe test_format_tools.py
"""
import os
import sys

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


def test_validacion_pura():
    """_validate_spec es pura: se prueba sin Excel."""
    from excel_mcp_2013.tools.cells import _validate_spec
    print("\n--- validacion (sin Excel) ---")
    _validate_spec({"bold": True, "border": "thin", "h_align": "center"}, "t")
    check("spec valida pasa", True)
    for spec, frag in (
        ({"negrita": True}, "desconocidas"),
        ({"h_align": "middle"}, "h_align"),
        ({"v_align": "abajo"}, "v_align"),
        ({"border": "dashed"}, "border"),
        ({"fill_color_rgb": "XX0000"}, "hex"),
        ({"font_color_rgb": "F00"}, "hex"),
    ):
        try:
            _validate_spec(spec, "t")
            check(f"rechaza {spec}", False)
        except ValueError as e:
            check(f"rechaza {spec}", frag in str(e), str(e)[:70])


def main():
    test_validacion_pura()

    pids_before = snapshot_excel_pids()
    pythoncom.CoInitialize()
    try:
        from excel_mcp_2013.session import SessionManager
        from excel_mcp_2013.tools.cells import (
            _apply_format, _apply_format_batch, _set_freeze_panes,
        )
        sm = SessionManager(visible=False)
        sm.start()
        app = sm.get_application()
        wb = app.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "Datos"
        ws.Range("A1:F1").Value = [("A", "B", "C", "D", "E", "F")]
        ws.Range("A2:F10").Value = tuple(tuple(r * 6 + c for c in range(6))
                                         for r in range(9))
        ws2 = wb.Worksheets.Add()
        ws2.Name = "Otra"
        ws2.Range("A1").Value = "x"

        # === apply_format legacy (sin regresion) ===
        print("\n--- apply_format legacy ---")
        r = _apply_format(sm, "Datos", "A1:F1",
                          {"bold": True, "fill_color_rgb": "D9E1F2",
                           "number_format": "@"})
        check("legacy aplica", "bold" in r and "fill_color" in r, r)
        check("legacy verificado por COM", bool(ws.Range("A1").Font.Bold))

        # === props nuevas, verificadas leyendo de vuelta ===
        print("\n--- props nuevas ---")
        _apply_format(sm, "Datos", "A1:F1",
                      {"font_name": "Arial", "h_align": "center",
                       "v_align": "center", "wrap_text": True, "row_height": 30})
        check("font_name", str(ws.Range("A1").Font.Name) == "Arial",
              str(ws.Range("A1").Font.Name))
        check("h_align center", int(ws.Range("A1").HorizontalAlignment) == -4108,
              ws.Range("A1").HorizontalAlignment)
        check("v_align center", int(ws.Range("A1").VerticalAlignment) == -4108)
        check("wrap_text", bool(ws.Range("A1").WrapText))
        check("row_height", float(ws.Rows(1).RowHeight) == 30.0,
              ws.Rows(1).RowHeight)

        _apply_format(sm, "Datos", "B1:D1", {"column_width": 14})
        check("column_width", float(ws.Columns("C").ColumnWidth) == 14.0,
              ws.Columns("C").ColumnWidth)
        check("column_width no toca columnas fuera del rango",
              float(ws.Columns("F").ColumnWidth) != 14.0)

        _apply_format(sm, "Datos", "A2:F5", {"border": "thin"})
        check("border thin", int(ws.Range("B3").Borders(9).LineStyle) == 1,
              ws.Range("B3").Borders(9).LineStyle)
        _apply_format(sm, "Datos", "A2:F5", {"border": "none"})
        check("border none limpia", int(ws.Range("B3").Borders(9).LineStyle) == -4142,
              ws.Range("B3").Borders(9).LineStyle)

        _apply_format(sm, "Datos", "H1:J1", {"merge": True})
        check("merge", bool(ws.Range("H1").MergeCells))

        # === batch ===
        print("\n--- apply_format_batch ---")
        r = _apply_format_batch(sm, "Datos", [
            {"range": "A7:F7", "bold": True, "fill_color_rgb": "FFF2CC"},
            {"range": "A8:F8", "number_format": "#,##0.00", "h_align": "right"},
            {"range": "A1", "sheet": "Otra", "italic": True},
        ])
        check("batch aplica 3", r["applied"] == 3 and r["failed"] == 0, str(r)[:120])
        check("batch verificado en hoja default", bool(ws.Range("A7").Font.Bold))
        check("batch item con sheet propia", bool(ws2.Range("A1").Font.Italic))
        check("batch details trae hoja!rango",
              r["details"][2]["range"] == "Otra!A1", str(r["details"][2]))

        # pre-vuelo: item invalido -> NADA se aplica
        antes = bool(ws.Range("A9").Font.Bold)
        try:
            _apply_format_batch(sm, "Datos", [
                {"range": "A9:F9", "bold": True},
                {"range": "A10:F10", "border": "punteado"},  # invalido
            ])
            check("batch invalido lanza", False)
        except ValueError as e:
            check("batch invalido lanza", "formats[1]" in str(e), str(e)[:80])
        check("pre-vuelo: el item valido NO se aplico",
              bool(ws.Range("A9").Font.Bold) == antes,
              f"bold A9 antes={antes} ahora={ws.Range('A9').Font.Bold}")
        try:
            _apply_format_batch(sm, "Datos", [{"bold": True}])  # sin range
            check("batch sin range lanza", False)
        except ValueError as e:
            check("batch sin range lanza", "'range'" in str(e), str(e)[:70])
        r_vacio = _apply_format_batch(sm, "Datos", [])
        check("batch vacio devuelve 0", r_vacio["applied"] == 0)
        # error COM post-validacion (rango malformado): registra y continua
        r_com = _apply_format_batch(sm, "Datos", [
            {"range": "ZZZZ99999999", "bold": True},
            {"range": "A10", "bold": True},
        ])
        check("error COM: falla 1 y aplica 1",
              r_com["failed"] == 1 and r_com["applied"] == 1, str(r_com)[:150])

        # === freeze panes ===
        print("\n--- set_freeze_panes ---")
        r = _set_freeze_panes(sm, "Datos", "C5")
        check("freeze C5", r["frozen"] and r["frozen_rows"] == 4
              and r["frozen_cols"] == 2, str(r))
        win = app.ActiveWindow
        check("freeze verificado por COM", bool(win.FreezePanes)
              and int(win.SplitRow) == 4 and int(win.SplitColumn) == 2,
              f"SplitRow={win.SplitRow} SplitCol={win.SplitColumn}")
        r = _set_freeze_panes(sm, "Datos", None)
        check("unfreeze", r["frozen"] is False and not bool(win.FreezePanes), str(r))
        r = _set_freeze_panes(sm, "Datos", "A3")
        check("freeze solo filas", r["frozen_rows"] == 2 and r["frozen_cols"] == 0,
              str(r))
        _set_freeze_panes(sm, "Datos", None)
        try:
            _set_freeze_panes(sm, "NoExiste", "C5")
            check("freeze hoja inexistente lanza", False)
        except Exception as e:
            check("freeze hoja inexistente lanza", "NoExiste" in str(e), str(e)[:70])

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
    print("[OK] Suite formato/layout completa")
    return 0


if __name__ == "__main__":
    sys.exit(main())
