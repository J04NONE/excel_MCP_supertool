# -*- coding: utf-8 -*-
"""Suite del paquete de introspeccion v1.5.0 (spec 2026-07-24).

Parte 1: funciones puras (sin Excel). Parte 2: Excel real via COM
(patron test_hardening: todo en el thread principal = STA implicito valido).

Uso:  .venv/Scripts/python.exe test_introspect.py
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

# Cadena REAL capturada del MULTIFORMATO (cubo de intranet inalcanzable).
CONN_SSAS = (
    "OLEDB;Provider=MSOLAP.5;Persist Security Info=True;Initial Catalog=VentasCorp;"
    "Data Source=OLAPSERVER;MDX Compatibility=1;Safety Options=2;"
    "MDX Missing Member Mode=Error;Update Isolation Level=2;Integrated Security=SSPI"
)


def check(name, cond, extra=""):
    tag = "[OK]" if cond else "[ERR]"
    print(f"{tag} {name}" + (f" -> {extra}" if extra else ""))
    if not cond:
        FAILED.append(name)


def snapshot_excel_pids():
    return {p.pid for p in psutil.process_iter(["name"]) if p.info["name"] == "EXCEL.EXE"}


# =============================================================================
# Parte 1: funciones puras
# =============================================================================


def test_puras():
    from excel_mcp_2013.tools.introspect import (
        _base_name, _is_builtin, _parse_conn_string, _reachability, _safe_conn_fields,
    )
    from excel_mcp_2013.utils.excel_utils import excel_error_name, to_jsonable

    print("\n--- funciones puras ---")
    # --- errores de celda (VT_ERROR llega como int con scode 0x800A0000|codigo) ---
    check("excel_error_name #N/A", excel_error_name(-2146826246) == "#N/A",
          excel_error_name(-2146826246))
    check("excel_error_name #DIV/0!", excel_error_name(-2146826281) == "#DIV/0!",
          excel_error_name(-2146826281))
    check("excel_error_name #GETTING_DATA", excel_error_name(-2146826245) == "#GETTING_DATA",
          excel_error_name(-2146826245))
    check("excel_error_name #REF!", excel_error_name(-2146826265) == "#REF!",
          excel_error_name(-2146826265))
    check("excel_error_name ignora numeros", excel_error_name(42) is None)
    check("excel_error_name ignora bool", excel_error_name(True) is None)
    check("excel_error_name ignora float", excel_error_name(3.14) is None)
    check("excel_error_name codigo fuera de mapa", excel_error_name(-2146826240) == "Error 2048",
          excel_error_name(-2146826240))
    check("to_jsonable traduce error", to_jsonable(-2146826246) == "#N/A")
    check("to_jsonable no rompe strings", to_jsonable("hola") == "hola")

    # --- parseo de cadena de conexion ---
    p = _parse_conn_string(CONN_SSAS)
    check("parse: provider", p.get("provider") == "MSOLAP.5", p.get("provider"))
    check("parse: data source", p.get("data source") == "OLAPSERVER", p.get("data source"))
    check("parse: initial catalog", p.get("initial catalog") == "VentasCorp",
          p.get("initial catalog"))
    check("parse: ignora prefijo OLEDB", "oledb" not in p, str(list(p)[:3]))
    campos = _safe_conn_fields(p)
    check("campos: integrated_security", campos["integrated_security"] is True)
    check("campos: had_credentials False", campos["had_credentials"] is False)
    check("campos: sin claves de credencial",
          not any(k in campos for k in ("password", "pwd", "user id", "uid")))

    con_pwd = _parse_conn_string('Provider=SQLOLEDB;Data Source=SRV;User ID=sa;Password=s3cr3t;')
    campos2 = _safe_conn_fields(con_pwd)
    check("campos: had_credentials True", campos2["had_credentials"] is True)
    check("campos: password NO se propaga", "s3cr3t" not in str(campos2), str(campos2))

    comillas = _parse_conn_string('Provider=Mashup;Location="C:\\a;b\\x.xlsx";Extended Properties=""')
    check("parse: respeta comillas (; interno)",
          comillas.get("location") == "C:\\a;b\\x.xlsx", comillas.get("location"))

    # --- alcanzabilidad (jamas toca la red) ---
    check("reach: servidor", _reachability("OLAPSERVER") == "no_verificable (servidor)",
          _reachability("OLAPSERVER"))
    check("reach: UNC", _reachability(r"\\servidor\share\x.xlsx") == "no_verificable (red/UNC)",
          _reachability(r"\\servidor\share\x.xlsx"))
    check("reach: local inexistente",
          _reachability(r"D:\pjimenez\cubo.odc") == "local:no_existe",
          _reachability(r"D:\pjimenez\cubo.odc"))
    check("reach: local existente",
          _reachability(os.path.abspath(__file__)) == "local:existe")
    check("reach: url", _reachability("https://x.com/a.csv") == "no_verificable (url)")
    check("reach: vacio", _reachability("") == "n/a")
    check("reach: $Embedded$", _reachability("$Embedded$") == "n/a")

    # --- nombres built-in ---
    check("base_name con hoja", _base_name("Datos!Print_Area") == "Print_Area")
    check("base_name con _xlnm", _base_name("_xlnm.Print_Area") == "Print_Area")
    check("is_builtin por COM", _is_builtin("Datos!Print_Area") is True)
    check("is_builtin por XML", _is_builtin("_xlnm._FilterDatabase") is True)
    check("is_builtin False en normal", _is_builtin("Area_Buena") is False)


# =============================================================================
# Parte 2: Excel real
# =============================================================================


def build_fixture(sm):
    """Libro con datos, formulas, celda en error y 4 nombres definidos."""
    app = sm.get_application()
    wb = app.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = "Datos"
    ws.Range("A1:J1").Value = [tuple(f"H{i}" for i in range(1, 11))]
    ws.Range("A2:J11").Value = tuple(tuple(r * 10 + c for c in range(10)) for r in range(10))
    ws.Range("L1").Value = "H1 total"          # substring: distingue whole_cell
    ws.Range("L2").Formula = '="H1 "&"total"'  # match en valor Y en formula

    ws2 = wb.Worksheets.Add()
    ws2.Name = "Formulas"
    ws2.Range("A1").Value = 2
    ws2.Range("A2").Value = 3
    ws2.Range("A3").Formula = "=SUM(A1:A2)"
    ws2.Range("A4").Formula = "=NA()"          # celda en error (sustituto seguro
                                               # de #GETTING_DATA)

    ws3 = wb.Worksheets.Add()
    ws3.Name = "Offset"
    ws3.Range("C4:E6").Value = ((1, 2, 3), (4, 5, 6), (7, 8, 9))

    # Nombres: valido, roto (via hoja borrada), de hoja y built-in.
    wb.Names.Add(Name="Area_Buena", RefersTo="=Datos!$A$1")
    tmp_ws = wb.Worksheets.Add()
    tmp_ws.Name = "Temporal"
    wb.Names.Add(Name="Area_Rota", RefersTo="=Temporal!$A$1")
    app.DisplayAlerts = False
    tmp_ws.Delete()  # deja Area_Rota en #REF!
    app.DisplayAlerts = True
    ws.Names.Add(Name="Local_Hoja", RefersTo="=Datos!$A$2")
    ws.PageSetup.PrintArea = "$A$1:$J$11"  # built-in Datos!Print_Area
    return wb


def test_search(sm):
    from excel_mcp_2013.tools.introspect import _search_workbook
    print("\n--- search_workbook ---")

    r = _search_workbook(sm, "H1", "values", False, False, False, None, 500)
    celdas = {(m["sheet"], m["cell"]) for m in r["matches"]}
    check("search nativo encuentra header", ("Datos", "$A$1") in celdas, str(sorted(celdas))[:120])
    check("search nativo incluye substring", ("Datos", "$L$1") in celdas)
    check("search engine=excel_find", r["engine"] == "excel_find")
    check("search recorre todas las hojas", r["sheets_searched"] == 3, r["sheets_searched"])

    r_whole = _search_workbook(sm, "H1", "values", False, False, True, None, 500)
    celdas_w = {(m["sheet"], m["cell"]) for m in r_whole["matches"]}
    check("whole_cell excluye 'H1 total'", ("Datos", "$L$1") not in celdas_w
          and ("Datos", "$A$1") in celdas_w, str(sorted(celdas_w))[:120])

    # El motor nativo busca sobre la formula LOCALIZADA ('=SUMA(A1:A2)' en un
    # Excel en espanol): se usa un token neutro al idioma.
    r_f = _search_workbook(sm, "A1:A2", "formulas", False, False, False, "Formulas", 500)
    check("search en formulas", any(m["cell"] == "$A$3" for m in r_f["matches"]),
          str(r_f["matches"])[:150])
    check("search formulas devuelve la formula en ingles",
          any("SUM(" in str(m["formula"]) for m in r_f["matches"]),
          str(r_f["matches"])[:150])
    r_fx = _search_workbook(sm, r"SUM\(", "formulas", True, False, False, "Formulas", 500)
    check("regex en formulas es independiente del idioma",
          any(m["cell"] == "$A$3" for m in r_fx["matches"]), str(r_fx["matches"])[:150])

    # Dedupe: L2 tiene "H1" en el valor calculado Y en la formula
    r_b = _search_workbook(sm, "H1", "both", False, False, False, "Datos", 500)
    l2 = [m for m in r_b["matches"] if m["cell"] == "$L$2"]
    check("dedupe: una sola entrada por celda", len(l2) == 1, f"{len(l2)} entradas")
    check("dedupe: matched_in value+formula",
          bool(l2) and l2[0]["matched_in"] == "value+formula",
          l2[0]["matched_in"] if l2 else "sin match")

    r_cap = _search_workbook(sm, "H", "values", False, False, False, "Datos", 2)
    check("max_results capa", r_cap["count"] == 2, r_cap["count"])
    check("max_results marca truncated", r_cap["truncated"] is True)

    # --- regex ---
    r_re = _search_workbook(sm, r"^H\d+$", "values", True, False, False, None, 500)
    check("regex matchea H1..H10", r_re["count"] == 10, r_re["count"])
    check("regex engine", r_re["engine"] == "regex")
    r_re2 = _search_workbook(sm, r"h1\s+total", "values", True, False, False, "Datos", 500)
    check("regex case-insensitive por defecto", r_re2["count"] >= 1, r_re2["count"])
    r_re3 = _search_workbook(sm, r"h1\s+total", "values", True, True, False, "Datos", 500)
    check("regex match_case=True filtra", r_re3["count"] == 0, r_re3["count"])
    r_off = _search_workbook(sm, r"^5$", "values", True, False, False, "Offset", 500)
    check("regex resuelve direccion con UsedRange desplazado",
          r_off["count"] == 1 and r_off["matches"][0]["cell"] == "$D$5",
          str(r_off["matches"])[:120])

    try:
        _search_workbook(sm, "[", "both", True, False, False, None, 500)
        check("regex invalido lanza error", False)
    except ValueError as e:
        check("regex invalido lanza error", "regex" in str(e).lower(), str(e)[:60])
    try:
        _search_workbook(sm, "x", "otro", False, False, False, None, 500)
        check("where invalido lanza error", False)
    except ValueError as e:
        check("where invalido lanza error", "where" in str(e), str(e)[:60])
    try:
        _search_workbook(sm, "", "both", False, False, False, None, 500)
        check("query vacio lanza error", False)
    except ValueError as e:
        check("query vacio lanza error", "vacio" in str(e), str(e)[:60])


def test_error_cells(sm):
    from excel_mcp_2013.tools.introspect import _search_workbook
    from excel_mcp_2013.tools.cells import _read_range
    print("\n--- celdas en error ---")

    vals = _read_range(sm, "Formulas", "A3:A4")
    check("read_range traduce el error a string",
          isinstance(vals[1][0], str) and vals[1][0].startswith("#"), str(vals))
    r = _search_workbook(sm, r"^#", "values", True, False, False, "Formulas", 500)
    check("regex sobre celda en error no rompe", r["count"] >= 1, str(r["matches"])[:120])
    # regex (no el motor nativo): =NA() se muestra como =NOD() en Excel espanol.
    r2 = _search_workbook(sm, r"NA\(\)", "formulas", True, False, False, "Formulas", 500)
    check("celda en error localizable por formula",
          any(m["cell"] == "$A$4" for m in r2["matches"]), str(r2["matches"])[:120])


def test_regex_ceiling(sm, wb):
    from excel_mcp_2013.tools.introspect import _search_workbook
    print("\n--- techo regex ---")
    ws = wb.Worksheets.Add()
    ws.Name = "Enorme"
    ws.Range("A1").Value = 1
    ws.Range("ZZ50000").Value = 1  # UsedRange 702x50000 = 35,1M celdas
    try:
        _search_workbook(sm, "1", "values", True, False, False, "Enorme", 500)
        check("techo regex lanza error", False)
    except ValueError as e:
        check("techo regex lanza error accionable",
              "regex=False" in str(e) and "Enorme" in str(e), str(e)[:110])
    r = _search_workbook(sm, "1", "values", False, False, False, "Enorme", 5)
    check("motor nativo sigue funcionando sin techo", r["count"] >= 1, r["count"])
    app = ws.Application
    app.DisplayAlerts = False
    ws.Delete()
    app.DisplayAlerts = True


def test_names(sm):
    from excel_mcp_2013.tools.introspect import _clean_defined_names, _list_defined_names
    print("\n--- list/clean defined_names ---")

    r = _list_defined_names(sm, False)
    por_nombre = {n["name"]: n for n in r["names"]}
    check("list: encuentra los 4 nombres", r["total"] >= 4, f"total={r['total']}")
    check("list: Area_Buena sana",
          por_nombre.get("Area_Buena", {}).get("broken") is False, str(por_nombre.get("Area_Buena")))
    check("list: Area_Rota marcada rota",
          por_nombre.get("Area_Rota", {}).get("broken") is True, str(por_nombre.get("Area_Rota")))
    check("list: Print_Area marcada builtin",
          por_nombre.get("Print_Area", {}).get("builtin") is True, str(por_nombre.get("Print_Area")))
    check("list: scope de hoja detectado",
          por_nombre.get("Local_Hoja", {}).get("scope") == "sheet:Datos",
          str(por_nombre.get("Local_Hoja")))
    check("list: conteos de scope cuadran",
          r["workbook_scoped"] + r["sheet_scoped"] == r["total"], str(r)[:100])
    check("list: broken>=1", r["broken"] >= 1, r["broken"])
    check("list: no expone 'raw'", "raw" not in r["names"][0], str(r["names"][0]))

    rb = _list_defined_names(sm, True)
    check("broken_only filtra la muestra",
          all(n["broken"] for n in rb["names"]) and len(rb["names"]) == rb["broken"],
          f"{len(rb['names'])} de {rb['broken']}")
    check("broken_only conserva conteos del universo", rb["total"] == r["total"],
          f"{rb['total']} vs {r['total']}")

    # --- clean: solo rotos, preservando built-in ---
    total_antes = r["total"]
    c = _clean_defined_names(sm, True, False, False)
    check("clean: borro el roto", c["removed"] == 1, str(c))
    check("clean: saved=False", c["saved"] is False)
    despues = _list_defined_names(sm, False)
    nombres = {n["name"] for n in despues["names"]}
    check("clean: Area_Rota desaparecio", "Area_Rota" not in nombres, str(sorted(nombres)))
    check("clean: Area_Buena preservada", "Area_Buena" in nombres)
    check("clean: Print_Area preservada", "Print_Area" in nombres)
    check("clean: total baja en 1", despues["total"] == total_antes - 1,
          f"{despues['total']} vs {total_antes}")

    # --- clean: todos menos built-in ---
    c2 = _clean_defined_names(sm, False, False, False)
    check("clean todos: borra los no-builtin", c2["removed"] >= 2, str(c2))
    final = _list_defined_names(sm, False)
    nombres_f = {n["name"] for n in final["names"]}
    check("clean todos: Print_Area SIGUE ahi", "Print_Area" in nombres_f, str(sorted(nombres_f)))
    check("clean todos: Area_Buena borrada", "Area_Buena" not in nombres_f)
    check("clean todos: remaining coincide", c2["remaining"] >= 1, str(c2))


def test_names_readonly_y_save(sm, tmpdir):
    from excel_mcp_2013.tools.introspect import _clean_defined_names, _list_defined_names
    print("\n--- clean: read-only y save ---")
    app = sm.get_application()
    path = os.path.join(tmpdir, "nombres.xlsx")
    wb2 = app.Workbooks.Add()
    ws = wb2.Worksheets(1)
    ws.Name = "H"
    ws.Range("A1").Value = 1
    wb2.Names.Add(Name="Buena_2", RefersTo="=H!$A$1")
    tmp_ws = wb2.Worksheets.Add()
    tmp_ws.Name = "Borrar"
    wb2.Names.Add(Name="Rota_2", RefersTo="=Borrar!$A$1")
    app.DisplayAlerts = False
    tmp_ws.Delete()
    app.DisplayAlerts = True
    wb2.SaveAs(path, FileFormat=51)
    wb2.Close(SaveChanges=False)

    # read_only -> error claro, sin borrar nada
    wb_ro = sm.open_workbook(path, read_only=True)
    try:
        _clean_defined_names(sm, True, False, False)
        check("clean rechaza read-only", False)
    except RuntimeError as e:
        check("clean rechaza read-only", "solo lectura" in str(e), str(e)[:70])
    wb_ro.Close(SaveChanges=False)

    # save=False no toca el archivo en disco
    mtime_antes = os.path.getmtime(path)
    wb_rw = sm.open_workbook(path, read_only=False)
    lista = _list_defined_names(sm, True)
    check("archivo guardado conserva el nombre roto", lista["broken"] == 1, str(lista)[:120])
    c = _clean_defined_names(sm, True, False, False)
    check("clean en archivo real borra 1", c["removed"] == 1, str(c))
    time.sleep(0.2)
    check("save=False NO escribe a disco", os.path.getmtime(path) == mtime_antes,
          f"{mtime_antes} -> {os.path.getmtime(path)}")
    # save=True si persiste
    c2 = _clean_defined_names(sm, False, False, True)
    check("save=True guarda", c2["saved"] is True, str(c2))
    check("save=True actualiza mtime", os.path.getmtime(path) > mtime_antes)
    wb_rw.Close(SaveChanges=False)


def test_connections(sm, wb, tmpdir):
    from excel_mcp_2013.tools.introspect import _connection_info, _list_connections
    print("\n--- list_connections ---")
    wb.Activate()
    r = _list_connections(sm)
    check("libro sin conexiones -> count 0", r["count"] == 0, str(r)[:120])
    check("LinkSources None -> lista vacia", r["excel_links"] == [], str(r["excel_links"]))

    # Conexion real de texto (QueryTable): valida el mapeo de XlConnectionType
    csv_path = os.path.join(tmpdir, "origen.csv")
    with open(csv_path, "w", encoding="utf-8") as f:
        f.write("a,b\n1,2\n")
    ws = wb.Worksheets.Add()
    ws.Name = "Externo"
    try:
        qt = ws.QueryTables.Add(Connection="TEXT;" + csv_path, Destination=ws.Range("A1"))
        qt.TextFileParseType = 1        # xlDelimited
        qt.TextFileCommaDelimiter = True
        qt.BackgroundQuery = False      # Excel 2013: refresh SIEMPRE sincrono
        qt.Refresh(BackgroundQuery=False)
        r2 = _list_connections(sm)
        check("conexion de texto listada", r2["count"] == 1, str(r2)[:200])
        conn = r2["connections"][0] if r2["connections"] else {}
        check("conexion trae nombre", bool(conn.get("name")), str(conn)[:120])
        check("conexion sin cadena cruda",
              not any("password" in str(k).lower() for k in conn), str(conn)[:120])
        check("conexion trae reachable", "reachable" in conn, str(conn)[:120])
    except Exception as e:
        check("conexion de texto creable", False, f"{type(e).__name__}: {str(e)[:90]}")

    # _connection_info tolera un objeto que no expone nada (tipo MODEL/WORKSHEET)
    class ConnMuda:
        Name = "Muda"

        def __getattr__(self, item):
            raise AttributeError(item)

    info = _connection_info(ConnMuda())
    check("_connection_info tolera conexion sin OLEDBConnection",
          info["name"] == "Muda" and info["reachable"] == "n/a", str(info)[:120])


def main():
    test_puras()  # no necesitan Excel

    pids_before = snapshot_excel_pids()
    pythoncom.CoInitialize()
    tmpdir = tempfile.mkdtemp(prefix="introspect_test_")
    try:
        from excel_mcp_2013.session import SessionManager
        sm = SessionManager(visible=False)
        sm.start()
        wb = build_fixture(sm)
        check("fixture creada", wb.Worksheets.Count == 3, f"{wb.Worksheets.Count} hojas")

        test_search(sm)
        test_error_cells(sm)
        test_regex_ceiling(sm, wb)
        test_names(sm)
        test_connections(sm, wb, tmpdir)
        test_names_readonly_y_save(sm, tmpdir)

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
    print("[OK] Suite introspeccion completa")
    return 0


if __name__ == "__main__":
    sys.exit(main())
