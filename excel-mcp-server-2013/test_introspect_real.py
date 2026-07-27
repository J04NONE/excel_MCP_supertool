# -*- coding: utf-8 -*-
"""E2E de introspeccion contra los archivos REALES del usuario (spec v1.5.0).

- MULTIFORMATO (.xlsb, 12 MB): list_connections debe mapear la conexion SSAS de
  intranet SIN colgarse, y search_workbook debe ubicar un codigo PT.
- Estimados Sebas.xlsx: valida la interaccion documentada con el auto-saneo
  (abierto = pocos rotos; el conteo real del ORIGINAL sale de scan_definedname_risk,
  funcion pura offline que NO abre el archivo).

Uso:  .venv/Scripts/python.exe test_introspect_real.py
"""
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

import psutil
import pythoncom

BASE = r"c:\Users\muril\Desktop\MCP Excel"
NUTRI = os.path.join(BASE, "Informe ventas Nutribella paises")
MULTIFORMATO = os.path.join(NUTRI, "MULTIFORMATO EVOLUCIÓN DE MARCAS JULIO (1)-1.xlsb")
ESTIMADOS = os.path.join(NUTRI, "Estimados Sebas.xlsx")

FAILED = []


def check(name, cond, extra=""):
    tag = "[OK]" if cond else "[ERR]"
    print(f"{tag} {name}" + (f" -> {extra}" if extra else ""), flush=True)
    if not cond:
        FAILED.append(name)


def snapshot_excel_pids():
    return {p.pid for p in psutil.process_iter(["name"]) if p.info["name"] == "EXCEL.EXE"}


def main():
    from excel_mcp_2013.tools.introspect import _list_connections, _list_defined_names, _search_workbook
    from excel_mcp_2013.utils.workbook_sanitize import scan_definedname_risk

    # --- offline: el ORIGINAL de Estimados sin abrirlo (no arriesga cuelgue) ---
    print("--- scan offline (sin COM) ---")
    riesgo = scan_definedname_risk(ESTIMADOS)
    check("Estimados: riesgo detectado offline", riesgo["risky"] is True, str(riesgo))
    check("Estimados: ~1886 nombres rotos en el ORIGINAL",
          riesgo["definednames_broken"] > 1000, riesgo["definednames_broken"])

    pids_before = snapshot_excel_pids()
    pythoncom.CoInitialize()
    try:
        from excel_mcp_2013.session import SessionManager
        sm = SessionManager(visible=False)
        sm.start()

        # --- MULTIFORMATO: conexiones de intranet ---
        print("\n--- MULTIFORMATO (.xlsb, 12 MB) ---")
        t0 = time.time()
        wb = sm.open_workbook(MULTIFORMATO, read_only=True)
        print(f"    abierto en {time.time() - t0:.1f}s ({wb.Worksheets.Count} hojas)")
        # Workbooks.Open RETORNA antes de que Excel termine de cargar el Data
        # Model: la PRIMERA llamada COM posterior se queda esperando 18-35s.
        # No es coste del tool — se mide aparte para no atribuirselo.
        t0 = time.time()
        _ = wb.Connections.Count
        print(f"    1a llamada COM tras open: {time.time() - t0:.1f}s "
              "(Excel terminando de cargar el modelo)")

        t0 = time.time()
        conns = _list_connections(sm)
        dur = time.time() - t0
        check("list_connections es rapido", dur < 5, f"{dur:.1f}s")
        check("list_connections encuentra conexiones", conns["count"] > 0, conns["count"])
        for c in conns["connections"]:
            print(f"    - {c.get('name')} | {c.get('type')} | prov={c.get('provider')} "
                  f"| src={c.get('data_source')} | cat={c.get('initial_catalog')} "
                  f"| {c.get('reachable')} | odc={c.get('odc_file')}")
        olap = [c for c in conns["connections"] if c.get("type") == "OLAP/Cubo (SSAS)"]
        check("conexion SSAS identificada", len(olap) > 0, str(conns["connections"])[:200])
        if olap:
            c = olap[0]
            check("SSAS: provider MSOLAP", str(c["provider"]).startswith("MSOLAP"), c["provider"])
            check("SSAS: no_verificable (servidor)",
                  c["reachable"] == "no_verificable (servidor)", c["reachable"])
            check("SSAS: sin credenciales expuestas", c["had_credentials"] is False,
                  str(c)[:150])
        check("ninguna conexion filtra la cadena cruda",
              not any("password" in str(c).lower() for c in conns["connections"]))
        print(f"    vinculos Excel->Excel: {conns['excel_link_count']}")

        t0 = time.time()
        pt = _search_workbook(sm, "5060094", "values", False, False, True, None, 20)
        dur = time.time() - t0
        check("search en libro real responde rapido", dur < 60, f"{dur:.1f}s")
        print(f"    '5060094': {pt['count']} matches en {dur:.1f}s "
              f"-> {[(m['sheet'], m['cell']) for m in pt['matches'][:5]]}")
        gd = _search_workbook(sm, "#GETTING_DATA", "values", False, False, False, None, 20)
        print(f"    '#GETTING_DATA': {gd['count']} matches "
              f"-> {[(m['sheet'], m['cell']) for m in gd['matches'][:5]]}")
        check("search sobre libro con celdas de cubo no rompe", isinstance(gd["count"], int))
        # No se cierra a mano: sm.close() cierra los libros que rastrea (cerrarlos
        # aqui deja proxies COM muertos y ruido de 'objeto desconectado').

        # --- Estimados Sebas: interaccion con el auto-saneo ---
        print("\n--- Estimados Sebas.xlsx (auto-saneado) ---")
        t0 = time.time()
        wb2 = sm.open_workbook(ESTIMADOS)
        info = sm.last_open_info()
        print(f"    abierto en {time.time() - t0:.1f}s; saneado={info.get('sanitized')} "
              f"({info.get('definednames_removed')} removidos)")
        check("open aplico el saneo", info.get("sanitized") is True, str(info)[:120])
        t0 = time.time()
        nombres = _list_defined_names(sm, False)
        dur = time.time() - t0
        print(f"    total={nombres['total']} broken={nombres['broken']} "
              f"hidden={nombres['hidden']} builtin={nombres['builtin']} en {dur:.1f}s")
        check("list_defined_names tolerable con ~3000 nombres", dur < 30, f"{dur:.1f}s")
        check("interaccion documentada: la copia saneada ya no tiene rotos",
              nombres["broken"] <= 5, nombres["broken"])
        check("los nombres sanos SI sobreviven al saneo", nombres["total"] > 0,
              nombres["total"])
        check("conteo coherente con el scan offline",
              nombres["total"] + riesgo["definednames_broken"] >= riesgo["definednames_total"],
              f"{nombres['total']} + {riesgo['definednames_broken']} vs "
              f"{riesgo['definednames_total']}")

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
    print("[OK] E2E introspeccion contra archivos reales")
    return 0


if __name__ == "__main__":
    sys.exit(main())
