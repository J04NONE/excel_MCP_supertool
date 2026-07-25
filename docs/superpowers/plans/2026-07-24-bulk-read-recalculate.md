# Plan: Lectura masiva + recalculate + diagnóstico Data Model (v1.4.0)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implementar el paquete #3+#4 del spec [2026-07-24-bulk-read-recalculate-design.md](../specs/2026-07-24-bulk-read-recalculate-design.md): `recalculate`, `read_table`, `export_sheet` (con range_addr/tsv), tope en `read_range` y `get_data_model_measures` honesto.

**Architecture:** Tools COM sobre el patrón existente `register(mcp, session, run)`. Lectura masiva en módulo nuevo `bulk.py` con fast path `.Value` (1 llamada COM por bloque, tamaño verificado ANTES de leer). Tests con Excel real estilo `test_hardening.py` (script standalone con `check()`, sin pytest).

**Tech Stack:** Python 3.12 + pywin32 (late binding), FastMCP. Todo API COM compatible Excel 2013.

**⛔ REGLA DURA DE ESTE REPO:** PROHIBIDO ejecutar `git commit`/`git push`/`git add`. Los pasos de commit de este plan significan: **mostrar a Michael el comando y mensaje sugerido y seguir**. Nunca ejecutarlos.

**Rutas base** (abreviadas en el plan):
- `SRC` = `excel-mcp-server-2013/src/excel_mcp_2013`
- Python del repo: `excel-mcp-server-2013/.venv/Scripts/python.exe`
- Tests corren DESDE `excel-mcp-server-2013/` (asumen `src/` importable vía `sys.path.insert`)

**Mapa de archivos:**

| Archivo | Acción | Responsabilidad |
|---|---|---|
| `SRC/tools/bulk.py` | Crear | read_table + export_sheet + helpers + constantes |
| `SRC/tools/workbook.py` | Modificar | tool recalculate |
| `SRC/tools/cells.py` | Modificar | tope 50k en read_range |
| `SRC/tools/power_pivot.py` | Modificar | get_data_model_measures → dict con diagnostic |
| `SRC/server.py` | Modificar | registrar bulk; versión 1.4.0 |
| `excel-mcp-server-2013/test_bulk_tools.py` | Crear | suite del paquete (Excel real) |
| `TOOLS.md`, `LIMITACIONES_MCP.md` | Modificar | docs y cierre |

---

### Task 1: Harness de test + fixture COM

**Files:**
- Create: `excel-mcp-server-2013/test_bulk_tools.py`

- [ ] **Step 1.1: Crear el harness con fixture (aún sin tests de tools)**

```python
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

        # === los tests de tools se agregan en las tasks siguientes ===

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
```

- [ ] **Step 1.2: Correr el harness**

Run: `cd excel-mcp-server-2013 && ./.venv/Scripts/python.exe test_bulk_tools.py`
Expected: `[OK] fixture creada -> 3 hojas` y `[OK] Suite bulk completa`, exit 0, sin EXCEL.EXE huérfanos.

- [ ] **Step 1.3: Sugerir commit a Michael (NO ejecutar)**

```text
test(bulk): harness + fixture COM para el paquete lectura masiva
```

---

### Task 2: `recalculate` en workbook.py (TDD)

**Files:**
- Modify: `excel-mcp-server-2013/test_bulk_tools.py` (añadir test)
- Modify: `SRC/tools/workbook.py`

- [ ] **Step 2.1: Añadir el test (falla: `_recalculate` no existe)**

En `test_bulk_tools.py`, después de `check("fixture creada", ...)`:

```python
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
```

- [ ] **Step 2.2: Correr y verificar que FALLA**

Run: `./.venv/Scripts/python.exe test_bulk_tools.py`
Expected: `ImportError: cannot import name '_recalculate'`

- [ ] **Step 2.3: Implementar en `SRC/tools/workbook.py`**

Ajustar imports de la cabecera:

```python
from ..utils.excel_utils import (
    LONG_OP_TIMEOUT,
    XL_CALCULATION_AUTOMATIC,
    XL_CELL_TYPE_FORMULAS,
    XL_FILE_FORMATS,
    get_active_workbook,
    get_sheet,
)
```

**Nota:** `XL_CALCULATION_AUTOMATIC = -4105` ya existe en `utils/excel_utils.py:10`.

Dentro de `register(mcp, session, run)` (junto a los otros tools):

```python
    @mcp.tool()
    def recalculate(full: bool = False, sheet: Optional[str] = None,
                    wait_async: bool = False) -> dict:
        """Recalcular el workbook activo (el MCP fuerza calculo MANUAL al abrir;
        sin esto las formulas nuevas no evaluan hasta guardar/reabrir).

        default: solo celdas sucias. full=True: todo desde cero.
        sheet='X': solo esa hoja (ignora full). wait_async=True: secuencia para
        formulas CUBE (#GETTING_DATA): pone calculo automatico temporal +
        CalculateUntilAsyncQueriesDone, restaura Manual al final (ignora sheet).
        Puede tardar minutos en libros grandes (timeout 600s)."""
        return run(_recalculate, session, full, sheet, wait_async,
                   timeout=LONG_OP_TIMEOUT)
```

Al final del archivo (zona de impls):

```python
def _recalculate(session, full: bool, sheet, wait_async: bool) -> dict:
    app = session.get_application()
    wb = get_active_workbook(session)  # error claro si no hay libro abierto
    try:
        app.ScreenUpdating = False  # blindaje: una macro previa pudo reactivarlo
    except Exception:
        pass
    if wait_async:
        # CUBE en Manual queda #GETTING_DATA aunque se espere: automatico temporal.
        prev = app.Calculation
        app.Calculation = XL_CALCULATION_AUTOMATIC
        try:
            app.CalculateFull()
            app.CalculateUntilAsyncQueriesDone()
        finally:
            app.Calculation = prev
        mode = "async_cube"
    elif sheet is not None:
        get_sheet(wb, sheet).Calculate()
        mode = f"sheet:{sheet}"
    elif full:
        app.CalculateFull()
        mode = "full"
    else:
        app.Calculate()
        mode = "dirty"
    state_map = {0: "done", 1: "calculating", 2: "pending"}
    try:
        state = state_map.get(int(app.CalculationState), "unknown")
    except Exception:
        state = "unknown"
    return {"calculated": True, "mode": mode, "calculation_state": state}
```

- [ ] **Step 2.4: Correr y verificar que PASA**

Run: `./.venv/Scripts/python.exe test_bulk_tools.py`
Expected: todos `[OK]` (incluye "A3 recalculada = 13", "Calculation sigue Manual tras wait_async"), exit 0.

- [ ] **Step 2.5: Sugerir commit a Michael (NO ejecutar)**

```text
feat(workbook): tool recalculate (dirty/full/sheet/async_cube, timeout 600s)

Con Calculation=Manual forzado, las formulas no evaluaban hasta
guardar+reabrir. wait_async pone automatico temporal para CUBE
(#GETTING_DATA) y restaura Manual siempre.
```

---

### Task 3: `bulk.py` con `read_table` (TDD)

**Files:**
- Modify: `excel-mcp-server-2013/test_bulk_tools.py`
- Create: `SRC/tools/bulk.py`

- [ ] **Step 3.1: Añadir tests de read_table (fallan: módulo no existe)**

En `test_bulk_tools.py`, después del bloque recalculate:

```python
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
        # cap inline: agrandar la tabla a 501x10... mas barato: tabla nueva 501x100
        ws_cap = wb.Worksheets.Add()
        ws_cap.Name = "Cap"
        ws_cap.Range("A1:CV1").Value = [tuple(f"C{i}" for i in range(1, 101))]
        bloque = tuple(tuple(1 for _ in range(100)) for _ in range(501))
        ws_cap.Range("A2:CV502").Value = bloque
        lo_cap = ws_cap.ListObjects.Add(SourceType=1, Source=ws_cap.Range("A1:CV502"),
                                        XlListObjectHasHeaders=1)
        lo_cap.Name = "TablaGrande"  # 501*100 = 50.100 > 50.000
        try:
            _read_table(sm, "TablaGrande", None)
            check("cap inline lanza error", False)
        except Exception as e:
            check("cap inline lanza error", str(MAX_INLINE_CELLS) in str(e) or "dest" in str(e),
                  str(e)[:80])
        tg = _read_table(sm, "TablaGrande", os.path.join(tmpdir, "grande.csv"))
        check("cap con dest OK", tg["rows"] == 501, str(tg["rows"]))
```

- [ ] **Step 3.2: Correr y verificar que FALLA**

Run: `./.venv/Scripts/python.exe test_bulk_tools.py`
Expected: `ModuleNotFoundError: No module named 'excel_mcp_2013.tools.bulk'`

- [ ] **Step 3.3: Crear `SRC/tools/bulk.py` completo (read_table + export_sheet + helpers)**

```python
"""Tools de lectura masiva: tablas completas y export de hojas/rangos a archivo.

Fast path COM: UNA llamada .Value por bloque (nunca celda a celda). El tamano
se verifica ANTES de materializar datos (Rows.Count/Columns.Count son baratos).
"""

import csv
import json
import logging
import os
from typing import Optional

from ..utils.excel_utils import get_active_workbook, get_sheet, matrix_to_jsonable

logger = logging.getLogger(__name__)

MAX_INLINE_CELLS = 50_000
MAX_EXPORT_CELLS = 5_000_000
SAMPLE_ROWS_DEFAULT = 5


def register(mcp, session, run):
    @mcp.tool()
    def read_table(table_name: str, dest: Optional[str] = None) -> dict:
        """Leer una tabla (ListObject) completa por nombre (case-insensitive).

        Sin dest: inline (max 50.000 celdas) -> {table, sheet, headers, rows,
        row_count, col_count}; 'rows' es la MATRIZ de datos.
        Con dest (.csv/.tsv/.json): escribe el archivo (headers en la primera
        fila) -> {file, rows, cols, headers, sample}; 'rows'/'cols' son CONTEOS."""
        return run(_read_table, session, table_name, dest)

    @mcp.tool()
    def export_sheet(sheet: str, dest: str, sample_rows: int = 5,
                     range_addr: Optional[str] = None) -> dict:
        """Exportar una hoja (UsedRange) o un rango (range_addr='B5:X9000') a
        .csv/.tsv/.json en UNA llamada COM. Techo: 5M celdas.

        OJO: UsedRange puede NO empezar en A1 — usa la clave 'range' de la
        respuesta para mapear offsets. Celdas combinadas: el valor queda solo
        en la celda superior-izquierda (el resto sale null)."""
        return run(_export_sheet, session, sheet, dest, sample_rows, range_addr)


def _find_table(wb, table_name: str):
    """(worksheet, ListObject) por nombre case-insensitive, o error con listado."""
    disponibles = []
    for ws in wb.Worksheets:
        for lo in ws.ListObjects:
            disponibles.append(f"{lo.Name} (hoja {ws.Name})")
            if str(lo.Name).lower() == table_name.lower():
                return ws, lo
    listado = ", ".join(disponibles) or "ninguna"
    raise ValueError(f"Tabla '{table_name}' no encontrada. Disponibles: {listado}")


def _headers_of(lo, n_cols: int) -> list:
    try:
        if lo.ShowHeaders:
            hdr = matrix_to_jsonable(lo.HeaderRowRange.Value)
            if hdr:
                return [str(h) if h is not None else f"col{i + 1}"
                        for i, h in enumerate(hdr[0])]
    except Exception:
        logger.debug("HeaderRowRange inaccesible; headers sinteticos")
    return [f"col{i + 1}" for i in range(n_cols)]


def _ext_of(dest: str) -> str:
    ext = os.path.splitext(dest)[1].lower()
    if ext not in (".csv", ".tsv", ".json"):
        raise ValueError(f"Extension no soportada: '{ext}' (usa .csv, .tsv o .json)")
    return ext


def _write_file(dest: str, matrix: list) -> str:
    """Escribe matriz 2D JSON-safe a dest segun extension. Devuelve ruta absoluta."""
    dest = os.path.abspath(dest)
    ext = _ext_of(dest)
    if ext == ".json":
        with open(dest, "w", encoding="utf-8") as f:
            json.dump(matrix, f, ensure_ascii=False)
    else:
        delim = "," if ext == ".csv" else "\t"
        # utf-8-sig: BOM para que Excel reabra el archivo con acentos correctos
        with open(dest, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f, delimiter=delim, quoting=csv.QUOTE_MINIMAL)
            for row in matrix:
                w.writerow(["" if c is None else c for c in row])
    return dest


def _read_table(session, table_name: str, dest) -> dict:
    wb = get_active_workbook(session)
    ws, lo = _find_table(wb, table_name)
    body = lo.DataBodyRange  # None si la tabla no tiene filas de datos
    if body is None:
        n_rows, n_cols = 0, int(lo.ListColumns.Count)
    else:
        n_rows, n_cols = int(body.Rows.Count), int(body.Columns.Count)
    headers = _headers_of(lo, n_cols)
    if not dest and n_rows * n_cols > MAX_INLINE_CELLS:
        raise ValueError(
            f"Tabla '{lo.Name}' tiene {n_rows * n_cols} celdas "
            f"(> {MAX_INLINE_CELLS}): pasa dest='ruta.csv|.tsv|.json'."
        )
    rows = matrix_to_jsonable(body.Value) if body is not None else []
    if dest:
        path = _write_file(dest, [headers] + rows)
        return {"file": path, "rows": n_rows, "cols": n_cols, "headers": headers,
                "sample": rows[:SAMPLE_ROWS_DEFAULT]}
    return {"table": str(lo.Name), "sheet": str(ws.Name), "headers": headers,
            "rows": rows, "row_count": n_rows, "col_count": n_cols}


def _export_sheet(session, sheet: str, dest: str, sample_rows: int,
                  range_addr) -> dict:
    wb = get_active_workbook(session)
    ws = get_sheet(wb, sheet)
    ur = ws.Range(range_addr) if range_addr else ws.UsedRange
    n_rows, n_cols = int(ur.Rows.Count), int(ur.Columns.Count)
    if n_rows * n_cols > MAX_EXPORT_CELLS:
        raise ValueError(
            f"{n_rows * n_cols} celdas (> {MAX_EXPORT_CELLS}): exporta por partes "
            "con range_addr o divide la hoja."
        )
    values = matrix_to_jsonable(ur.Value)
    path = _write_file(dest, values)
    return {"file": path, "rows": n_rows, "cols": n_cols,
            "range": str(ur.Address), "sample": values[:max(0, int(sample_rows))]}
```

- [ ] **Step 3.4: Correr y verificar que PASA**

Run: `./.venv/Scripts/python.exe test_bulk_tools.py`
Expected: todos `[OK]` hasta "cap con dest OK". (El import de `_export_sheet` ya resuelve porque se creó en el mismo archivo.)

- [ ] **Step 3.5: Sugerir commit a Michael (NO ejecutar)**

```text
feat(bulk): read_table — ListObject completo en 2 llamadas COM

Inline hasta 50k celdas o dest .csv/.tsv/.json (archivo + muestra).
Tamano verificado ANTES de leer; bordes: tabla vacia, headers sinteticos.
```

---

### Task 4: `export_sheet` (tests de offset, tsv, range_addr)

**Files:**
- Modify: `excel-mcp-server-2013/test_bulk_tools.py`

*(La implementación ya quedó en Task 3; esta task la verifica.)*

- [ ] **Step 4.1: Añadir tests de export_sheet**

Después del bloque read_table:

```python
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
```

- [ ] **Step 4.2: Correr y verificar que PASA**

Run: `./.venv/Scripts/python.exe test_bulk_tools.py`
Expected: todos `[OK]` (offset, tsv, range_addr, sample, extensión).

- [ ] **Step 4.3: Sugerir commit a Michael (NO ejecutar)**

```text
feat(bulk): export_sheet — UsedRange o range_addr a .csv/.tsv/.json

1 llamada COM, techo 5M celdas, clave 'range' para mapear offsets
(UsedRange no siempre empieza en A1).
```

---

### Task 5: Tope 50k en `read_range` (cells.py)

**Files:**
- Modify: `excel-mcp-server-2013/test_bulk_tools.py`
- Modify: `SRC/tools/cells.py`

- [ ] **Step 5.1: Añadir test (falla: hoy no hay tope)**

```python
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
```

- [ ] **Step 5.2: Correr y verificar que FALLA**

Run: `./.venv/Scripts/python.exe test_bulk_tools.py`
Expected: `[ERR] read_range cap lanza error` (hoy lee 3.5M celdas inline o revienta).

- [ ] **Step 5.3: Implementar el tope en `SRC/tools/cells.py`**

`cells.py` YA tiene su propio helper local `_get_sheet(session, sheet)` (cells.py:63-65)
— usarlo, NO importar `get_sheet` de utils. Solo hay que añadir 1 import:

```python
from .bulk import MAX_INLINE_CELLS
```

(No hay ciclo: `bulk.py` importa solo de `..utils.excel_utils`, nunca de `cells`.)

Reemplazar `_read_range` (hoy exactamente en cells.py:68-70):

```python
def _read_range(session, sheet: str, range_addr: str) -> list:
    ws = _get_sheet(session, sheet)
    rng = ws.Range(range_addr)
    n_cells = int(rng.Rows.Count) * int(rng.Columns.Count)
    if n_cells > MAX_INLINE_CELLS:
        raise ValueError(
            f"Rango {range_addr} tiene {n_cells} celdas (> {MAX_INLINE_CELLS}). "
            "Usa export_sheet(sheet, dest, range_addr=...) para volcarlo a archivo."
        )
    return matrix_to_jsonable(rng.Value)
```

Actualizar también el docstring del tool `read_range` en `register`:

```python
    @mcp.tool()
    def read_range(sheet: str, range_addr: str) -> list:
        """Leer VALORES de un rango (ej: sheet='Hoja1', range_addr='A1:C10').
        Maximo 50.000 celdas inline; para rangos mayores usa
        export_sheet(sheet, dest, range_addr=...)."""
        return run(_read_range, session, sheet, range_addr)
```

- [ ] **Step 5.4: Correr y verificar que PASA**

Run: `./.venv/Scripts/python.exe test_bulk_tools.py`
Expected: todos `[OK]` incluidos los 2 nuevos.

- [ ] **Step 5.5: Sugerir commit a Michael (NO ejecutar)**

```text
feat(cells): tope 50k celdas en read_range (cambio de comportamiento)

Un rango gigante inline revienta el contexto del agente. El error
apunta a export_sheet(range_addr=...). Rangos chicos: sin cambios.
```

---

### Task 6: `get_data_model_measures` honesto (power_pivot.py)

**Files:**
- Modify: `excel-mcp-server-2013/test_bulk_tools.py`
- Modify: `SRC/tools/power_pivot.py`

- [ ] **Step 6.1: Añadir test (falla: hoy devuelve list)**

```python
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
```

- [ ] **Step 6.2: Correr y verificar que FALLA**

Run: `./.venv/Scripts/python.exe test_bulk_tools.py`
Expected: `[ERR] measures devuelve dict` (hoy retorna list).

- [ ] **Step 6.3: Reescribir `_get_data_model_measures` en power_pivot.py (líneas 143-171)**

```python
def _get_data_model_measures(session) -> dict:
    wb = get_active_workbook(session)
    diagnostic = {"model_present": False, "model_tables": 0}
    try:
        model = _get_model(wb)
        diagnostic["model_present"] = True
        try:
            diagnostic["model_tables"] = int(model.ModelTables.Count)
        except Exception:
            pass
    except Exception as e:
        diagnostic["note"] = f"Sin Data Model accesible: {e}"
        return {"measures": [], "diagnostic": diagnostic}

    measures = []
    try:
        # Via COM (Excel 2016+): ModelMeasures expone nombre + formula directamente
        for m in model.ModelMeasures:
            measures.append({
                "name": str(m.Name),
                "expression": str(m.Formula),
                "table": str(m.AssociatedTable.Name),
            })
    except Exception:
        logger.debug("ModelMeasures no disponible (host 2013); fallback DMV")
        try:
            # Fallback DMV (funciona tambien en 2013): medidas visibles del cubo
            result = _evaluate_dax_query(
                session,
                "SELECT [MEASURE_NAME], [EXPRESSION], [MEASURE_CAPTION] "
                "FROM $SYSTEM.MDSCHEMA_MEASURES WHERE [MEASURE_IS_VISIBLE]",
                1000,
            )
            measures = [
                {"name": r[0], "expression": r[1], "caption": r[2]}
                for r in result["rows"]
            ]
        except Exception as e:
            diagnostic["note"] = f"ModelMeasures y DMV fallaron: {e}"

    if not measures and diagnostic["model_tables"] > 0 and "note" not in diagnostic:
        diagnostic["note"] = (
            f"Modelo presente con {diagnostic['model_tables']} tablas y 0 medidas "
            "explicitas: probablemente pivots con agregaciones implicitas."
        )
    return {"measures": measures, "diagnostic": diagnostic}
```

Actualizar el docstring y tipo del tool en `register` (power_pivot.py:35):

```python
    @mcp.tool()
    def get_data_model_measures() -> dict:
        """Medidas del Data Model con su expresion DAX. Devuelve
        {measures, diagnostic}: diagnostic distingue 'sin medidas' (modelo
        presente, medidas implicitas) de 'sin modelo'."""
```

- [ ] **Step 6.4: Correr y verificar que PASA**

Run: `./.venv/Scripts/python.exe test_bulk_tools.py`
Expected: todos `[OK]`.

- [ ] **Step 6.5: Sugerir commit a Michael (NO ejecutar)**

```text
fix(power_pivot): get_data_model_measures distingue sin-medidas de sin-modelo

Antes devolvia [] mudo con modelo presente (confundio el analisis del
MULTIFORMATO). Ahora {measures, diagnostic} con model_tables y nota.
```

---

### Task 7: Cableado en server.py (registro + versión)

**Files:**
- Modify: `SRC/server.py`

- [ ] **Step 7.1: Registrar bulk y subir versión**

En el bloque de imports de tools (server.py, cerca de la línea 258):

```python
from .tools import (  # noqa: E402
    bulk,
    cells,
    discovery,
    documentation,
    elt,
    pivots,
    power_pivot,
    power_query,
    semantics,
    shapes,
    vba,
    workbook,
)
```

Junto a los registros:

```python
bulk.register(mcp, session, run_with_excel)
```

Y en el constructor `FastMCP(...)`: `version="1.3.0"` → `version="1.4.0"`.

- [ ] **Step 7.2: Verificar que el server importa y registra**

Run:
```bash
cd excel-mcp-server-2013 && ./.venv/Scripts/python.exe -c "import sys; sys.path.insert(0,'src'); from excel_mcp_2013 import server; print('version', server.mcp.version if hasattr(server.mcp,'version') else 'n/a'); print('bulk OK')"
```
Expected: imprime sin traceback (el import ejecuta todos los register). Si `mcp.version` no es atributo público, basta con que no explote.

- [ ] **Step 7.3: Compilar todo**

Run: `./.venv/Scripts/python.exe -m py_compile src/excel_mcp_2013/tools/bulk.py src/excel_mcp_2013/tools/workbook.py src/excel_mcp_2013/tools/cells.py src/excel_mcp_2013/tools/power_pivot.py src/excel_mcp_2013/server.py`
Expected: exit 0.

- [ ] **Step 7.4: Sugerir commit a Michael (NO ejecutar)**

```text
feat(server): registrar tools bulk; version 1.4.0
```

---

### Task 8: E2E contra archivos reales (Nutribella)

**Files:**
- Create: `scratchpad/_e2e_bulk.py` (temporal, se borra al final)

- [ ] **Step 8.1: Script E2E**

Crear en el scratchpad de la sesión (NO en el repo):

```python
# -*- coding: utf-8 -*-
"""E2E del paquete bulk contra el MULTIFORMATO real (solo lectura)."""
import os, sys, csv, tempfile
sys.path.insert(0, r"C:\Users\muril\Desktop\MCP Excel\excel-mcp-server-2013\src")
import pythoncom

MULTI = (r"C:\Users\muril\Desktop\MCP Excel\Informe ventas Nutribella paises"
         r"\MULTIFORMATO EVOLUCIÓN DE MARCAS JULIO (1)-1.xlsb")

pythoncom.CoInitialize()
try:
    from excel_mcp_2013.session import SessionManager
    from excel_mcp_2013.tools.bulk import _export_sheet, _read_table
    from excel_mcp_2013.tools.workbook import _recalculate
    sm = SessionManager(visible=False)
    sm.open_workbook(MULTI, read_only=True)
    tmp = tempfile.mkdtemp(prefix="e2e_bulk_")

    # 1) read_table VENTAS (tamano real desconocido: probar inline, caer a dest)
    try:
        t = _read_table(sm, "VENTAS", None)
        print("VENTAS inline:", t["row_count"], "filas x", t["col_count"], "cols")
    except ValueError as e:
        print("VENTAS excede inline (esperado si es ancha):", str(e)[:80])
        t = _read_table(sm, "VENTAS", os.path.join(tmp, "ventas.csv"))
        print("VENTAS a csv:", t["rows"], "filas x", t["cols"], "cols ->", t["file"])

    # 2) export_sheet BD completa (876k celdas)
    import time
    t0 = time.time()
    e = _export_sheet(sm, "BD", os.path.join(tmp, "bd.csv"), 3, None)
    print(f"BD: {e['rows']}x{e['cols']} range={e['range']} en {time.time()-t0:.1f}s")
    assert e["range"].startswith("$C$4"), f"range inesperado: {e['range']}"

    # 3) validar contra cifra conocida: suma Ene-2026 hoja VENTA col BC = 2885.69
    t0 = time.time()
    ev = _export_sheet(sm, "VENTA", os.path.join(tmp, "venta.csv"), 3, "BC12:BC3406")
    with open(ev["file"], encoding="utf-8-sig") as f:
        total = sum(float(r[0]) for r in csv.reader(f) if r and r[0])
    print(f"Suma Ene-2026 = {total:.2f} (esperado 2885.69) en {time.time()-t0:.1f}s")
    assert abs(total - 2885.69) < 0.01, "NO cuadra con la cifra validada"

    # 4) recalculate smoke (no rompe en libro real, read-only)
    r = _recalculate(sm, False, None, False)
    print("recalculate:", r)
    sm.close()
    print("E2E OK")
finally:
    pythoncom.CoUninitialize()
```

- [ ] **Step 8.2: Correr**

Run: `./.venv/Scripts/python.exe <scratchpad>/_e2e_bulk.py`
Expected: `Suma Ene-2026 = 2885.69`, `E2E OK`, BD exportada con `range=$C$4:$CU$9036`. Si `read_table("VENTAS")` revela que la tabla solo cubre columnas A:B (posible: `_FilterDatabase` apuntaba a A11:B2911), anotar el hallazgo en el resumen — no es fallo del tool.

- [ ] **Step 8.3: Borrar el script y temporales E2E**

Run: `rm <scratchpad>/_e2e_bulk.py` y borrar el dir temp `e2e_bulk_*`.

---

### Task 9: Regresión completa

- [ ] **Step 9.1: Correr todo el suite**

Run (desde `excel-mcp-server-2013/`):
```bash
./.venv/Scripts/python.exe test_bulk_tools.py && ./.venv/Scripts/python.exe test_sanitize.py && ./.venv/Scripts/python.exe test_guard_wedge.py && ./.venv/Scripts/python.exe test_hardening.py
```
Expected: los 4 exit 0, cero `[ERR]`, sin EXCEL.EXE huérfanos al final.

---

### Task 10: Documentación y cierre

**Files:**
- Modify: `TOOLS.md`
- Modify: `LIMITACIONES_MCP.md`

- [ ] **Step 10.1: TOOLS.md** — leer el formato existente del archivo y añadir, siguiendo ese formato, las entradas: `recalculate`, `read_table`, `export_sheet` (nuevos) y las notas de cambio en `read_range` (tope 50k) y `get_data_model_measures` (ahora devuelve `{measures, diagnostic}`).

- [ ] **Step 10.2: LIMITACIONES_MCP.md** — marcar #3 y #4 como ✅ RESUELTO en la tabla resumen y en sus secciones, con la evidencia del E2E (BD 876k celdas exportada en 1 llamada; suma Ene-2026 = 2.885,69 verificada; recalculate operativo). Mantener la nota del backlog (`refresh_all`).

- [ ] **Step 10.3: Sugerir commit final a Michael (NO ejecutar)**

```text
docs: TOOLS.md + LIMITACIONES (paquete #3+#4 resuelto, v1.4.0)
```

- [ ] **Step 10.4: Recordar a Michael** reiniciar el servidor MCP para que los 3 tools nuevos aparezcan en vivo, y que `get_data_model_measures` cambió de forma.
