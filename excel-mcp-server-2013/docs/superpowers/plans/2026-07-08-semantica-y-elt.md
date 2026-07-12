# Bloques A (Semántica) + B (ELT) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans para ejecutar
> tarea por tarea. Checkboxes `- [ ]` para tracking. Spec: `PROPUESTAS_MEJORAS.md`.

**Goal:** 10 tools nuevas (45 total): comprensión semántica de libros (perfil de fórmulas
R1C1, trace de celda, guardián 2013, dependencias, VBA, expediente) + kit ELT (medidas DAX,
queries PQ, CUBEVALUE, macro refresh) + piloto ida-vuelta 2021→2013.

**Architecture:** Mismos patrones del repo: módulos en `tools/` con `register(mcp, session, run)`,
impls como funciones planas `_impl(session, ...)` ejecutadas en el STA thread vía `run`.
Verificación E2E con `fastmcp.Client` + `StdioTransport` (no hay git ni pytest en este repo).

**Tech Stack:** Python 3.12, FastMCP 3.4.3, pywin32 (late binding SIEMPRE — ver CLAUDE.md), re.

**Convenciones obligatorias (de CLAUDE.md):**
- Nunca `Range().Resize(r,c)` → usar `utils.excel_utils.target_range`.
- Todo COM dentro de funciones pasadas a `run` (STA affinity).
- NUNCA `Application.Run` desde el STA (deadlock probado) — `setup_refresh_macro` solo INYECTA, no ejecuta.
- Logging a stderr; verificación final siempre: `tasklist | findstr EXCEL` sin procesos nuevos.

---

### Task 1: `utils/compat_2013.py` — datos del guardián 2013

**Files:** Create `src/excel_mcp_2013/utils/compat_2013.py`

- [ ] Diccionario `MODERN_FUNCTIONS: dict[str, str]` (función → versión que la introdujo):
  2016: TEXTJOIN, CONCAT, IFS, SWITCH, MAXIFS, MINIFS, FORECAST.ETS;
  2019/365: XLOOKUP, XMATCH, FILTER, SORT, SORTBY, UNIQUE, SEQUENCE, RANDARRAY, LET;
  365: LAMBDA, MAP, REDUCE, SCAN, BYROW, BYCOL, MAKEARRAY, TEXTSPLIT, TEXTBEFORE,
  TEXTAFTER, VSTACK, HSTACK, TOCOL, TOROW, WRAPROWS, WRAPCOLS, TAKE, DROP, CHOOSEROWS,
  CHOOSECOLS, EXPAND, GROUPBY, PIVOTBY, IMAGE.
- [ ] `SPILL_REF_RE = re.compile(r"\b[A-Z]{1,3}\$?\d+#")` (referencias de derrame `A1#`).
- [ ] `FUNC_RES: dict[str, re.Pattern]` precompilados `\b<FN>\s*\(` (escapar el punto de FORECAST.ETS).

### Task 2: `tools/semantics.py` — profile_formulas

**Files:** Create `src/excel_mcp_2013/tools/semantics.py`; Modify `server.py` (registro al final, Task 8)

- [ ] `profile_formulas(sheet: str, max_patterns: int = 200)`:
  leer `ur = ws.UsedRange`; `f = ur.FormulaR1C1` (tuple 2D; una sola llamada COM);
  `base_row, base_col = ur.Row, ur.Column`; agrupar celdas cuyo texto empieza con `=`
  por su string R1C1 en `dict[str, {count, min/max row/col, first_cell}]`.
- [ ] Por patrón devolver: `formula_r1c1`, `example_cell` (dirección A1 de la primera),
  `example_a1` (leer `ws.Range(example_cell).Formula` — 1 llamada por patrón),
  `count`, `covers` (`"D5:D3594"` aproximado con min/max), `columns` (letras).
- [ ] Ordenar por count desc, truncar a `max_patterns`, devolver también
  `{total_formula_cells, unique_patterns, truncated}`.
- [ ] Helper `_col_letter(n)` local (26-ario).
- [ ] Verificar: sobre `'Seguimiento OC filiales'` del archivo real debe devolver <100 patrones
  para 114.898 fórmulas.

### Task 3: `tools/semantics.py` — trace_cell

- [ ] `trace_cell(sheet: str, cell: str, max_precedents: int = 25)`:
  `rng = ws.Range(cell)` → `{formula, formula_local, value, number_format}`.
- [ ] Precedentes por PARSEO (no `DirectPrecedents`, que no cruza hojas):
  `REF_RE = re.compile(r"(?:'([^']+)'|([A-Za-z0-9_. ]+))?!?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?")`
  — más simple y robusto: tokenizar con dos regex separados:
  `SHEET_REF = r"(?:'[^']+'|[\w.]+)!\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?"` y
  `BARE_REF = r"(?<![\w!:$])\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?(?![\w(])"`.
  Filtrar falsos positivos: token seguido de `(` es función, no ref.
- [ ] Nombres definidos: tokens `[A-Za-z_][\w.]*` de la fórmula ∩ `wb.Names` → resolver `RefersTo`.
- [ ] Por cada ref (límite `max_precedents`): `{ref, sheet, values}` con values solo si el
  rango es ≤3x3 (leer `.Value`); si es mayor: `{rows, cols}` y esquina.
- [ ] Dependientes: `try: rng.DirectDependents` (limitación: misma hoja; capturar error si no hay)
  → direcciones. Añadir `"note": "dependientes solo de la misma hoja (limitacion COM)"`.
- [ ] Verificar: sobre una celda con fórmula del archivo real devuelve precedentes cross-hoja.

### Task 4: `tools/semantics.py` — check_2013_compatibility

- [ ] `check_2013_compatibility()`: por cada hoja, leer `ur.Formula` UNA vez; concatenar solo
  celdas string que empiecen `=`; correr `FUNC_RES` y `SPILL_REF_RE` por hoja acumulando
  `{function, introduced, sheet, count, example_cell}` (primera celda encontrada por función).
- [ ] Nombres rotos: `wb.Names` donde `.Name` empiece `_xlfn.` o `.RefersTo` contenga `#NAME?`.
- [ ] Power Query: si `wb.Queries` accesible, correr `validate_m_expression` sobre cada `.Formula`.
- [ ] Devolver `{compatible: bool, findings: [...], broken_names: [...], pq_issues: [...],
  verdict: "El libro {no} es seguro para Excel 2013"}`.
- [ ] Verificar: en el archivo real debe reportar los nombres `_xlfn.SUMIFS`, `_xlfn.IFERROR`,
  `_xleta.N` y cero funciones modernas (fue hecho EN 2013).

### Task 5: `tools/semantics.py` — map_dependencies

- [ ] `map_dependencies()`: por hoja, extraer del `ur.Formula` los nombres de hoja referenciados
  (`re.findall(r"'([^']+)'!|(\b[\w.]+)!", ...)` filtrando los que existan en
  `{ws.Name for ws in wb.Worksheets}`); sumar aristas de pivots (`pt.PivotCache().SourceData`
  parsea `Hoja!Rango` o R1C1 `'INV'!F7C2...` → tomar texto antes de `!`).
- [ ] Clasificar cada hoja: `entrada` (0 fórmulas y ≥1 dependiente), `calculo` (fórmulas>0),
  `salida` (tiene pivots y pocas fórmulas), `estatica` (sin fórmulas ni dependientes).
- [ ] Devolver `{edges: [{from, to, via: "formulas|pivot"}], classification: {hoja: tipo},
  hidden_sheets: [...]}`.

### Task 6: `tools/vba.py` — analyze_vba_project

**Files:** Modify `src/excel_mcp_2013/tools/vba.py`

- [ ] Tool `analyze_vba_project()`: recorrer `vbp.VBComponents`, extraer código completo.
- [ ] Regex por módulo: procedimientos `^\s*(?:Public |Private )?(Sub|Function)\s+(\w+)`,
  llamadas (identificadores que coinciden con nombres de otros procedimientos, con
  `\b<name>\b` fuera de comentarios `'`), hojas tocadas
  (`(?:Worksheets|Sheets)\(\s*"([^"]+)"`), rangos (`Range\(\s*"([^"]+)"`),
  eventos (nombre empieza `Workbook_` o `Worksheet_`).
- [ ] Devolver `[{module, type, procedures: [{name, kind, is_event, calls, sheets, ranges}]}]`
  + `call_graph: [{caller: "Mod.Proc", callee: "Mod2.Proc2"}]`.
- [ ] Verificar contra el archivo real: Módulo2/3/5 tienen código; debe listar sus Subs.

### Task 7: `tools/documentation.py` — document_workbook

**Files:** Create `src/excel_mcp_2013/tools/documentation.py`

- [ ] `document_workbook(output_path: Optional[str] = None)`: ejecuta EN UNA sola pasada STA
  (una única función interna que llama a las impls planas): `workbook._analyze_workbook`,
  `semantics._profile_formulas` (por hoja con fórmulas, max_patterns=40),
  `semantics._map_dependencies`, `semantics._check_2013_compatibility`,
  `pivots._list_pivot_tables`, `power_query._list_power_queries`, y
  `vba._analyze_vba_project` (envuelto en try → si no hay trust, sección "no disponible").
- [ ] Renderizar markdown (secciones: Resumen, Hojas, Dependencias, Fórmulas por hoja,
  Tablas dinámicas, Power Query, VBA, Compatibilidad 2013) — f-strings simples.
- [ ] `output_path` default: junto al workbook → `<carpeta>/<nombre>_DOCUMENTACION.md`.
  La escritura del archivo va FUERA del guard (es I/O local, no COM).
- [ ] Devolver `{path, sections, sheet_count, warnings}`.

### Task 8: `tools/elt.py` — kit ELT (4 tools) + registro en server.py

**Files:** Create `src/excel_mcp_2013/tools/elt.py`; Modify `src/excel_mcp_2013/server.py`

- [ ] `add_data_model_measure(table_name, measure_name, dax, number_format="general")`:
  `model.ModelMeasures.Add(measure_name, model.ModelTables(table_name), dax, fmt)` donde
  `fmt = model.ModelFormatGeneral` (mapear "general|decimal|whole|percentage|currency" a
  `ModelFormat*`). Si `ModelMeasures` no existe (host 2013) → RuntimeError explicando que
  la medida debe crearse en el host 2021 y viaja en el archivo.
- [ ] `add_power_query(query_name, m_code, load_to="connection_only", target_sheet=None)`:
  1. `validate_m_expression(m_code)` → si `blocked_found`, ValueError con alternativas.
  2. `wb.Queries.Add(query_name, m_code)` (host 2016+; capturar → RuntimeError "API Queries
     no existe en 2013").
  3. `load_to="sheet"`: `qt = ws.QueryTables.Add(Connection="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=<name>", Destination=ws.Range("A1"), Sql="SELECT * FROM [<name>]")`;
     `qt.RefreshStyle=1; qt.Refresh(False)`. (QueryTables.Add es la API vieja confiable;
     NO usar ListObjects.Add(xlSrcExternal) — falla cross-process como xlSrcModel.)
  4. `load_to="data_model"`: `wb.Connections.Add2(f"Query - {name}", desc, <conn_str>, f'"{name}"', 6, True, False)`.
- [ ] `write_cube_formulas(sheet, start_cell, model_name="ThisWorkbookDataModel", title, rows: [{caption, member}], values: [{caption, measure}])`:
  escribe encabezados con `write` normal y fórmulas `=CUBEMEMBER(...)/=CUBEVALUE(...)`
  vía `.Formula` celda a celda (pocas celdas). CUBEVALUE referencia la celda del
  CUBEMEMBER (`=CUBEVALUE("<model>",$A5,B$4)` patrón clásico: fila = member, columna =
  celda con `=CUBEMEMBER("<model>","[Measures].[X]")`).
- [ ] `setup_refresh_macro(macro_name="ActualizarTodo")`: usa `vba._inject_vba_code` con
  plantilla VBA: recorre `ThisWorkbook.Connections` con `BackgroundQuery=False` + `.Refresh`,
  luego `Model.Refresh` (en `On Error Resume Next` por si no hay modelo), luego
  `PivotTable.RefreshTable` de todas las hojas. SOLO inyecta — la ejecución es del usuario
  (F5/botón) o `execute_vba_macro` en libro abierto con `enable_macros=True`.
- [ ] En `server.py`: importar y registrar `semantics`, `documentation`, `elt`.

### Task 9: E2E v5 — semántica contra el archivo real (read-only)

**Files:** Create `scratchpad/e2e_test_v5.py`

- [ ] Abrir `0607 Seguimiento OC F.xlsm` read-only sin macros; llamar:
  `profile_formulas("Seguimiento OC filiales")` (esperar <100 patrones),
  `trace_cell` sobre el example_cell del patrón más frecuente,
  `check_2013_compatibility` (esperar 3 nombres rotos, 0 funciones modernas),
  `map_dependencies` (esperar arista INV→TD),
  `analyze_vba_project` (esperar Subs en Módulo2/3/5),
  `document_workbook` → verificar que el .md existe y tiene las 8 secciones.
- [ ] Cerrar todo; `tasklist | findstr EXCEL` sin procesos nuevos.

### Task 10: E2E v6 — piloto ELT ida-vuelta

**Files:** Create `scratchpad/e2e_test_v6.py`; Produce `c:\Users\muril\Desktop\MCP Excel\PILOTO_IDA_VUELTA.xlsm`

- [ ] Crear libro nuevo (script make + open): hoja `Datos` con tabla 6x3 (Filial/Producto/Cajas).
- [ ] `add_table_to_data_model("Datos","A1:C6","VentasPiloto")` (guardar antes).
- [ ] `add_data_model_measure("VentasPiloto","Total Cajas","SUM(VentasPiloto[Cajas])")`.
- [ ] `add_power_query("QPiloto", "let Origen = Excel.CurrentWorkbook(){[Name=\"VentasPiloto\"]}[Content] in Origen", load_to="sheet", target_sheet="PQ")` — M 100% legacy-safe.
- [ ] `write_cube_formulas("Dashboard","B2",...)` con 3 filiales + medida Total Cajas.
- [ ] `setup_refresh_macro()` (libro abierto con enable_macros=True para poder inyectar y guardar .xlsm).
- [ ] `save_workbook(r"c:\Users\muril\Desktop\MCP Excel\PILOTO_IDA_VUELTA.xlsm")`,
  `check_2013_compatibility` sobre él (debe salir limpio), cerrar sin zombies.
- [ ] Instrucción al usuario: llevar el .xlsm al Excel 2013 del trabajo, abrir, habilitar
  macros, correr `ActualizarTodo`, verificar PQ + CUBEVALUE + modelo.

### Task 11: Documentación

**Files:** Modify `TOOLS.md`, `MANUAL.md`, `README.md`; republish Artifact (mismo file path)

- [ ] TOOLS.md: título "(45)", secciones nuevas "Comprensión semántica" y "ELT".
- [ ] MANUAL.md: caso de uso F (documentar herramienta) y G (piloto ELT); flujo actualizado
  con `check_2013_compatibility` como paso obligatorio pre-entrega.
- [ ] README.md: contador y comment de desglose.
- [ ] Artifact: actualizar contadores, añadir sección ELT + guardián 2013; mismo path/favicon.
- [ ] Memoria: actualizar `excel-mcp-2013-project.md` (45 tools, piloto pendiente de validar en el trabajo).

---

## Self-review

- Cobertura spec: A1→T2, A2→T3, A3→T4(+T1), A4→T5, A5→T6, A6→T7, B5→T8, B7 piloto→T10,
  B6 (PBI) es guía de uso (MANUAL, T11), sin código. ✓
- Sin placeholders: cada task tiene API exacta y valores concretos. ✓
- Consistencia de nombres: impls planas `_snake_case` importables entre módulos
  (documentation importa de semantics/vba/pivots/power_query/workbook). ✓
- Riesgo conocido: `wb.Queries.Add` / `ModelMeasures.Add` / `QueryTables.Add` con Mashup
  provider cross-process no están probados — T10 los valida y si fallan se documenta
  el fallback (crear en UI) sin bloquear el resto. ✓
