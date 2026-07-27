# Changelog

## v1.5.0 — 2026-07-26

Paquete de introspección para desenmarañar libros ajenos sin salir del MCP.
Resuelve el gap de introspección de `LIMITACIONES_MCP.md` (buscar sin tool,
nombres definidos leídos a mano del XML, conexiones vistas con `unzip`).

### Nuevos tools (`tools/introspect.py`)

- **`search_workbook`**: el "grep" de Excel sobre celdas (valores y/o fórmulas).
  Motor nativo `Range.Find` (sin techo de tamaño) o regex de Python
  (techo 2M celdas/hoja). Dedupe por celda con `matched_in` (`value` /
  `formula` / `value+formula`), tope de 500 matches con `truncated`.

- **`list_defined_names`**: conteos del universo completo
  (total/rotos/ocultos/built-in, scope libro vs hoja) + muestra de 300.

- **`clean_defined_names`** (escritura): borra los nombres rotos preservando
  los built-in de Excel. `save=False` por defecto; rechaza libros en solo
  lectura y avisa cuando el libro abierto es la copia saneada temporal.

- **`list_connections`**: mapea conexiones (OLE DB, ODBC, Power Query, OLAP/SSAS,
  Data Model, texto) y vínculos Excel→Excel. **Nunca sondea la red**: las
  fuentes de servidor/UNC se marcan `no_verificable` en vez de intentar
  contactarlas. Las cadenas se devuelven parseadas y sin credenciales.

### Fixes

- **Celdas en error** (`utils/excel_utils.py`): COM entrega un `#N/A` como el
  entero `-2146826246` (VT_ERROR, scode `0x800A0000|código`). `to_jsonable` lo
  traducía como número → `read_range`, `export_sheet`, `read_table` y
  `read_formulas` devolvían basura numérica en celdas con error. Ahora se
  traduce a `#N/A`, `#REF!`, `#GETTING_DATA`, etc. (`excel_error_name`).

### Mejoras internas

- `SessionManager.is_temp_copy()`: permite a las tools de escritura detectar
  que el libro abierto es la copia saneada y no un archivo real.
- Bump de versión: `1.4.0 → 1.5.0`.

### Hallazgos verificados contra Excel real (documentados en el código)

- `Find` con `LookIn=xlFormulas` busca sobre la fórmula **localizada**
  (`=SUMA(...)` en un Excel en español); el motor regex lee `.Formula`, siempre
  en inglés. Ambos comportamientos quedan en el docstring del tool.
- `wb.Names` **ya incluye** los nombres de scope hoja (llegan como
  `Hoja!Nombre` con `Parent` = Worksheet), y los built-in llegan **sin** el
  prefijo `_xlnm.` que sí tienen en el XML.
- Borrar un built-in (`Print_Area`) lanza 1004 aunque se pida explícitamente.
- `XlConnectionType`: OLEDB=1, ODBC=2 (el orden inverso al que se suele asumir).
- **Rendimiento**: `Workbooks.Open` sobre un `.xlsb` de 12 MB con Data Model
  retorna en ~1,6 s pero la PRIMERA llamada COM posterior espera 18-35 s
  (Excel sigue cargando el modelo); ese coste no es del tool que se invoque.
  `list_connections` sobre 14 conexiones tarda 0,1 s.
- El scope de un nombre se resuelve por el string (`Hoja!Nombre` contra la
  lista de hojas), no por `nm.Parent`: preguntarle el padre a cada uno
  triplicaba el tiempo sobre libros con miles de nombres (~10 s para 3.000).

### Tests

- `test_introspect.py`: 77 checks (funciones puras sin Excel + Excel real:
  search nativo/regex/dedupe/techo, nombres, borrado seguro, read-only,
  `save=False` no escribe a disco, conexiones).
- `test_introspect_real.py`: E2E contra el MULTIFORMATO (conexión SSAS de
  intranet mapeada sin colgarse) y `Estimados Sebas.xlsx` (interacción con el
  auto-saneo).
- Regresión verde: `test_bulk_tools`, `test_sanitize`, `test_guard_wedge`,
  `test_hardening`.

## v1.4.0 — 2026-07-24

Paquete de lectura masiva + recálculo + diagnóstico de Data Model.
Resuelve los ítems #3 y #4 de `LIMITACIONES_MCP.md`.

### Nuevos tools

- **`recalculate`** (`tools/workbook.py`): recalcula fórmulas sin guardar/reabrir.
  Soporta `dirty` (`Calculate`), `full` (`CalculateFull`), `sheet` y `async_cube`
  (secuencia CUBE con auto temporal + `CalculateUntilAsyncQueriesDone`).
  Timeout: 600s (`LONG_OP_TIMEOUT`).

- **`read_table`** (`tools/bulk.py`): lee un `ListObject` completo por nombre
  (case-insensitive) en 2 llamadas COM. Inline ≤ 50k celdas, con `dest` a
  `.csv`/`.json` para tablas grandes. Maneja tabla vacía, sin headers y 1×1.

- **`export_sheet`** (`tools/bulk.py`): vuelca hoja completa (`UsedRange`) o rango
  arbitrario (`range_addr`) a `.csv`/`.tsv`/`.json` en 1 llamada COM.
  Techo: 5M celdas. Devuelve muestra + metadata del rango.

### Cambios en tools existentes

- **`read_range`** (`tools/cells.py`): ahora rechaza rangos > 50k celdas inline
  con error accionable que apunta a `export_sheet(range_addr=...)`.

- **`get_data_model_measures`** (`tools/power_pivot.py`): cambio de forma —
  ahora devuelve `{"measures": [...], "diagnostic": {"model_present": bool, ...}}`
  distinguiendo "sin medidas" de "sin modelo".

### Fixes

- `close_workbook` ahora limpia la copia temporal saneada al cerrar el libro
  (no espera a `close_excel`).

### Mejoras internas

- Normalización float→int en valores COM: evita exportar códigos PT como `5060094.0`.
- Bump de versión: `1.3.0 → 1.4.0`.
- Nuevo módulo `tools/bulk.py` con lógica de serialización compartida.

### Tests

- `test_bulk_tools.py`: 43 checks (recalculate, read_table, export_sheet,
  read_range cap, bordes vacíos, get_data_model_measures, limpieza).
- E2E contra MULTIFORMATO real: 876k celdas exportadas en 2.0 s con fidelidad
  perfecta (11.932,2160 ≈ COM 11.932,2160).
- Regresión: `test_sanitize.py`, `test_guard_wedge.py`, `test_hardening.py` verdes.
