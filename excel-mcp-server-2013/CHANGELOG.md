# Changelog

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
