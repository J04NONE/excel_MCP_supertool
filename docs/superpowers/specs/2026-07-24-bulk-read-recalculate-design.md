# Spec: Lectura masiva + recálculo + diagnóstico de Data Model (paquete #3+#4)

**Fecha:** 2026-07-24
**Origen:** `LIMITACIONES_MCP.md` ítems #3 (lectura masiva) y #4 (recálculo), detectados
en la sesión del informe Nutribella. Diseño revisado (5 correcciones incorporadas tras
revisión crítica).
**Repo:** `excel-mcp-server-2013/` — versión objetivo **1.4.0** (desde 1.3.0).

## Objetivo

Cuatro capacidades que hoy obligan a salir del MCP (a `pyxlsb`/`openpyxl`) o a
guardar+reabrir libros:

1. `recalculate` — recalcular sin guardar/reabrir (el MCP fuerza `Calculation=Manual`).
2. `read_table` — leer un `ListObject` completo por nombre en 2 llamadas COM.
3. `export_sheet` — volcar una hoja completa (`UsedRange.Value`, 1 llamada COM) a archivo.
4. `get_data_model_measures` honesto — distinguir "sin medidas" de "sin modelo".

**Restricción global:** todo el API COM usado debe existir en Excel 2013 (target real).

---

## Componente 1: `recalculate` — en `tools/workbook.py`

### Firma del tool

```python
recalculate(full: bool = False, sheet: Optional[str] = None, wait_async: bool = False) -> dict
```

### Comportamiento

| Parámetros | Acción COM |
|---|---|
| default | `Application.Calculate()` (solo celdas sucias) |
| `full=True` | `Application.CalculateFull()` (todo desde cero) |
| `sheet="X"` | `Worksheets("X").Calculate()` — ignora `full`; error claro si la hoja no existe |
| `wait_async=True` | ver secuencia CUBE abajo (aplica a nivel aplicación; ignora `sheet`) |

### Secuencia CUBE (corrección #1 de la revisión)

Con `Calculation=Manual` las fórmulas CUBE quedan `#GETTING_DATA` aunque se llame
`CalculateUntilAsyncQueriesDone()` (trampa documentada del proyecto). Con `wait_async=True`:

```python
prev = app.Calculation
app.Calculation = xlCalculationAutomatic   # -4105
try:
    app.CalculateFull()
    app.CalculateUntilAsyncQueriesDone()
finally:
    app.Calculation = prev                 # restaurar SIEMPRE (Manual normalmente)
```

### Timeout (corrección #2)

El tool corre con `timeout=LONG_OP_TIMEOUT` (600 s, constante ya existente para
`execute_vba_macro`), NO el default de 120 s: un `CalculateFull` de un libro corporativo
puede tardar minutos y no es un cuelgue (no hay diálogo; el fail-fast de diálogo modal no
se dispara, correcto).

### Respuesta

```json
{"calculated": true, "mode": "full|dirty|sheet:<nombre>|async_cube",
 "calculation_state": "done|calculating|pending"}
```

`calculation_state` mapea `Application.CalculationState` (0=xlDone→"done",
1=xlCalculating→"calculating", 2=xlPending→"pending"). Si la propiedad falla (sin libro
abierto), error claro "no hay workbook abierto".

### Compatibilidad 2013

`Calculate`, `CalculateFull`, `CalculateUntilAsyncQueriesDone` (2010+),
`CalculationState` (2007+): todos OK en 2013.

---

## Componente 2: `read_table` — en `tools/bulk.py` (módulo NUEVO)

### Firma del tool

```python
read_table(table_name: str, dest: Optional[str] = None) -> dict
```

### Búsqueda

Recorrer `wb.Worksheets` → `ws.ListObjects` comparando nombre **case-insensitive**
(patrón existente en el repo). Si no existe: error listando las tablas disponibles
(nombre + hoja) para que el agente se autocorrija.

### Orden de operaciones (corrección #3: tamaño ANTES de leer)

1. Localizar `ListObject`.
2. `n_rows = DataBodyRange.Rows.Count`, `n_cols = DataBodyRange.Columns.Count`
   (propiedades baratas; NO materializan datos). Si `DataBodyRange` es `None` (tabla
   vacía): `rows=[]` y responder sin leer (corrección #4).
3. Decidir entrega según `dest` y `n_rows*n_cols` vs `MAX_INLINE_CELLS = 50_000`.
4. Solo entonces leer: `HeaderRowRange.Value` (1 llamada) + `DataBodyRange.Value`
   (1 llamada).

### Bordes (corrección #4)

- Tabla vacía → `{"rows": [], "row_count": 0, ...}` sin tocar `.Value`.
- `ShowHeaders=False` / `HeaderRowRange` inaccesible → headers sintéticos
  `["col1", ..., "colN"]`.
- Rango 1×1 / 1×N / N×1: COM devuelve escalar o tupla 1-D → normalizar a matriz 2-D
  (helper `_as_matrix` en bulk.py si `matrix_to_jsonable` no lo cubre ya; verificar en
  implementación y reusar lo existente).

### Contrato de entrega (formas fijas, sin sorpresas)

- **Sin `dest`:** inline si `celdas ≤ 50.000`:
  `{"table", "sheet", "headers", "rows", "row_count", "col_count"}`.
  Si excede: **error accionable** — "tabla X tiene N celdas (> 50.000): pasa dest=ruta
  .csv/.json".
- **Con `dest`:** escribe archivo y devuelve
  `{"file", "rows", "cols", "headers", "sample"}` (sample = primeras 5 filas de datos).

**Nota anti-ambigüedad:** en la forma inline, `rows` es la MATRIZ de datos y los conteos
van en `row_count`/`col_count`; en la forma archivo, `rows`/`cols` son CONTEOS (misma
convención que `export_sheet`). Las dos formas se distinguen por la presencia de `file`.

---

## Componente 3: `export_sheet` — en `tools/bulk.py`

### Firma del tool

```python
export_sheet(sheet: str, dest: str, sample_rows: int = 5) -> dict
```

`dest` es OBLIGATORIO (el propósito del tool es volcar a archivo sin quemar contexto).

### Comportamiento

1. `ur = ws.UsedRange`; `n_rows/n_cols` por `Rows.Count/Columns.Count` (sin leer).
2. **Techo (corrección #5):** `MAX_EXPORT_CELLS = 5_000_000`. Si excede → error
   accionable ("exporta por rangos con read_range o divide la hoja"). La BD real del
   multiformato (9.036×97 ≈ 876k celdas) pasa holgada.
3. `ur.Value` en **1 llamada COM** (fast path).
4. Escribir `dest` según extensión: `.csv` o `.json` (otra extensión → error).
5. Respuesta: `{"file", "rows", "cols", "range", "sample"}` donde `range` es
   `ur.Address` (ej. `"$C$4:$CU$9036"`).

### Reglas de serialización (ambos tools de bulk)

- Fechas COM (`pywintypes.datetime`) → ISO 8601, reusando `matrix_to_jsonable` de
  `utils/excel_utils` (consistencia con `read_range`).
- CSV: `utf-8-sig` (BOM para que Excel lo reabra bien), separador coma, `csv.QUOTE_MINIMAL`,
  `newline=""`.
- JSON: lista de listas (una hoja no garantiza fila de encabezado), `ensure_ascii=False`.
- `None` → celda vacía en CSV / `null` en JSON.

### Documentado, no "arreglado"

- `UsedRange` puede NO empezar en A1 (la BD real empieza en C4): el consumidor debe usar
  la clave `range` para mapear offsets. Va en el docstring del tool.
- Celdas combinadas: `.Value` trae el valor solo en la celda superior-izquierda; el resto
  sale `None`. Va en el docstring.

---

## Componente 4: `get_data_model_measures` honesto — editar `tools/power_pivot.py`

### Cambio de forma (breaking aceptado)

Antes: `list` (y `[]` mudo que confunde "sin medidas" con "sin modelo").
Ahora:

```json
{"measures": [...],
 "diagnostic": {"model_present": true, "model_tables": 3,
                "note": "Modelo presente con 3 tablas y 0 medidas explícitas: "
                        "probablemente pivots con agregaciones implícitas."}}
```

- Con medidas: `diagnostic` mínimo (`model_present`, `model_tables`).
- Sin modelo: `{"measures": [], "diagnostic": {"model_present": false, ...}}`.
- El conteo de tablas usa `Model.ModelTables.Count` (2013+); si el propio acceso al
  modelo falla, `model_present: false` con el motivo en `note`.
- El fallback DMV existente se conserva; solo se envuelve la respuesta en la nueva forma.

---

## Cableado y transversales

1. `tools/bulk.py` nuevo con `register(mcp, session, run)`; añadirlo al bloque de
   imports y registros de `server.py` (junto a `cells`, `workbook`, etc.).
2. Constantes en `bulk.py`: `MAX_INLINE_CELLS = 50_000`, `MAX_EXPORT_CELLS = 5_000_000`,
   `SAMPLE_ROWS_DEFAULT = 5`.
3. `recalculate` usa la constante de timeout larga existente (`LONG_OP_TIMEOUT`/600 s);
   si hoy vive en otro módulo, importarla, no duplicarla.
4. Bump de versión en `server.py`: `1.3.0 → 1.4.0`.
5. Actualizar `TOOLS.md` (4 tools nuevos/cambiados) y `MANUAL.md` si aplica; anotar el
   cambio de forma de `get_data_model_measures`.
6. Marcar #3 y #4 como resueltos en `LIMITACIONES_MCP.md` al cerrar, con evidencia.

## Plan de pruebas

### `test_bulk_tools.py` (nuevo, con Excel real — patrón de test_hardening)

1. Crear libro vía COM: hoja con tabla 100×10 (`ListObjects.Add`) + hoja con fórmulas.
2. `recalculate`: en Manual, escribir fórmula nueva (valor vacío/desactualizado) →
   `recalculate()` → valor correcto; `full=True` y `sheet=` también; `calculation_state`
   = "done".
3. `read_table` inline: headers + 100 filas correctas; tipos (número, texto, fecha ISO).
4. `read_table` con `dest` .csv y .json: archivo existe, sample correcto, contenido
   coincide (releer el csv y comparar 2-3 celdas).
5. Cap inline: tabla sintética > 50k celdas (ej. 600×100) → error accionable sin `dest`;
   OK con `dest`.
6. Bordes: tabla vacía (`rows=[]`), rango 1×1 normalizado.
7. `export_sheet`: csv y json de una hoja con `UsedRange` que NO empieza en A1
   (escribir desde C4) → `range` correcto, offsets verificables.
8. `get_data_model_measures`: en libro sin modelo → `model_present: false` (sin explotar).
9. Limpieza total (sin EXCEL.EXE huérfanos, sin archivos temp).

### E2E contra archivos reales de la carpeta Nutribella

- `read_table("VENTAS")` en el MULTIFORMATO (3.406 filas → requiere `dest`) → comparar
  totales de una columna contra los valores ya validados en la sesión (2.885,69 t enero).
- `export_sheet("BD", dest=...csv)` (876k celdas) → filas/cols correctos, `range`
  `$C$4:...`, tiempo < 120 s.
- `recalculate(wait_async=True)` sobre un libro con fórmulas CUBE si hay uno a mano
  (el dashboard PT-CUOTA sirve); mínimo: verificar que no rompe sin CUBE.

### Regresión

`test_sanitize.py`, `test_guard_wedge.py`, `test_hardening.py` — todos verdes.

## Fuera de alcance (YAGNI)

- Export a .xlsx/parquet; filtros/proyecciones en read_table; export de múltiples hojas
  en una llamada; streaming; sanitización CSV-injection (datos propios, uso analítico).

## Riesgos aceptados

- `.Value` de 876k celdas ≈ decenas de MB en el proceso Python durante segundos: OK
  (una sola operación, se libera al escribir el archivo).
- Cambio de forma en `get_data_model_measures`: consumidores son agentes; documentado
  en TOOLS.md.
- `recalculate` con 600 s puede "sentirse" colgado desde el cliente: la respuesta incluye
  `calculation_state` y el docstring lo advierte.
