# Referencia de Tools (48)

Convenciones:
- Todas las tools que tocan Excel lo arrancan solas si no está corriendo (lazy init).
- "Workbook activo" = el último abierto con `open_workbook`.
- Rangos en notación A1 (`"A1"`, `"A1:C10"`). Matrices = lista de filas (lista de listas).

## Sesión

| Tool | Argumentos | Devuelve | Notas |
|---|---|---|---|
| `ping` | — | `{status, timestamp}` | No requiere Excel |
| `get_session_info` | — | `{pid, version, visible, workbook_count}` | Arranca Excel |
| `close_excel` | — | `{closed}` | Descarta cambios sin guardar |

## Workbook

| Tool | Argumentos | Devuelve | Notas |
|---|---|---|---|
| `open_workbook` | `path`, `read_only=False`, `password=None`, `enable_macros=False` | hojas + metadata | Con `enable_macros=False` las macros NO se cargan (seguro para .xlsm ajenos). Para `execute_vba_macro` abrir con `True` |
| `save_workbook` | `path=None` | `{saved, path}` | Sin `path` → Save; con `path` → SaveAs (formato por extensión: .xlsx/.xlsm/.xlsb/.xls) |
| `close_workbook` | `save_changes=False` | `{closed, saved}` | Por defecto NO guarda |
| `list_sheets` | — | `[{name, type}]` | |
| `create_sheet` | `name`, `after=None` | `{name, index}` | Sin `after` la agrega al final |
| `delete_sheet` | `name` | `{deleted}` | Irreversible dentro del libro |
| `analyze_workbook` | — | radiografía completa | **Punto de partida para desenmarañar**: hojas (dimensiones, nº fórmulas, visibilidad), tablas, pivots, nombres definidos, conexiones, si tiene VBA |

## Celdas y fórmulas

| Tool | Argumentos | Devuelve | Notas |
|---|---|---|---|
| `read_range` | `sheet`, `range_addr` | matriz 2D de valores | Fechas → ISO 8601 |
| `write_range` | `sheet`, `range_addr`, `values` | confirmación | `range_addr` puede ser solo la esquina (`"A1"`); el tamaño lo da `values` |
| `read_formulas` | `sheet`, `range_addr`, `local=False` | matriz 2D de fórmulas | Celdas sin fórmula devuelven su valor. `local=True` → `=SUMA(...)` (idioma UI) |
| `write_formulas` | `sheet`, `range_addr`, `formulas`, `local=False` | confirmación | `local=False` espera inglés (`=SUM`), portable entre idiomas |
| `apply_format` | `sheet`, `range_addr`, `bold`, `italic`, `font_size`, `font_color_rgb`, `fill_color_rgb`, `number_format` | confirmación | Colores hex `"RRGGBB"`; `number_format` ej. `"#,##0.00"`, `"dd/mm/yyyy"` |
| `auto_fit_columns` | `sheet`, `range_addr=None` | confirmación | Sin rango ajusta todo el UsedRange |

## VBA

| Tool | Argumentos | Devuelve | Notas |
|---|---|---|---|
| `list_vba_modules` | — | `[{name, type, lines}]` | Requiere trust del VBA object model (†) |
| `get_vba_code` | `module_name` | `{module, type, lines, code}` | Extrae el código fuente completo (†) |
| `execute_vba_macro` | `macro_name`, `arguments=None`, `timeout_s=120` | `{macro, result}` | NO requiere trust, pero el libro debe abrirse con `enable_macros=True`. `timeout_s` (cap 600) para macros largas |
| `inject_vba_code` | `module_name`, `code`, `replace=False` | `{module, created, lines}` | Crea el módulo si no existe (†) |

(†) Requiere: Centro de confianza → Configuración de macros → *"Confiar en el acceso al modelo de objetos de proyectos de VBA"*. Si falta, la tool devuelve las instrucciones exactas.

## Power Query

| Tool | Argumentos | Devuelve | Notas |
|---|---|---|---|
| `list_power_queries` | — | `{queries, connections, queries_api_available}` | En Excel 2013 `Workbook.Queries` no existe → solo `connections` |
| `get_power_query_m` | `query_name` | `{name, m_code}` | Solo host 2016+; en 2013 el M no es accesible vía COM |
| `refresh_power_query` | `connection_name=None` | `{refreshed, failed}` | SIEMPRE síncrono (`BackgroundQuery=False`); sin nombre refresca todas |
| `validate_m_code` | `m_code` | `{valid, blocked_found, alternatives}` | Valida contra el motor M v2.62 de Excel 2013 |
| `m_function_compatible` | `function_name` | `{compatible, alternative, ...}` | Consulta puntual de una función M |

## Power Pivot / Data Model (DAX)

| Tool | Argumentos | Devuelve | Notas |
|---|---|---|---|
| `list_data_model_tables` | — | `[{name, record_count, source}]` | Tablas del Data Model |
| `add_table_to_data_model` | `sheet`, `range_addr`, `table_name` | tabla + estado | Convierte un rango en ListObject y lo carga al modelo. Base del ELT. Requiere libro guardado |
| `evaluate_dax_query` | `dax`, `max_rows=1000` | `{columns, rows, row_count, truncated}` | Ej. `EVALUATE 'Tabla'` o `EVALUATE SUMMARIZECOLUMNS(...)`. Vía ADO in-process (†PP) |
| `get_data_model_measures` | — | `[{name, expression, table}]` | Medidas DAX; fallback DMV si el host no expone `ModelMeasures` |
| `refresh_data_model` | — | `{refreshed, tables}` | Refresca todas las conexiones del modelo |

(†PP) `evaluate_dax_query` usa el proveedor **MSOLAP** vía la conexión ADO del modelo. En **Excel 2013 Professional Plus (MSI)** MSOLAP está registrado a nivel máquina y funciona directo. En hosts **Office Click-to-Run** (dev 2016+) MSOLAP puede no estar visible para el proceso; entonces la tool **falla rápido con un mensaje accionable** (no cuelga). Alternativa portable sin MSOLAP: `create_pivot_table` agrega contra el mismo modelo.

## Tablas dinámicas

| Tool | Argumentos | Devuelve | Notas |
|---|---|---|---|
| `list_pivot_tables` | — | `[{name, sheet, location, source, fields}]` | `fields` separa filas/columnas/filtros/valores |
| `create_pivot_table` | `source_sheet`, `source_range`, `dest_sheet`, `dest_cell`, `rows`, `columns`, `filters`, `values`, `name` | pivot creada | `values=[{"field": "Ventas", "agg": "sum"}]`; agg: sum, count, average, max, min, product, stdev, var. 100% nativa (PivotCache) |
| `refresh_pivot_tables` | `sheet=None` | `{refreshed, failed}` | Sin `sheet` refresca todo el libro |

## Comprensión semántica (desenmarañar)

| Tool | Argumentos | Devuelve | Notas |
|---|---|---|---|
| `profile_formulas` | `sheet`, `max_patterns=200` | patrones únicos de fórmula | Agrupa por R1C1: 114.898 fórmulas → ~34 patrones. LA forma de leer hojas masivas |
| `trace_cell` | `sheet`, `cell`, `max_precedents=25` | fórmula + precedentes (con valores) + dependientes + nombres | El "porqué" de una celda con evidencia |
| `check_2013_compatibility` | — | funciones post-2013, refs de derrame `A1#`, nombres `_xlfn.*` (bloqueantes vs residuales), M inválido | **Ejecutar SIEMPRE antes de entregar al Excel 2013 del trabajo** |
| `map_dependencies` | — | grafo hoja→hoja + clasificación entrada/cálculo/salida + hojas ocultas | La arquitectura del libro |
| `analyze_vba_project` | — | procedimientos, call graph, hojas/rangos tocados, eventos | Requiere trust VBA (†) |
| `document_workbook` | `output_path?`, `max_patterns_per_sheet=40` | genera el expediente markdown completo | Orquesta todo el análisis en una pasada |

## ELT (Extract-Load-Transform con solo Excel)

| Tool | Argumentos | Devuelve | Notas |
|---|---|---|---|
| `add_data_model_measure` | `table_name`, `measure_name`, `dax`, `number_format` | medida + referencia CUBE | La medida viaja en el archivo y funciona en 2013. API de creación requiere host 2016+ |
| `add_power_query` | `query_name`, `m_code`, `load_to`, `target_sheet?` | query creada y cargada | Valida el M contra 2013 ANTES de crear. `load_to`: connection_only / sheet / data_model |
| `write_cube_formulas` | `sheet`, `start_cell`, `title`, `rows`, `values`, `model_name` | mini-dashboard CUBE | CUBEVALUE/CUBEMEMBER: nativas desde 2007, perfectas para 2013. Fuerza el recálculo asíncrono |
| `setup_refresh_macro` | `macro_name="ActualizarTodo"` | macro inyectada + instrucciones | Refresca PQ → modelo → pivots en orden síncrono. Solo inyecta (no ejecuta) |

## Introspección visual (dashboards)

| Tool | Argumentos | Devuelve | Notas |
|---|---|---|---|
| `list_shapes` | `sheet` | `{top_level_count, total_count, shapes}` árbol con `children` | Ve lo que `read_range` no ve: gráficos, slicers, rectángulos, imágenes. Recursa dentro de grupos (los gauges reales suelen estar agrupados) |
| `list_charts` | `sheet` | `{chart_count, charts}` con tipo legible, series y origen | `group_path` marca charts anidados; `is_pivot_chart` + `pivot_source` identifican el pivot que lo alimenta; `-4111` = xlCombination (típico de gauges) |
| `list_slicers` | `include_items=False` | `{slicer_cache_count, slicer_caches}` | Por cache: OLAP, origen, pivots que controla y posición de cada slicer visible. **Items de caches OLAP jamás se listan** (iterar items del Data Model cuelga la sesión COM — MANUAL §5.2); no-OLAP solo con `include_items=True` (tope 50) |

## Diagnóstico del entorno

| Tool | Argumentos | Devuelve | Notas |
|---|---|---|---|
| `discover_capabilities` | — | versión, arquitectura x86/x64, add-ins PQ/PP, acceso VBA | Protocolo de autodescubrimiento del informe técnico |
| `validate_environment` | — | `{status, warnings, errors, recommendations}` | Advertencias accionables según restricciones de Excel 2013 |
