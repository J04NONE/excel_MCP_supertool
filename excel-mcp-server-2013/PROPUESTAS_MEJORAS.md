# Propuestas de mejora — brainstorming 2026-07-08

Contexto: el flujo real es **Excel 2013 (trabajo) → PC personal (Excel 2021 + MCP) → Excel 2013 (trabajo)**.
Todo lo que el MCP produzca debe volver funcionando a 2013. Dos bloques de propuestas:

---

## Bloque A — Comprensión semántica: que el MCP "entienda" el libro

El problema real: explicar *por qué* existe una fórmula en una celda, *qué* calcula un
bloque, *cómo* funciona una macro. Un LLM explica bien cuando recibe el contexto correcto;
hoy tendría que leer 120.387 fórmulas celda por celda. Estas tools recolectan la evidencia
mínima suficiente:

### A1. `profile_formulas(sheet)` — huella de fórmulas por patrón ⭐ recomendada

La observación clave: en hojas masivas casi todas las fórmulas son **la misma fórmula
arrastrada**. En notación R1C1 (relativa), las 114.898 fórmulas de "Seguimiento OC
filiales" colapsan a ~60-80 patrones únicos (≈1 por columna).

- Lee `UsedRange.FormulaR1C1` de una pasada, agrupa por patrón único.
- Devuelve por patrón: fórmula ejemplo en A1 y R1C1, rango que cubre, cantidad de celdas,
  columnas involucradas.
- Con esto el agente explica una hoja de 114k fórmulas leyendo 80 filas de resumen.

### A2. `trace_cell(sheet, cell)` — el "porqué" de una celda ⭐ recomendada

- Fórmula de la celda (A1 + local), su valor actual.
- **Precedentes**: qué celdas/rangos/nombres definidos alimentan la fórmula (parseo de
  referencias + `DirectPrecedents` cuando aplica), con sus valores actuales.
- **Dependientes**: qué celdas usan esta (quién se rompe si la cambias).
- Resolución de nombres definidos (`BD!_FilterDatabase` → rango real).
- Con esa evidencia el agente responde "esta celda calcula X porque toma Y de la hoja Z".

### A3. `check_2013_compatibility()` — el guardián del viaje de vuelta ⭐ recomendada

Tu riesgo número 1 (ya visible en el archivo real: nombres `_xlfn.SUMIFS`, `_xlfn.IFERROR`
rotos con `#NAME?`). Al editar en Excel 2021 es facilísimo introducir funciones que 2013
no tiene y el archivo revienta al volver al trabajo.

- Escanea todas las fórmulas buscando funciones post-2013: XLOOKUP, XMATCH, TEXTJOIN,
  CONCAT, IFS, SWITCH, MAXIFS, MINIFS, LET, LAMBDA, FILTER, SORT, UNIQUE, SEQUENCE,
  RANDARRAY, TOCOL/TOROW, VSTACK/HSTACK, TEXTSPLIT, GROUPBY…
- Detecta referencias de arrays dinámicos (`A1#`) que 2013 no soporta.
- Reporta nombres `_xlfn.*` ya rotos.
- Complementa `validate_m_code` (que ya cubre el lado Power Query).
- Se ejecuta SIEMPRE antes de `save_workbook` final como checklist de entrega.

### A4. `map_dependencies()` — arquitectura del libro

- Grafo hoja→hoja: qué hojas alimentan a cuáles (referencias `'Hoja'!` en fórmulas,
  orígenes de pivots, rangos de nombres definidos).
- Clasifica: hojas de ENTRADA (datos, sin fórmulas), CÁLCULO (muchas fórmulas),
  SALIDA (pivots/reportes), OCULTAS con lógica.
- Para el archivo real produciría: `INV → {TD, TD 2, CFECHA} → Seguimiento OC filiales → PTE PROD.`

### A5. `analyze_vba_project()` — el "cómo y porqué" de las macros

- Sobre `get_vba_code` existente: análisis estático de todos los módulos.
- Call graph (qué macro llama a cuál), qué hojas/rangos toca cada una, eventos
  (Workbook_Open, Worksheet_Change), referencias externas.
- El agente explica el proyecto VBA completo con evidencia, no adivinando.

### A6. `document_workbook(output_path)` — el expediente completo

- Orquesta: analyze_workbook + profile_formulas (todas las hojas) + map_dependencies +
  list_pivot_tables + analyze_vba_project + list_power_queries + check_2013_compatibility.
- Genera un markdown: "Documentación técnica de <herramienta>.xlsm" — el entregable que
  hoy nadie tiene tiempo de escribir a mano.
- Es la tool que responde de una vez "¿de qué está compuesto este archivo?".

**Orden recomendado:** A1 + A2 + A3 primero (máximo valor por esfuerzo, y A3 protege cada
entrega), luego A4 + A5, y A6 al final porque orquesta a las demás.

---

## Bloque B — Estrategia ETL/ELT con solo Excel (y Power BI Desktop opcional)

Restricciones aceptadas: sin licencias nuevas, sin servidores, sin salir de Excel;
el runtime final es Excel 2013 Professional Plus (PQ 2.62 add-in + Power Pivot).

### B1. El patrón: de "libro de fórmulas" a mini data-warehouse

```text
HOY (Seguimiento OC):                      PROPUESTO (ELT en Excel):
pegar 27k filas a mano en INV        →     EXTRACT:  Power Query desde archivo/carpeta
114.898 fórmulas encadenadas         →     LOAD:     al Data Model (add_table_to_data_model
4 pivots con origen columna-completa →               o PQ "Agregar al modelo de datos")
                                           TRANSFORM: medidas DAX en el modelo
                                           SALIDA:   pivots del modelo + CUBEVALUE
```

Por qué ELT y no solo ETL: el Data Model (motor columnar xVelocity, el mismo de Power BI)
maneja millones de filas en el mismo .xlsx; las medidas DAX se recalculan al refrescar
sin una sola fórmula en celdas. 114k fórmulas → ~10 medidas DAX.

### B2. Extract con PQ 2.62 (compatible 2013)

- `Csv.Document`, `Excel.Workbook`, `Folder.Files` (ingesta de carpeta: "todos los TXT
  del mes"), `Web.Contents` — todo disponible en 2.62.
- Regla: cada query pasa por `validate_m_code` antes de entregarse.
- Staging: queries "solo conexión" (no cargar a hoja) → directo al modelo.

### B3. Transform: DAX básico en 2013

- Medidas con `SUM`, `CALCULATE`, `FILTER`, `DIVIDE`, `DISTINCTCOUNT` — todas en 2013.
- El MCP ya puede: `add_table_to_data_model`, `evaluate_dax_query` (ver caveat MSOLAP),
  `refresh_data_model`.
- Falta (propuesta B5): crear medidas persistentes desde el MCP.

### B4. Salida sin Power BI: pivots del modelo + fórmulas CUBE ⭐ el hallazgo

Las funciones **CUBEVALUE / CUBEMEMBER / CUBESET existen desde Excel 2007** y en 2013
consultan el Data Model del propio libro (`"ThisWorkbookDataModel"`):

```text
=CUBEVALUE("ThisWorkbookDataModel";
           CUBEMEMBER("ThisWorkbookDataModel";"[Ventas].[Filial].[Norte]");
           "[Measures].[Total Cajas]")
```

Con esto se arman dashboards en la hoja HOME con celdas sueltas alimentadas por el
modelo — sin pivots gigantes, sin fórmulas encadenadas, 100 % nativo 2013, y el MCP
puede escribirlas con `write_formulas`.

### B5. Nuevas tools propuestas para cerrar el ciclo ELT

| Tool | Qué hace |
|---|---|
| `add_data_model_measure(table, name, dax)` | Crea medida persistente (Model.ModelMeasures en 2016+; en 2013 la medida viaja en el archivo al crearse en el host — verificar ida y vuelta) |
| `add_power_query(name, m_code, load_to)` | Crea query PQ (valida con m_constraints; `load_to`: hoja / modelo / solo-conexión). En 2013 la API Queries no existe: la query debe crearse en el host 2021 y VIAJA en el archivo — verificar compatibilidad de versión de mashup |
| `write_cube_formulas(sheet, layout)` | Genera el dashboard CUBEVALUE/CUBEMEMBER desde una spec simple |
| `setup_refresh_macro()` | Inyecta `ActualizarTodo()`: refresca PQ → modelo → pivots en orden, con manejo de errores |

### B6. Power BI Desktop (capa opcional, solo visualización)

- **Gratis y legal para uso local**: .pbix en el escritorio no requiere licencia;
  lo que requiere Pro es PUBLICAR al servicio (y las políticas de la empresa
  probablemente lo bloquean de todos modos).
- Importa el .xlsx/.xlsm directamente, incluido su Data Model y medidas.
- Úsalo en TU PC para explorar/prototipar visuales; el entregable al trabajo sigue
  siendo el .xlsm autosuficiente (no dependas de que el trabajo instale PBI Desktop).
- Riesgo a validar: si en el trabajo no se puede instalar software, PBI es solo
  herramienta personal de análisis, nunca parte del entregable.

### B7. Migración concreta del "0607 Seguimiento OC F.xlsm" (caso piloto)

1. `INV` deja de ser pegado manual → PQ desde el archivo fuente del día (Folder.Files).
2. Cargar INV y "Seguimiento" crudos al Data Model.
3. Las columnas calculadas críticas → medidas DAX (validar cuáles de las 62 columnas
   son agregaciones vs. lógica fila a fila; las fila-a-fila se quedan como fórmulas
   o pasan a columnas calculadas DAX).
4. Pivots recreados contra el modelo (adiós orígenes de columna-completa).
5. HOME con CUBEVALUE.
6. `ActualizarTodo()` inyectada: un clic al día.
7. `check_2013_compatibility()` + prueba de ida y vuelta antes de entregar.

**Riesgo principal del bloque B (a validar primero):** que un Data Model + queries PQ
creados en Excel 2021 abran y refresquen bien en Excel 2013 (versión del motor de mashup
y del modelo). Propuesta: crear un libro piloto mínimo aquí, llevarlo al trabajo y
validar ANTES de migrar la herramienta real.

---

## Bloque C — Introspección visual (shapes, gráficos, segmentadores) ✅ IMPLEMENTADO 2026-07-12

Origen: al analizar el dashboard de "Seguimiento PT-CUOTA Julio 8.xlsb" (2026-07-08) no
hubo forma de listar objetos flotantes con las tools existentes — `analyze_workbook` y
`read_range` solo ven valores de celda. Hubo que rodear con `inject_vba_code` +
`execute_vba_macro` (ver MANUAL.md §5.1). Funcionó, pero cuesta una ida y vuelta de VBA
por cada pregunta nueva.

> **Estado: implementado** como `tools/shapes.py` (`list_shapes`, `list_charts`,
> `list_slicers`), verificado contra libro sintético (23/23) y contra el dashboard
> real (20/20, `list_slicers` en ~1 s sin cuelgue OLAP). Ver TOOLS.md.

### C1. `list_shapes(sheet)`

Nombre, tipo (`msoChart`/`msoSlicer`/`msoGroup`/rectángulo/etc.), posición y tamaño de
cada shape de una hoja, recursando dentro de `Group`s (el gauge real del caso de estudio
estaba anidado dentro de un `Group` y no aparecía en una pasada plana).

### C2. `list_charts(sheet)`

Por cada `ChartObject`/gráfico embebido: tipo (`ChartType`), tabla dinámica de origen
(si es PivotChart, con hoja), y fórmula de cada serie (`Series.Formula`). Esto es lo que
permitió detectar los gráficos de dona (`ChartType=-4111`) usados como "gauge" con una
serie real + un anillo plano decorativo.

### C3. `list_slicers()`

Nombre de cada `SlicerCache`, si es OLAP, `SourceName`, y qué tablas dinámicas controla
— **sin** iterar `SlicerCacheItems` (ver MANUAL.md §5.2: eso puede colgar >120s en
slicers conectados al Data Model). Con esto se detectó que los 3 segmentadores visibles
del dashboard filtran una tabla dinámica de otra hoja, no los gráficos del propio
dashboard — hallazgo que hoy solo sale escribiendo VBA ad hoc.

---

## Resumen ejecutivo

| Prioridad | Item | Valor |
|---|---|---|
| 1 | A1 + A2 + A3 (perfil de fórmulas, trace de celda, guardián 2013) | Entender libros masivos + proteger cada entrega |
| 2 | A4 + A5 + A6 (dependencias, VBA, expediente) | Documentación automática completa |
| 3 | B5 + B7 (tools ELT + piloto Seguimiento) | Automatizar el proceso diario de verdad |
| 4 | B6 (PBI Desktop personal) | Exploración visual sin licencia |
| 5 | ~~C1 + C2 + C3 (shapes/charts/slicers)~~ ✅ implementado 2026-07-12 | Diagnosticar dashboards sin rodeo de VBA ad hoc |
