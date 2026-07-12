# Manual de uso — excel-mcp-server-2013

Este manual muestra cómo usar el servidor MCP para el flujo completo de trabajo con
"herramientas" (libros de Excel complejos): **entender → desarmar → mejorar → armar → entregar**.

Todos los ejemplos son instrucciones en lenguaje natural que le das a Claude Code (u otro
agente conectado al MCP); entre paréntesis, las tools que el agente usará.

---

## 1. Flujo de trabajo recomendado

```
1. validate_environment          ← ¿el entorno cumple? (x64, add-ins, trust VBA)
2. open_workbook (read_only=True, enable_macros=False)   ← inspección SEGURA
3. analyze_workbook              ← radiografía: hojas, fórmulas, pivots, VBA, conexiones
4. list_pivot_tables / list_vba_modules / get_vba_code / read_formulas
                                 ← desarmar pieza por pieza
5. (reabrir en escritura)        ← open_workbook sin read_only
6. write_formulas / create_pivot_table / inject_vba_code / apply_format
                                 ← armar/mejorar con piezas NATIVAS
7. check_2013_compatibility      ← GUARDIÁN: ¿el libro sobrevive el regreso a 2013?
8. save_workbook (path nuevo)    ← nunca pisar el original hasta validar
9. close_workbook / close_excel
```

Reglas de oro:
- **Primera pasada siempre `read_only=True, enable_macros=False`**: ves todo sin riesgo.
- **Entrega con fórmulas, no valores**: `write_formulas` escribe la fórmula real en la celda;
  el resultado lo calcula Excel. El libro queda como si lo hubiera hecho un humano.
- **Guardar en copia** (`save_workbook` con path nuevo) hasta que el cambio esté validado.

---

## 2. Casos de uso

### Caso A — Radiografía de una herramienta desconocida (real)

> *"Ábreme '0607 Seguimiento OC F.xlsm' en solo lectura y dime de qué está compuesto."*

Resultado real de `analyze_workbook` + `list_pivot_tables` sobre ese archivo (4,66 MB):

| Hoja | Rango usado | Filas | Fórmulas | Contenido |
|---|---|---|---|---|
| HOME | B3:N18 | 16 | 0 | Portada/menú |
| 1. Actualización | B5:D14 | 10 | 0 | Instrucciones |
| INV | A1:Q27188 | 27.188 | 0 | **Datos crudos de inventario** |
| TD | A4:Q1075 | 1.072 | 2.128 | Pivot "Tabla dinámica2" + fórmulas |
| TD 2 | A2:K1126 | 1.125 | 2.242 | Pivot "Tabla dinámica1" + fórmulas |
| Seguimiento OC filiales | A1:BJ3594 | 3.594 | **114.898** | **El corazón: 62 columnas de fórmulas** |
| CFECHA | A2:C35 | 34 | 0 | **Hoja OCULTA** con pivot auxiliar |
| PTE PROD. | B2:I171 | 170 | 278 | Pivot de pendientes de producción |
| BD | B2:AE989 | 988 | 841 | Base auxiliar |
| INFO Nutri | A1:N395 | 395 | 0 | Datos de referencia |

**Total: 120.387 celdas con fórmula, 4 tablas dinámicas, proyecto VBA con 3 módulos con código**
(Módulo2: 11 líneas, Módulo3: 38, Módulo5: 41 — extraíbles con `get_vba_code`).

Diagnósticos que salen gratis del análisis:
- Dos pivots (CFECHA y Seguimiento) usan **origen de columna completa** (`F1048576` = fila 1.048.576):
  infla el PivotCache y ralentiza el refresco → acotar el origen o usar tabla.
- Nombres definidos `_xlfn.SUMIFS`, `_xlfn.IFERROR`, `_xleta.N` con `#NAME?`: huella de que el
  archivo circuló entre **versiones distintas de Excel** — exactamente el riesgo 2013 vs 2021
  que este servidor valida.
- Hoja oculta `CFECHA`: lógica escondida que un "mejoramiento" manual podría romper sin saberlo.

### Caso B — Mejorar una herramienta existente ("mejoramientos")

> *"En la copia de la herramienta, agrega en 'Seguimiento OC filiales' una columna
> 'Alerta' que marque las OC con más de 30 días, y una hoja Resumen con tabla dinámica
> por filial."*

El agente hace:
1. `open_workbook` (escritura) → `save_workbook` a `..._v2.xlsm` (trabajar sobre copia).
2. `read_formulas` de las columnas vecinas para **imitar el estilo de fórmula existente**.
3. `write_formulas` con la fórmula nueva (ej. `=IF(TODAY()-D5>30,"ALERTA","")` — en inglés,
   Excel la muestra en español) + `apply_format` para el encabezado.
4. `create_sheet("Resumen")` + `create_pivot_table(rows=["FILIAL"], values=[{"field":"Pend (Cjs)","agg":"sum"}])`.
5. `save_workbook` + `analyze_workbook` de control (¿cuántas fórmulas quedaron?).

Todo el resultado es **nativo**: fórmulas reales, pivot con PivotCache real. Nada de valores pegados.

### Caso C — Armar un entregable desde cero con fórmulas

> *"Crea un libro de control de ventas: hoja Datos con estas columnas, hoja Dashboard
> con totales por SUMIFS y una tabla dinámica por categoría."*

`write_range` (datos) → `write_formulas` (`=SUMIFS(...)`, `=IFERROR(...)`) →
`apply_format` + `auto_fit_columns` → `create_pivot_table` → `save_workbook("control.xlsx")`.

El libro se entrega **con las fórmulas vivas**: quien lo reciba puede auditar y extender cada celda.

### Caso D — Automatizar el proceso diario

> *"Todos los días pego el inventario nuevo en INV y refresco todo a mano. Automatízalo."*

Dos vías (combinables):
- **VBA**: `inject_vba_code` con una macro `ActualizarTodo()` que limpia INV, importa el archivo
  del día y refresca pivots; luego un solo `execute_vba_macro("ActualizarTodo")` diario.
  (Abrir con `enable_macros=True`; requiere trust VBA para inyectar, no para ejecutar.)
- **Sin VBA**: el agente ejecuta la secuencia `write_range` (datos nuevos) →
  `refresh_power_query` → `refresh_pivot_tables` en un solo pedido.

### Caso E — Generar Power Query compatible con Excel 2013

> *"Necesito una query M que separe el código de artículo por el guion, pero la herramienta
> se usa en Excel 2013."*

1. El agente propone el M y lo pasa por `validate_m_code`.
2. Si usa funciones post-2013 (ej. `Text.BeforeDelimiter`), la tool las detecta y devuelve la
   **alternativa legacy** (`Text.PositionOf + Text.Start`).
3. `m_function_compatible("List.Generate")` para dudas puntuales.

Restricción de fondo: Excel 2013 usa el motor M **v2.62 legacy** y Power Query es un
**add-in COM** (no nativo). `validate_environment` avisa si el add-in no está activo.

### Caso F — Entender y documentar una herramienta (pack semántico, verificado)

> *"Explícame de qué está hecho este libro y por qué esa fórmula está ahí."*

Resultados reales sobre `0607 Seguimiento OC F.xlsm`:

- `profile_formulas("Seguimiento OC filiales")`: **114.898 fórmulas → 34 patrones únicos**
  (el patrón top: `=AP5-TODAY()` repetido 7.180 veces en AR5:AS3594). Leer 34 resúmenes
  en vez de 114 mil celdas.
- `trace_cell(...)`: para cualquier celda devuelve su fórmula, qué la alimenta (con
  valores) y qué depende de ella. Detectó que `AR5` está en estado de error `#VALUE!`.
- `map_dependencies()`: el grafo completo — `INV → {TD, TD 2, CFECHA} → Seguimiento
  OC filiales → PTE PROD.`, con `BD` e `INFO Nutri` alimentando a Seguimiento, y
  `CFECHA` marcada como oculta.
- `analyze_vba_project()`: 5 procedimientos con call graph; hallazgo real:
  `Módulo5.ORDER` referencia la hoja "CALCULO DISPONIBLE" **que ya no existe** (código
  muerto o bug latente).
- `document_workbook()`: genera `<nombre>_DOCUMENTACION.md` con todo lo anterior — el
  expediente que nadie escribe a mano.

### Caso G — Piloto ELT ida-vuelta 2021 → 2013 (generado y verificado)

El archivo `PILOTO_IDA_VUELTA.xlsm` (en la carpeta del proyecto) fue construido 100 %
por el MCP y contiene el ciclo ELT completo:

| Pieza | Contenido | Verificado |
|---|---|---|
| Hoja `Datos` | Tabla `VentasPiloto` (5 filas) cargada al Data Model | ✅ |
| Medida DAX | `Total Cajas = SUM(VentasPiloto[Cajas])` (formato entero) | ✅ |
| Power Query | `QPiloto` (M 100 % legacy 2.62) cargada a hoja `PQ` | ✅ 6 filas |
| Dashboard | CUBEMEMBER/CUBEVALUE contra el modelo | ✅ Norte 200 · Sur 260 · Centro 90 |
| Macro | `ActualizarTodo()` (PQ → modelo → pivots, síncrono) | ✅ inyectada |
| Guardián | `check_2013_compatibility` | ✅ seguro para 2013 |

**Protocolo de validación en el trabajo (Excel 2013 Pro Plus):**

1. Copia `PILOTO_IDA_VUELTA.xlsm` al PC del trabajo y ábrelo en Excel 2013.
2. Habilita macros. Verifica que la hoja `Dashboard` muestre los totales (CUBEVALUE).
3. Cambia un valor en `Datos` (ej. Norte/Cola: 120 → 500) y ejecuta `Alt+F8 → ActualizarTodo`.
4. Debe refrescar: hoja `PQ` (query), el modelo y el Dashboard (Norte pasa a 580).
5. Si los 4 pasos funcionan, la arquitectura ELT completa es viable para migrar la
   herramienta real (caso B7 de PROPUESTAS_MEJORAS.md).

Si algo falla, anota el mensaje exacto: identifica qué capa no viajó (query, modelo,
medida o CUBE) y se ajusta la estrategia con ese dato.

---

## 3. Seguridad y límites

| Tema | Comportamiento |
|---|---|
| Excel del usuario | Nunca se toca: el servidor crea su PROPIA instancia (DispatchEx) |
| Macros al abrir | Deshabilitadas por defecto (`enable_macros=False`) — un .xlsm ajeno no ejecuta nada |
| Guardado | Siempre explícito; `close_workbook` por defecto descarta |
| Procesos zombie | Limpieza de 3 niveles al apagar: Quit → GC → kill del PID exacto |
| Excel 2013 | `Workbook.Queries` no existe (M no extraíble vía COM); PQ/PP son add-ins; M v2.62; DAX/Power Pivot pendiente (Fase 5) |
| VBA | Leer/escribir código requiere trust del VBA object model; ejecutar macros no |
| Timeout | 120 s default, configurable por operación (refresh/document usan 600 s; `execute_vba_macro` expone `timeout_s`) |
| Crash de Excel | Recovery automático: pre-flight reinicia y procede; mid-flight lanza `ExcelCrashedError` (nunca retry silencioso) |
| Errores COM | HRESULTs traducidos a mensajes accionables (proceso muerto / ocupado / 1004 / elemento inexistente) |

## 4. Estado vs. plan maestro

| Fase | Estado |
|---|---|
| 1. Esqueleto + sesión COM | ✅ Completa |
| 2. Workbook + celdas/fórmulas/formato | ✅ Completa |
| 3. VBA (listar, extraer, ejecutar, inyectar) | ✅ Completa |
| 4. Power Query (listar, refrescar, validar M 2013) | ✅ Completa (extracción de M solo host 2016+) |
| 5. Power Pivot / DAX | ✅ Completa (Data Model + DAX vía ADO `$Embedded$`; ver caveat MSOLAP) |
| 6. Autodescubrimiento | ✅ Completa |
| 7. Hardening (recovery ante crash, batch, cache) | ✅ Núcleo completo (recovery automático tras crash, HRESULT traducidos, timeout por operación; excluido a propósito: cache COM y estrés 1000 ops) |
| 8. Documentación | ✅ Este manual + README + TOOLS.md |
| Ext. A: Comprensión semántica | ✅ profile_formulas, trace_cell, check_2013_compatibility, map_dependencies, analyze_vba_project, document_workbook |
| Ext. B: Kit ELT | ✅ add_data_model_measure, add_power_query, write_cube_formulas, setup_refresh_macro + `PILOTO_IDA_VUELTA.xlsm` (pendiente validar en el 2013 del trabajo) |

### Caveat DAX / MSOLAP (importante para tu flujo)

`evaluate_dax_query` consulta el Data Model vía el proveedor **MSOLAP**. En tu
**Excel 2013 Professional Plus del trabajo (instalación MSI)**, MSOLAP está registrado
a nivel de máquina y las consultas DAX funcionan directamente. En tu **PC personal con
Office 2021 Click-to-Run**, MSOLAP a veces no es visible para el proceso del servidor;
en ese caso la tool **falla en < 1 s con un mensaje claro** (nunca cuelga ni deja Excel
zombie). Las demás tools del modelo (`add_table_to_data_model`, `list_data_model_tables`,
`get_data_model_measures`, `refresh_data_model`) funcionan en ambos entornos porque usan
el modelo de objetos COM de Excel, no MSOLAP externo.

Para agregaciones portables que funcionan en ambos hosts sin depender de MSOLAP, usa
`create_pivot_table` (tabla dinámica nativa contra el mismo Data Model).

---

## 5. Limitaciones conocidas (bitácora)

Registro de limitaciones reales encontradas en uso (no hipótesis). Se agrega una entrada
cada vez que aparece una nueva, con la mitigación que funcionó.

### 5.1 Sin tools nativas para objetos flotantes (shapes, gráficos, segmentadores)

`analyze_workbook` / `read_range` solo ven **valores de celda**. Un dashboard con
gráficos, segmentadores (slicers) o formas flotantes es invisible para esas tools
(`read_range` sobre un rango lleno de gráficos devuelve `null` en casi todo).

**Workaround usado (funciona, sin tool dedicada todavía):**

1. `open_workbook(enable_macros=True)` — reabrir con macros.
2. `inject_vba_code` en un módulo scratch (ej. `ClaudeInspect`) con una `Function` que
   arma un `String` recorriendo `Worksheet.Shapes`, `ChartObjects`,
   `ThisWorkbook.SlicerCaches`.
3. `execute_vba_macro` para leer el resultado (`Application.Run` soporta funciones con
   valor de retorno).
4. `close_workbook(save_changes=False)` al terminar — nada se escribe al archivo real.

Caso real (2026-07-08): dashboard de "Seguimiento PT-CUOTA" con 26 shapes, 3 slicers y
7 gráficos (incluido 1 "gauge" anidado dentro de un `Group`) diagnosticado con este rodeo.

> **RESUELTO (2026-07-12):** existen tools nativas `list_shapes(sheet)`,
> `list_charts(sheet)` y `list_slicers()` (Bloque C). El rodeo VBA queda solo como
> referencia histórica / para objetos aún no cubiertos. Bonus: la primera corrida de
> las tools sobre el dashboard real corrigió 3 errores del análisis manual VBA
> (eran 7 charts, los gauges son xlCombination y los slicers del DashBoard SÍ
> filtran el dashboard vía las pivots de TD CUBO) — ver findings del plan 2026-07-12.

### 5.2 Iterar los items de un slicer OLAP puede colgar >120 s

Extiende el caveat de MSOLAP de arriba a un caso nuevo: no es solo `evaluate_dax_query`
vía ADO — el **modelo de objetos COM normal** también puede quedarse esperando el
proveedor MSOLAP al introspeccionar los items de un slicer conectado al Data Model
(`OLAP=True`) si el proceso no lo tiene bien registrado. A diferencia de
`evaluate_dax_query`, esta vía **no falla rápido**: colgó la llamada COM hasta el
timeout de 120 s.

Precisión de nombres (corregido 2026-07-12): la colección se llama
**`SlicerCache.SlicerItems`** (`SlicerCacheItems` no existe en el modelo de objetos;
en caches OLAP los items viven en `SlicerCacheLevels(n).SlicerItems`, y
`SlicerCacheLevels` en un cache no-OLAP lanza 0x800A03EC).

**Mitigación (aplicada en `list_slicers`):** leer solo metadata barata (`Name`, `OLAP`,
`SourceName`, `PivotTables` — todos seguros y rápidos) y **jamás** tocar los items de
un cache OLAP; los de caches no-OLAP solo bajo `include_items=True` con tope de 50.
Si la lectura de `.OLAP` falla, se asume OLAP (conservador).

### 5.3 `get_data_model_measures` / `list_data_model_tables` solo ven el Data Model LOCAL

Si un libro combina el Data Model embebido (Power Pivot, `ThisWorkbookDataModel`) **con**
conexiones en vivo a un cubo OLAP corporativo externo (SSAS, ej. conexiones
`OLAPSERVER ...`), estas tools solo devuelven las tablas/medidas del modelo local (en la
práctica, las que vinieron de Power Query). Las "medidas" visibles en pivots que usan
dimensiones tipo `[Producto].[Marcas]`, `[Fecha Factura].[Año]` pueden venir del cubo
externo y no van a aparecer acá — no es un bug, es un límite de alcance a tener presente
al diagnosticar "¿por qué esta tool no me muestra la medida que veo en el pivot?".

### 5.4 Al escribir VBA de inspección: no combines varias propiedades falibles en una sola sentencia

Gotcha de VBA en sí (no del MCP), pero costó un ciclo de debugging: `On Error Resume Next`
protege la *sentencia completa*, no cada sub-expresión. Si concatenás dos propiedades en
una sola línea (`s = s & a.PropOk & a.PropQueFalla`) y la segunda lanza error, **toda la
asignación se descarta**, incluida la primera propiedad que sí funcionaba. Se detectó así:
un gráfico sin tabla dinámica (`PivotLayout.PivotTable.Name` inválido) hizo desaparecer
también su `ChartType`, que sí era accesible. Regla: una propiedad riesgosa por sentencia,
con su propio `On Error Resume Next` / `Err.Clear`.

### 5.5 El timeout de 120 s (tabla §3) aplica también a macros propias inyectadas

No es solo para las tools nativas: cualquier `execute_vba_macro` sobre código inyectado
corre bajo el mismo límite. Si una macro larga (recorrer miles de shapes, refrescar un
cubo externo lento) puede superar 120 s, conviene dividirla en pasos más chicos en vez
de asumir que el timeout se ajusta solo.

> **RESUELTO (2026-07-12):** el timeout ahora es configurable por operación.
> `execute_vba_macro` acepta `timeout_s` (default 120, cap 600) y las tools lentas
> (`refresh_power_query`, `refresh_pivot_tables`, `refresh_data_model`,
> `document_workbook`) usan 600 s automáticamente. Además, si EXCEL.EXE muere, la
> sesión se reinicia sola en la siguiente llamada (recovery pre-flight) o lanza
> `ExcelCrashedError` accionable si murió a mitad de operación — ya no queda rota
> hasta un `close_excel` manual. Limitación que persiste: tras un TimeoutError el
> STA sigue ocupado con la tarea colgada; `close_excel` es la vía de escape (y si
> el STA está bloqueado, el shutdown mata el proceso por PID).
