# Limitaciones del MCP excel-2013 detectadas en producción

**Origen:** sesión del 2026-07-24, construcción del *Informe ventas Nutribella países*
(fuentes: `MULTIFORMATO ... JULIO.xlsb` 12,6 MB + `Estimados2.xlsx` con 37 vínculos externos).
**Objetivo:** resolver estos puntos en el desarrollo del MCP. Cada ítem trae
*síntoma → causa raíz (archivo:línea) → fix propuesto → prioridad*.

Ruta del código: `excel-mcp-server-2013/src/excel_mcp_2013/`

---

## Resumen (prioridad)

| # | Limitación | Impacto en la sesión | Estado |
|---|------------|----------------------|--------|
| 1 | `open_workbook` se cuelga: diálogo **"Conflicto de nombres"** por nombres definidos rotos (NO era update-links) | Bloqueó Excel 120 s; hubo que matar el proceso | ✅ **RESUELTO** (saneo auto) |
| 2 | Una llamada COM colgada tumba TODA la sesión (STA único) | `get_session_info` también expiró; se perdió el libro abierto | ✅ **RESUELTO** (fail-fast por diálogo + kill inmediato + timeout dedicado) |
| 3 | Sin lectura masiva de `.xlsb` / tabla por nombre / Data Model | Hubo que extraer con `pyxlsb` fuera del MCP | ✅ **RESUELTO** (v1.4.0: read_table, export_sheet, diagnóstico) |
| 4 | No hay herramienta de recálculo; cálculo forzado a MANUAL | Fórmulas no evaluaban hasta guardar+reabrir | ✅ **RESUELTO** (v1.4.0: recalculate) |
| 5 | Autoría de libros (hojas+estilos) inviable por COM en volumen | Se construyó con `openpyxl`, MCP solo validó | 🟡 Media (workflow) |
| 6 | Nota: la codificación Unicode del MCP funcionó bien | (no es bug — ver abajo) | ⚪ Info |

> **Corrección tras verificar (2026-07-24):** el diagnóstico inicial de #1 (falta de
> `UpdateLinks=0`) resultó **incompleto**. Al probar el fix, el archivo SEGUÍA colgando.
> Enumerando las ventanas del proceso Excel durante el cuelgue apareció el verdadero
> culpable: el diálogo modal `bosa_sdm_XL9` **"Conflicto de nombres"**, disparado por
> **4.887 nombres definidos** (1.886 rotos: `#REF!` / vínculos `[n]`) que colisionan al
> importar los de los 37 vínculos externos. Ese diálogo **no lo suprime
> `DisplayAlerts=False`** ni un watchdog de teclas (probado: 152 Enter no lo cerraron).
> Verificar antes de dar por bueno el fix evitó entregar una solución falsa.

---

## 1. 🔴 `open_workbook` se cuelga indefinidamente con vínculos externos

**Síntoma.** Al abrir `Estimados Sebas.xlsx` / `Estimados2.xlsx` (37 referencias
externas en `xl/externalLinks/`), `open_workbook` expiró a los 120 s con
*"Operacion COM excedio 120s"*. El archivo pesa solo 341 KB.

**Causa raíz (verificada).** El bloqueo es un **diálogo modal** que aparece durante
`Workbooks.Open`, NO la actualización de vínculos. Diagnóstico: se lanzó el `Open` en un
hilo daemon y desde otro hilo se enumeraron las ventanas del PID de Excel. Apareció:

```
('bosa_sdm_XL9', 'Conflicto de nombres', visible=1, enabled=1)
```

`bosa_sdm_XL9` es un diálogo propio de Office (no un `#32770` estándar). Se dispara porque
el archivo arrastra **4.887 nombres definidos** (3.072 ocultos; **1.886 rotos**: valor
`#REF!` o referencia externa `[n]` — basura heredada tipo Lotus `__123Graph_*`, `MRP2`…).
Al abrir con 37 vínculos externos, Excel intenta importar/reconciliar esos nombres y
encuentra conflictos uno tras otro → **un diálogo por conflicto** → cuelgue efectivo.

Por qué los intentos previos NO bastan (todo verificado empíricamente):
- `DisplayAlerts=False` **no** suprime este diálogo (ya estaba activo y colgó igual).
- `UpdateLinks=0` + `AskToUpdateLinks=False` **no** lo evitan (el diálogo reapareció).
- Watchdog de teclas: **152 `Enter`** enviados con `SetForegroundWindow`+`keybd_event` no
  lo cerraron (es un edit-box que re-prompta; frágil e inútil aquí).

**Fix IMPLEMENTADO (saneo determinista, sin UI).** Nuevo módulo
`utils/workbook_sanitize.py` + integración en `session.open_workbook`:

1. **Pre-vuelo estático (sin COM):** `scan_definedname_risk(path)` abre el `.xlsx/.xlsm`
   como ZIP y cuenta vínculos externos y nombres rotos. `risky = OOXML + vínculos>0 +
   rotos≥1`.
2. **Copia saneada:** si hay riesgo (o `sanitize_names="always"`), `make_sanitized_copy`
   crea una copia temporal quitando del `xl/workbook.xml` **solo** los nombres rotos
   (los legítimos —áreas de impresión, `_FilterDatabase`— se conservan) y se abre ESA
   copia. El **original nunca se modifica**; la copia se borra en `close()`.
3. El tool `open_workbook` gana el parámetro `sanitize_names` (`auto`|`always`|`never`,
   default `auto`) y devuelve un bloque `sanitized` con nº de nombres removidos.
4. Se mantiene `UpdateLinks=0` + `AskToUpdateLinks=False` como **higiene** (evitan el
   prompt de vínculos, que es un cuelgue distinto y también común), aunque no eran la
   causa de ESTE caso.

**Verificación (evidencia).**
- End-to-end contra el código real del MCP (`SessionManager.open_workbook`):
  `Estimados Sebas.xlsx` pasó de **colgar >120 s** a **abrir en 5,7 s**, `sanitized=true`,
  **1.864 nombres removidos**, leyendo `A1='CODIGO'`, `A2=5060094`. Archivo normal →
  `sanitized=false` (auto no lo toca). Copia temporal borrada en `close()`.
- `test_sanitize.py` (14 checks, Python puro) y `test_hardening.py` (recovery/timeout)
  pasan sin regresiones.

---

## 2. 🔴 Una llamada COM colgada wedgea toda la sesión; recuperar exige matar Excel

**Síntoma.** Tras el open colgado, la siguiente llamada (`get_session_info`) **también**
expiró a 120 s. La única salida fue `close_excel` (mató el PID), y con ello se perdió
el `MULTIFORMATO.xlsb` que ya tenía abierto en solo-lectura (hubo que reabrirlo).

**Causa raíz.** `com_guard.py` ejecuta todo COM en **un único hilo STA** con cola de
tareas (`execute()`, línea 52-77). Cuando una llamada COM no retorna (diálogo modal /
resolución de vínculos), el hilo STA queda bloqueado *dentro* de esa llamada; el timeout
de Python (`result_q.get(timeout=...)`, línea 69) **no interrumpe la llamada COM**, solo
deja de esperar. Toda tarea encolada detrás hereda el bloqueo y expira una tras otra.

**Fix (defensa en capas).**

1. ✅ **Prevenir (ver #1, IMPLEMENTADO):** el pre-vuelo estático sin COM
   (`scan_definedname_risk`) + saneo a copia temporal elimina el cuelgue de ESTE caso
   (diálogo "Conflicto de nombres") antes de tocar Excel.
2. ✅ **Pre-vuelo estático sin COM (IMPLEMENTADO):** el archivo se inspecciona como ZIP
   antes de `Workbooks.Open` (vínculos externos + nombres rotos). Extensible a detectar
   otros perfiles de riesgo y devolver `warning` en la respuesta.
3. ✅ **Timeout dedicado y corto para `open_workbook` (IMPLEMENTADO):**
   `OPEN_WORKBOOK_TIMEOUT = 45 s` (vs 120 s default). Un open sano tarda segundos; pasarse
   casi siempre es un diálogo. Acota la PRIMERA llamada colgada.
4. ✅ **Fail-fast por detección de diálogo modal (IMPLEMENTADO, mejor que la "bandera"):**
   en vez de una bandera basada en timeout (que daría falsos positivos con operaciones
   lentas legítimas), se usa una **señal positiva**: `session.has_modal_dialog()` enumera
   las ventanas del PID de Excel (clases `#32770` / `bosa_sdm_XL9`) — corre en el hilo del
   caller, sin COM, así que funciona aunque el STA esté colgado. `run_with_excel` hace
   **pre-vuelo**: si hay una tarea en curso (`guard.inflight_info()`) **y** un diálogo
   modal, lanza `STAWedgedError` al instante (*"corré close_excel"*) en vez de esperar.
   Verificado: la llamada siguiente falló en **0,00 s** en vez de esperar su timeout. Una
   operación lenta SIN diálogo (p. ej. un recálculo) sigue encolándose normal (probado en
   `test_hardening`).
5. ✅ **`close_excel` mata de inmediato si detecta diálogo (IMPLEMENTADO):** antes esperaba
   30 s al cierre cooperativo; ahora, si `has_modal_dialog()`, mata por PID directo
   (verificado en **0,28 s**). Sigue siendo todo-o-nada (un solo proceso Excel): matarlo
   cierra también los libros sanos abiertos — documentado en el docstring del tool.

**Verificación (evidencia).** Nuevo `com_guard.inflight_info()` + `session.has_modal_dialog()`;
`test_guard_wedge.py` (8 checks, Excel-free) + e2e contra el módulo real del server
(diálogo detectado, fail-fast 0,00 s, kill 0,28 s, sin huérfanos) + `test_hardening.py`
sin regresiones.

> **Pendiente (menor):** el pre-vuelo cubre diálogos que abren ventana. Un bloqueo COM sin
> ventana (raro) todavía esperaría el timeout de esa llamada antes de que el siguiente lo
> note; `close_excel` sigue siendo la vía de escape universal.

---

## 3. 🟠 Sin lectura masiva de `.xlsb`, tabla por nombre, ni Data Model utilizable

**Síntoma.** Para extraer la hoja `VENTA` (3.406 filas) y `BD` (9.036 filas) del `.xlsb`
tuve que usar `pyxlsb` en Python, fuera del MCP:
- `read_range` sobre miles de filas serían muchísimos round-trips COM lentos.
- No existe "volcar hoja/tabla completa" en una sola llamada.
- `get_data_model_measures` devolvió `[]` pese a que el libro tiene un Data Model
  tabular real (`ThisWorkbookDataModel`) que alimenta fórmulas CUBE
  (`[ANALISIS].[Marca]`, etc.). Las medidas no fueron enumerables.

**Causa raíz.** (a) No hay tool de exportación masiva; `read_range` está pensado para
ventanas, no para volcados. (b) La lectura del Data Model no cubre el caso de medidas
implícitas / modelo no materializado vía COM.

**Fix propuesto.**

1. **`read_table(table_name)`** — lee el `DataBodyRange` de un `ListObject` completo en
   UNA llamada. La hoja VENTA tiene la tabla `VENTAS`; habría sido 1 sola llamada.
2. **`export_sheet(sheet, format=csv|json)`** — usar `Worksheet.UsedRange.Value2`, que
   marshalla todo el array 2D en **una** operación COM (mucho más rápido que leer por
   trozos). Es el "fast path" que hoy no está expuesto.
3. **Diagnóstico en `get_data_model_measures`.** Enumerar `Model.ModelMeasures` y, si
   sale vacío pero `Model.ModelTables` tiene tablas, devolver algo como
   *"modelo presente, sin medidas explícitas (posibles medidas implícitas/cubo)"* en
   lugar de un `[]` silencioso. Considerar introspección vía el propio motor tabular.
4. Documentar que para `.xlsb` masivo el camino más rápido es `Value2` sobre `UsedRange`.

**✅ RESUELTO (v1.4.0, 2026-07-24).** Nuevo `tools/bulk.py`: `read_table(name, dest)`
(ListObject completo por nombre, 2 llamadas COM) y `export_sheet(sheet, dest, range_addr)`
(`UsedRange`/rango en 1 llamada `.Value`). Entrega archivo+muestra (no revienta contexto),
tope 50k inline / 5M export, floats enteros → int (códigos PT sin `.0`). `get_data_model_measures`
ahora devuelve `{measures, diagnostic}` distinguiendo *sin medidas* de *sin modelo*.
Verificado E2E: **BD del MULTIFORMATO (9.033×97 ≈ 876k celdas) exportada en 2,0 s**, suma
fiel (export == COM). `test_bulk_tools.py` 43 checks + regresión sin cambios.

---

## 4. 🟠 No hay herramienta de recálculo y el cálculo se fuerza a MANUAL

**Síntoma.** Tras escribir las fórmulas `SUMAR.SI.CONJUNTO` de la hoja VERIFICACION
(referencias externas al `.xlsb`), los valores no evaluaban; hubo que **guardar +
reabrir con ambos libros abiertos** para obtener resultados. No hay un tool `recalculate`.

**Causa raíz.** `session.py:69-75` fuerza `Application.Calculation = xlCalculationManual`
al abrir cada libro (correcto por performance/seguridad), pero **no existe tool
compañero** que dispare el recálculo. El operador debe adivinar que hay que guardar/reabrir.

**Fix propuesto.**

1. **`recalculate(full: bool = False, sheet: str | None = None)`** →
   `Application.CalculateFull()` / `Application.Calculate()` / `Worksheet.Calculate()`.
   Es prácticamente obligatorio dado que el modo manual está forzado.
2. (Opcional) **`calculate_until_stable`** — iterar `CalculateFull()` hasta
   `Application.CalculationState == xlDone` (útil con dependencias/volátiles).
3. Documentar de forma prominente el default de cálculo manual: sorprende ver
   fórmulas "en blanco/desactualizadas" hasta guardar y reabrir.

**✅ RESUELTO (v1.4.0, 2026-07-24).** `recalculate(full, sheet, wait_async)` en
`tools/workbook.py`: `Calculate` (sucias) / `CalculateFull` / `Worksheet.Calculate`.
`wait_async` maneja la trampa CUBE (`#GETTING_DATA` con Manual): pone automático temporal
+ `CalculateUntilAsyncQueriesDone` y restaura Manual. Timeout 600s (un CalculateFull grande
no es cuelgue). Verificado: en Manual, fórmula desactualizada (A3=5) → `recalculate()` →
A3=13; restaura Manual tras `wait_async`. Docstring advierte el default manual.

---

## 5. 🟡 Autoría de libros con estilos en volumen: hoy se resuelve con openpyxl

**Síntoma.** El libro de salida (6 hojas; fuentes, rellenos, formatos numéricos, anchos
de columna, `freeze_panes`, 441×14 fórmulas + una hoja de 3.105 filas de datos) se
construyó con **openpyxl**, usando el MCP solo para abrir, recalcular y validar. Hacerlo
por COM (`apply_format`/`write_range` celda a celda) serían cientos de llamadas lentas.

**Causa raíz / matiz.** No es tanto un bug como una **división de trabajo**:
- `openpyxl` es ideal para *autoría* de `.xlsx` (rápido, sin COM), pero **no lee `.xlsb`**
  ni preserva el Data Model.
- El MCP (COM) es imprescindible para lo que openpyxl no puede: `.xlsb`, Data Model,
  tablas dinámicas nativas, VBA, Power Query y **recálculo real**.

**Fix propuesto (enhancement).**

1. Documentar el flujo recomendado: *openpyxl para armar el `.xlsx`; MCP para lo COM-only*.
2. Si se quiere cubrir estilos en el MCP, permitir que `apply_format` reciba un **rango**
   y aplique el estilo con un único objeto `Range` (evita el bucle por celda).
   `write_range`/`write_formulas` ya aceptan array 2D (bien); el hueco es el **estilo masivo**.

---

## 6. ⚪ Nota (no es bug): la codificación Unicode del MCP funcionó bien

Los acentos (`PERÚ`, `Línea`) se vieron corruptos (`PER�`) **solo** al leer con
`pyxlsb`/Bash en Latin-1. El MCP (`read_range`) devolvió el Unicode correcto. Se anota
para no perseguir un falso bug: es un punto a favor del MCP, la corrupción estaba en la
ruta alternativa fuera de él.

---

## Backlog (ideas evaluadas, NO abordar sin diseño propio)

### `refresh_all` orquestador (PQ → Data Model → pivots)

Sugerido en la revisión del spec del paquete #3+#4 (2026-07-24). **Rechazado para ese
paquete** por estas razones, que cualquier implementación futura debe resolver:

- **Footgun en nuestro entorno:** los libros del trabajo tienen conexiones a fuentes
  corporativas inaccesibles desde el PC personal. Refrescar a ciegas cuelga o, peor,
  destruye los valores cacheados (en la sesión Nutribella se evitó refrescar a propósito).
- La detección de "terminó" difiere por tipo (`BackgroundQuery`, refresh asíncrono del
  modelo, `CalculateUntilAsyncQueriesDone`) y el refresh es de las zonas COM más
  propensas a cuelgues (quirks Mashup/QueryTables en 2013).
- Diseño mínimo requerido: **dry-run** que liste conexiones y su alcanzabilidad ANTES de
  tocar nada, timeouts por paso, política skip-on-fail explícita, y resumen por conexión.

### Re-assert defensivo de `ScreenUpdating` (descartado como feature)

`ScreenUpdating=False`/`EnableEvents=False` ya se fijan globalmente en `session.start()`
— no hay ganancia por-operación (sugerencia evaluada y descartada; solo se acepta un
re-assert de 1 línea en operaciones largas como blindaje contra macros que lo reactiven).

---

## Quick wins (orden sugerido de implementación)

1. **#1** — `UpdateLinks=0` + `AskToUpdateLinks=False` (2 líneas; elimina el bloqueante más común).
2. **#2.2 + #2.4** — pre-vuelo ZIP de vínculos externos + bandera "wedged" con fallo rápido.
3. **#4.1** — tool `recalculate` (imprescindible dado el cálculo manual forzado).
4. **#3.1 + #3.2** — `read_table` y `export_sheet` (Value2) para lectura masiva.
5. **#3.3** — diagnóstico honesto en `get_data_model_measures`.

*Los ítems #1, #2 y #4 son de bajo esfuerzo y alto impacto; con ellos, esta misma
sesión se habría hecho end-to-end dentro del MCP sin caer a `pyxlsb`/`openpyxl`.*
