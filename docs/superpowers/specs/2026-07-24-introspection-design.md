# Spec: Introspección para desenmarañar libros (paquete v1.5.0)

**Fecha:** 2026-07-24
**Origen:** `LIMITACIONES_MCP.md` — gaps de introspección detectados en la sesión Nutribella
(buscar sin tool, 4.887 nombres definidos descubiertos por XML a mano, vínculos/conexiones
vistos con `unzip`). Diseño revisado (self-review externo + análisis de conexión OLAP real).
**Repo:** `excel-mcp-server-2013/` — versión objetivo **1.5.0** (desde 1.4.0).

## Objetivo

Cuatro tools para entender un libro ajeno sin salir del MCP:

1. `search_workbook` — grep de Excel (valores/fórmulas, nativo + regex).
2. `list_defined_names` — inventario de nombres definidos, con marca de rotos.
3. `clean_defined_names` — borrar permanentemente los nombres rotos (escritura).
4. `list_connections` — mapear de dónde sale la data (conexiones + vínculos), incluidas
   las fuentes de intranet **inalcanzables** (SSAS/OLAP, UNC), SIN intentar conectarse.

**Módulo nuevo:** `tools/introspect.py`. Reusa `_BROKEN_RE` de `utils/workbook_sanitize`
(criterio de "roto" = `#REF!` o referencia externa `[n]`), `matrix_to_jsonable`,
`get_active_workbook`, `get_sheet`.

**Restricción global:** API COM compatible Excel 2013. **Ningún tool debe sondear la red**
(evita el cuelgue tipo SMB/OLAP que ya blindamos en #1/#2).

---

## Componente 1: `search_workbook`

### Firma

```python
search_workbook(query: str, where: str = "both", regex: bool = False,
                match_case: bool = False, whole_cell: bool = False,
                sheet: Optional[str] = None, max_results: int = 500) -> dict
```

- `where`: `"values"` | `"formulas"` | `"both"`.
- `regex=False` → motor nativo `Cells.Find`. `regex=True` → leer + escanear (Python `re`).
- `sheet=None` → todas las hojas (incluidas ocultas/muy ocultas — para desenmarañar sí).

### Alcance (declarar en docstring)

Busca SOLO en celdas (valores y/o fórmulas). NO busca en: nombres de hoja, títulos de
shapes/charts, comentarios, ni código VBA. Para nombres → `list_defined_names`; para VBA →
`get_vba_code`.

### Motor nativo (`regex=False`)

Constantes: `xlValues=-4163`, `xlFormulas=-4123`, `xlWhole=1`, `xlPart=2`, `xlByRows=1`.

Por cada hoja (o la indicada) y por cada `LookIn` según `where`:

```python
first = None
cell = ws.Cells.Find(What=query, LookIn=look_in, LookAt=(xlWhole if whole_cell else xlPart),
                     MatchCase=match_case, SearchOrder=xlByRows)
while cell is not None:
    addr = str(cell.Address)          # $A$1
    if first is None:
        first = addr
    elif addr == first:
        break                          # dio la vuelta -> cortar (evita bucle infinito)
    # registrar (dedupe por (hoja, addr) entre pasadas values/formulas)
    ...
    if len(matches) >= max_results:
        break
    cell = ws.Cells.FindNext(cell)
```

- `where="both"` → dos pasadas (xlValues + xlFormulas), **dedupe por (hoja, addr)**,
  `matched_in` refleja en cuál se encontró (si en ambas → `"value+formula"`).
- `Find` devuelve `None` si no hay match (no error).

### Motor regex (`regex=True`)

- Por hoja: `ur = ws.UsedRange`; `n = Rows.Count*Columns.Count`. Si `n > MAX_SEARCH_CELLS
  (2_000_000)` → error accionable ("hoja X excede el techo regex; acota con sheet= o rango").
- Leer `.Value` (para values) y/o `.Formula` (para formulas). La dirección de un match se
  obtiene con `ws.Cells(ur.Row + i, ur.Column + j).Address` (1 llamada COM **solo por match**,
  no por celda; los matches están capados) — evita construir un conversor col→letra y NO
  asume que UsedRange empiece en A1.
- `re.search(query, str(cell), flags=re.IGNORECASE if not match_case else 0)`.
- `whole_cell` se ignora en regex (el patrón manda; usar `^...$` si se quiere ancla).

### Celdas en error (self-review 🟠)

Celdas `#GETTING_DATA` (cubos intranet caídos) o `#REF!` devuelven un `VT_ERROR` en COM.
`matrix_to_jsonable`/`to_jsonable` deben convertirlo a su string (`"#GETTING_DATA"`,
`"#REF!"`, `"Error 2042"` como fallback) sin lanzar. **El spec exige un test de esto**
(hoy `to_jsonable` no está probado contra errores).

### Salida

```json
{"query": "...", "count": 12, "truncated": false,
 "matches": [{"sheet": "VENTA", "cell": "$BC$12", "value": "1.19...",
              "formula": "=...", "matched_in": "value"}]}
```

`value`/`formula` truncados a 200 chars. `truncated=true` si se cortó por `max_results`.

---

## Componente 2: `list_defined_names`

### Firma

```python
list_defined_names(broken_only: bool = False) -> dict
```

### Enumeración

- Nombres de libro: iterar `wb.Names`.
- Nombres de hoja (scope local): iterar `ws.Names` de cada hoja.
- **Verificar en implementación** si `wb.Names` ya incluye los sheet-scoped (en algunos
  hosts sí, con `.Name` tipo `Hoja1!Local`). Dedupe por (scope, name) para no duplicar.
- Por cada `Name`: `{name, refers_to, visible, scope, broken, builtin}`.
  - `refers_to` = `str(nm.RefersTo)` envuelto en try/except (algunos lanzan al leer).
  - `visible` = `bool(nm.Visible)`.
  - `scope` = `"workbook"` o `"sheet:<hoja>"` (por `nm.Parent`; si es Worksheet → sheet).
  - `broken` = `_BROKEN_RE.search(refers_to)` (=`#REF!` o `[n]`).
  - `builtin` = el nombre empieza por `_xlnm.` (Print_Area, _FilterDatabase, etc.).

### Salida (conteos + muestra — NO volcar 4.887)

```json
{"total": 4887, "broken": 1886, "hidden": 3072, "builtin": 5,
 "workbook_scoped": 4880, "sheet_scoped": 7,
 "names": [ {..hasta 300..} ], "truncated": true}
```

`broken_only=True` → `names` trae solo los rotos (igualmente capado a 300 con `truncated`).
Los conteos (`total/broken/hidden/builtin`) SIEMPRE son sobre el universo completo, no la
muestra.

---

## Interacción CRÍTICA con el saneo automático (afecta #2 y #3)

`open_workbook` con `sanitize_names="auto"` (default) abre una **copia temporal SIN los
nombres rotos** cuando el archivo tiene vínculos externos + nombres rotos (ej.
`Estimados Sebas.xlsx`). Consecuencia obligada de fijar en el spec:

- Sobre un archivo **auto-saneado**, `list_defined_names(broken_only=True)` verá **~0 rotos**
  (ya se quitaron en la copia) y `clean_defined_names` no tendrá qué borrar. Esto NO es un
  bug: es coherente — el saneo ya hizo el trabajo en la copia de trabajo; para **persistirlo**
  al archivo real se usa `save_workbook(<ruta original>)`.
- `list_defined_names`/`clean_defined_names` son útiles y operan directo cuando el archivo
  **NO dispara auto-saneo**: libros con nombres rotos pero **sin vínculos externos** (abren
  sin colgar) — ahí sí hay rotos que inspeccionar y limpiar en vivo.
- Para inspeccionar los rotos del archivo ORIGINAL sin abrirlo por COM, existe la función
  pura `scan_definedname_risk` (offline, sobre el XML) — la usa el E2E para validar conteos
  sin depender del COM ni arriesgar cuelgue.
- `list_connections` NO se ve afectado: el saneo solo toca `<definedNames>`, no las
  conexiones — funciona igual sobre la copia saneada.

El docstring de `clean_defined_names` debe declarar esta interacción.

## Componente 3: `clean_defined_names` ⚠️ escritura destructiva

### Firma

```python
clean_defined_names(broken_only: bool = True, include_builtin: bool = False,
                    save: bool = False) -> dict
```

### Reglas de seguridad (self-review 🔴/🟠)

1. **Read-only:** si `wb.ReadOnly` es True → error claro ("el libro está en solo lectura;
   reábrelo con read_only=False para limpiar"). No intentar borrar.
2. **Preservar built-in por defecto:** los `_xlnm.*` (áreas de impresión, _FilterDatabase)
   NO se borran salvo `include_builtin=True`, aunque estén `#REF!` — borrarlos puede romper
   el área de impresión/filtro o lanzar.
3. **Patrón de borrado seguro (colección que muta):** primero capturar la LISTA de nombres
   a borrar como strings + su scope; luego borrar uno a uno resolviendo de nuevo el objeto:
   ```python
   objetivos = [ (nm.Name, scope) for nm in _iter_names(wb)
                 if _match(nm, broken_only, include_builtin) ]
   removed = 0
   for name, scope in objetivos:
       try:
           _resolve_name(wb, name, scope).Delete()
           removed += 1
       except Exception:
           logger.debug("No se pudo borrar %s", name)
   ```
   NUNCA iterar objetos `Name` vivos mientras se borra (la colección se reindexa).
4. **`save=False` por defecto:** el usuario decide guardar (respeta el control de escritura;
   ver [[no-commit-no-push]] — misma filosofía para escrituras de archivos).

### Salida

```json
{"removed": 1886, "remaining": 3001, "broken_only": true,
 "include_builtin": false, "saved": false}
```

`remaining` = `wb.Names.Count` tras el borrado (recuento real, no cálculo).

---

## Componente 4: `list_connections`

### Firma

```python
list_connections() -> dict
```

### Qué enumera

- `wb.Connections` (OLE DB, ODBC, Power Query/Mashup, Data Model, texto…). Si
  `wb.Connections.Count == 0` → `connections: []`.
- Vínculos Excel→Excel: `wb.LinkSources(xlExcelLinks=1)` (rutas de libros vinculados).
  **Devuelve `None`** (no lista vacía) cuando no hay vínculos → tratar como `[]`.

### Por conexión (parseo estructurado, NO string crudo)

Constantes de tipo (verificar exactas en implementación): `xlConnectionTypeODBC=1`,
`xlConnectionTypeOLEDB=2`, `xlConnectionTypeXMLMAP=4`, `xlConnectionTypeWORKSHEET=5`,
`xlConnectionTypeTEXT=6`, `xlConnectionTypeMODEL=7`.

```python
{
  "name": conn.Name,
  "description": conn.Description or "",
  "type": "OLAP/Cubo (SSAS)" | "OLE DB" | "ODBC" | "Power Query" | "Data Model" | ...,
  "provider": "MSOLAP.5",            # de la cadena
  "data_source": "OLAPSERVER",       # de la cadena
  "initial_catalog": "VentasCorp",   # de la cadena
  "command_type": "cube|sql|table|default",
  "command_text": "VentasCorp Indicadores Ventas Diario Logistica",
  "odc_file": conn.OLEDBConnection.SourceConnectionFile or None,
  "refresh_on_open": bool,
  "reachable": "local:existe" | "local:no_existe" | "no_verificable (servidor)" |
               "no_verificable (red/UNC)" | "n/a",
  "integrated_security": true,
  "had_credentials": false
}
```

- Acceso: `conn.OLEDBConnection.Connection` (cadena), `.CommandText`, `.CommandType`,
  `.SourceConnectionFile`, `.RefreshOnFileOpen`. Envolver TODO en try/except por tipo
  (una conexión WORKSHEET/MODEL/TEXT no tiene `OLEDBConnection`).
- Clasificación `type`: si `Provider` empieza por `MSOLAP` → "OLAP/Cubo (SSAS)"; si por
  `Microsoft.Mashup` → "Power Query"; conexión al modelo → "Data Model"; según `conn.Type`
  el resto.

### Parseo y sanitización de la cadena (self-review 🔴 seguridad)

`_parse_conn_string(s)` → dict case-insensitive de `key=value` (split por `;`, respetar
comillas). De ahí se toman provider/data source/initial catalog/location.

- **Eliminar** los valores de claves de credencial: `Password`, `PWD`, `User ID`, `UID`.
- `integrated_security` = clave `Integrated Security` presente (p.ej. `SSPI`).
- `had_credentials` = existe `Password`/`PWD` con valor no vacío.
- NUNCA devolver la cadena cruda (evita fuga de credenciales/servidores internos).

### Alcanzabilidad SIN sondeo de red (self-review 🔴 cuelgue)

Función `_reachability(target)` que **jamás toca la red**:

- `target` vacío / server name sin separador de ruta → `"no_verificable (servidor)"`
  (ej. `Data Source=OLAPSERVER`).
- Empieza por `\\` (UNC) → `"no_verificable (red/UNC)"` — **NO** `os.path.exists`.
- Ruta con letra de unidad local (`C:\`, `D:\`) → `os.path.exists` (rápido; un `D:\pjimenez`
  inexistente da `local:no_existe` al instante).
- Se aplica tanto al `data_source` (si es archivo, p.ej. `.accdb`/`.xlsx`) como al `odc_file`.

### Salida

```json
{"count": 6, "connections": [ {..} ],
 "excel_links": [ {"source": "C:\\...\\Libro.xlsx", "reachable": "local:no_existe"} ]}
```

### Validación con datos reales (de la captura analizada)

La conexión `OLAPSERVER VentasCorp … Logistica31` debe salir con `type=OLAP/Cubo (SSAS)`,
`provider=MSOLAP.5`, `data_source=OLAPSERVER`, `initial_catalog=VentasCorp`,
`command_type=cube`, `reachable="no_verificable (servidor)"`, `integrated_security=true`,
`had_credentials=false`. Es lo que explica los `#GETTING_DATA` del MULTIFORMATO.

---

## Cableado y transversales

1. `tools/introspect.py` nuevo con `register(mcp, session, run)`; añadir al bloque de
   imports/registros de `server.py`.
2. Constantes en `introspect.py`: `MAX_SEARCH_CELLS = 2_000_000`, `MATCH_CAP = 500`,
   `NAMES_SAMPLE = 300`, `VALUE_TRUNC = 200`.
3. Bump de versión en `server.py`: `1.4.0 → 1.5.0`.
4. `TOOLS.md`: 4 tools nuevos (sección "Introspección").
5. `LIMITACIONES_MCP.md`: documentar la frontera **intranet/OLAP fuera de alcance por red
   (limitación externa, no del MCP)** — el tool las mapea y explica, no las cruza; evidencia:
   la conexión SSAS de la captura. Marcar el gap de introspección como resuelto.

## Plan de pruebas

### `test_introspect.py` (Excel real, patrón test_hardening)

1. **Fixture:** libro con: hoja de datos con valores repetidos y fórmulas; 3 nombres
   definidos (1 válido `Area_Buena`, 1 roto `#REF!`, 1 built-in-roto `_xlnm.Print_Area`);
   una celda con error real (`=NA()` → `#N/D`, sustituto seguro de `#GETTING_DATA`).
2. **search nativo:** `search_workbook("H1", where="both")` encuentra header; en fórmulas
   `search_workbook("SUMA"/"SUM", where="formulas")`; `whole_cell=True` vs substring;
   `max_results` capa y marca `truncated`.
3. **search dedupe:** un valor que también aparece en una fórmula → `matched_in="value+formula"`,
   una sola entrada por celda.
4. **search regex:** `search_workbook(r"H\d+", regex=True)` matchea H1..H10; techo: hoja
   sintética > 2M celdas → error accionable.
5. **search error-cell:** la celda `#N/D` no rompe; se puede buscar `"#N/D"` / `"#N/A"`.
6. **list_defined_names:** `total/broken/builtin` correctos; `broken_only=True` filtra;
   conteos sobre el universo (no la muestra).
7. **clean_defined_names:** borra el roto `Area_Rota`, **preserva** `_xlnm.*` (con
   `include_builtin=False`), preserva `Area_Buena`; `remaining` correcto; `save=False` no
   escribe a disco (mtime del archivo sin cambios si el libro estaba guardado); read-only →
   error.
8. **list_connections:** con una conexión OLE DB creada por COM (o al menos sin conexiones →
   `count:0`); `_parse_conn_string` y `_reachability` probados como funciones puras
   (unit, sin Excel) con la cadena EXACTA de la captura → verifica parseo, credenciales,
   `no_verificable (servidor)`.
9. Limpieza total (sin EXCEL.EXE huérfanos, sin temporales).

### E2E contra archivos reales

- `search_workbook("5060094")` en el MULTIFORMATO → ubica el código PT (hoja/celda).
- `search_workbook("#GETTING_DATA")` → celdas cubo afectadas (si las hay tras abrir).
- `list_defined_names` en `Estimados Sebas.xlsx` abierto normal (auto-saneado) → dará
  **pocos rotos** (el saneo ya los quitó) — eso VALIDA la interacción documentada, no es
  fallo. El conteo real de ~1.886 rotos del ORIGINAL se valida con `scan_definedname_risk`
  (función pura, offline) — NO abrir con `sanitize_names="never"` (colgaría).
- `list_connections` en el MULTIFORMATO → debe listar la conexión SSAS con
  `reachable="no_verificable (servidor)"` sin colgarse.

### Regresión

`test_bulk_tools.py`, `test_sanitize.py`, `test_guard_wedge.py`, `test_hardening.py` — verdes.

## Fuera de alcance (YAGNI)

- Búsqueda en shapes/charts/comentarios/VBA (ya hay `get_vba_code`); reemplazo (find&replace);
  conectar/refrescar/validar conexiones (footgun, backlog `refresh_all`); resolver credenciales
  o alcanzar fuentes de intranet (limitación externa).

## Riesgos aceptados

- `list_defined_names`/`clean` sobre 4.887 nombres: la enumeración COM de ~5k nombres puede
  tardar 1-3 s (aceptable). El borrado de ~1.886 uno a uno, algunos segundos.
- `search` regex "both" lee `.Value` y `.Formula` (doble) — acotado por el techo 2M/hoja.
- `clean_defined_names` es destructivo: mitigado por `save=False` (nada se persiste sin que
  el usuario guarde) y por preservar built-ins.
