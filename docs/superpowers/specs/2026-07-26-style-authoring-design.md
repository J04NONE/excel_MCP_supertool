# Spec: Autoría de estilos en volumen (paquete v1.6.0)

**Fecha:** 2026-07-26
**Origen:** `LIMITACIONES_MCP.md` §5 — el informe Nutribella se estilizó con openpyxl
porque el MCP no cubría layout (anchos, freeze) y estilizar N rangos costaba N tool calls.
**Repo:** `excel-mcp-server-2013/` — versión objetivo **1.6.0** (desde 1.5.0).

## Corrección al diagnóstico de §5

`apply_format` NO aplica "celda a celda": ya opera sobre el objeto `Range` completo.
Los huecos reales, medidos contra el script `_build_informe.py` que sí se usó:

1. **Propiedades ausentes**: `column_width` (usado 4 veces), `freeze_panes` (usado),
   `font_name`, alineación, wrap, bordes, merge, `row_height`. Ninguna existe en el MCP.
2. **Volumen**: cada rango estilizado = 1 tool call (1 viaje STA + JSON-RPC). Un informe
   con ~40 zonas de estilo son ~40 llamadas. Medido: 60 rangos × 3 props por COM directo
   = 0,6 s — el coste está en los viajes, no en COM.

## Diseño (3 cambios, módulo `tools/cells.py`)

### 1. `apply_format` extendido (compatible hacia atrás)

Props nuevas, todas opcionales: `font_name`, `h_align` (left|center|right|justify),
`v_align` (top|center|bottom), `wrap_text`, `border` (thin|medium|thick|none — aplica
a todas las aristas del rango), `column_width`, `row_height`, `merge`.

### 2. `apply_format_batch(sheet, formats)` — nuevo

`formats` = lista de dicts `{"range": "A1:F1", ...mismas claves que apply_format...}`.
Un solo tool call / un solo viaje STA para todo el estilo de una hoja.

- **Validación pre-vuelo de TODO el batch** (claves y valores) ANTES de tocar la hoja:
  una spec inválida no deja el formato a medias.
- Error COM al aplicar un item (raro, ya validado): se registra y se continúa; la
  respuesta trae `{applied, failed, details}`.
- `sheet` a nivel de llamada; un item puede sobreescribir con `"sheet"` propia
  (informes multi-hoja en un solo call).

### 3. `set_freeze_panes(sheet, at=None)` — nuevo

`at="C5"` congela filas 1-4 y columnas A-B (semántica openpyxl/Excel). `at=None`
descongela. Implementación verificada en instancia oculta SIN `Select`:
`ws.Activate()` + `win.SplitRow/SplitColumn` + `win.FreezePanes=True`.
Resolver "C5" → (row, col) vía `ws.Range(at).Row/.Column` (no parsear A1 a mano).

### Constantes (verificadas contra Excel real)

`xlContinuous=1`, `xlLineStyleNone=-4142`; Weight: thin=2, medium=-4138, thick=4;
HAlign: left=-4131, center=-4108, right=-4152, justify=-4130;
VAlign: top=-4160, center=-4108, bottom=-4107.

## Fuera de alcance (YAGNI)

Estilos con nombre, formato condicional, outline de caja (4 aristas por separado),
temas, y autoría masiva de `.xlsx` sin Excel — para eso la división de trabajo con
openpyxl sigue siendo válida y queda documentada en §5.

## Pruebas (`test_format_tools.py`, patrón test_hardening)

1. `apply_format` legacy sin regresión (bold/fill/number_format).
2. Props nuevas aplican y se VERIFICAN leyendo de vuelta por COM
   (ColumnWidth, HorizontalAlignment, WrapText, Borders.LineStyle, MergeCells).
3. `border="none"` limpia bordes existentes.
4. Batch: N items en un call; item con `sheet` propia; validación pre-vuelo
   (clave desconocida / valor inválido → ValueError SIN aplicar nada, verificado).
5. Freeze: `at="C5"` → SplitRow=4/SplitColumn=2/FreezePanes=True; `at=None` descongela;
   hoja inexistente → error claro.
6. Regresión: 5 suites previas verdes. Limpieza: sin EXCEL.EXE huérfanos.
