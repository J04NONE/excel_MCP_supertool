"""
Restricciones del lenguaje M para Power Query en Excel 2013 (v2.62.5222.701)
=======================================================================
Power Query para Excel 2013 es un add-in legacy congelado en la version 2.x del motor M.
Muchas funciones modernas introducidas post-2016 NO estan disponibles.

Uso:
    from utils.m_constraints import (
        BLOCKED_FUNCTIONS,
        SYSTEM_PROMPT_EXTENSION,
        validate_m_expression,
        get_alternative,
    )
"""

# =============================================================================
# LISTA COMPLETA DE FUNCIONES M NO DISPONIBLES EN EXCEL 2013
# =============================================================================
# Estas funciones producen el error:
# "Expression.Error: The name 'X' wasn't recognized"
# =============================================================================

BLOCKED_FUNCTIONS = {
    # -------------------------------------------------------------------------
    # FUNCIONES DE TEXTO (introducidas 2016-2019)
    # -------------------------------------------------------------------------
    # Alternativa legacy: usar Text.PositionOf + Text.Start/End/Range
    "Text.BeforeDelimiter": {
        "alternative": "Text.PositionOf + Text.Start",
        "example_legacy": 'Text.Start(text, Text.PositionOf(text, delimiter) - 1)',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.AfterDelimiter": {
        "alternative": "Text.PositionOf + Text.Range",
        "example_legacy": 'Text.Range(text, Text.PositionOf(text, delimiter) + Text.Length(delimiter))',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.BetweenDelimiters": {
        "alternative": "Text.PositionOf (inicio y fin) + Text.Range",
        "example_legacy": 'let\n  start = Text.PositionOf(text, startDelim) + Text.Length(startDelim),\n  end = Text.PositionOf(text, endDelim, start),\n  result = Text.Range(text, start, end - start)\nin result',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.Select": {
        "alternative": "Text.Remove + List.Select",
        "example_legacy": 'Text.Remove(text, (c) => not List.Contains(allowedChars, c))',
        "introduced": "2018",
        "category": "Text"
    },
    "Text.Remove": {
        "alternative": "Text.Replace + cada caracter a remover con ''",
        "example_legacy": 'Text.Replace(Text.Replace(text, "a", ""), "b", "")',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.RemoveRange": {
        "alternative": "Text.Start + Text.Range (offset+count al final)",
        "example_legacy": 'Text.Start(text, offset) & Text.Range(text, offset + count)',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.PadStart": {
        "alternative": "Text.Repeat + Text.Length + concatenacion",
        "example_legacy": 'Text.Repeat(padChar, desiredLength - Text.Length(text)) & text',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.PadEnd": {
        "alternative": "Text.Repeat + Text.Length + concatenacion",
        "example_legacy": 'text & Text.Repeat(padChar, desiredLength - Text.Length(text))',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.At": {
        "alternative": "Text.Range(text, position, 1)",
        "example_legacy": 'Text.Range(text, position, 1)',
        "introduced": "2018",
        "category": "Text"
    },
    "Text.Contains": {
        "alternative": "Text.PositionOf(text, substring) <> -1",
        "example_legacy": 'Text.PositionOf(text, substring) <> -1',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.EndsWith": {
        "alternative": "Text.Range(text, Text.Length(text) - Text.Length(substring)) = substring",
        "example_legacy": 'Text.Range(text, Text.Length(text) - Text.Length(substring)) = substring',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.StartsWith": {
        "alternative": "Text.Start(text, Text.Length(substring)) = substring",
        "example_legacy": 'Text.Start(text, Text.Length(substring)) = substring',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.Reverse": {
        "alternative": "List.Accumulate(List.Reverse(Text.ToList(text)), '', (s,c) => s & c)",
        "example_legacy": 'List.Accumulate(List.Reverse(Text.ToList(text)), "", (s,c) => s & c)',
        "introduced": "2018",
        "category": "Text"
    },
    "Text.SplitAny": {
        "alternative": "List.Transform + Splitter.SplitTextByEachDelimiter",
        "example_legacy": 'Splitter.SplitTextByAnyDelimiter(delimiters)(text)',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.TrimEnd": {
        "alternative": "Text.Trim (solo quita espacios al final, no funciona con caracteres)",
        "example_legacy": 'Text.Trim(text)   -- NOTA: solo espacios, no chars personalizados',
        "introduced": "2017",
        "category": "Text"
    },
    "Text.TrimStart": {
        "alternative": "Text.Trim (solo quita espacios al inicio, no funciona con caracteres)",
        "example_legacy": 'Text.Trim(text)   -- NOTA: solo espacios, no chars personalizados',
        "introduced": "2017",
        "category": "Text"
    },

    # -------------------------------------------------------------------------
    # FUNCIONES DE TABLA (introducidas 2016-2019)
    # -------------------------------------------------------------------------
    # Alternativa legacy: usar Table.SelectColumns + Table.TransformColumns
    "Table.Profile": {
        "alternative": "Table.TransformColumns + estadisticas manuales",
        "example_legacy": 'manually compute min/max/count per column using List.Min, List.Max, List.Count',
        "introduced": "2017",
        "category": "Table"
    },
    "Table.Schema": {
        "alternative": "Type.TableSchema o Table.TransformColumnTypes + metadatos manuales",
        "example_legacy": 'Table.TransformColumnTypes(tabla, {{"col", type text}})',
        "introduced": "2017",
        "category": "Table"
    },
    "Table.Buffer": {
        "alternative": "Table.Buffer NO disponible; usar Table.StopFolding si existe",
        "example_legacy": 'NO DISPONIBLE - no hay alternativa directa',
        "introduced": "2017",
        "category": "Table",
        "note": "Table.Buffer en 2013 no existe. Comportamiento diferente en legacy."
    },
    "Table.View": {
        "alternative": "No disponible - no reemplazable",
        "example_legacy": 'NO DISPONIBLE',
        "introduced": "2018",
        "category": "Table"
    },
    "Table.ViewError": {
        "alternative": "No disponible - no reemplazable",
        "example_legacy": 'NO DISPONIBLE',
        "introduced": "2018",
        "category": "Table"
    },
    "Table.ReplaceErrorValues": {
        "alternative": "Table.ReplaceValue con campo error detectado",
        "example_legacy": 'Table.ReplaceValue(tabla, null, valor, Replacer.ReplaceValue, {"col"})',
        "introduced": "2017",
        "category": "Table"
    },
    "Table.MatchesAnyRows": {
        "alternative": "Table.SelectRows + Table.RowCount > 0",
        "example_legacy": 'Table.RowCount(Table.SelectRows(tabla, condicion)) > 0',
        "introduced": "2019",
        "category": "Table"
    },
    "Table.MatchesAllRows": {
        "alternative": "Table.SelectRows + Table.RowCount = Table.RowCount(original)",
        "example_legacy": 'Table.RowCount(Table.SelectRows(tabla, condicion)) = Table.RowCount(tabla)',
        "introduced": "2019",
        "category": "Table"
    },
    "Table.AddRankColumn": {
        "alternative": "Table.AddColumn + Table.Sort + List.PositionOf manual",
        "example_legacy": 'let\n  sorted = Table.Sort(tabla, {{"col", Order.Descending}}),\n  ranked = Table.AddIndexColumn(sorted, "Rank", 1, 1)\nin ranked',
        "introduced": "2017",
        "category": "Table"
    },

    # -------------------------------------------------------------------------
    # FUNCIONES DE LISTA (introducidas 2016-2019)
    # -------------------------------------------------------------------------
    "List.Buffer": {
        "alternative": "No disponible - no hay alternativa directa",
        "example_legacy": 'NO DISPONIBLE',
        "introduced": "2017",
        "category": "List"
    },
    "List.Combine": {
        "alternative": "@(list1 & list2) -- operador de concatenacion",
        "example_legacy": 'list1 & list2',
        "introduced": "2017",
        "category": "List"
    },
    "List.Range": {
        "alternative": "List.FirstN + List.Skip",
        "example_legacy": 'List.FirstN(List.Skip(lista, offset), count)',
        "introduced": "2017",
        "category": "List"
    },
    "List.Split": {
        "alternative": "List.Transform con List.Range manual",
        "example_legacy": 'List.Transform({0..List.Count(lista)/size-1}, each List.Range(lista, _*size, size))',
        "introduced": "2017",
        "category": "List"
    },
    "List.TransformMany": {
        "alternative": "List.Transform + List.Accumulate anidado",
        "example_legacy": 'List.Accumulate(lista, {}, (s, c) => s & List.Transform(otraLista, cada transformacion))',
        "introduced": "2017",
        "category": "List"
    },
    "List.Accumulate": {
        "alternative": "List.Generate **NO DISPONIBLE** -- usar List.Transform + recursion manual",
        "example_legacy": 'funcion recursiva con @ operador',
        "introduced": "2017",
        "category": "List",
        "note": "List.Accumulate NO esta en 2013. Usar funciones de List.Transform."
    },
    "List.Generate": {
        "alternative": "No disponible - sin alternativa directa. Usar @ recursivo.",
        "example_legacy": 'funcion recursiva manual con @fxName',
        "introduced": "2017",
        "category": "List"
    },

    # -------------------------------------------------------------------------
    # FUNCIONES DE FECHA Y DURACION (introducidas post-2016)
    # -------------------------------------------------------------------------
    "Date.ToText": {
        "alternative": "DateTime.ToText (version legacy)",
        "example_legacy": 'DateTime.ToText(datetime)  -- formato limitado',
        "introduced": "2017",
        "category": "Date"
    },
    "Duration.TotalDays": {
        "alternative": "Duration.Days * 1.0",
        "example_legacy": 'Duration.Days(duration) * 1.0 + Duration.Hours(duration) / 24',
        "introduced": "2016",
        "category": "Date"
    },

    # -------------------------------------------------------------------------
    # FUNCIONES DE BINARIO (introducidas post-2016)
    # -------------------------------------------------------------------------
    "Binary.Buffer": {
        "alternative": "No disponible - no hay alternativa",
        "example_legacy": 'NO DISPONIBLE',
        "introduced": "2017",
        "category": "Binary"
    },
    "Binary.View": {
        "alternative": "No disponible - no hay alternativa",
        "example_legacy": 'NO DISPONIBLE',
        "introduced": "2018",
        "category": "Binary"
    },

    # -------------------------------------------------------------------------
    # FUNCIONES DE TIPO Y META (introducidas post-2016)
    # -------------------------------------------------------------------------
    "Function.From": {
        "alternative": "No disponible - no hay alternativa",
        "example_legacy": 'NO DISPONIBLE',
        "introduced": "2017",
        "category": "Function"
    },
    "Function.ScalarVector": {
        "alternative": "No disponible - no hay alternativa",
        "example_legacy": 'NO DISPONIBLE',
        "introduced": "2018",
        "category": "Function"
    },

    # -------------------------------------------------------------------------
    # FUNCIONES DE URI (introducidas post-2016)
    # -------------------------------------------------------------------------
    "Uri.Parts": {
        "alternative": "Text.Split + parseo manual de la URL",
        "example_legacy": 'parseo manual con Text.Split(url, "/")',
        "introduced": "2017",
        "category": "Uri"
    },
    "Uri.Port": {
        "alternative": "Text.Split + Text.PositionOf para extraer puerto",
        "example_legacy": 'parseo manual con Text.Split(url, ":")',
        "introduced": "2017",
        "category": "Uri"
    },

    # -------------------------------------------------------------------------
    # OTRAS FUNCIONES (introducidas post-2016)
    # -------------------------------------------------------------------------
    "Diagnostics.Trace": {
        "alternative": "No disponible - no hay alternativa",
        "example_legacy": 'NO DISPONIBLE',
        "introduced": "2018",
        "category": "Diagnostics"
    },
    "Diagnostics.EnterScope": {
        "alternative": "No disponible - no hay alternativa",
        "example_legacy": 'NO DISPONIBLE',
        "introduced": "2018",
        "category": "Diagnostics"
    },

    # =========================================================================
    # NOTAS ADICIONALES
    # =========================================================================
    # - Value.ReplaceType, Value.NativeQuery: NO disponibles
    # - Lines.FromText, Lines.ToText: NO disponibles
    # - Csv.Document: disponible pero con menos opciones que version moderna
    # - Web.Page: NO disponible (introducido 2017)
    # - Xml.Document: disponible pero legacy
    # - Json.Document: disponible pero legacy
    # - Excel.Workbook: disponible pero legacy
    # - Record.Combine, Record.Select, Record.TransformFields: limitados
}

# =============================================================================
# FUNCIONES SEGURAS Y RECOMENDADAS PARA EXCEL 2013
# =============================================================================

SAFE_FUNCTIONS = {
    # Texto basico
    "Text.Length", "Text.Start", "Text.End", "Text.Range",
    "Text.PositionOf", "Text.Upper", "Text.Lower", "Text.Trim",
    "Text.Replace", "Text.Reverse", "Text.ToList", "Text.From",
    "Text.Repeat", "Text.Remove", "Text.Split", "Text.NewGuid",
    "Text.PadStart", "Text.PadEnd",

    # Listas basicas
    "List.Count", "List.First", "List.Last", "List.FirstN",
    "List.LastN", "List.Skip", "List.Reverse", "List.Sort",
    "List.Distinct", "List.Transform", "List.Select",
    "List.Contains", "List.PositionOf", "List.RemoveItems",
    "List.RemoveFirstN", "List.RemoveLastN", "List.ReplaceValue",
    "List.StandardDeviation", "List.Sum", "List.Average",
    "List.Min", "List.Max", "List.Median", "List.Mode",
    "List.Percentile", "List.Product", "List.Zip",
    "List.RemoveNulls", "List.Empty",

    # Tablas basicas
    "Table.Column", "Table.ColumnNames", "Table.ColumnCount",
    "Table.RowCount", "Table.First", "Table.Last", "Table.FirstN",
    "Table.Skip", "Table.SelectRows", "Table.SelectColumns",
    "Table.RemoveColumns", "Table.TransformColumnTypes",
    "Table.TransformColumns", "Table.ReplaceValue",
    "Table.ReplaceColumns", "Table.AddColumn", "Table.RenameColumns",
    "Table.RemoveRows", "Table.Distinct", "Table.Sort",
    "Table.Group", "Table.Combine", "Table.Join",
    "Table.NestedJoin", "Table.ExpandTableColumn",
    "Table.ExpandRecordColumn", "Table.ToList", "Table.FromList",
    "Table.Pivot", "Table.Unpivot",
    "Table.FillDown", "Table.FillUp", "Table.ReplaceRows",
    "Table.Split", "Table.AddKey", "Table.Keys",
    "Table.AddJoinColumn", "Table.AggregateTableColumn",
    "Table.FromColumns", "Table.FromRows", "Table.FromRecords",
    "Table.ToColumns", "Table.ToRows", "Table.ToRecords",
    "Table.Transpose", "Table.ReverseRows", "Table.PrefixColumns",
    "Table.RemoveRowsWithErrors",

    # Fecha/Hora basico
    "Date.Year", "Date.Month", "Date.Day", "Date.DayOfWeek",
    "Date.DayOfYear", "Date.DaysInMonth", "Date.From",
    "Date.FromText", "Date.ToText", "Date.AddDays", "Date.AddMonths",
    "Date.AddQuarters", "Date.AddWeeks", "Date.AddYears",
    "Date.IsInCurrentDay", "Date.IsInCurrentMonth",
    "Date.IsInCurrentYear", "Date.IsInNextDay", "Date.IsInNextMonth",
    "Date.IsInNextYear", "Date.IsInPreviousDay",
    "Date.IsInPreviousMonth", "Date.IsInPreviousYear",
    "Date.IsInYearToDate", "Date.QuarterOfYear", "Date.StartOfDay",
    "Date.StartOfMonth", "Date.StartOfQuarter", "Date.StartOfWeek",
    "Date.StartOfYear", "Date.EndOfDay", "Date.EndOfMonth",
    "Date.EndOfQuarter", "Date.EndOfWeek", "Date.EndOfYear",
    "Time.Hour", "Time.Minute", "Time.Second", "Time.From",
    "Time.FromText", "Time.ToText", "Time.StartOfHour",
    "Time.EndOfHour",
    "DateTime.LocalNow", "DateTime.FixedLocalNow",
    "DateTime.From", "DateTime.FromText", "DateTime.ToText",
    "DateTime.Date", "DateTime.Time", "DateTime.AddZone",
    "Duration.Days", "Duration.Hours", "Duration.Minutes",
    "Duration.Seconds", "Duration.From",

    # Registros basicos
    "Record.Field", "Record.FieldCount", "Record.FieldNames",
    "Record.FieldValues", "Record.HasFields", "Record.SelectFields",
    "Record.RemoveFields", "Record.RenameFields",
    "Record.TransformFields", "Record.AddField", "Record.Combine",
    "Record.FromList", "Record.ToList",
    "Record.FieldOrDefault",

    # Numeros
    "Number.From", "Number.FromText", "Number.ToText",
    "Number.IsEven", "Number.IsNaN", "Number.IsOdd",
    "Number.Abs", "Number.Round", "Number.RoundAwayFromZero",
    "Number.RoundDown", "Number.RoundUp", "Number.Exp",
    "Number.Power", "Number.Sqrt", "Number.Sign",
    "Number.Floor", "Number.Ceiling", "Number.IntegerDivide",
    "Number.Mod", "Number.Combinations", "Number.Permutations",
    "Number.Random", "Number.RandomBetween",
    "Number.BitwiseAnd", "Number.BitwiseNot",
    "Number.BitwiseOr", "Number.BitwiseShift",
    "Number.Factorial", "Number.Ln", "Number.Log", "Number.Log10",

    # Conectores basicos
    "Excel.Workbook", "Csv.Document", "Json.Document",
    "Xml.Document", "Odbc.DataSource", "OleDb.DataSource",
    "Web.Contents", "Binary.Combine",

    # Type system basico
    "Type.Is", "Type.IsOpenRecord", "Type.RecordFields",
    "Type.TableKeys", "Type.TableColumn", "Type.TableColumns",
    "Type.TableSchema", "Type.FunctionParameters",
    "Type.FunctionRequiredParameters",
    "Type.FunctionReturn", "Type.AddTableKey",
    "Type.ReplaceTableKeys", "Type.ForFunction",
    "Type.ForRecord", "Type.ForTable",
}


def get_blocked_list() -> list:
    """Devuelve lista simple de funciones bloqueadas."""
    return list(BLOCKED_FUNCTIONS.keys())


def get_alternative(function_name: str) -> str:
    """Devuelve la alternativa legacy para una funcion bloqueada."""
    func = BLOCKED_FUNCTIONS.get(function_name)
    if not func:
        return f"Funcion '{function_name}' no esta en la lista de bloqueadas"
    return f"Alternativa: {func['alternative']}\nEjemplo: {func['example_legacy']}"


def validate_m_expression(m_code: str) -> dict:
    """
    Valida que una expresion M no use funciones bloqueadas.
    
    Args:
        m_code: Codigo M a validar
    
    Returns:
        dict con {valid: bool, blocked_found: list, safe: bool}
    """
    import re
    blocked_found = []
    
    for func_name in BLOCKED_FUNCTIONS:
        # Buscar el nombre de la funcion en el codigo
        # Pattern: func_name( o func_name (
        pattern = rf'\b{re.escape(func_name)}\s*\('
        if re.search(pattern, m_code):
            blocked_found.append(func_name)
    
    return {
        "valid": len(blocked_found) == 0,
        "blocked_found": blocked_found,
        "total_blocked": len(BLOCKED_FUNCTIONS),
        "safe": len(blocked_found) == 0
    }


def get_system_prompt_extension() -> str:
    """
    Genera el bloque de SYSTEM PROMPT para inyectar en el LLM.
    Esto asegura que el asistente de IA genere codigo M compatible con Excel 2013.
    """
    blocked_list = "\n".join(f"      - {name}: {info['category']} - {info['alternative']}"
                            for name, info in BLOCKED_FUNCTIONS.items())
    
    safe_categories = {
        "Text": ["Text.Start", "Text.End", "Text.Range", "Text.PositionOf",
                 "Text.Length", "Text.Split", "Text.Replace", "Text.Trim"],
        "Table": ["Table.SelectRows", "Table.SelectColumns", "Table.AddColumn",
                  "Table.Group", "Table.TransformColumns", "Table.Combine"],
        "List": ["List.Transform", "List.Select", "List.FirstN", "List.Skip",
                 "List.Count", "List.Sort", "List.Distinct"],
    }
    
    safe_text = ""
    for cat, funcs in safe_categories.items():
        safe_text += f"   {cat}: {', '.join(funcs)}\n"
    
    return f"""
## SYSTEM PROMPT EXTENSION: EXCEL_2013_PQ_COMPATIBILITY

IMPORTANTE: El codigo M que generes debe ser COMPATIBLE con Power Query v2.62.5222.701
(Microsoft Excel 2013 Professional Plus). Sigue estas reglas ESTRICTAMENTE:

### REGLA 1: Funciones PROHIBIDAS
Las siguientes funciones NO existen en Excel 2013. NO las uses bajo ninguna circunstancia:

{blocked_list}

### REGLA 2: Funciones RECOMENDADAS
Usa estas funciones LEGACY que SI estan disponibles:

{safe_text}

### REGLA 3: Verificacion
Antes de generar codigo M, verifica cada funcion contra esta lista.
Si una funcion no esta en la lista de seguras ni en la de bloqueadas,
NO la uses - asume que no esta disponible.

### REGLA 4: Alternativas
Para cada funcion bloqueada, usa su alternativa legacy:
- Text.BeforeDelimiter -> Text.PositionOf + Text.Start
- Text.AfterDelimiter -> Text.PositionOf + Text.Range
- Text.Select -> usar List.Select + Text.Remove con condicion
- Table.Profile -> calculos manuales con List.Min, List.Max, List.Count
- Table.Schema -> Table.TransformColumnTypes manual
- List.Combine -> operador & de concatenacion
- List.Range -> List.FirstN + List.Skip
- List.TransformMany -> List.Transform + List.Accumulate anidado
"""


# =============================================================================
# EJEMPLOS DE CODIGO M COMPATIBLE CON EXCEL 2013
# =============================================================================

EXAMPLES_COMPATIBLES = {
    "split_column_by_delimiter": """
let
    Source = Excel.CurrentWorkbook(){[Name="Tabla1"]}[Content],
    #"Columna Dividida" = Table.TransformColumns(Source, {{"Columna", each 
        let
            parts = Text.Split(_, "-"),
            first = List.First(parts),
            rest = Text.Range(_, Text.Length(first) + 2)
        in
            rest, type text
    }})
in
    #"Columna Dividida"
""",

    "filter_rows_basic": """
let
    Source = Excel.CurrentWorkbook(){[Name="Tabla1"]}[Content],
    #"Filas Filtradas" = Table.SelectRows(Source, each [Columna] <> null and [Columna] > 100)
in
    #"Filas Filtradas"
""",

    "add_calculated_column": """
let
    Source = Excel.CurrentWorkbook(){[Name="Tabla1"]}[Content],
    #"Columna Calculada" = Table.AddColumn(Source, "NuevaCol", each 
        if [Precio] > 1000 then "Caro" else "Barato"
    )
in
    #"Columna Calculada"
""",

    "group_and_aggregate": """
let
    Source = Excel.CurrentWorkbook(){[Name="Tabla1"]}[Content],
    #"Filas Agrupadas" = Table.Group(Source, {"Categoria"}, {{
        "Total", each List.Sum([Ventas]), type number
    }})
in
    #"Filas Agrupadas"
""",

    "merge_queries": """
let
    Source = Excel.CurrentWorkbook(){[Name="Ventas"]}[Content],
    Productos = Excel.CurrentWorkbook(){[Name="Productos"]}[Content],
    #"Combinacion" = Table.NestedJoin(Source, {"ProductoID"}, Productos, {"ID"}, "Producto", JoinKind.LeftOuter),
    #"Expandido" = Table.ExpandTableColumn(#"Combinacion", "Producto", {"Nombre", "Precio"})
in
    #"Expandido"
""",
}


def get_example(name: str) -> str:
    """Devuelve un ejemplo de codigo M compatible."""
    return EXAMPLES_COMPATIBLES.get(name, "Ejemplo no encontrado")
