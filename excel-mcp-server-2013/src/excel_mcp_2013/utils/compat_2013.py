"""Funciones de hoja de calculo NO disponibles en Excel 2013.

Datos para check_2013_compatibility: al editar en el host 2021 es facil
introducir funciones que revientan con #NAME? al volver al Excel 2013 del
trabajo (huella tipica: nombres definidos _xlfn.* rotos).
"""

import re

# funcion -> version que la introdujo (todas POSTERIORES a 2013)
MODERN_FUNCTIONS = {
    # Excel 2016
    "TEXTJOIN": "2016",
    "CONCAT": "2016",
    "IFS": "2016",
    "SWITCH": "2016",
    "MAXIFS": "2016",
    "MINIFS": "2016",
    "FORECAST.ETS": "2016",
    "FORECAST.ETS.SEASONALITY": "2016",
    "FORECAST.ETS.CONFINT": "2016",
    "FORECAST.LINEAR": "2016",
    # Excel 2019 / 365 (arrays dinamicos)
    "XLOOKUP": "2021/365",
    "XMATCH": "2021/365",
    "FILTER": "2021/365",
    "SORT": "2021/365",
    "SORTBY": "2021/365",
    "UNIQUE": "2021/365",
    "SEQUENCE": "2021/365",
    "RANDARRAY": "2021/365",
    "LET": "2021/365",
    # 365
    "LAMBDA": "365",
    "MAP": "365",
    "REDUCE": "365",
    "SCAN": "365",
    "BYROW": "365",
    "BYCOL": "365",
    "MAKEARRAY": "365",
    "ISOMITTED": "365",
    "TEXTSPLIT": "365",
    "TEXTBEFORE": "365",
    "TEXTAFTER": "365",
    "VSTACK": "365",
    "HSTACK": "365",
    "TOCOL": "365",
    "TOROW": "365",
    "WRAPROWS": "365",
    "WRAPCOLS": "365",
    "TAKE": "365",
    "DROP": "365",
    "CHOOSEROWS": "365",
    "CHOOSECOLS": "365",
    "EXPAND": "365",
    "GROUPBY": "365",
    "PIVOTBY": "365",
    "IMAGE": "365",
    "STOCKHISTORY": "365",
}

# Referencia de derrame de arrays dinamicos (A1#) — no existe en 2013
SPILL_REF_RE = re.compile(r"\b[A-Z]{1,3}\$?\d+#")

# Patrones precompilados: la funcion seguida INMEDIATAMENTE de "(" — asi
# CONCAT( no matchea CONCATENATE( ni FILTER( matchea FILTERXML( (ambas OK en 2013).
# (?<![\w.]) evita falsos positivos en sufijos como _xlfn.XLOOKUP ya contados aparte.
FUNC_RES = {
    name: re.compile(rf"(?<![\w.]){re.escape(name)}\s*\(")
    for name in MODERN_FUNCTIONS
}
