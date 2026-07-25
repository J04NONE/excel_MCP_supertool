# -*- coding: utf-8 -*-
"""Saneo de nombres definidos rotos para evitar el cuelgue al abrir.

Problema: un libro OOXML (.xlsx/.xlsm) con referencias externas y muchos
nombres definidos ROTOS (valor `#REF!` o referencia a un origen `[n]`
inaccesible) dispara al abrir el dialogo modal 'Conflicto de nombres'
(clase Office `bosa_sdm_XL9`), UNO por cada nombre en conflicto. Ese dialogo
NO lo suprime `Application.DisplayAlerts=False`, y como el MCP usa un unico
hilo STA, el `Workbooks.Open` bloqueado tumba TODA la sesion (solo
`close_excel` recupera).

Solucion: antes de abrir, si el archivo tiene ese perfil de riesgo, se crea
una COPIA temporal a la que se le quitan del `xl/workbook.xml` unicamente los
nombres definidos rotos (los legitimos —areas de impresion, _FilterDatabase—
se conservan). El original NUNCA se modifica. Verificado: el mismo archivo
que colgaba >120s abre en <1s tras el saneo.

Solo aplica a OOXML empaquetado (.xlsx/.xlsm). El .xlsb guarda el libro en un
stream binario (`xl/workbook.bin`) y no es saneable por esta via; para esos se
devuelve risky=False (no se pudo, no cuelga por este motivo en la practica).
"""
from __future__ import annotations

import logging
import os
import re
import tempfile
import time
import zipfile

logger = logging.getLogger(__name__)

# Un nombre definido es "roto" si su valor es #REF! o referencia un origen
# externo por indice [n] (un libro que aqui no esta disponible).
_BROKEN_RE = re.compile(r"#REF!|\[\d+\]")
_DEFINEDNAME_RE = re.compile(r"<definedName\b[^>]*>.*?</definedName>", re.S)
_BLOCK_RE = re.compile(r"<definedNames>.*?</definedNames>", re.S)
_TAG_RE = re.compile(r"<[^>]+>")

# Umbral: por debajo de esto no vale la pena (y evita tocar libros normales).
_MIN_BROKEN = 1


def scan_definedname_risk(path: str) -> dict:
    """Escaneo ESTATICO (sin COM/Excel) del riesgo de cuelgue por nombres.

    Devuelve dict con: ooxml, external_links, definednames_total,
    definednames_broken, risky.
    """
    info = {
        "ooxml": False,
        "external_links": 0,
        "definednames_total": 0,
        "definednames_broken": 0,
        "risky": False,
    }
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if "xl/workbook.xml" not in names:
                return info  # .xlsb u otro: no saneable por XML
            info["ooxml"] = True
            info["external_links"] = sum(
                1 for n in names
                if re.match(r"xl/externalLinks/externalLink\d+\.xml$", n)
            )
            wbxml = z.read("xl/workbook.xml").decode("utf-8", "replace")
    except (zipfile.BadZipFile, OSError) as e:
        logger.debug("scan_definedname_risk: no es ZIP OOXML (%s): %s", path, e)
        return info

    block = _BLOCK_RE.search(wbxml)
    if block:
        dn = _DEFINEDNAME_RE.findall(block.group(0))
        info["definednames_total"] = len(dn)
        info["definednames_broken"] = sum(
            1 for d in dn if _BROKEN_RE.search(_TAG_RE.sub("", d))
        )
    # El cuelgue necesita el import de nombres desde los vinculos externos:
    # riesgo = OOXML + >=1 vinculo externo + >=1 nombre roto.
    info["risky"] = (
        info["ooxml"]
        and info["external_links"] > 0
        and info["definednames_broken"] >= _MIN_BROKEN
    )
    return info


def make_sanitized_copy(path: str) -> tuple[str, int]:
    """Crea una copia temporal sin los nombres definidos ROTOS.

    Devuelve (ruta_temporal, nombres_removidos). El original no se toca.
    Lanza excepcion si el saneo falla (el caller decide; nunca colgar en silencio).
    """
    with zipfile.ZipFile(path) as zin:
        names = zin.namelist()
        payload = {n: zin.read(n) for n in names}

    wbxml = payload["xl/workbook.xml"].decode("utf-8", "replace")
    removed = 0
    block = _BLOCK_RE.search(wbxml)
    if block:
        kept = [
            d for d in _DEFINEDNAME_RE.findall(block.group(0))
            if not _BROKEN_RE.search(_TAG_RE.sub("", d))
        ]
        removed = len(_DEFINEDNAME_RE.findall(block.group(0))) - len(kept)
        new_block = "<definedNames>" + "".join(kept) + "</definedNames>" if kept else ""
        wbxml = wbxml.replace(block.group(0), new_block)
        payload["xl/workbook.xml"] = wbxml.encode("utf-8")

    ext = os.path.splitext(path)[1] or ".xlsx"
    fd, tmp = tempfile.mkstemp(prefix="mcp_san_", suffix=ext)
    os.close(fd)
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, payload[n])
    logger.info("Copia saneada creada (%d nombres rotos removidos): %s", removed, tmp)
    return tmp, removed
