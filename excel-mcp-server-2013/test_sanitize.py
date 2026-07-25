# -*- coding: utf-8 -*-
"""Regresion del saneo de nombres definidos (utils/workbook_sanitize).

Reproduce en un OOXML sintetico el perfil que colgaba el Open real
(Estimados Sebas.xlsx: vinculos externos + nombres #REF!/[n]) y verifica que
el escaneo estatico lo marca de riesgo y que la copia saneada quita SOLO los
rotos, dejando los legitimos. Puro Python, sin Excel.

Uso:  .venv/Scripts/python.exe test_sanitize.py
"""
import io
import os
import sys
import zipfile

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from excel_mcp_2013.utils.workbook_sanitize import (  # noqa: E402
    make_sanitized_copy,
    scan_definedname_risk,
)

FAILED = []


def check(name, cond, extra=""):
    tag = "[OK]" if cond else "[ERR]"
    print(f"{tag} {name}" + (f" -> {extra}" if extra else ""))
    if not cond:
        FAILED.append(name)


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.'
    'openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    "</Types>"
)

DEFINED_NAMES = (
    "<definedNames>"
    '<definedName name="Area_Buena">Hoja1!$A$1:$B$2</definedName>'      # legit
    '<definedName name="_xlnm.Print_Area">Hoja1!$A$1</definedName>'      # legit
    '<definedName name="Roto1">#REF!</definedName>'                       # roto
    '<definedName name="Roto2">[1]Otro!#REF!</definedName>'               # roto (externo+ref)
    '<definedName name="Externo">[3]Hoja!$A$1</definedName>'             # roto (externo)
    "</definedNames>"
)

WORKBOOK_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<sheets><sheet name="Hoja1" sheetId="1" r:id="rId1"/></sheets>'
    + DEFINED_NAMES +
    "</workbook>"
)


def build_xlsx(path, with_external_links=True):
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("xl/workbook.xml", WORKBOOK_XML)
        if with_external_links:
            z.writestr("xl/externalLinks/externalLink1.xml", "<x/>")


def main():
    tmpdir = os.path.join(os.path.dirname(__file__), "_tmp_sanitize_test")
    os.makedirs(tmpdir, exist_ok=True)
    risky = os.path.join(tmpdir, "risky.xlsx")
    clean = os.path.join(tmpdir, "clean.xlsx")
    build_xlsx(risky, with_external_links=True)
    build_xlsx(clean, with_external_links=False)

    # 1) escaneo del archivo de riesgo
    r = scan_definedname_risk(risky)
    check("scan detecta OOXML", r["ooxml"] is True, str(r["ooxml"]))
    check("scan cuenta vinculos externos", r["external_links"] == 1, str(r["external_links"]))
    check("scan cuenta 5 nombres", r["definednames_total"] == 5, str(r["definednames_total"]))
    check("scan cuenta 3 rotos", r["definednames_broken"] == 3, str(r["definednames_broken"]))
    check("scan marca RISKY", r["risky"] is True, str(r["risky"]))

    # 2) archivo sin vinculos externos -> NO risky (aunque tenga rotos)
    r2 = scan_definedname_risk(clean)
    check("sin vinculos externos NO es risky", r2["risky"] is False, str(r2["risky"]))

    # 3) saneo quita solo los rotos y conserva los legitimos
    tmp, removed = make_sanitized_copy(risky)
    try:
        check("saneo removio 3 rotos", removed == 3, str(removed))
        with zipfile.ZipFile(tmp) as z:
            wbxml = z.read("xl/workbook.xml").decode("utf-8")
            # las partes externas y demas se conservan
            has_links = "xl/externalLinks/externalLink1.xml" in z.namelist()
        check("conserva nombre legitimo", "Area_Buena" in wbxml)
        check("conserva print area", "Print_Area" in wbxml)
        check("elimina Roto1", "Roto1" not in wbxml)
        check("elimina Roto2", "Roto2" not in wbxml)
        check("elimina Externo", '"Externo"' not in wbxml)
        check("no rompe otras partes (externalLinks intacto)", has_links)
        # el resultado sigue siendo un zip valido
        check("copia es zip valido", zipfile.is_zipfile(tmp))
    finally:
        os.remove(tmp)

    # limpieza
    os.remove(risky)
    os.remove(clean)
    os.rmdir(tmpdir)

    print()
    if FAILED:
        print(f"[FALLARON {len(FAILED)}]: {FAILED}")
        sys.exit(1)
    print("[OK] Saneo de nombres: escaneo + copia quirurgica verificados")


if __name__ == "__main__":
    main()
