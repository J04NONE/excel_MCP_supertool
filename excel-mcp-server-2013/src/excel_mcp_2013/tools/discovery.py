"""Tools de autodescubrimiento del entorno (plan maestro Fase 6 / informe 5.2)."""

import logging

logger = logging.getLogger(__name__)

PQ_ADDIN_PROGID = "Microsoft.Data.Mashup.ExcelAddin"
PP_ADDIN_PROGID = "Microsoft.Office.PowerPivot.ExcelAddin"


def register(mcp, session, run):
    @mcp.tool()
    def discover_capabilities() -> dict:
        """Autodescubrir capacidades del host: version de Excel, arquitectura
        (x86/x64), add-ins COM de Power Query y Power Pivot, acceso VBA."""
        return run(_discover_capabilities, session)

    @mcp.tool()
    def validate_environment() -> dict:
        """Validar el entorno contra las restricciones de Excel 2013 y devolver
        advertencias y recomendaciones accionables."""
        return run(_validate_environment, session)


def _addin_state(app, progid: str) -> dict:
    try:
        addin = app.COMAddIns(progid)
        return {"installed": True, "connected": bool(addin.Connect)}
    except Exception:
        return {"installed": False, "connected": False}


def _vba_trust_state(app) -> bool:
    try:
        wb = app.ActiveWorkbook
        if wb is None:
            return None
        _ = wb.VBProject.VBComponents.Count
        return True
    except Exception:
        return False


def _discover_capabilities(session) -> dict:
    app = session.get_application()
    os_string = str(app.OperatingSystem)
    version = str(app.Version)
    return {
        "excel_version": version,
        "excel_version_name": {
            "15.0": "Excel 2013",
            "16.0": "Excel 2016+ / Office 2019/2021/365 (LTSC dev host)",
        }.get(version, "desconocida"),
        "operating_system": os_string,
        "is_64bit": "64-bit" in os_string,
        "power_query_addin": _addin_state(app, PQ_ADDIN_PROGID),
        "power_pivot_addin": _addin_state(app, PP_ADDIN_PROGID),
        "vba_project_access": _vba_trust_state(app),
        "queries_api_available": version >= "16.0",
        "session_pid": session.get_pid(),
    }


def _validate_environment(session) -> dict:
    caps = _discover_capabilities(session)
    warnings, errors, recommendations = [], [], []

    if caps["excel_version"] == "15.0":
        if not caps["power_query_addin"]["connected"]:
            warnings.append("Power Query add-in no esta activo (en 2013 es un complemento COM externo).")
            recommendations.append(
                "Activar: Archivo > Opciones > Complementos > Complementos COM > 'Microsoft Power Query for Excel'. "
                "Descarga: https://www.microsoft.com/en-us/download/details.aspx?id=39379"
            )
        if not caps["power_pivot_addin"]["connected"]:
            warnings.append("Power Pivot add-in no esta activo (solo SKU Professional Plus / Stand-alone / 365 ProPlus).")
            recommendations.append(
                "Activar: Archivo > Opciones > Complementos > Complementos COM > "
                "'Microsoft Office PowerPivot for Excel 2013'."
            )
    elif caps["excel_version"] == "16.0":
        warnings.append(
            "Host de desarrollo es Office 2016+/LTSC: los archivos generados aqui deben validarse "
            "contra las restricciones de Excel 2013 (usa validate_m_code para el M)."
        )

    if not caps["is_64bit"]:
        warnings.append("Excel x86 (32 bits): limite de 4 GB; riesgo de desbordamiento con modelos grandes.")
        recommendations.append("Para volumenes masivos usar Excel x64 o delegar visualizacion a Power BI Desktop.")

    if caps["vba_project_access"] is False:
        warnings.append("Sin acceso al modelo de objetos VBA: list_vba_modules/get_vba_code fallaran (execute_vba_macro SI funciona).")
        recommendations.append(
            "Activar: Centro de confianza > Configuracion de macros > 'Confiar en el acceso al modelo de objetos de proyectos de VBA'."
        )

    return {
        "status": "error" if errors else ("warning" if warnings else "ok"),
        "capabilities": caps,
        "warnings": warnings,
        "errors": errors,
        "recommendations": recommendations,
    }
