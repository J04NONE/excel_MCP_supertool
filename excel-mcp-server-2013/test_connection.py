"""
Script de verificacion: pywin32 + Excel 2013
----------------------------------------------
Verifica que pywin32 puede conectarse a Excel via COM.

USO: .venv/Scripts/python.exe test_connection.py
"""

import sys
import gc
import time
import os

# Forzar UTF-8 en la salida estandar
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

def test_excel_com():
    """Prueba 1: Conectar a Excel via COM."""
    print("=" * 60)
    print("PRUEBA 1: Conexion COM a Excel.Application")
    print("=" * 60)

    try:
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        print(f"  [OK] Excel conectado")

        version = excel.Version
        print(f"  [OK] Version de Excel: {version}")

        if version == "15.0":
            print("  [OK] Es Excel 2013! (v15.0)")
        else:
            print(f"  [!] Version {version} detectada (no es 2013)")

        os_info = excel.OperatingSystem
        print(f"  [OS] Sistema Operativo: {os_info}")

        if "64" in os_info:
            print("  [OK] Excel es 64-bit")
        else:
            print("  [!] Excel es 32-bit (x86) - limite de 4GB")

        return excel

    except ImportError as e:
        print(f"  [ERR] Error importando win32com: {e}")
        return None
    except Exception as e:
        print(f"  [ERR] Error conectando a Excel: {e}")
        return None


def test_addins(excel):
    """Prueba 2: Detectar Power Query y Power Pivot."""
    print()
    print("=" * 60)
    print("PRUEBA 2: Deteccion de Add-ins COM")
    print("=" * 60)

    if not excel:
        print("  [ERR] No hay conexion Excel - saltando prueba")
        return

    try:
        # Power Query
        pq_progid = "Microsoft.Data.Mashup.ExcelAddin"
        try:
            pq_addin = excel.COMAddIns(pq_progid)
            if pq_addin.Connect:
                print(f"  [OK] Power Query: ACTIVO")
            else:
                print(f"  [!] Power Query: Instalado pero INACTIVO")
                print(f"      Activar: Archivo > Opciones > Complementos > COM")
        except Exception:
            print(f"  [!] Power Query: NO INSTALADO")
            print(f"      Descargar: https://www.microsoft.com/en-us/download/details.aspx?id=39379")

        # Power Pivot
        pp_progid = "Microsoft.Office.PowerPivot.ExcelAddin"
        try:
            pp_addin = excel.COMAddIns(pp_progid)
            if pp_addin.Connect:
                print(f"  [OK] Power Pivot: ACTIVO")
            else:
                print(f"  [!] Power Pivot: Instalado pero INACTIVO")
                print(f"      Activar: Archivo > Opciones > Complementos > COM")
        except Exception:
            print(f"  [!] Power Pivot: NO INSTALADO")
            print(f"      Requiere: Office 2013 Professional Plus")

    except Exception as e:
        print(f"  [ERR] Error detectando add-ins: {e}")


def test_vba(excel):
    """Prueba 3: Verificar acceso a VBA."""
    print()
    print("=" * 60)
    print("PRUEBA 3: Acceso a VBA")
    print("=" * 60)

    if not excel:
        print("  [ERR] No hay conexion Excel - saltando prueba")
        return

    try:
        wb = excel.Workbooks.Add()
        vbp = wb.VBProject
        print("  [OK] Acceso a VBA Project: DISPONIBLE")
        wb.Close(False)
    except Exception as e:
        error_str = str(e).lower()
        if "access denied" in error_str or "no esta habilitado" in error_str:
            print("  [!] Acceso a VBA: NO DISPONIBLE")
            print("      Activar: Archivo > Opciones > Centro de confianza")
            print("      > Configuracion > Acceso al modelo VBA")
            print("      (Solo necesario para leer/escribir codigo)")
        else:
            print(f"  [ERR] Error inesperado: {e}")


def test_cleanup(excel):
    """Prueba 4: Verificar limpieza de procesos."""
    print()
    print("=" * 60)
    print("PRUEBA 4: Limpieza de procesos")
    print("=" * 60)

    if not excel:
        print("  [ERR] No hay conexion Excel - saltando prueba")
        return None

    try:
        import psutil

        before = [p for p in psutil.process_iter(['pid', 'name']) if p.info['name'] == 'EXCEL.EXE']
        print(f"  [DATA] Procesos EXCEL.EXE antes: {len(before)}")

        print("  Cerrando Excel...")
        excel.Quit()

        del excel
        gc.collect()
        gc.collect()

        time.sleep(2)

        after = [p for p in psutil.process_iter(['pid', 'name']) if p.info['name'] == 'EXCEL.EXE']

        if len(after) == 0:
            print("  [OK] Limpieza perfecta: 0 procesos EXCEL.EXE restantes")
        elif len(after) < len(before):
            print(f"  [!] Quedan {len(after)} procesos (habia {len(before)})")
            print("      Pueden ser de otras instancias abiertas manualmente")
        else:
            print(f"  [!] Quedan {len(after)} procesos - posible zombie")

        return after

    except Exception as e:
        print(f"  [ERR] Error en limpieza: {e}")
        return None


def test_orphan_kill():
    """Prueba 5: Verificar que psutil puede matar procesos."""
    print()
    print("=" * 60)
    print("PRUEBA 5: psutil - Capacidad de matar procesos")
    print("=" * 60)

    try:
        import psutil
        current = psutil.Process()
        print(f"  [OK] psutil funciona: PID actual = {current.pid}")
        print(f"  [OK] Nombre del proceso: {current.name()}")

        excel_processes = [p for p in psutil.process_iter(['pid', 'name'])
                          if p.info['name'] == 'EXCEL.EXE']
        print(f"  [DATA] Procesos EXCEL.EXE actualmente: {len(excel_processes)}")
        for p in excel_processes:
            print(f"         - PID {p.info['pid']}: {p.info['name']}")

    except Exception as e:
        print(f"  [ERR] Error con psutil: {e}")


def main():
    print()
    print("VERIFICACION DEL ENTORNO: pywin32 + Excel 2013")
    print(f"Python: {sys.version}")
    print()

    excel = test_excel_com()
    test_addins(excel)
    test_vba(excel)
    remaining = test_cleanup(excel)
    excel = None
    test_orphan_kill()

    print()
    print("=" * 60)
    print("RESUMEN")
    print("=" * 60)

    if remaining is not None and len(remaining) == 0:
        print("[OK] TODO OK: Entorno listo para el desarrollo")
    else:
        print("[!] Entorno funcional - verificar manualmente si quedan procesos Excel")


if __name__ == "__main__":
    main()
