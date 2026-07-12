"""Prueba de hardening Fase 7: timeout por operacion, recovery tras crash,
errores COM traducidos.

Uso:  .venv/Scripts/python.exe test_hardening.py
Requiere Excel instalado. Mata instancias PROPIAS (creadas por este script)
a proposito para probar recovery; nunca toca un Excel del usuario.
"""

import sys
import threading
import time

sys.stdout.reconfigure(encoding="utf-8")

import psutil
import pywintypes

from src.excel_mcp_2013.com_guard import ExcelWriteGuard
from src.excel_mcp_2013.session import SessionManager
from src.excel_mcp_2013.utils.com_errors import (
    ExcelCrashedError,
    is_dead_server,
    translate_com_error,
)

FAILED = []


def check(name: str, cond: bool, extra: str = "") -> None:
    tag = "[OK]" if cond else "[ERR]"
    print(f"{tag} {name}" + (f" -> {extra}" if extra else ""))
    if not cond:
        FAILED.append(name)


def snapshot_excel_pids() -> set:
    return {
        p.pid
        for p in psutil.process_iter(["name"])
        if p.info["name"] == "EXCEL.EXE"
    }


def make_runner(session, guard):
    """Copia local del patron run_with_excel del server (sin FastMCP)."""

    def run(func, *args, timeout=None, **kwargs):
        def runner():
            if session.get_application() is None:
                session.start()
            elif not session.is_alive():
                session.reset_after_crash()
                session.start()
            try:
                return func(*args, **kwargs)
            except pywintypes.com_error as e:
                if is_dead_server(e):
                    session.reset_after_crash()
                    raise ExcelCrashedError(
                        "Excel murio durante la operacion."
                    ) from e
                raise

        try:
            return guard.execute(runner, timeout=timeout)
        except pywintypes.com_error as e:
            raise RuntimeError(translate_com_error(e)) from e

    return run


def main() -> int:
    pids_before = snapshot_excel_pids()
    session = SessionManager()
    guard = ExcelWriteGuard()
    run = make_runner(session, guard)

    print("=== 1. Timeout por operacion ===")
    try:
        run(time.sleep, 4, timeout=2)
        check("timeout corto lanza TimeoutError", False)
    except TimeoutError as e:
        check("timeout corto lanza TimeoutError", "excedio 2s" in str(e), str(e)[:80])
    r = run(lambda: "listo", timeout=30)
    check("timeout holgado OK", r == "listo")

    print("=== 2. Recovery pre-flight (Excel muere ENTRE llamadas) ===")
    v = run(lambda: session.get_application().Version)
    pid1 = session.get_pid()
    check("sesion arranca", v is not None and pid1 is not None, f"PID {pid1} v{v}")
    psutil.Process(pid1).kill()
    time.sleep(1.5)
    check("proceso muerto", not psutil.pid_exists(pid1) or not session.is_alive())
    v2 = run(lambda: session.get_application().Version)
    pid2 = session.get_pid()
    check(
        "siguiente llamada reinicia sola",
        v2 is not None and pid2 is not None and pid2 != pid1,
        f"PID nuevo {pid2}",
    )

    print("=== 3. Recovery mid-flight (Excel muere DURANTE la operacion) ===")
    pid3 = session.get_pid()

    def slow_com_op():
        app = session.get_application()
        time.sleep(2.5)  # el killer dispara a mitad
        return app.Workbooks.Count

    killer = threading.Timer(1.0, lambda: psutil.Process(pid3).kill())
    killer.start()
    try:
        slow = run(slow_com_op, timeout=30)
        check("mid-flight lanza ExcelCrashedError", False, f"devolvio {slow}")
    except ExcelCrashedError as e:
        check("mid-flight lanza ExcelCrashedError", True, str(e)[:60])
    except Exception as e:  # noqa: BLE001
        check("mid-flight lanza ExcelCrashedError", False, f"{type(e).__name__}: {e}")
    finally:
        killer.join()
    v3 = run(lambda: session.get_application().Version)
    check("llamada post-crash arranca limpia", v3 is not None, f"PID {session.get_pid()}")

    print("=== 4. Error COM traducido (0x800A03EC) ===")
    def bad_sheet():
        return session.get_application().ActiveWorkbook.Worksheets("NO_EXISTE").Name

    run(lambda: session.get_application().Workbooks.Add())
    try:
        run(bad_sheet)
        check("hoja inexistente lanza error", False)
    except RuntimeError as e:
        msg = str(e)
        legible = (
            "rechazo la operacion" in msg
            or "Elemento no encontrado" in msg
            or "Error COM" in msg
        )
        check("error traducido legible (no com_error crudo)", legible, msg[:110])
    except pywintypes.com_error as e:
        check("error traducido legible (no com_error crudo)", False, f"crudo: {e}")

    print("=== 5. Limpieza ===")
    guard.execute(session.close)
    guard.shutdown()
    time.sleep(1)
    leftover = snapshot_excel_pids() - pids_before
    check("0 procesos EXCEL.EXE huerfanos", not leftover, str(leftover or "ninguno"))

    print()
    if FAILED:
        print(f"[ERR] {len(FAILED)} fallos: {FAILED}")
        return 1
    print("[OK] Hardening completo: timeout + recovery + traduccion de errores")
    return 0


if __name__ == "__main__":
    sys.exit(main())
