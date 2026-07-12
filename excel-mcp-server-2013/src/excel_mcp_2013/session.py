"""SessionManager: ciclo de vida de Excel.Application via COM.

Todas las llamadas a metodos de esta clase que tocan COM deben ejecutarse
en el thread STA dedicado (ver com_guard.ExcelWriteGuard) — COM en modo
apartment tiene afinidad de thread.
"""

import gc
import logging
import os
import time
from typing import Optional

import psutil
import win32com.client

logger = logging.getLogger(__name__)

# Constantes XlCalculation (Excel 2013+)
XL_CALCULATION_MANUAL = -4135


class SessionManager:
    """Manage Excel Application lifecycle via COM."""

    def __init__(self, visible: bool = False):
        self._app: Optional[win32com.client.CDispatch] = None
        self._excel_pid: Optional[int] = None
        self._excel_create_time: Optional[float] = None
        self._visible = visible
        self._workbooks = {}  # path -> workbook reference

    def start(self) -> None:
        """Initialize a NEW Excel application instance.

        Usa DispatchEx (no Dispatch) para forzar una instancia nueva:
        Dispatch puede adjuntarse a un Excel que el usuario ya tenga abierto,
        y en close() lo cerrariamos/matariamos.
        """
        if self._app is not None:
            logger.debug("Excel ya inicializado (PID %s), start() ignorado", self._excel_pid)
            return
        self._app = win32com.client.DispatchEx("Excel.Application")
        # Capturar el PID ANTES de configurar: si la configuracion falla,
        # necesitamos el PID para no dejar el proceso huerfano.
        self._capture_process_id()
        # create_time distingue NUESTRO proceso de un PID reciclado por Windows
        # (is_alive lo compara antes de dar el proceso por vivo).
        try:
            if self._excel_pid:
                self._excel_create_time = psutil.Process(self._excel_pid).create_time()
        except Exception as e:
            self._excel_create_time = None
            logger.warning("No se pudo capturar create_time del PID %s: %s", self._excel_pid, e)
        try:
            self._app.Visible = self._visible
            self._app.DisplayAlerts = False
            self._app.ScreenUpdating = False
            self._app.EnableEvents = False
            self._app.UserControl = False
            # NOTA: Application.Calculation NO se puede asignar sin un workbook
            # abierto (com_error -2146827284); se aplica en open_workbook().
        except Exception:
            logger.exception("Fallo configurando Excel.Application; limpiando proceso")
            self.close()
            raise
        logger.info("Excel initialized. PID: %s, Version: %s", self._excel_pid, self._app.Version)

    def _apply_manual_calculation(self) -> None:
        """Poner calculo manual. Solo es asignable con >=1 workbook abierto."""
        try:
            self._app.Calculation = XL_CALCULATION_MANUAL
            logger.debug("Calculation = xlCalculationManual aplicado")
        except Exception as e:
            logger.warning("No se pudo poner Calculation manual: %s", e)

    def _capture_process_id(self) -> None:
        """Find the Excel process ID of OUR instance.

        Metodo primario: Hwnd de la ventana de la instancia -> PID exacto.
        Fallback: el EXCEL.EXE mas reciente del usuario actual.
        """
        try:
            import win32process

            hwnd = self._app.Hwnd
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid:
                self._excel_pid = pid
                return
        except Exception as e:
            logger.warning("PID via Hwnd fallo (%s); usando fallback psutil", e)

        try:
            current_user = psutil.Process(os.getpid()).username()
            excel_procs = [
                p
                for p in psutil.process_iter(["pid", "name", "username", "create_time"])
                if p.info["name"] == "EXCEL.EXE" and p.info["username"] == current_user
            ]
            if excel_procs:
                # El mas reciente es (casi seguro) el que acabamos de crear
                newest = max(excel_procs, key=lambda p: p.info["create_time"])
                self._excel_pid = newest.info["pid"]
        except Exception:
            logger.exception("No se pudo capturar el PID de Excel")

    def open_workbook(
        self,
        path: str,
        read_only: bool = False,
        password: Optional[str] = None,
        enable_macros: bool = False,
    ):
        """Open a workbook and track it.

        enable_macros=False (default) abre con msoAutomationSecurityForceDisable:
        las macros del libro NO se cargan (inspeccion segura de .xlsm ajenos).
        Para ejecutar macros hay que reabrir con enable_macros=True.
        """
        if not self._app:
            self.start()
        path = os.path.abspath(path)
        if not os.path.exists(path):
            raise FileNotFoundError(f"Workbook no encontrado: {path}")
        # msoAutomationSecurityLow=1 (macros activas), ForceDisable=3
        automation_security = 1 if enable_macros else 3
        try:
            prev_security = self._app.AutomationSecurity
            self._app.AutomationSecurity = automation_security
        except Exception as e:
            prev_security = None
            logger.warning("No se pudo fijar AutomationSecurity: %s", e)
        try:
            kwargs = {"ReadOnly": read_only}
            if password is not None:
                kwargs["Password"] = password
            wb = self._app.Workbooks.Open(path, **kwargs)
        except Exception:
            logger.exception("Fallo abriendo workbook: %s", path)
            raise
        finally:
            if prev_security is not None:
                try:
                    self._app.AutomationSecurity = prev_security
                except Exception:
                    pass
        self._workbooks[path] = wb
        self._apply_manual_calculation()
        logger.info("Workbook abierto: %s (read_only=%s)", path, read_only)
        return wb

    def close(self) -> None:
        """Close Excel and cleanup COM."""
        if self._app is None:
            self._kill_if_alive()
            return
        try:
            for path, wb in list(self._workbooks.items()):
                try:
                    wb.Close(SaveChanges=False)
                    logger.debug("Workbook cerrado: %s", path)
                except Exception as e:
                    logger.warning("No se pudo cerrar workbook %s: %s", path, e)
        finally:
            self._workbooks.clear()

        try:
            self._app.Quit()
            logger.info("Excel Quit() OK")
        except Exception as e:
            logger.warning("Quit failed: %s", e)
        finally:
            self._app = None
            self._force_gc()
            self._kill_if_alive()

    def _force_gc(self) -> None:
        """Force garbage collection so COM references are released."""
        gc.collect()
        gc.collect()

    def _kill_if_alive(self) -> None:
        """Kill Excel if it survived Quit()."""
        if not self._excel_pid:
            return
        try:
            time.sleep(1)
            proc = psutil.Process(self._excel_pid)
            if proc.is_running() and proc.name() == "EXCEL.EXE":
                logger.warning("Killing orphan Excel PID %s", self._excel_pid)
                proc.kill()
                proc.wait(timeout=3)
        except psutil.NoSuchProcess:
            pass  # ya murio, perfecto
        except (psutil.AccessDenied, psutil.TimeoutExpired) as e:
            logger.error("No se pudo matar Excel PID %s: %s", self._excel_pid, e)
        finally:
            self._excel_pid = None

    def is_alive(self) -> bool:
        """Health-check barato SIN COM: ¿el proceso de NUESTRA instancia vive?

        Solo psutil (no puede colgarse ni dar falso positivo con Excel ocupado,
        como haria un probe COM tipo app.Version). Si no hay sesion iniciada
        devuelve False; si no tenemos PID (captura fallo) asumimos vivo para
        no reiniciar por falta de evidencia.
        """
        if self._app is None:
            return False
        if self._excel_pid is None:
            return True  # sin PID no podemos verificar; no matar por la duda
        try:
            proc = psutil.Process(self._excel_pid)
            if not (proc.is_running() and proc.name() == "EXCEL.EXE"):
                return False
            if self._excel_create_time is not None:
                # PID reciclado por Windows -> no es nuestro proceso
                return abs(proc.create_time() - self._excel_create_time) < 1.0
            return True
        except psutil.NoSuchProcess:
            return False
        except Exception as e:
            logger.warning("is_alive() no pudo verificar PID %s: %s", self._excel_pid, e)
            return True  # en la duda, no disparar recovery

    def reset_after_crash(self) -> None:
        """Limpieza de estado tras muerte de EXCEL.EXE.

        NO llama metodos COM (el servidor ya no existe; Quit() colgaria o
        fallaria): suelta las referencias colgantes, limpia _workbooks, mata
        el proceso si quedo zombie y deja la sesion lista para un start()
        lazy. DEBE ejecutarse en el thread STA (afinidad de las referencias).
        """
        logger.warning("reset_after_crash: limpiando sesion (PID %s)", self._excel_pid)
        self._workbooks.clear()
        self._app = None
        self._excel_create_time = None
        self._force_gc()
        self._kill_if_alive()  # resetea _excel_pid; tolera NoSuchProcess

    def get_application(self):
        return self._app

    def get_pid(self) -> Optional[int]:
        return self._excel_pid
