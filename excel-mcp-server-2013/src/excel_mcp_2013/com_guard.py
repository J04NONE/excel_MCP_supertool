"""ExcelWriteGuard: ejecuta operaciones COM en un thread STA dedicado.

COM en single-threaded apartment exige que todos los accesos al objeto
Excel.Application ocurran en el MISMO thread que lo creo. FastMCP ejecuta
tools sincronos en un threadpool, asi que serializamos todo por aqui.
"""

import logging
import queue
import threading
import time
from typing import Any, Callable, Optional

import pythoncom

logger = logging.getLogger(__name__)


class STAWedgedError(RuntimeError):
    """El hilo STA quedo bloqueado (p.ej. un dialogo modal de Excel) y no puede
    procesar mas tareas. La recuperacion es via close_excel (mata por PID)."""


class ExcelWriteGuard:
    """Execute COM operations on a dedicated STA thread."""

    def __init__(self, timeout: int = 120):
        self._timeout = timeout
        self._task_queue: queue.Queue = queue.Queue()
        # Estado de la tarea EN CURSO en el STA (para detectar bloqueos). Se
        # escribe/lee bajo lock desde dos hilos: el STA y el que llama execute().
        self._state_lock = threading.Lock()
        self._inflight_started: Optional[float] = None
        self._inflight_timeout: Optional[float] = None
        self._inflight_desc: Optional[str] = None
        self._sta_thread = threading.Thread(
            target=self._sta_loop, name="excel-sta", daemon=True
        )
        self._sta_thread.start()
        logger.debug("ExcelWriteGuard STA thread started")

    def _sta_loop(self) -> None:
        """Dedicated STA thread loop."""
        pythoncom.CoInitialize()
        try:
            while True:
                task = self._task_queue.get()
                if task is None:
                    break
                func, args, kwargs, result_q, task_timeout, desc = task
                with self._state_lock:
                    self._inflight_started = time.monotonic()
                    self._inflight_timeout = task_timeout
                    self._inflight_desc = desc
                try:
                    result_q.put(("ok", func(*args, **kwargs)))
                except BaseException as e:  # el caller decide; nunca matar el loop
                    logger.debug("Tarea COM fallo: %r", e)
                    result_q.put(("error", e))
                finally:
                    with self._state_lock:
                        self._inflight_started = None
                        self._inflight_timeout = None
                        self._inflight_desc = None
                    # Soltar referencias antes de bloquear en el proximo get():
                    # si no, el resultado/closure del ultimo call quedan vivos
                    # (memoria duplicada en reposo entre llamadas).
                    del task, func, args, kwargs, result_q
        finally:
            pythoncom.CoUninitialize()
            logger.debug("STA thread finished")

    def inflight_info(self) -> Optional[dict]:
        """Info de la tarea en curso en el STA, o None si esta ocioso.

        Devuelve {'desc', 'elapsed', 'timeout'}. La usa la capa superior para,
        junto con una senal positiva de bloqueo (p.ej. un dialogo modal de
        Excel), decidir un fail-fast en vez de esperar el timeout completo.
        No decide 'wedged' por si sola: una tarea larga legitima tambien
        aparece aqui.
        """
        with self._state_lock:
            started = self._inflight_started
            tmo = self._inflight_timeout
            desc = self._inflight_desc
        if started is None:
            return None
        return {"desc": desc, "elapsed": time.monotonic() - started, "timeout": tmo}

    def execute(self, func: Callable, *args, timeout: float | None = None, **kwargs) -> Any:
        """Execute a function on the STA thread and return its result.

        Cada llamada usa su propia cola de resultado: una cola compartida
        mezclaria resultados entre llamadas concurrentes.

        `timeout` (keyword-only, segundos) sobreescribe el default de la
        instancia para ESTA llamada; nunca se reenvia a `func`. Ojo: el
        tiempo corre desde el encolado, asi que incluye la espera detras
        de una tarea larga previa en el STA.
        """
        if not self._sta_thread.is_alive():
            raise RuntimeError("ExcelWriteGuard ya fue apagado (STA thread muerto)")
        effective = timeout if timeout is not None else self._timeout
        desc = getattr(func, "__name__", None) or "operacion"
        result_q: queue.Queue = queue.Queue(maxsize=1)
        self._task_queue.put((func, args, kwargs, result_q, effective, desc))
        try:
            status, result = result_q.get(timeout=effective)
        except queue.Empty:
            raise TimeoutError(
                f"Operacion COM excedio {effective}s (Excel puede estar bloqueado "
                "por un dialogo). Si queda colgado, close_excel reinicia la sesion."
            )
        if status == "error":
            raise result
        return result

    def shutdown(self) -> None:
        """Shutdown the STA thread."""
        if not self._sta_thread.is_alive():
            return
        self._task_queue.put(None)
        self._sta_thread.join(timeout=5)
        if self._sta_thread.is_alive():
            logger.warning("STA thread no termino en 5s (daemon: morira con el proceso)")
