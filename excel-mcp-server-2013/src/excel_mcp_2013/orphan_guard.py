"""OrphanProcessGuard: estrategia de 3 niveles para no dejar EXCEL.EXE zombie.

Nivel 1: app.Quit() cooperativo.
Nivel 2: doble gc.collect() para soltar referencias COM colgantes.
Nivel 3: kill del proceso por PID si sobrevivio.
"""

import gc
import logging
import time

import psutil

logger = logging.getLogger(__name__)


class OrphanProcessGuard:
    """Three-level strategy to kill orphan Excel processes."""

    @staticmethod
    def safe_shutdown(app) -> bool:
        """Level 1: cooperative Quit()."""
        if app:
            try:
                app.Quit()
                logger.info("Level 1 (Quit) OK")
                return True
            except Exception as e:
                logger.warning("Level 1 (Quit) failed: %s", e)
        return False

    @staticmethod
    def force_gc() -> None:
        """Level 2: release dangling COM references."""
        gc.collect()
        time.sleep(0.5)
        gc.collect()
        logger.debug("Level 2 (GC) done")

    @staticmethod
    def kill_process(pid: int) -> bool:
        """Level 3: kill by PID if still alive. True si lo mato."""
        try:
            proc = psutil.Process(pid)
            if proc.is_running() and proc.name() == "EXCEL.EXE":
                logger.warning("Level 3: killing Excel PID %s", pid)
                proc.kill()
                proc.wait(timeout=3)
                return True
        except psutil.NoSuchProcess:
            logger.debug("Level 3: PID %s ya no existe", pid)
        except (psutil.AccessDenied, psutil.TimeoutExpired) as e:
            logger.warning("Level 3 (Kill) failed: %s", e)
        return False

    @classmethod
    def full_cleanup(cls, app, pid: int) -> dict:
        """Run all three levels. Returns what each level did."""
        result = {"quit": False, "gc": False, "kill": False}
        result["quit"] = cls.safe_shutdown(app)
        cls.force_gc()
        result["gc"] = True
        if pid:
            result["kill"] = cls.kill_process(pid)
        logger.info("Full cleanup: %s", result)
        return result
