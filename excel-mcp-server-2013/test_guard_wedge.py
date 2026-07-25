# -*- coding: utf-8 -*-
"""Regresion del seguimiento de tarea en curso del guard (com_guard).

Verifica, SIN Excel, que:
- inflight_info() es None cuando el STA esta ocioso.
- Mientras una tarea corre (aunque el caller ya hizo timeout), inflight_info()
  reporta la tarea, y execute() sigue lanzando TimeoutError como antes.
- Al terminar la tarea, inflight_info() vuelve a None (auto-recuperacion), de
  modo que la senal de bloqueo NO queda pegada tras una operacion lenta legitima.

Uso:  .venv/Scripts/python.exe test_guard_wedge.py
"""
import os
import sys
import threading
import time

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))

from excel_mcp_2013.com_guard import ExcelWriteGuard, STAWedgedError  # noqa: E402

FAILED = []


def check(name, cond, extra=""):
    tag = "[OK]" if cond else "[ERR]"
    print(f"{tag} {name}" + (f" -> {extra}" if extra else ""))
    if not cond:
        FAILED.append(name)


def main():
    guard = ExcelWriteGuard(timeout=120)

    # 1) ocioso -> sin tarea en curso
    check("ocioso: inflight_info None", guard.inflight_info() is None)

    # 2) tarea bloqueante controlada por Event
    release = threading.Event()

    def blocking_task():
        release.wait(timeout=15)
        return "terminado"

    # el caller hace timeout corto, pero la tarea sigue viva en el STA
    t0 = time.time()
    try:
        guard.execute(blocking_task, timeout=1)
        check("execute lanza TimeoutError", False)
    except TimeoutError as e:
        check("execute lanza TimeoutError", "excedio 1s" in str(e), str(e)[:60])
    check("timeout respeta el limite (~1s)", time.time() - t0 < 3, f"{time.time()-t0:.2f}s")

    # mientras sigue bloqueada, inflight_info la reporta
    info = guard.inflight_info()
    check("durante bloqueo: inflight_info no es None", info is not None, str(info))
    if info:
        check("inflight_info trae desc/elapsed/timeout",
              info["desc"] == "blocking_task" and info["elapsed"] > 0 and info["timeout"] == 1,
              str(info))

    # STAWedgedError es un RuntimeError (para que la capa de tools lo propague)
    check("STAWedgedError es RuntimeError", issubclass(STAWedgedError, RuntimeError))

    # 3) liberar la tarea -> el STA se desocupa -> inflight_info vuelve a None
    release.set()
    ok = False
    for _ in range(50):  # hasta ~5s
        if guard.inflight_info() is None:
            ok = True
            break
        time.sleep(0.1)
    check("tras completar: inflight_info vuelve a None (auto-recuperacion)", ok)

    # 4) el STA sigue vivo y procesa una tarea nueva normalmente
    r = guard.execute(lambda: 42, timeout=5)
    check("STA sigue operativo tras la tarea lenta", r == 42, str(r))

    guard.shutdown()

    print()
    if FAILED:
        print(f"[FALLARON {len(FAILED)}]: {FAILED}")
        sys.exit(1)
    print("[OK] Seguimiento de tarea en curso del guard verificado")


if __name__ == "__main__":
    main()
