"""Shared helper: run Office COM calls on ONE dedicated thread (COM objects can't cross threads)."""

import json
import queue
import threading
from concurrent.futures import Future

MAX_RESULT_CHARS = 12000


class ComWorker:
    def __init__(self):
        self._q: queue.Queue = queue.Queue()
        threading.Thread(target=self._loop, daemon=True).start()

    def _loop(self):
        try:
            import pythoncom

            pythoncom.CoInitialize()
        except ImportError:
            pass
        while True:
            fn, fut = self._q.get()
            try:
                fut.set_result(fn())
            except Exception as e:  # noqa: BLE001 - everything is reported back to the caller
                fut.set_exception(e)

    def run(self, fn, timeout=180):
        fut: Future = Future()
        self._q.put((fn, fut))
        return fut.result(timeout=timeout)


def dispatch(worker: ComWorker, tools: dict, name: str, args: dict) -> str:
    """Run tools[name](**args) on the worker thread and return a JSON string (errors included)."""
    fn = tools.get(name)
    if fn is None:
        return json.dumps({"error": f"Unknown tool '{name}'. Available: {list(tools)}"})
    try:
        result = worker.run(lambda: fn(**args))
    except TypeError as e:
        return json.dumps({"error": f"Bad arguments for {name}: {e}"})
    except Exception as e:  # noqa: BLE001
        return json.dumps({"error": f"{type(e).__name__}: {e}"})
    return json.dumps(result, default=str, ensure_ascii=False)[:MAX_RESULT_CHARS]