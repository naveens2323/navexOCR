import concurrent.futures
import threading
import os
from navexOCR.logger import get_logger

log = get_logger(__name__)

_executor = None
_inflight = 0
_inflight_lock = threading.Lock()


def init_threadpool(max_workers=None):
    global _executor
    if _executor is None:
        if max_workers is None:
            # concurrent.futures.thread._MAX_WORKERS is an internal attribute
            # and may not exist on some Python builds. Use a safe heuristic
            # based on CPU count with reasonable caps.
            cpu_count = os.cpu_count() or 4
            max_workers = min(32, max(2, cpu_count))
        _executor = concurrent.futures.ThreadPoolExecutor(max_workers=max_workers)
        log.info(f"Initialized thread pool with workers={max_workers}")
    return _executor


def submit_task(fn, *args, **kwargs):
    global _executor
    if _executor is None:
        init_threadpool()
    with _inflight_lock:
        global _inflight
        _inflight += 1
    fut = _executor.submit(fn, *args, **kwargs)

    def _done(_):
        global _inflight
        with _inflight_lock:
            if _inflight > 0:
                _inflight -= 1

    fut.add_done_callback(_done)
    return fut


def get_inflight():
    with _inflight_lock:
        return _inflight


def get_pool_size():
    if _executor is None:
        return 0
    # ThreadPoolExecutor doesn't expose worker count; return configured max_workers if available
    try:
        return _executor._max_workers
    except Exception:
        return 0
