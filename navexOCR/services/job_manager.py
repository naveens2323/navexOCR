import io
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from typing import Any, Dict

from navexOCR.config import TEMP_DIR

# Simple in-memory job store. Thread-safe.
_lock = threading.Lock()
_jobs: Dict[str, Dict[str, Any]] = {}

# Executor for background processing
EXECUTOR = ThreadPoolExecutor(max_workers=4)


def create_job(payload: Dict[str, Any]) -> str:
    job_id = str(uuid.uuid4())
    job = {
        "id": job_id,
        "status": "queued",
        "result": None,
        "error": None,
        "payload": payload,
    }

    with _lock:
        _jobs[job_id] = job

    return job_id


def set_job_status(job_id: str, status: str):
    with _lock:
        if job_id in _jobs:
            _jobs[job_id]["status"] = status


def set_job_result(job_id: str, result: Any):
    with _lock:
        if job_id in _jobs:
            _jobs[job_id]["status"] = "completed"
            _jobs[job_id]["result"] = result


def set_job_error(job_id: str, error: str):
    with _lock:
        if job_id in _jobs:
            _jobs[job_id]["status"] = "failed"
            _jobs[job_id]["error"] = error


def get_job(job_id: str) -> Dict[str, Any]:
    with _lock:
        return _jobs.get(job_id)


def submit_job(fn, *args, **kwargs):
    # fn should accept job_id as first arg if it wants to update status
    return EXECUTOR.submit(fn, *args, **kwargs)
