import os
import time
import threading
import shutil
import queue
from datetime import datetime, timedelta
from navexOCR.config import TEMP_DIR
from navexOCR.logger import get_logger

log = get_logger(__name__)


def cleanup_folder(folder):
    """Enqueue a folder for background cleanup.

    This returns quickly; actual deletion is performed by the cleanup worker
    using robust deletion logic (_robust_rmtree).
    """
    try:
        if os.path.exists(folder):
            _cleanup_queue.put(folder)
            log.info(f"Enqueued folder for cleanup: {folder}")
    except Exception:
        log.exception(f"Failed to enqueue folder for cleanup: {folder}")


def _remove_old_dirs(ttl_minutes=10):
    now = datetime.utcnow()
    ttl = timedelta(minutes=ttl_minutes)
    removed = []
    try:
        for name in os.listdir(TEMP_DIR):
            path = os.path.join(TEMP_DIR, name)
            # include tombstoned dirs (ending with .delete_me) and skip loose marker files
            if not (os.path.isdir(path) or name.endswith('.delete_me')):
                # if it's a tombstone marker file created when rename failed, try removing the base dir
                if name.endswith('.delete_marker'):
                    # marker files point to a dir that couldn't be renamed; attempt removal of the target
                    try:
                        with open(path, 'r', encoding='utf-8') as fh:
                            target = fh.read().strip()
                        if target and os.path.exists(target):
                            try:
                                _robust_rmtree(target)
                            except Exception:
                                log.exception(f"Failed to remove target from marker: {target}")
                        # remove the marker file itself
                        try:
                            os.remove(path)
                        except Exception:
                            log.exception(f"Failed to remove marker file: {path}")
                    except Exception:
                        log.exception(f"Failed to read/delete marker file: {path}")
                continue
            try:
                # For tombstoned dirs (ending with .delete_me) use the folder mtime if present
                check_path = path
                if name.endswith('.delete_me') and not os.path.exists(path):
                    # path may not exist if rename previously failed; skip
                    continue
                mtime = datetime.utcfromtimestamp(os.path.getmtime(check_path))
                if now - mtime > ttl:
                    try:
                        _robust_rmtree(check_path)
                        removed.append(check_path)
                    except Exception:
                        log.exception(f"Failed to remove temp dir: {check_path}")
            except Exception:
                log.exception(f"Failed to stat temp dir: {path}")
    except Exception:
        log.exception("Failed to sweep temp dir")
    return removed


def _robust_rmtree(path, max_attempts=3):
    """Attempt to remove a directory tree with retries.

    On Windows PermissionError often means files are in use or read-only.
    This helper will attempt to clear read-only flags, retry a few times,
    and if still failing, rename the folder to a tombstone name so the
    sweeper can try again later without blocking new jobs from creating
    temp dirs with the same name.
    """
    for attempt in range(1, max_attempts + 1):
        try:
            shutil.rmtree(path)
            log.info(f"Removed folder: {path} (attempt {attempt})")
            return
        except PermissionError as pe:
            log.warning(f"PermissionError removing {path} (attempt {attempt}): {pe}")
            # Try to clear read-only flags on files
            try:
                for root, dirs, files in os.walk(path):
                    for fname in files:
                        fp = os.path.join(root, fname)
                        try:
                            os.chmod(fp, 0o666)
                        except Exception:
                            pass
                    for dname in dirs:
                        dp = os.path.join(root, dname)
                        try:
                            os.chmod(dp, 0o777)
                        except Exception:
                            pass
            except Exception:
                pass

            time.sleep(0.5 * attempt)
            continue
        except FileNotFoundError:
            # Already removed concurrently
            log.info(f"Folder already removed: {path}")
            return
        except Exception as e:
            log.exception(f"Error removing folder {path}: {e}")
            raise

    # If still exists, attempt to rename to tombstone so next sweeper pass can try again
    try:
        tomb = f"{path}.delete_me"
        try:
            os.rename(path, tomb)
            log.warning(f"Renamed stuck temp dir to tombstone: {tomb}")
        except PermissionError:
            # If rename is denied (e.g., file handles still open), create a lightweight marker
            try:
                parent = os.path.dirname(path)
                marker = os.path.join(parent, f"{os.path.basename(path)}.delete_marker")
                with open(marker, 'w', encoding='utf-8') as fh:
                    fh.write(path)
                log.warning(f"Could not rename stuck temp dir; created marker: {marker}")
            except Exception:
                log.exception(f"Failed to create tombstone marker for stuck temp dir: {path}")
        except Exception:
            log.exception(f"Failed to rename stuck temp dir: {path}")
    except Exception:
        log.exception(f"Final attempt failed to handle stuck temp dir: {path}")


def start_sweeper(interval_seconds=60, ttl_minutes=10):
    """Start a background thread that removes old temp dirs periodically."""

    def _run():
        log.info(f"Temp sweeper started: scanning {TEMP_DIR} every {interval_seconds}s, TTL={ttl_minutes}m")
        while True:
            try:
                removed = _remove_old_dirs(ttl_minutes=ttl_minutes)
                if removed:
                    for p in removed:
                        log.info(f"Removed old temp dir: {p}")
            except Exception:
                log.exception("Sweeper error")
            time.sleep(interval_seconds)

    t = threading.Thread(target=_run, daemon=True)
    t.start()
    return t


# Background cleanup queue and worker
_cleanup_queue = queue.Queue()
_cleanup_worker_thread = None


def _cleanup_worker():
    log.info("Cleanup worker started")
    while True:
        try:
            path = _cleanup_queue.get()
            if path is None:
                break
            try:
                _robust_rmtree(path)
            except Exception:
                log.exception(f"Cleanup worker failed to remove: {path}")
            finally:
                _cleanup_queue.task_done()
        except Exception:
            log.exception("Cleanup worker encountered an error")
            time.sleep(1)


def start_cleanup_worker():
    """Start the background thread that processes the cleanup queue."""
    global _cleanup_worker_thread
    if _cleanup_worker_thread and _cleanup_worker_thread.is_alive():
        return _cleanup_worker_thread
    _cleanup_worker_thread = threading.Thread(target=_cleanup_worker, daemon=True)
    _cleanup_worker_thread.start()
    return _cleanup_worker_thread