import os
import time
import fitz
import traceback
from navexOCR.logger import get_logger

log = get_logger(__name__)
import win32com.client
import pythoncom
import pywintypes
import win32process
import threading
import asyncio
from concurrent.futures import ThreadPoolExecutor
import subprocess
import shlex
import sys
from pathlib import Path

# Word COM is not thread-safe across multiple threads. Use a global lock
# to serialize Word automation calls so the RPC server doesn't get into
# an inconsistent state when multiple threads dispatch Word concurrently.
_word_lock = threading.Lock()

# Dedicated single-thread executor for Word COM operations. Running all
# Word automation on a single persistent thread reduces COM RPC/connection
# issues caused by multiple threads creating/destroying Word instances.
_word_executor = ThreadPoolExecutor(max_workers=1)


def pdf_to_word(pdf_file, output_docx, temp_dir):

    # Create an isolated working subdirectory for this in-process conversion.
    # This prevents collisions with process-worker subdirs and locked files.
    inproc_dir = os.path.join(temp_dir, f"inproc_{os.getpid()}_{int(time.time()*1000)}")
    os.makedirs(inproc_dir, exist_ok=True)

    pdf_doc = fitz.open(pdf_file)

    total_pages = pdf_doc.page_count

    page_pdfs = []

    for i in range(total_pages):

        page_pdf_path = os.path.join(
            inproc_dir,
            f"page_{i+1:04d}.pdf"
        )

        single = fitz.open()

        single.insert_pdf(
            pdf_doc,
            from_page=i,
            to_page=i
        )

        single.save(page_pdf_path)

        single.close()

        page_pdfs.append(page_pdf_path)

    pdf_doc.close()

    page_docx_files = []

    # Serialize access to Word automation
    with _word_lock:
        # Initialize COM on this thread and create Word application
        pythoncom.CoInitialize()
        word = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            # capture Word process info when available to aid debugging on other machines
            try:
                word_pid = None
                # Prefer DispatchEx to create a private Word instance; capture pid from window handle if present
                hwnd = getattr(word, 'Hwnd', None)
                if hwnd:
                    try:
                        _tid, _pid = win32process.GetWindowThreadProcessId(hwnd)
                        word_pid = _pid
                    except Exception:
                        word_pid = hwnd
                log.info(f"Launched Word COM object; hwnd/pid={word_pid}")
            except Exception:
                log.debug("Could not obtain Word window handle/pid")

            word.Visible = False
            word.DisplayAlerts = 0

            # helper to restart Word if COM connection is lost
            def _restart_word():
                nonlocal word
                try:
                    if word:
                        try:
                            word.Quit()
                        except Exception:
                            pass
                except Exception:
                    pass
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
                time.sleep(1)
                pythoncom.CoInitialize()
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                word.DisplayAlerts = 0

            CM_TO_POINTS = 28.35

            MARGIN = 0.5 * CM_TO_POINTS

            for idx, page_pdf in enumerate(page_pdfs):

                page_num = idx + 1

                # store per-page docx files inside the inproc working dir to avoid collisions
                out_docx = os.path.join(
                    inproc_dir,
                    f"page_{page_num:04d}.docx"
                )


                doc = None

                try:

                    # Documents.Open can fail with COM 'application is busy' errors
                    # Retry a few times with backoff to let Word become responsive.
                    # Additionally handle -2147418111 (Call was rejected by callee) by
                    # pumping messages briefly and retrying.
                    doc = None
                    max_attempts = 8
                    for attempt in range(max_attempts):
                        try:
                            doc = word.Documents.Open(page_pdf)
                            break
                        except pywintypes.com_error as e:
                            code = e.args[0] if e.args else None
                            # RPC_E_SERVERCALL_RETRYLATER: application busy
                            if code == -2147417846 and attempt < (max_attempts - 1):
                                time.sleep(0.5 + attempt * 0.5)
                                continue
                            # Call was rejected by callee: pump messages and retry
                            if code == -2147418111 and attempt < (max_attempts - 1):
                                log.warning(f"Word call rejected (code={code}); pumping messages and retrying (attempt {attempt+1})")
                                # Allow COM to process windows messages on this thread
                                for _ in range(10):
                                    pythoncom.PumpWaitingMessages()
                                    time.sleep(0.05)
                                time.sleep(0.2 + attempt * 0.2)
                                continue
                            # Other disconnect/RPC errors: try restarting Word and retry
                            if code in (-2147220995, -2147023174, -2147417836) and attempt < (max_attempts - 1):
                                try:
                                    log.warning(f"Word COM error {code}; restarting Word (attempt {attempt+1})")
                                    _restart_word()
                                except Exception:
                                    log.exception("Failed to restart Word during per-page open")
                                time.sleep(0.5 + attempt * 0.5)
                                continue
                            # otherwise re-raise
                            raise

                    # if doc is still None, raise
                    if doc is None:
                        raise Exception(f"Failed to open page PDF in Word: {page_pdf}")

                    time.sleep(4)

                    for i in range(1, doc.Sections.Count + 1):

                        ps = doc.Sections(i).PageSetup

                        ps.TopMargin = MARGIN
                        ps.BottomMargin = MARGIN
                        ps.LeftMargin = MARGIN
                        ps.RightMargin = MARGIN

                    for p in doc.Paragraphs:

                        try:
                            p.Format.SpaceBefore = 0
                            p.Format.SpaceAfter = 0
                            p.Format.LineSpacingRule = 0
                        except:
                            pass

                    for t in doc.Tables:

                        try:
                            t.AutoFitBehavior(2)
                            t.AllowAutoFit = True
                            t.Rows.AllowBreakAcrossPages = False
                        except:
                            pass

                    doc.Repaginate()

                    time.sleep(2)

                    scale = 100

                    min_scale = 75

                    wdStatisticPages = 2

                    while scale >= min_scale:

                        pages_now = doc.ComputeStatistics(
                            wdStatisticPages
                        )

                        if pages_now <= 1:
                            break

                        doc.Content.Font.Scaling = scale

                        doc.Repaginate()

                        time.sleep(1)

                        scale -= 2

                    # Save with retries to handle intermittent COM 'call rejected' errors
                    max_save_attempts = 6
                    for save_attempt in range(max_save_attempts):
                        try:
                            doc.SaveAs(out_docx, FileFormat=16)
                            break
                        except pywintypes.com_error as e:
                            code = e.args[0] if e.args else None
                            log.exception(f"SaveAs attempt {save_attempt+1} failed with COM error {code}")
                            # Pump messages and retry for 'Call was rejected by callee'
                            if code == -2147418111 and save_attempt < (max_save_attempts - 1):
                                for _ in range(20):
                                    pythoncom.PumpWaitingMessages()
                                    time.sleep(0.05)
                                time.sleep(0.5 + save_attempt * 0.2)
                                continue
                            # Try restarting Word for certain fatal codes then retry
                            if code in (-2147220995, -2147023174, -2147417836) and save_attempt < (max_save_attempts - 1):
                                try:
                                    log.warning(f"Restarting Word due to SaveAs COM error (attempt {save_attempt+1})")
                                    _restart_word()
                                except Exception:
                                    log.exception("Failed to restart Word during SaveAs retry")
                                time.sleep(0.5 + save_attempt * 0.5)
                                continue
                            raise

                    # Attempt safe close; some RPC errors may occur here — try pumping and retrying
                    try:
                        try:
                            doc.Close(False)
                        except pywintypes.com_error as e_close:
                            code_close = e_close.args[0] if e_close.args else None
                            log.exception(f"doc.Close failed with COM error {code_close}; attempting pump-and-retry")
                            for _ in range(10):
                                pythoncom.PumpWaitingMessages()
                                time.sleep(0.05)
                            try:
                                doc.Close(False)
                            except Exception:
                                log.exception("Second attempt to close doc failed")
                    except Exception:
                        # swallow close errors after logging — we want to continue converting other pages
                        log.exception("Failed to close Word doc after SaveAs")

                    page_docx_files.append(out_docx)

                except Exception:
                    log.exception(f"Failed to convert page {page_num} -> {page_pdf}")
                    try:
                        if doc:
                            try:
                                doc.Close(False)
                            except Exception:
                                log.exception("Failed to close Word doc after error")
                    except Exception:
                        log.exception("Error while handling failed doc close")
                        pass

            # If no per-page DOCX files were produced, surface a clear error instead
            if not page_docx_files:
                log.error("Word conversion failed: no page docx files were produced")
                raise Exception("Word conversion failed: no page docx files were produced")

            # Merge per-page docx files while Word is still available
            max_merge_attempts = 6
            for merge_attempt in range(max_merge_attempts):
                try:
                    merged_doc = word.Documents.Open(page_docx_files[0])

                    wdCollapseEnd = 0
                    wdPageBreak = 7

                    for docx_path in page_docx_files[1:]:
                        end_range = merged_doc.Content
                        end_range.Collapse(wdCollapseEnd)
                        end_range.InsertBreak(wdPageBreak)
                        end_range = merged_doc.Content
                        end_range.Collapse(wdCollapseEnd)
                        end_range.InsertFile(docx_path)

                    merged_doc.SaveAs(output_docx, FileFormat=16)
                    merged_doc.Close(False)

                    return output_docx

                except pywintypes.com_error as e:
                    code = e.args[0] if e.args else None
                    log.exception(f"Merge attempt {merge_attempt+1} failed with COM error {code}")
                    # If call was rejected by callee, try pumping messages then retry
                    if merge_attempt < (max_merge_attempts - 1) and code == -2147418111:
                        log.warning(f"Merge call rejected (code={code}); pumping messages and retrying (attempt {merge_attempt+1})")
                        for _ in range(20):
                            pythoncom.PumpWaitingMessages()
                            time.sleep(0.05)
                        time.sleep(0.5 + merge_attempt * 0.2)
                        continue
                    # Try to recover by restarting Word for certain error codes
                    if merge_attempt < (max_merge_attempts - 1) and code in (-2147220995, -2147023174, -2147417836):
                        try:
                            log.warning(f"Restarting Word before merge retry (attempt {merge_attempt+1})")
                            _restart_word()
                        except Exception:
                            log.exception("Failed to restart Word during merge retry")
                        time.sleep(0.5 + merge_attempt * 0.5)
                        continue
                    raise

        finally:
            # Ensure Word quits and COM is uninitialized even on errors
            try:
                if word:
                    try:
                        word.Quit()
                    except Exception:
                        log.exception("Failed to quit Word application")
            except Exception:
                pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


async def run_pdf_to_word(pdf_file, output_docx, temp_dir):
    """Async helper that runs pdf_to_word on a single-threaded executor.

    Use this from async code so all Word COM operations execute on the same
    dedicated worker thread.
    """
    # Allow opting into a process-isolated Word worker via env var.
    worker_mode = os.environ.get('NAVEXOCR_WORD_WORKER', 'process')

    # If user explicitly wants the thread-based worker, use the existing executor
    if worker_mode.lower() == 'thread':
        loop = asyncio.get_running_loop()
        return await loop.run_in_executor(_word_executor, pdf_to_word, pdf_file, output_docx, temp_dir)

    # Otherwise try process worker first. This isolates Word COM in a separate process
    # which is often more robust on packaged/other-machine deployments.
    try:
        cmd = [sys.executable, '-m', 'navexOCR.services.word_worker_process', pdf_file, output_docx, temp_dir]
        log.info(f"Spawning Word worker process: {cmd}")
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        if proc.returncode == 0:
            log.info("Word worker process completed successfully")
            return output_docx
        else:
            log.error(f"Word worker process failed rc={proc.returncode}; stdout={proc.stdout}; stderr={proc.stderr}")
            # Fallback to thread-based worker for best-effort
    except Exception as e:
        log.exception("Failed to run process-isolated Word worker; falling back to thread worker")

    # Before falling back, attempt to clean any worker-specific artifacts which may
    # cause PermissionErrors when fitz tries to write page PDFs into the same temp_dir.
    try:
        import glob
        import shutil

        def _list_winword_processes():
            try:
                # Use tasklist to report WINWORD.EXE processes; this is diagnostic only
                out = subprocess.check_output('tasklist /FI "IMAGENAME eq WINWORD.EXE" /FO LIST', shell=True, text=True)
                return out
            except Exception:
                return None

        def _try_rmtree(path):
            try:
                shutil.rmtree(path)
                log.info(f"Removed leftover worker dir before fallback: {path}")
                return True
            except Exception as e:
                log.debug(f"rmtree failed for {path}: {e}")
                return False

        # Try multiple times with backoff because child Word processes may still be releasing handles
        worker_dirs = glob.glob(os.path.join(temp_dir, 'worker_*'))
        for w in worker_dirs:
            removed = False
            for attempt in range(8):
                if _try_rmtree(w):
                    removed = True
                    break
                # log diagnostics about WINWORD processes
                winword_info = _list_winword_processes()
                if winword_info:
                    log.warning(f"WINWORD processes present while cleaning worker dir (attempt {attempt+1}):\n{winword_info}")
                else:
                    log.debug(f"No WINWORD info returned on attempt {attempt+1}")
                time.sleep(0.5 + attempt * 0.5)
            if not removed:
                # attempt to tombstone the directory for later sweeper removal
                try:
                    tomb = w + '.delete_me'
                    os.rename(w, tomb)
                    log.info(f"Tombstoned worker dir for later removal: {tomb}")
                except Exception:
                    log.exception(f"Failed to remove worker dir: {w}")

        # remove page_*.pdf and page_*.docx leftovers in parent temp_dir (these are less likely to be locked now)
        for p in glob.glob(os.path.join(temp_dir, 'page_*.pdf')) + glob.glob(os.path.join(temp_dir, 'page_*.docx')):
            for attempt in range(6):
                try:
                    os.remove(p)
                    log.info(f"Removed leftover page file before fallback: {p}")
                    break
                except Exception as e:
                    log.debug(f"Failed to remove page file {p} on attempt {attempt+1}: {e}")
                    time.sleep(0.3 + attempt * 0.3)
    except Exception:
        log.exception("Cleanup before fallback failed")

    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(_word_executor, pdf_to_word, pdf_file, output_docx, temp_dir)
