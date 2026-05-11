import os
import time
import fitz
import traceback
from navexOCR.logger import get_logger

log = get_logger(__name__)
import win32com.client
import pythoncom
import pywintypes
import threading
import asyncio
from concurrent.futures import ThreadPoolExecutor

# Word COM is not thread-safe across multiple threads. Use a global lock
# to serialize Word automation calls so the RPC server doesn't get into
# an inconsistent state when multiple threads dispatch Word concurrently.
_word_lock = threading.Lock()

# Dedicated single-thread executor for Word COM operations. Running all
# Word automation on a single persistent thread reduces COM RPC/connection
# issues caused by multiple threads creating/destroying Word instances.
_word_executor = ThreadPoolExecutor(max_workers=1)


def pdf_to_word(pdf_file, output_docx, temp_dir):

    os.makedirs(temp_dir, exist_ok=True)

    pdf_doc = fitz.open(pdf_file)

    total_pages = pdf_doc.page_count

    page_pdfs = []

    for i in range(total_pages):

        page_pdf_path = os.path.join(
            temp_dir,
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

                out_docx = os.path.join(
                    temp_dir,
                    f"page_{page_num:04d}.docx"
                )

                doc = None

                try:

                    # Documents.Open can fail with COM 'application is busy' errors
                    # Retry a few times with backoff to let Word become responsive
                    doc = None
                    max_attempts = 6
                    for attempt in range(max_attempts):
                        try:
                            doc = word.Documents.Open(page_pdf)
                            break
                        except pywintypes.com_error as e:
                            # common recoverable HRESULTs: busy or RPC/server disconnect
                            code = e.args[0] if e.args else None
                            # RPC_E_SERVERCALL_RETRYLATER
                            if code == -2147417846 and attempt < (max_attempts - 1):
                                time.sleep(1 + attempt)  # backoff
                                continue
                            # Other disconnect/RPC errors: try restarting Word and retry
                            if code in (-2147220995, -2147023174, -2147417836) and attempt < (max_attempts - 1):
                                try:
                                    log.warning(f"Word COM error {code}; restarting Word (attempt {attempt+1})")
                                    _restart_word()
                                except Exception:
                                    log.exception("Failed to restart Word during per-page open")
                                time.sleep(1 + attempt)
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

                    doc.SaveAs(
                        out_docx,
                        FileFormat=16
                    )

                    # Attempt safe close; some RPC errors may occur here, catch and continue
                    try:
                        doc.Close(False)
                    except Exception:
                        try:
                            # best-effort: attempt to abort closing errors
                            pass
                        except Exception:
                            pass

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
            max_merge_attempts = 4
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
                    # Try to recover by restarting Word for certain error codes
                    if merge_attempt < (max_merge_attempts - 1) and code in (-2147220995, -2147023174, -2147417836):
                        try:
                            log.warning(f"Restarting Word before merge retry (attempt {merge_attempt+1})")
                            _restart_word()
                        except Exception:
                            log.exception("Failed to restart Word during merge retry")
                        time.sleep(1 + merge_attempt)
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
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(_word_executor, pdf_to_word, pdf_file, output_docx, temp_dir)
