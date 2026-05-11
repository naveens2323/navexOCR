import os
import sys
import time
import fitz
import traceback
import pythoncom
import win32com.client
import pywintypes
import win32process
from navexOCR.logger import get_logger

log = get_logger(__name__)


def _split_pdf_to_pages(pdf_file, temp_dir):
    os.makedirs(temp_dir, exist_ok=True)
    pdf_doc = fitz.open(pdf_file)
    total_pages = pdf_doc.page_count
    page_files = []
    for i in range(total_pages):
        page_pdf_path = os.path.join(temp_dir, f"page_{i+1:04d}.pdf")
        single = fitz.open()
        single.insert_pdf(pdf_doc, from_page=i, to_page=i)
        single.save(page_pdf_path)
        single.close()
        page_files.append(page_pdf_path)
    pdf_doc.close()
    return page_files


def worker_main(pdf_file, output_docx, temp_dir):
    """Run Word COM conversion in this process. Exit codes:
    0 = success, non-zero = failure
    """
    try:
        # operate in a worker-specific subdirectory to avoid clobbering parent temp
        worker_dir = os.path.join(temp_dir, f"worker_{os.getpid()}_{int(time.time()*1000)}")
        os.makedirs(worker_dir, exist_ok=True)
        page_pdfs = _split_pdf_to_pages(pdf_file, worker_dir)

        pythoncom.CoInitialize()
        word = None
        def _start_word():
            # Start a private Word instance for this process
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass
            w = win32com.client.DispatchEx("Word.Application")
            w.Visible = False
            w.DisplayAlerts = 0
            try:
                hwnd = getattr(w, 'Hwnd', None)
                # Do not log window handle/pid to avoid noisy output
            except Exception:
                pass
            return w

        try:
            word = _start_word()

            page_docx_files = []

            for idx, page_pdf in enumerate(page_pdfs):
                page_num = idx + 1
                out_docx = os.path.join(worker_dir, f"page_{page_num:04d}.docx")
                doc = None
                # try open with retries (handle busy/call rejected)
                for attempt in range(8):
                    try:
                        doc = word.Documents.Open(page_pdf)
                        break
                    except pywintypes.com_error as e:
                        code = e.args[0] if e.args else None
                        # Open page COM error: suppress verbose warning
                        if code == -2147418111 and attempt < 7:
                            for _ in range(12):
                                pythoncom.PumpWaitingMessages()
                                time.sleep(0.05)
                            time.sleep(0.3 + attempt * 0.2)
                            continue
                        if code in (-2147220995, -2147023174, -2147417836) and attempt < 7:
                            # restart Word instance and retry
                            try:
                                # Restarting Word due to COM open error
                                try:
                                    if word:
                                        word.Quit()
                                except Exception:
                                    pass
                                try:
                                    pythoncom.CoUninitialize()
                                except Exception:
                                    pass
                                time.sleep(0.5)
                                word = _start_word()
                            except Exception:
                                log.exception("Failed to restart Word during open retries")
                            time.sleep(0.5 + attempt * 0.2)
                            continue
                        if attempt < 7:
                            time.sleep(0.4 + attempt * 0.3)
                            continue
                        raise

                if doc is None:
                    raise Exception(f"Failed to open page PDF in Word: {page_pdf}")

                # Apply the same formatting and scaling used in the in-process worker
                try:
                    CM_TO_POINTS = 28.35
                    MARGIN = 0.5 * CM_TO_POINTS

                    for i in range(1, doc.Sections.Count + 1):
                        try:
                            ps = doc.Sections(i).PageSetup
                            ps.TopMargin = MARGIN
                            ps.BottomMargin = MARGIN
                            ps.LeftMargin = MARGIN
                            ps.RightMargin = MARGIN
                        except Exception:
                            pass

                    for p in doc.Paragraphs:
                        try:
                            p.Format.SpaceBefore = 0
                            p.Format.SpaceAfter = 0
                            p.Format.LineSpacingRule = 0
                        except Exception:
                            pass

                    for t in doc.Tables:
                        try:
                            t.AutoFitBehavior(2)
                            t.AllowAutoFit = True
                            t.Rows.AllowBreakAcrossPages = False
                        except Exception:
                            pass

                    doc.Repaginate()
                    time.sleep(1)

                    scale = 100
                    min_scale = 75
                    wdStatisticPages = 2
                    while scale >= min_scale:
                        try:
                            pages_now = doc.ComputeStatistics(wdStatisticPages)
                        except Exception:
                            pages_now = 1
                        if pages_now <= 1:
                            break
                        try:
                            doc.Content.Font.Scaling = scale
                            doc.Repaginate()
                        except Exception:
                            pass
                        time.sleep(0.5)
                        scale -= 2

                    # now perform SaveAs with retries
                    save_ok = False
                    
                except Exception:
                    log.exception(f"Failed to apply formatting to page {page_num}")
                    # continue to save/close attempt even if formatting failed
                # Wrap per-page save in a try/finally so we always attempt to close the doc
                try:
                    save_ok = False
                    for s_attempt in range(6):
                        try:
                            doc.SaveAs(out_docx, FileFormat=16)
                            save_ok = True
                            break
                        except pywintypes.com_error as e:
                            scode = e.args[0] if e.args else None
                            # SaveAs COM error: suppress verbose warning
                            if scode == -2147418111 and s_attempt < 5:
                                for _ in range(12):
                                    pythoncom.PumpWaitingMessages()
                                    time.sleep(0.05)
                                time.sleep(0.3 + s_attempt * 0.2)
                                continue
                            if scode in (-2147220995, -2147023174, -2147417836) and s_attempt < 5:
                                # restart Word and retry the page save
                                try:
                                    # Restarting Word due to SaveAs COM error
                                    try:
                                        if word:
                                            word.Quit()
                                    except Exception:
                                        pass
                                    try:
                                        pythoncom.CoUninitialize()
                                    except Exception:
                                        pass
                                    time.sleep(0.5)
                                    word = _start_word()
                                except Exception:
                                    log.exception("Failed to restart Word during SaveAs retry")
                                time.sleep(0.5 + s_attempt * 0.3)
                                # try to re-open the same document before saving
                                try:
                                    if os.path.exists(page_pdf):
                                        doc = word.Documents.Open(page_pdf)
                                except Exception:
                                    log.exception("Failed to re-open page after restarting Word")
                                continue
                            if s_attempt < 5:
                                time.sleep(0.6 + s_attempt * 0.3)
                                continue
                            raise
                    if not save_ok:
                        raise Exception(f"Failed to SaveAs page docx: {out_docx}")
                finally:
                    # Ensure we always attempt to close the document, with a pump-and-retry if needed
                    if doc:
                        try:
                            try:
                                doc.Close(False)
                            except pywintypes.com_error as e_close:
                                ccode = e_close.args[0] if e_close.args else None
                                # doc.Close COM error; will pump messages and retry
                                for _ in range(8):
                                    pythoncom.PumpWaitingMessages()
                                    time.sleep(0.05)
                                try:
                                    doc.Close(False)
                                except Exception:
                                    log.exception("Second attempt to close doc failed in worker")
                        except Exception:
                            pass

                page_docx_files.append(out_docx)

            if not page_docx_files:
                raise Exception("No page docx files produced")

            # Merge
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

            # Save merged doc into worker_dir first, then copy to requested output
            merged_out = os.path.join(worker_dir, 'merged.docx')
            save_ok = False
            for s_attempt in range(6):
                try:
                    merged_doc.SaveAs(merged_out, FileFormat=16)
                    save_ok = True
                    break
                except pywintypes.com_error as e:
                    scode = e.args[0] if e.args else None
                    if scode == -2147418111 and s_attempt < 5:
                        for _ in range(12):
                            pythoncom.PumpWaitingMessages()
                            time.sleep(0.05)
                        time.sleep(0.2 + s_attempt * 0.2)
                        continue
                    if s_attempt < 5:
                        time.sleep(0.5 + s_attempt * 0.3)
                        continue
                    raise

            try:
                merged_doc.Close(False)
            except Exception:
                pass

            if not save_ok:
                raise Exception("Failed to save merged docx")

            # Copy merged result back to requested output path
            try:
                import shutil
                shutil.copyfile(merged_out, output_docx)
            except Exception:
                log.exception("Failed to copy merged docx to output location")
                raise

            return 0

        finally:
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

    except Exception:
        log.exception("Word worker process failed")
        return 2


if __name__ == '__main__':
    # Called as: python -m navexOCR.services.word_worker_process <pdf> <out.docx> <temp_dir>
    if len(sys.argv) < 4:
        print("Usage: word_worker_process <pdf> <output_docx> <temp_dir>", file=sys.stderr)
        sys.exit(3)
    pdf_file = sys.argv[1]
    output_docx = sys.argv[2]
    temp_dir = sys.argv[3]
    rc = worker_main(pdf_file, output_docx, temp_dir)
    sys.exit(rc)
