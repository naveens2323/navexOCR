import os
import time
import fitz
import traceback
import win32com.client


def pdf_to_word(pdf_file, filename=None, temp_dir=None):

    """
    Convert a searchable PDF to DOCX using MS Word COM.

    pdf_file: can be (bytes or path string). If bytes, a single temp file will be created.
    filename: original filename (used for naming result)
    temp_dir: directory to hold short-lived temp files (will be cleaned up)

    Returns: bytes of the resulting .docx
    """

    if temp_dir is None:
        raise ValueError("temp_dir is required")

    os.makedirs(temp_dir, exist_ok=True)

    # Ensure input is a file on disk for Word COM. Create one temp file if needed.
    input_pdf_path = None
    cleanup_files = []

    if isinstance(pdf_file, (bytes, bytearray)):
        input_pdf_path = os.path.join(temp_dir, f"input_{int(time.time())}.pdf")
        with open(input_pdf_path, "wb") as f:
            f.write(pdf_file)
        cleanup_files.append(input_pdf_path)
    else:
        input_pdf_path = pdf_file

    pdf_doc = fitz.open(input_pdf_path)

    # We'll ask Word to convert the entire PDF at once
    output_docx_path = os.path.join(temp_dir, f"{os.path.splitext(filename or 'output')[0]}_{int(time.time())}.docx")

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        # Word requires file paths
        doc = None
        doc = word.Documents.Open(input_pdf_path)

        time.sleep(2)

        # try to adjust layout heuristics (best-effort)
        try:
            for i in range(1, doc.Sections.Count + 1):
                ps = doc.Sections(i).PageSetup
                CM_TO_POINTS = 28.35
                MARGIN = 0.5 * CM_TO_POINTS
                ps.TopMargin = MARGIN
                ps.BottomMargin = MARGIN
                ps.LeftMargin = MARGIN
                ps.RightMargin = MARGIN
        except Exception:
            pass

        doc.Repaginate()
        time.sleep(1)

        doc.SaveAs(output_docx_path, FileFormat=16)
        doc.Close(False)

    finally:
        try:
            word.Quit()
        except Exception:
            pass

    # read result bytes and cleanup temp files
    result_bytes = None
    try:
        with open(output_docx_path, "rb") as f:
            result_bytes = f.read()
    except Exception:
        result_bytes = None

    cleanup_files.append(output_docx_path)

    # cleanup temporary files
    for p in cleanup_files:
        try:
            if os.path.exists(p):
                os.remove(p)
        except Exception:
            pass

    return result_bytes