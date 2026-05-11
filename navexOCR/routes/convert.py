# navexOCR/routes/convert.py

import io
import os
import shutil
import zipfile
import asyncio
import tempfile
from typing import List
from navexOCR.config import TEMP_DIR

from fastapi import APIRouter
from fastapi import UploadFile
from fastapi import File
from fastapi import BackgroundTasks
from fastapi.responses import StreamingResponse
from fastapi.responses import FileResponse
from starlette.concurrency import run_in_threadpool

from navexOCR.services.ocr_service import create_searchable_pdf
from navexOCR.services.word_service import pdf_to_word, run_pdf_to_word
from navexOCR.logger import get_logger
from navexOCR.ocr_threadpool import init_threadpool, submit_task, get_inflight as thread_get_inflight, get_pool_size as thread_get_pool_size
import logging
import pathlib
import os
from navexOCR.ocr_threadpool import get_inflight as thread_get_inflight, get_pool_size as thread_get_pool_size
from navexOCR.services.cleanup_service import start_sweeper

from fastapi import Depends

# =========================================================
# ROUTER
# =========================================================

router = APIRouter(
    prefix="/convert",
    tags=["PDF Converter"]
)

# initialize logging
log = get_logger(__name__)

# Use thread-pool OCR by default (global model loaded once in ocr_service)
USE_PROCESS_OCR = False
try:
    init_threadpool()
    log.info("Using thread-pool OCR worker")
except Exception:
    log.exception("Failed to initialize thread pool")

# sweeper will be started at application startup (navexOCR.main)

# =========================================================
# Concurrency model
# Per-request semaphore is created inside the handler so each request
# can choose to run all uploaded files in parallel (default) or be
# limited by NAVEXOCR_CONCURRENCY environment variable.
# =========================================================
MAX_CONCURRENCY = 64

# =========================================================
# CLEANUP
# =========================================================

def cleanup_folder(path: str):

    try:

        if os.path.exists(path):
            shutil.rmtree(path)

    except Exception:
        pass

# =========================================================
# PROCESS SINGLE PDF
# =========================================================

async def process_pdf(
    uploaded_file: UploadFile,
    index: int,
    semaphore: asyncio.Semaphore
):

    async with semaphore:

        log.info(f"Start processing uploaded file index={index}; filename={uploaded_file.filename}")

        # =================================================
        # Build safe filenames based on original filename + 1-based index
        # =================================================

        original_name = os.path.basename(
            uploaded_file.filename or f"file_{index + 1}.pdf"
        )

        base, ext = os.path.splitext(original_name)
        if not ext:
            ext = ".pdf"

        # ensure 1-based index in filenames for ordering and uniqueness
        pdf_name = f"{base}_{index + 1}{ext}"
        searchable_pdf_name = f"{base}_{index + 1}_searchable.pdf"
        docx_name = f"{base}_{index + 1}.docx"

        # =================================================
        # READ PDF
        # =================================================

        pdf_bytes = await uploaded_file.read()

        if not pdf_bytes:

            raise Exception(
                f"{pdf_name} is empty"
            )

        # =================================================
        # TEMP DIRECTORY
        # =================================================

        # create per-job temp dir inside configured TEMP_DIR so it's easier to find
        temp_dir = tempfile.mkdtemp(prefix="navexocr_", dir=TEMP_DIR)

        input_pdf_path = os.path.join(
            temp_dir,
            pdf_name
        )

        searchable_pdf_path = os.path.join(
            temp_dir,
            searchable_pdf_name
        )

        output_docx_path = os.path.join(
            temp_dir,
            docx_name
        )

        # =================================================
        # SAVE ORIGINAL PDF
        # =================================================

        with open(input_pdf_path, "wb") as f:

            f.write(pdf_bytes)
        log.info(f"Saved uploaded PDF to temporary path: {input_pdf_path}")

        # =================================================
        # OCR -> SEARCHABLE PDF
        # Use process-based worker if enabled; otherwise use create_searchable_pdf
        # =================================================

        # Submit OCR work to the thread pool which uses the single global model
        # We pass bytes to create_searchable_pdf so it uses stream mode
        log.info(f"Submitting OCR job for {pdf_name}")
        fut = submit_task(create_searchable_pdf, input_pdf_bytes=pdf_bytes)
        searchable_pdf_bytes = await asyncio.get_event_loop().run_in_executor(None, fut.result)
        log.info(f"OCR job completed for {pdf_name}")

        # =================================================
        # VALIDATE OCR OUTPUT
        # =================================================

        if searchable_pdf_bytes is None:

            raise Exception(
                "OCR returned None"
            )

        if isinstance(searchable_pdf_bytes, list):

            # Fix for:
            # list index out of range

            if len(searchable_pdf_bytes) == 0:

                raise Exception(
                    "OCR returned empty list"
                )

            searchable_pdf_bytes = (
                searchable_pdf_bytes[0]
            )

        if not isinstance(
            searchable_pdf_bytes,
            (bytes, bytearray)
        ):

            raise Exception(
                f"Unexpected OCR output type: "
                f"{type(searchable_pdf_bytes)}"
            )

        # =================================================
        # SAVE SEARCHABLE PDF
        # =================================================

        with open(searchable_pdf_path, "wb") as f:

            f.write(searchable_pdf_bytes)
        log.info(f"Wrote searchable PDF: {searchable_pdf_path}")

        # =================================================
        # PDF -> WORD
        # =================================================
        log.info(f"Starting Word conversion for: {searchable_pdf_path}")
        await run_pdf_to_word(searchable_pdf_path, output_docx_path, temp_dir)
        log.info(f"Word conversion finished: {output_docx_path}")

        # =================================================
        # VALIDATE DOCX EXISTS
        # =================================================

        if not os.path.exists(output_docx_path):

            raise Exception(
                "DOCX conversion failed"
            )

        # =================================================
        # READ DOCX
        # =================================================

        with open(output_docx_path, "rb") as f:

            docx_bytes = f.read()

        return {
            "filename": docx_name,
            "docx_bytes": docx_bytes,
            "temp_dir": temp_dir
        }

# =========================================================
# CONVERT API
# =========================================================

@router.post("/")
async def convert_pdf(
    background_tasks: BackgroundTasks,

    files: List[UploadFile] = File(
        ...,
        description="Upload one or more PDF files"
    )
):

    try:

        # =================================================
        # VALIDATE INPUT
        # =================================================

        if not files:

            return {
                "success": False,
                "error": "No files uploaded"
            }

        # =================================================
        # PROCESS CONCURRENTLY
        # =================================================

    # Determine concurrency: if NAVEXOCR_CONCURRENCY is set and > 0 use it,
        # otherwise allow all files to be processed in parallel up to MAX_CONCURRENCY
        try:
            conf = int(os.environ.get("NAVEXOCR_CONCURRENCY", "0"))
        except Exception:
            conf = 0

        if conf > 0:
            concurrency = max(1, min(MAX_CONCURRENCY, conf))
        else:
            concurrency = max(1, min(MAX_CONCURRENCY, len(files)))

        semaphore = asyncio.Semaphore(concurrency)

        # per-request logging
        log.info(f"Request: {len(files)} files uploaded; concurrency={concurrency}; use_process_ocr={USE_PROCESS_OCR}")

        results = await asyncio.gather(
            *[
                process_pdf(file, index, semaphore)
                for index, file in enumerate(files)
            ]
        )

        # =================================================
        # CLEANUP TEMP FOLDERS
        # =================================================

        for result in results:
            # enqueue cleanup; worker runs in background
            log.info(f"Enqueuing cleanup for temp dir: {result['temp_dir']}")
            cleanup_folder(result["temp_dir"])

        # =================================================
        # SINGLE FILE -> DOCX
        # =================================================

        if len(results) == 1:

            result = results[0]

            return StreamingResponse(
                io.BytesIO(result["docx_bytes"]),
                media_type=(
                    "application/vnd.openxmlformats-"
                    "officedocument.wordprocessingml.document"
                ),
                headers={
                    "Content-Disposition":
                        f'attachment; filename="{result["filename"]}"'
                }
            )

        # =================================================
        # MULTIPLE FILES -> ZIP
        # =================================================

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(
            zip_buffer,
            mode="w",
            compression=zipfile.ZIP_DEFLATED
        ) as zip_file:

            for result in results:

                zip_file.writestr(
                    result["filename"],
                    result["docx_bytes"]
                )

        zip_buffer.seek(0)

        return StreamingResponse(
            zip_buffer,
            media_type="application/zip",
            headers={
                "Content-Disposition":
                    'attachment; filename="converted_files.zip"'
            }
        )

    except Exception as e:

        log.exception("Conversion request failed")

        return {
            "success": False,
            "error": str(e)
        }


@router.get('/dashboard')
async def dashboard():
    try:
        pool_size = thread_get_pool_size()
        inflight = thread_get_inflight()

        # compute temp dir usage
        total_size = 0
        for root, dirs, files in os.walk(TEMP_DIR):
            for f in files:
                try:
                    fp = os.path.join(root, f)
                    total_size += os.path.getsize(fp)
                except Exception:
                    pass

        return {
            "use_process_ocr": USE_PROCESS_OCR,
            "pool_size": pool_size,
            "inflight_tasks": inflight,
            "temp_dir": TEMP_DIR,
            "temp_size_bytes": total_size
        }
    except Exception as e:
        log.exception("Dashboard error")
        return {"success": False, "error": str(e)}


@router.get('/health')
async def health_check():
    """Check OCR worker and Word COM availability."""
    ok = True
    details = {}
    try:
        # OCR worker health: try a very small dummy image via process or thread
        try:
            # call with a tiny invalid pdf to ensure worker responds (will likely raise but we check responsiveness)
            # test threadpool OCR responsiveness
            try:
                fut = submit_task(create_searchable_pdf, input_pdf_bytes=b'%PDF-1.4\n')
                # use small timeout
                _ = fut.result(timeout=3)
                details['ocr'] = 'ok'
            except Exception as e:
                details['ocr'] = f'error: {str(e)}'
                ok = False
            details['ocr'] = 'ok'
        except Exception as e:
            details['ocr'] = f'error: {str(e)}'
            ok = False

        # Word COM health: try to dispatch Word and quit quickly
        try:
            import pythoncom
            import win32com.client
            pythoncom.CoInitialize()
            w = win32com.client.Dispatch('Word.Application')
            w.Quit()
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            details['word'] = 'ok'
        except Exception as e:
            details['word'] = f'error: {str(e)}'
            ok = False

    except Exception as e:
        log.exception("Health check failed")
        return {"success": False, "error": str(e)}

    return {"ok": ok, "details": details}



@router.get("/logs")
async def get_logs():
    """Return the current log file for quick inspection."""
    try:
        # locate log file used by logger module
        from navexOCR.logger import LOG_FILE
        if not os.path.exists(LOG_FILE):
            return {"success": False, "error": "log file not found"}
        return FileResponse(LOG_FILE, media_type="text/plain", filename=os.path.basename(LOG_FILE))
    except Exception as e:
        log.exception("Failed to read log file")
        return {"success": False, "error": str(e)}