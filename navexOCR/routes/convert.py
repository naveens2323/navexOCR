import os
import uuid
import shutil

from fastapi import APIRouter
from fastapi import UploadFile
from fastapi import File
from fastapi import BackgroundTasks
from fastapi import HTTPException

from fastapi.responses import FileResponse
from fastapi.responses import JSONResponse

from navexOCR.config import TEMP_DIR

from navexOCR.services.ocr_service import create_searchable_pdf
from navexOCR.services.word_service import pdf_to_word
from navexOCR.services.job_manager import (
    create_job,
    submit_job,
    set_job_status,
    set_job_result,
    set_job_error,
    get_job,
)

router = APIRouter()


# =========================================================
# CLEANUP FUNCTION
# =========================================================

def cleanup_folder(folder_path):

    try:

        if os.path.exists(folder_path):

            shutil.rmtree(folder_path)

            print(f"🗑 Temp folder deleted: {folder_path}")

    except Exception as e:

        print("Cleanup Error:", e)


# =========================================================
# CONVERT ROUTE
# =========================================================

@router.post("/convert")
async def convert_pdf(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    """
    Accept upload, create a job that processes the PDF in background and returns a job_id.
    Client should poll /status/{job_id} and download from /result/{job_id} when completed.
    """
    content = await file.read()

    job = create_job({"filename": file.filename})

    def _worker(job_id, filename, pdf_bytes):
        try:
            set_job_status(job_id, "running")

            # Create searchable PDF in-memory
            searchable_pdf_bytes = create_searchable_pdf(input_pdf_bytes=pdf_bytes)

            # Word conversion: pdf_to_word can accept bytes or will fallback to a temp file
            result_path_or_bytes = pdf_to_word(
                searchable_pdf_bytes,
                filename=filename,
                temp_dir=TEMP_DIR,
            )

            set_job_result(job_id, result_path_or_bytes)

        except Exception as e:
            set_job_error(job_id, str(e))

    submit_job(_worker, job["id"], file.filename, content)

    return JSONResponse({"job_id": job["id"]})


@router.get("/status/{job_id}")
def job_status(job_id: str):
    job = get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="job not found")
    return {"id": job_id, "status": job["status"], "error": job.get("error")}


@router.get("/result/{job_id}")
def job_result(job_id: str):
    job = get_job(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="job not found")

    if job["status"] != "completed":
        raise HTTPException(status_code=400, detail="job not completed")

    res = job["result"]

    # If result is a file path, stream it; if bytes, return as FileResponse via temp file
    if isinstance(res, str) and os.path.exists(res):
        return FileResponse(res, filename=os.path.basename(res), media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    if isinstance(res, (bytes, bytearray)):
        # send bytes as a response
        return FileResponse(_bytes_to_tempfile(res, job_id), filename=f"{job_id}.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    return JSONResponse({"detail": "no result available"}, status_code=500)


def _bytes_to_tempfile(b: bytes, job_id: str):
    # helper: write bytes once to a temp hidden file for FileResponse to serve
    import os
    path = os.path.join(TEMP_DIR, f"{job_id}.docx")
    with open(path, "wb") as f:
        f.write(b)
    try:
        # attempt to mark hidden on Windows
        import ctypes

        FILE_ATTRIBUTE_HIDDEN = 0x02
        ctypes.windll.kernel32.SetFileAttributesW(path, FILE_ATTRIBUTE_HIDDEN)
    except Exception:
        pass
    return path