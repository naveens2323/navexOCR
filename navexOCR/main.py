from fastapi import FastAPI

from navexOCR.routes.convert import router
from navexOCR.services.cleanup_service import start_sweeper
from navexOCR.services.cleanup_service import start_cleanup_worker
from navexOCR.logger import get_logger
import os

log = get_logger(__name__)

import pyfiglet


banner = pyfiglet.figlet_format(
    "navexOCR",
    font="slant"
)

print(banner)


app = FastAPI(

    title="navexOCR API",

    version="1.0.0"
)

app.include_router(router)


@app.on_event("startup")
async def on_startup():
    try:
        sweeper_ttl = int(os.environ.get("NAVEXOCR_TEMP_TTL_MINUTES", "10"))
        sweeper_interval = int(os.environ.get("NAVEXOCR_SWEEPER_INTERVAL_SECONDS", "60"))
        start_sweeper(interval_seconds=sweeper_interval, ttl_minutes=sweeper_ttl)
        start_cleanup_worker()
        log.info("Temp sweeper started on startup")
    except Exception:
        log.exception("Failed to start temp sweeper on startup")


@app.get("/")
async def root():
    return {
        "message": "Navex OCR API Running"
    }