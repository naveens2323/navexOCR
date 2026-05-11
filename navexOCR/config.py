import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

import tempfile as _tempfile

PROJECT_ROOT = os.path.abspath(os.path.join(BASE_DIR, os.pardir))
LOCAL_TEMP = os.path.join(PROJECT_ROOT, "temp")

# Prefer system temp dir for process-local temp files so temp data isn't
# stored inside the project folder. Allow override via NAVEXOCR_TEMP_DIR and
# fall back to project-local `temp/` only if necessary.
SYSTEM_TEMP = os.path.join(_tempfile.gettempdir(), "navexocr_temp")
TEMP_DIR = os.environ.get("NAVEXOCR_TEMP_DIR") or SYSTEM_TEMP or LOCAL_TEMP
os.makedirs(TEMP_DIR, exist_ok=True)


MODELS_DIR = os.environ.get("NAVEXOCR_MODELS_DIR") or os.path.join(BASE_DIR, "models")

DET_MODEL_DIR = os.path.join(MODELS_DIR, "det")

REC_MODEL_DIR = os.path.join(MODELS_DIR, "rec")

CLS_MODEL_DIR = os.path.join(MODELS_DIR, "cls")

for p in (DET_MODEL_DIR, REC_MODEL_DIR, CLS_MODEL_DIR):
	if not os.path.exists(p):
		print(f"WARNING: model path not found: {p}")