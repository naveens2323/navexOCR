import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# =========================================================
# TEMP DIRECTORY
# =========================================================

TEMP_DIR = os.path.join(BASE_DIR, "temp")

os.makedirs(TEMP_DIR, exist_ok=True)

# =========================================================
# MODEL DIRECTORIES
# =========================================================

MODELS_DIR = os.path.join(BASE_DIR, "models")

DET_MODEL_DIR = os.path.join(MODELS_DIR, "det")

REC_MODEL_DIR = os.path.join(MODELS_DIR, "rec")

CLS_MODEL_DIR = os.path.join(MODELS_DIR, "cls")