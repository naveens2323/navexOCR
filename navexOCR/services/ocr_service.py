import os
import io
import fitz

from PIL import Image
from paddleocr import PaddleOCR

from navexOCR.config import (
    DET_MODEL_DIR,
    REC_MODEL_DIR,
    CLS_MODEL_DIR
)

# =========================================================
# DISABLE ONLINE CHECKS
# =========================================================

os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"

os.environ["REQUESTS_CA_BUNDLE"] = ""

os.environ["CURL_CA_BUNDLE"] = ""

# =========================================================
# LOAD OCR
# =========================================================
# =========================================================
# ABSOLUTE MODEL PATHS
# =========================================================

DET_MODEL_PATH = os.path.abspath(DET_MODEL_DIR)

REC_MODEL_PATH = os.path.abspath(REC_MODEL_DIR)

CLS_MODEL_PATH = os.path.abspath(CLS_MODEL_DIR)
print("DET_MODEL_DIR =", DET_MODEL_DIR)
print("REC_MODEL_DIR =", REC_MODEL_DIR)
print("CLS_MODEL_DIR =", CLS_MODEL_DIR)
# =========================================================
# LOAD OCR
# =========================================================

ocr = PaddleOCR(

    use_angle_cls=True,

    lang="en",

    det_model_dir=DET_MODEL_PATH,

    rec_model_dir=REC_MODEL_PATH,

    cls_model_dir=CLS_MODEL_PATH,

    use_gpu=False
)
# =========================================================
# CREATE SEARCHABLE PDF
# =========================================================

def create_searchable_pdf(input_pdf_path=None, output_pdf_path=None, input_pdf_bytes=None):
    """
    Create a searchable PDF.

    Accepts either an input file path (input_pdf_path) or bytes (input_pdf_bytes).
    If output_pdf_path is provided, saves to disk and returns that path.
    Otherwise returns bytes of the generated PDF.
    """

    if input_pdf_bytes is not None:
        input_doc = fitz.open(stream=input_pdf_bytes, filetype="pdf")
    elif input_pdf_path is not None:
        input_doc = fitz.open(input_pdf_path)
    else:
        raise ValueError("Either input_pdf_path or input_pdf_bytes must be provided")

    output_doc = fitz.open()

    for page_index in range(len(input_doc)):

        page = input_doc[page_index]

        pix = page.get_pixmap(dpi=300)

        img_data = pix.tobytes("png")

        # PaddleOCR can accept numpy array or image bytes; pass bytes via PIL->numpy
        image = Image.open(io.BytesIO(img_data)).convert("RGB")

        import numpy as np

        arr = np.array(image)

        # run OCR on numpy array directly
        result = ocr.ocr(arr, cls=True)

        pdf_page = output_doc.new_page(
            width=page.rect.width,
            height=page.rect.height
        )

        # insert image from bytes stream (no disk)
        pdf_page.insert_image(
            page.rect,
            stream=img_data
        )

        if result and result[0]:
            try:
                for line in result[0]:
                    bbox = line[0]
                    text = line[1][0]
                    confidence = float(line[1][1])

                    if confidence < 0.50:
                        continue

                    # bbox coords are in image pixels
                    x = bbox[0][0]
                    y = bbox[0][1]

                    pdf_x = (x / pix.width) * page.rect.width
                    pdf_y = (y / pix.height) * page.rect.height

                    pdf_page.insert_text(
                        (pdf_x, pdf_y),
                        text,
                        fontsize=10,
                        render_mode=3
                    )
            except Exception:
                pass

    # produce output
    if output_pdf_path:
        output_doc.save(
            output_pdf_path,
            garbage=4,
            deflate=True
        )

        output_doc.close()
        input_doc.close()

        return output_pdf_path
    else:
        bio = io.BytesIO()
        output_doc.save(bio, garbage=4, deflate=True)
        output_doc.close()
        input_doc.close()
        bio.seek(0)
        return bio.getvalue()