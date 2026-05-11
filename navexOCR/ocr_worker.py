import io
import multiprocessing
import traceback
import threading


# Worker function will import PaddleOCR inside the process to avoid model duplication in the main process.
def _worker_process(pdf_bytes):
    try:
        # Import here so each process loads its own Paddle models
        import fitz
        from paddleocr import PaddleOCR
        from PIL import Image
        import numpy as np
        import io as _io

        # Use environment/model paths from config
        from navexOCR.config import DET_MODEL_DIR, REC_MODEL_DIR, CLS_MODEL_DIR
        DET_MODEL_PATH = DET_MODEL_DIR
        REC_MODEL_PATH = REC_MODEL_DIR
        CLS_MODEL_PATH = CLS_MODEL_DIR

        ocr = PaddleOCR(use_angle_cls=True, lang='en', det_model_dir=DET_MODEL_PATH, rec_model_dir=REC_MODEL_PATH, cls_model_dir=CLS_MODEL_PATH, use_gpu=False)

        # open PDF
        input_doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        output_doc = fitz.open()

        for page_index in range(len(input_doc)):
            page = input_doc[page_index]
            pix = page.get_pixmap(dpi=300)
            img_data = pix.tobytes('png')
            image = Image.open(_io.BytesIO(img_data)).convert('RGB')
            arr = np.array(image)

            result = ocr.ocr(arr, cls=True)

            pdf_page = output_doc.new_page(width=page.rect.width, height=page.rect.height)
            pdf_page.insert_image(page.rect, stream=img_data)

            if result and result[0]:
                try:
                    for line in result[0]:
                        bbox = line[0]
                        text = line[1][0]
                        confidence = float(line[1][1])
                        if confidence < 0.50:
                            continue
                        x = bbox[0][0]
                        y = bbox[0][1]
                        pdf_x = (x / pix.width) * page.rect.width
                        pdf_y = (y / pix.height) * page.rect.height
                        pdf_page.insert_text((pdf_x, pdf_y), text, fontsize=10, render_mode=3)
                except Exception:
                    pass

        bio = _io.BytesIO()
        output_doc.save(bio, garbage=4, deflate=True)
        output_doc.close()
        input_doc.close()
        bio.seek(0)
        return bio.getvalue()

    except Exception:
        traceback.print_exc()
        raise


# Pool and inflight tracking
_pool = None
_inflight = 0
_inflight_lock = threading.Lock()


def init_pool(processes=None):
    global _pool
    if _pool is None:
        if processes is None:
            processes = max(1, multiprocessing.cpu_count() - 1)
        _pool = multiprocessing.Pool(processes=processes)
    return _pool


def _inflight_inc():
    global _inflight
    with _inflight_lock:
        _inflight += 1


def _inflight_dec():
    global _inflight
    with _inflight_lock:
        if _inflight > 0:
            _inflight -= 1


def get_inflight():
    global _inflight
    with _inflight_lock:
        return _inflight


def get_pool_size():
    if _pool is None:
        return 0
    try:
        return getattr(_pool, '_processes', 0)
    except Exception:
        return 0


def ocr_in_process(pdf_bytes, timeout=None):
    global _pool
    if _pool is None:
        init_pool()
    _inflight_inc()
    try:
        res = _pool.apply_async(_worker_process, (pdf_bytes,))
        return res.get(timeout=timeout)
    finally:
        _inflight_dec()
