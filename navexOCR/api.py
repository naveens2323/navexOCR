import os
import uuid

from navexOCR.services.ocr_service import (
    create_searchable_pdf
)

from navexOCR.services.word_service import (
    pdf_to_word
)

from navexOCR.config import TEMP_DIR


def convert_pdf_file(input_pdf):

    job_id = str(uuid.uuid4())

    work_dir = os.path.join(
        TEMP_DIR,
        job_id
    )

    os.makedirs(work_dir, exist_ok=True)

    searchable_pdf = os.path.join(
        work_dir,
        "searchable.pdf"
    )

    output_docx = os.path.join(
        work_dir,
        "output.docx"
    )

    create_searchable_pdf(
        input_pdf,
        searchable_pdf
    )

    pdf_to_word(
        searchable_pdf,
        output_docx,
        work_dir
    )

    return output_docx