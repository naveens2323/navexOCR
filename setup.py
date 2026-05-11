from setuptools import setup, find_packages

setup(
    name="navexOCR",
    version="1.1.1",
    author="Naveen S",
    author_email="ansl6283@gmail.com",

    packages=find_packages(),

    include_package_data=True,

    package_data={
        "navexOCR": [
            "models/**/*",
        ],
    },

    install_requires=[
        "fastapi==0.110.0",
        "starlette==0.36.3",
        "uvicorn==0.29.0",
        "python-multipart==0.0.6",

        "pymupdf==1.20.2",
        "pillow==12.2.0",

        "paddleocr==2.7.0.3",
        "paddlepaddle==2.6.2",

        "pywin32==311",

        "numpy==1.26.4",
        "pyfiglet==1.0.4",
    ],
    description="Advanced OCR PDF to Word Engine",
)