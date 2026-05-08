from setuptools import setup, find_packages

setup(
    name="navexOCR",
    version="1.0.7",

    packages=find_packages(),

    include_package_data=True,

    package_data={
        "navexOCR": [
            "models/**/*",
        ],
    },

    install_requires=[

        "fastapi",
        "uvicorn",
        "python-multipart",
        "pymupdf",
        "pillow",
        "paddleocr==2.7.0.3",
        "paddlepaddle==2.6.2",
        "pywin32",
        "numpy==1.26.4",
        "pyfiglet"
    ],

    author="Naveen",
    description="Advanced OCR PDF to Word Engine",
)