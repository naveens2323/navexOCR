from fastapi import FastAPI

from navexOCR.routes.convert import router

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


@app.get("/")
def home():

    return {

        "message": "navexOCR API Running"
    }