from fastapi import FastAPI, UploadFile
from parse_excel import parse_excel_v7

app = FastAPI()


@app.get("/")
def read_root():
    return {"Hello": "World"}


@app.post("/parse/excel/")
async def create_upload_file(file: UploadFile):
    return parse_excel_v7(await file.read())
