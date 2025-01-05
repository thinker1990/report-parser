from typing import Union
from fastapi import FastAPI, File, UploadFile
from parse_excel import parse_excel_v7

app = FastAPI()


@app.get("/")
def read_root():
    return {"Hello": "World"}

@app.post("/parse/excel/")
async def create_upload_file(file: UploadFile):
    file_content = await file.read()
    return parse_excel_v7(file_content)