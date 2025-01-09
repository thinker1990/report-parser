from fastapi import FastAPI, UploadFile
from parse_excel import parse_excel_v7

app = FastAPI()


@app.get("/")
def read_root():
    """
    Returns a greeting message.

    Returns:
        dict: A dictionary containing a greeting message.

    """
    return {"Hello": "World"}


@app.post("/parse/excel/")
async def create_upload_file(file: UploadFile):
    """
    Asynchronously reads the content of an uploaded file and parses it.
    
    Args:
        file (UploadFile): The file to be uploaded and parsed.
    
    Returns:
        Parsed data from the file content.

    """
    return parse_excel_v7(await file.read())
