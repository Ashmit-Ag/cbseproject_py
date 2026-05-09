from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse
import tempfile
import shutil
import os

from processor import process_result

app = FastAPI()

UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

SUBJECT_FILE = "subcode12.xlsx"


@app.post("/analyze_12th")
async def analyze_result(file: UploadFile = File(...)):

    if not file.filename.endswith(".TXT") and not file.filename.endswith(".txt"):
        return {"error": "Only TXT files are allowed"}


    original_name = os.path.splitext(file.filename)[0]

    with tempfile.NamedTemporaryFile(delete=False, suffix=".txt") as tmp_txt:
        shutil.copyfileobj(file.file, tmp_txt)
        txt_path = tmp_txt.name

    output_path = tempfile.mktemp(suffix=".xlsx")

    process_result(
        txt_path=txt_path,
        subject_path=SUBJECT_FILE,
        output_path=output_path
    )

    return FileResponse(
        path=output_path,
        filename=f"{original_name}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )