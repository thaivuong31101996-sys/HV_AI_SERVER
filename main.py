from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
from docx.shared import Cm
import io

app = FastAPI()

# THƯ VIỆN QUY ĐỊNH CỦA CÁC TRƯỜNG (AI TỰ HỌC)
SCHOOL_RULES = {
    "VAN LANG": {"left": 3.0, "right": 2.0, "top": 2.0, "bottom": 2.0},
    "BACH KHOA": {"left": 3.5, "right": 2.0, "top": 2.0, "bottom": 2.0},
    "DEFAULT": {"left": 3.0, "right": 2.0, "top": 2.0, "bottom": 2.0}
}

@app.get("/")
async def home():
    return {"message": "Shop Hai Vuong AI is Ready!"}

@app.post("/process")
async def process_word(file: UploadFile = File(...)):
    try:
        content = await file.read()
        doc = Document(io.BytesIO(content))
        
        # AI TỰ ĐỌC VĂN BẢN ĐỂ BIẾT TRƯỜNG NÀO
        full_text = "\n".join([p.text for p in doc.paragraphs[:10]]) # Chỉ đọc 10 dòng đầu cho nhanh
        
        rule = SCHOOL_RULES["DEFAULT"]
        if "VĂN LANG" in full_text.upper():
            rule = SCHOOL_RULES["VAN LANG"]
            print("Phát hiện đồ án Văn Lang!")
        elif "BÁCH KHOA" in full_text.upper():
            rule = SCHOOL_RULES["BACH KHOA"]

        # TỰ ĐỘ FILE THEO LUẬT CỦA TRƯỜNG
        for section in doc.sections:
            section.left_margin = Cm(rule["left"])
            section.right_margin = Cm(rule["right"])
            section.top_margin = Cm(rule["top"])
            section.bottom_margin = Cm(rule["bottom"])

        out_io = io.BytesIO()
        doc.save(out_io)
        out_io.seek(0)
        return FileResponse(out_io, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="HV_DONE.docx")
    except Exception as e:
        return {"error": str(e)}
