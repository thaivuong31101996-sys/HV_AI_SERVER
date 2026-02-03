from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse, JSONResponse
from docx import Document
from docx.shared import Cm
import io
import traceback

app = FastAPI()

# TRÍ NHỚ: LUẬT CỦA CÁC TRƯỜNG
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
        # 1. Đọc file từ máy tiệm gửi lên
        content = await file.read()
        try:
            doc = Document(io.BytesIO(content))
        except Exception as e:
            return JSONResponse(status_code=400, content={"message": "File lỗi không đọc được, anh kiểm tra lại file Word nhé!"})
        
        # 2. AI TỰ ĐỌC TÊN TRƯỜNG (CHẾ ĐỘ AN TOÀN)
        rule = SCHOOL_RULES["DEFAULT"]
        try:
            # Đọc 10 đoạn đầu tiên để tìm tên trường
            full_text = ""
            for p in doc.paragraphs[:10]:
                full_text += " " + p.text.upper()
            
            if "VĂN LANG" in full_text:
                print("Phát hiện: VĂN LANG")
                rule = SCHOOL_RULES["VAN LANG"]
            elif "BÁCH KHOA" in full_text:
                print("Phát hiện: BÁCH KHOA")
                rule = SCHOOL_RULES["BACH KHOA"]
        except:
            # Nếu lỗi phần đọc tên, cứ dùng lề mặc định, không được báo lỗi 500
            print("Không nhận diện được tên trường, dùng Default")
            pass

        # 3. ÁP DỤNG LỀ
        for section in doc.sections:
            section.left_margin = Cm(rule["left"])
            section.right_margin = Cm(rule["right"])
            section.top_margin = Cm(rule["top"])
            section.bottom_margin = Cm(rule["bottom"])

        # 4. Xuất file
        out_io = io.BytesIO()
        doc.save(out_io)
        out_io.seek(0)
        return FileResponse(out_io, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="HV_DONE.docx")

    except Exception as e:
        # Bắt lỗi hệ thống và in ra để debug
        error_msg = f"Lỗi Server: {str(e)}\n{traceback.format_exc()}"
        print(error_msg)
        return JSONResponse(status_code=500, content={"message": str(e)})
