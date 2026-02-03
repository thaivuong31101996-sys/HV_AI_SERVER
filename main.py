from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn
import io, re

app = FastAPI()

@app.get("/")
async def home():
    return {"message": "Shop Hai Vuong AI is Ready!"}

@app.post("/process")
async def process_word(
    file: UploadFile = File(...),
    le_trai: float = Form(3.0), le_phai: float = Form(2.0),
    le_tren: float = Form(2.0), le_duoi: float = Form(2.0)
):
    try:
        content = await file.read()
        doc = Document(io.BytesIO(content))
        
        # 1. AI Tự động chỉnh lề chuẩn đóng cuốn
        for section in doc.sections:
            section.left_margin, section.right_margin = Cm(le_trai), Cm(le_phai)
            section.top_margin, section.bottom_margin = Cm(le_tren), Cm(le_duoi)

        # 2. AI Tự học cách ép font và xử lý mục lục (Bản siêu bền)
        for p in doc.paragraphs:
            try:
                txt = p.text.strip()
                # Tự nhận diện Chương/Mục để căn giữa
                if re.match(r'^(CHƯƠNG|MỞ ĐẦU|KẾT LUẬN|DANH MỤC)', txt.upper()):
                    p.alignment = 1 
                
                # Ép font Times New Roman an toàn
                for run in p.runs:
                    run.font.name = 'Times New Roman'
                    rPr = run._element.get_or_add_rPr()
                    rFonts = rPr.get_or_add_rFonts()
                    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
                    rFonts.set(qn('w:ascii'), 'Times New Roman')
                    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
            except:
                continue # Nếu lỗi một đoạn nhỏ, AI tự bỏ qua để làm tiếp các đoạn khác

        out_io = io.BytesIO()
        doc.save(out_io)
        out_io.seek(0)
        return FileResponse(out_io, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="HV_DONE.docx")
    except Exception as e:
        return {"error": str(e)}
