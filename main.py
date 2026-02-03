# Lưu file này là main.py trên Server
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
import docx
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os, io, re

app = FastAPI()

def has_image(p):
    return True if p._element.xpath('.//w:drawing') or p._element.xpath('.//w:inline') else False

@app.post("/process")
async def process_word(
    file: UploadFile = File(...),
    le_trai: float = Form(...), le_phai: float = Form(...),
    le_tren: float = Form(...), le_duoi: float = Form(...),
    h1_size: int = Form(...), h2_size: int = Form(...), h3_size: int = Form(...),
    line_spacing: float = Form(...), indent: float = Form(...),
    h1_upper: bool = Form(...), chua_trang: int = Form(...)
):
    # 1. ĐỌC FILE NÁT CỦA KHÁCH
    content = await file.read()
    old_doc = Document(io.BytesIO(content))
    
    # 2. TẠO FILE MỚI TRẮNG TINH (RESET 100%)
    new_doc = Document()
    
    # Ép lề chuẩn Hải Vương
    section = new_doc.sections[0]
    section.left_margin, section.right_margin = Cm(le_trai), Cm(le_phai)
    section.top_margin, section.bottom_margin = Cm(le_tren), Cm(le_duoi)

    idx = 1
    for p in old_doc.paragraphs:
        if 'w:br' in p._element.xml or 'lastRenderedPageBreak' in p._element.xml:
            idx += 1
        
        txt = p.text.strip()
        if not txt and not has_image(p): continue # Dọn rác dòng trống

        # Tạo đoạn mới trong file sạch
        new_p = new_doc.add_paragraph()
        
        # --- CHIẾN LƯỢC PHÂN CẤP AI ---
        # Cấp 1: Chương
        if re.match(r'^(CHƯƠNG|LỜI MỞ ĐẦU|KẾT LUẬN|DANH MỤC|PHỤ LỤC|TÀI LIỆU|MỤC LỤC|PHẦN)', txt.upper()):
            new_p.text = txt.upper() if h1_upper else txt
            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            new_p.paragraph_format.space_before = Pt(18)
            new_p.paragraph_format.space_after = Pt(12)
            cur_size = h1_size
            is_bold = True; cur_indent = 0
            
        # Cấp 2: 1.1, 1.2
        elif re.match(r'^\d+\.\d+(\s|$)', txt):
            new_p.text = txt
            new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            new_p.paragraph_format.space_before = Pt(12)
            cur_size = h2_size
            is_bold = True; cur_indent = 0 # 1.2 sát lề theo ý anh
            
        # Cấp 3: 1.2.1 (QUY TẮC BẬC THANG)
        elif re.match(r'^\d+\.\d+\.\d+', txt):
            new_p.text = txt
            new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            new_p.paragraph_format.left_indent = Cm(indent) # Thụt bằng nội dung
            cur_size = h3_size
            is_bold = True; cur_indent = 0

        # Ảnh & Chú thích
        elif has_image(p) or re.match(r'^(Bảng|Hình|Ảnh|Table|Figure)', txt, re.IGNORECASE):
            new_p.text = txt
            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cur_size = 12; is_bold = False; cur_indent = 0
            new_p.paragraph_format.space_before = Pt(6)

        # Văn bản thường
        else:
            new_p.text = txt
            new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            new_p.paragraph_format.line_spacing = line_spacing
            new_p.paragraph_format.first_line_indent = Cm(indent)
            new_p.paragraph_format.space_after = Pt(6)
            cur_size = 13; is_bold = False; cur_indent = 0

        # ÉP FONT TIMES NEW ROMAN TRIỆT ĐỂ
        for run in new_p.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(cur_size)
            run.bold = is_bold
            run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    # Trả file về cho anh Vương
    out_io = io.BytesIO()
    new_doc.save(out_io)
    out_io.seek(0)
    return FileResponse(out_io, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="HV_DONE.docx")