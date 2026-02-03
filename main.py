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
    """Kiểm tra xem paragraph có chứa ảnh không"""
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
    # 1. ĐỌC FILE TỪ CLIENT GỬI LÊN
    content = await file.read()
    old_doc = Document(io.BytesIO(content))
    
    # 2. TẠO FILE MỚI SẠCH 100% ĐỂ RE-BUILD
    new_doc = Document()
    
    # Thiết lập lề chuẩn shop Hải Vương
    section = new_doc.sections[0]
    section.left_margin = Cm(le_trai)
    section.right_margin = Cm(le_phai)
    section.top_margin = Cm(le_tren)
    section.bottom_margin = Cm(le_duoi)

    for p in old_doc.paragraphs:
        txt = p.text.strip()
        # Loại bỏ dòng trống rác, chỉ giữ lại ảnh hoặc chữ có nội dung
        if not txt and not has_image(p): 
            continue 

        new_p = new_doc.add_paragraph()
        
        # --- CHIẾN LƯỢC PHÂN CẤP AI (BẬC THANG) ---
        
        # Cấp 1: Chương và các mục lớn tương đương
        if re.match(r'^(CHƯƠNG|LỜI MỞ ĐẦU|KẾT LUẬN|DANH MỤC|PHỤ LỤC|TÀI LIỆU|MỤC LỤC|PHẦN)', txt.upper()):
            new_p.text = txt.upper() if h1_upper else txt
            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            new_p.paragraph_format.space_before = Pt(18)
            new_p.paragraph_format.space_after = Pt(12)
            cur_size, is_bold, cur_indent = h1_size, True, 0
            
        # Cấp 2: Mục 1.1, 1.2 (Sát lề gáy)
        elif re.match(r'^\d+\.\d+(\s|$)', txt):
            new_p.text = txt
            new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            new_p.paragraph_format.space_before = Pt(12)
            cur_size, is_bold, cur_indent = h2_size, True, 0
            
        # Cấp 3: Mục 1.2.1 (Thụt vào bằng Thụt dòng nội dung)
        elif re.match(r'^\d+\.\d+\.\d+', txt):
            new_p.text = txt
            new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            new_p.paragraph_format.left_indent = Cm(indent)
            cur_size, is_bold, cur_indent = h3_size, True, 0

        # Ảnh và Chú thích ảnh/bảng (Căn giữa)
        elif has_image(p) or re.match(r'^(Bảng|Hình|Ảnh|Table|Figure)', txt, re.IGNORECASE):
            new_p.text = txt
            new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            new_p.paragraph_format.space_before = Pt(6)
            cur_size, is_bold, cur_indent = 12, False, 0

        # Văn bản thường (Body text)
        else:
            new_p.text = txt
            new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            new_p.paragraph_format.line_spacing = line_spacing
            new_p.paragraph_format.first_line_indent = Cm(indent)
            new_p.paragraph_format.space_after = Pt(6)
            cur_size, is_bold, cur_indent = 13, False, 0

        # ÉP FONT TIMES NEW ROMAN CHO TOÀN BỘ FILE
        run = new_p.runs[0] if new_p.runs else new_p.add_run(txt)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(cur_size)
        run.bold = is_bold
        # Đảm bảo hỗ trợ gõ tiếng Việt không bị nhảy font
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

    # Xuất file trả về cho anh Vương
    out_io = io.BytesIO()
    new_doc.save(out_io)
    out_io.seek(0)
    return FileResponse(
        out_io, 
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
        filename="HV_AI_DONE.docx"
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=10000)
