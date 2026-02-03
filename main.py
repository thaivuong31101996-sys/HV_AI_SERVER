from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
import docx
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import os, io, re, logging

app = FastAPI()

def has_image(p):
    return True if p._element.xpath('.//w:drawing') or p._element.xpath('.//w:inline') else False

@app.get("/")
async def home():
    return {"message": "Shop Hai Vuong AI is Ready!"}

@app.post("/process")
async def process_word(
    file: UploadFile = File(...),
    le_trai: float = Form(3.0), le_phai: float = Form(2.0),
    le_tren: float = Form(2.0), le_duoi: float = Form(2.0),
    h1_size: int = Form(14), h2_size: int = Form(13), h3_size: int = Form(13),
    line_spacing: float = Form(1.3), indent: float = Form(1.27),
    h1_upper: bool = Form(True)
):
    try:
        content = await file.read()
        old_doc = Document(io.BytesIO(content))
        new_doc = Document()
        
        # Thiet lap le chuan Hai Vuong
        section = new_doc.sections[0]
        section.left_margin, section.right_margin = Cm(le_trai), Cm(le_phai)
        section.top_margin, section.bottom_margin = Cm(le_tren), Cm(le_duoi)

        for p in old_doc.paragraphs:
            txt = p.text.strip()
            if not txt and not has_image(p): continue 

            new_p = new_doc.add_paragraph()
            
            # Quy tac bac thang AI
            if re.match(r'^(CHƯƠNG|LỜI MỞ ĐẦU|KẾT LUẬN|DANH MỤC|PHỤ LỤC|TÀI LIỆU|MỤC LỤC|PHẦN)', txt.upper()):
                new_p.text = txt.upper() if h1_upper else txt
                new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cur_size, is_bold = h1_size, True
            elif re.match(r'^\d+\.\d+(\s|$)', txt):
                new_p.text = txt
                cur_size, is_bold = h2_size, True
            elif re.match(r'^\d+\.\d+\.\d+', txt):
                new_p.text = txt
                new_p.paragraph_format.left_indent = Cm(indent)
                cur_size, is_bold = h3_size, True
            else:
                new_p.text = txt
                new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                new_p.paragraph_format.line_spacing = line_spacing
                new_p.paragraph_format.first_line_indent = Cm(indent)
                cur_size, is_bold = 13, False

            # Ep font Times New Roman - Cach sua loi 500
            for run in new_p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = Pt(cur_size)
                run.bold = is_bold
                # Dong nay quan trong de khong bi loi 500:
                rfonts = run._element.get_or_add_rPr().get_or_add_rFonts()
                rfonts.set(qn('w:eastAsia'), 'Times New Roman')
                rfonts.set(qn('w:ascii'), 'Times New Roman')
                rfonts.set(qn('w:hAnsi'), 'Times New Roman')

        out_io = io.BytesIO()
        new_doc.save(out_io)
        out_io.seek(0)
        return FileResponse(out_io, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="HV_DONE.docx")
    
    except Exception as e:
        return {"error": str(e)}
