import os, io
from flask import Flask, request, jsonify, send_file, render_template_string
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
import fitz
import requests
import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

app = Flask(__name__)

# === مفاتيح آمنة من متغيرات البيئة ===
DEEPL_KEY = os.environ.get('DEEPL_API_KEY')
OCR_KEY = os.environ.get('OCR_API_KEY')

# تحميل خط عربي مرة واحدة
FONT_PATH = "/tmp/Amiri-Regular.ttf"
if not os.path.exists(FONT_PATH):
    try:
        import urllib.request
        urllib.request.urlretrieve("https://github.com/aliftype/amiri/raw/master/Amiri-Regular.ttf", FONT_PATH)
        pdfmetrics.registerFont(TTFont('Amiri', FONT_PATH))
    except: pass

def fix_arabic(t):
    try: return get_display(arabic_reshaper.reshape(t))
    except: return t

def translate_text(text):
    if not text or not DEEPL_KEY: return text
    try:
        r = requests.post("https://api-free.deepl.com/v2/translate",
            headers={"Authorization": f"DeepL-Auth-Key {DEEPL_KEY}"},
            json={"text": [text], "target_lang": "AR"}, timeout=20)
        if r.status_code==200:
            return r.json()['translations'][0]['text']
    except: pass
    return text

def ocr_bytes(img_bytes):
    if not OCR_KEY: return ""
    try:
        r = requests.post("https://api.ocr.space/parse/image",
            files={"file": ("img.png", img_bytes, "image/png")},
            data={"apikey": OCR_KEY, "language": "ara", "OCREngine": 2}, timeout=30)
        j=r.json()
        if j.get('ParsedResults'): return j['ParsedResults'][0].get('ParsedText','').strip()
    except: pass
    return ""

def translate_pptx_bytes(data: bytes):
    prs = Presentation(io.BytesIO(data))
    for slide in prs.slides:
        for shape in list(slide.shapes):
            # 1. ترجمة النص
            if shape.has_text_frame and shape.text.strip():
                orig = shape.text.strip()
                # تجنب ترجمة النص المترجم مسبقا
                if len(orig) > 1:
                    trans = translate_text(orig)
                    if trans != orig:
                        left, top, width, height = shape.left, shape.top + shape.height + Pt(10), shape.width, Inches(0.6)
                        box = slide.shapes.add_textbox(left, top, width, height)
                        tf = box.text_frame
                        tf.word_wrap=True
                        p=tf.paragraphs[0]
                        p.text=trans
                        p.font.size=Pt(18)
                        p.alignment=PP_ALIGN.RIGHT
            # 2. ترجمة الصورة
            if shape.shape_type == 13: # picture
                try:
                    ocr_t = ocr_bytes(shape.image.blob)
                    if ocr_t:
                        trans = translate_text(ocr_t)
                        left, top = shape.left, shape.top + shape.height + Inches(0.1)
                        box = slide.shapes.add_textbox(left, top, shape.width, Inches(0.8))
                        tf=box.text_frame
                        tf.word_wrap=True
                        p=tf.paragraphs[0]
                        p.text=trans
                        p.font.size=Pt(16)
                except: pass
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

def translate_pdf_bytes(data: bytes):
    doc = fitz.open(stream=data, filetype="pdf")
    out_bio = io.BytesIO()
    doc_out = SimpleDocTemplate(out_bio, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    style_orig = ParagraphStyle('orig', fontName='Helvetica', fontSize=11, leading=14)
    style_ar = ParagraphStyle('ar', fontName='Amiri' if os.path.exists(FONT_PATH) else 'Helvetica', fontSize=12, leading=16, alignment=2)
    story=[]
    for page in doc:
        txt = page.get_text("text").strip()
        if txt:
            for line in txt.split('\n'):
                if line.strip():
                    story.append(Paragraph(line, style_orig))
                    story.append(Paragraph(fix_arabic(translate_text(line)), style_ar))
                    story.append(Spacer(1,8))
        else:
            pix = page.get_pixmap(dpi=200)
            ocr_t = ocr_bytes(pix.tobytes("png"))
            if ocr_t:
                story.append(Paragraph(fix_arabic(translate_text(ocr_t)), style_ar))
                story.append(Spacer(1,12))
    doc_out.build(story)
    out_bio.seek(0)
    return out_bio

HTML = """<!DOCTYPE html><html lang=ar dir=rtl><head><meta charset=UTF-8><meta name=viewport content="width=device-width,initial-scale=1"><title>الريان</title>
<style>body{background:#eef2ff;font-family:sans-serif;padding:20px}.card{background:#fff;border-radius:16px;padding:16px;max-width:500px;margin:20px auto}</style></head><body>
<div class=card><h2>الريان للترجمة - الحل المثالي</h2><textarea id=t style="width:100%;height:80px"></textarea><button onclick="tr()" style="width:100%;height:45px;background:#7c5cff;color:#fff;border:none;border-radius:10px;margin-top:10px">ترجم</button><div id=r style="margin-top:10px;background:#f5f5ff;padding:10px;border-radius:10px"></div></div>
<div class=card><form action=/upload method=post enctype=multipart/form-data><input type=file name=file accept=.pdf,.pptx required><button style="width:100%;height:45px;background:#5b5bd6;color:#fff;border:none;border-radius:10px;margin-top:10px">ترجم ملف PDF/PPTX</button></form></div>
<script>async function tr(){let txt=document.getElementById('t').value;let res=await fetch('/translate',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({text:txt})});let j=await res.json();document.getElementById('r').innerText=j.translation}</script></body></html>"""

@app.route('/')
def home(): return render_template_string(HTML)

@app.route('/translate', methods=['POST'])
def api_tr():
    d=request.get_json()
    return jsonify({"translation": translate_text(d.get('text',''))})

@app.route('/upload', methods=['POST'])
def upl():
    f=request.files.get('file')
    if not f: return "لا يوجد ملف",400
    data=f.read()
    if f.filename.lower().endswith('.pptx'):
        out=translate_pptx_bytes(data)
        return send_file(out, as_attachment=True, download_name="translated.pptx", mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    elif f.filename.lower().endswith('.pdf'):
        out=translate_pdf_bytes(data)
        return send_file(out, as_attachment=True, download_name="translated.pdf", mimetype="application/pdf")
    return "pdf or pptx only",400

if __name__=='__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT',8080)))
