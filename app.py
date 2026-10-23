import os, io, uuid
from flask import Flask, request, jsonify, send_file, render_template_string
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
import fitz, arabic_reshaper
from bidi.algorithm import get_display
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PIL import Image
import google.generativeai as genai

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 120 * 1024 * 1024

GEMINI_KEY = os.environ.get('GEMINI_API_KEY')
if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)
    gemini_model = genai.GenerativeModel('gemini-1.5-flash')

FONT_PATH = "/tmp/Amiri-Regular.ttf"
if not os.path.exists(FONT_PATH):
    try:
        import urllib.request
        urllib.request.urlretrieve("https://github.com/aliftype/amiri/raw/master/Amiri-Regular.ttf", FONT_PATH)
        pdfmetrics.registerFont(TTFont('Amiri', FONT_PATH))
    except: pass

def fix_ar(t):
    try: return get_display(arabic_reshaper.reshape(t))
    except: return t

def translate_text(text):
    if not GEMINI_KEY or len(text.strip()) < 2: return text
    try:
        prompt = f"ترجم هذا النص من الإنجليزية للعربية الفصحى الأكاديمية بدقة، بدون شرح، فقط الترجمة:\n\n{text}"
        resp = gemini_model.generate_content(prompt)
        return resp.text.strip()
    except: return text

def ocr_and_translate_image(image_bytes):
    if not GEMINI_KEY: return ""
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        prompt = "استخرج النص من هذه الصورة، ثم ترجمه إلى العربية الفصحى. اكتب النص الأصلي ثم الترجمة."
        resp = gemini_model.generate_content([prompt, img])
        return resp.text.strip()
    except Exception as e:
        print(e)
        return ""

def translate_pptx_bytes(data):
    prs = Presentation(io.BytesIO(data))
    slides = list(prs.slides)
    total = len(slides)
    if total == 0: return None
    for i, slide in enumerate(slides):
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip():
                orig = shape.text.strip()
                if len(orig) > 1:
                    tr = translate_text(orig)
                    if tr and tr != orig:
                        box = slide.shapes.add_textbox(shape.left, shape.top + shape.height + Pt(10), shape.width, Inches(0.6))
                        box.text_frame.word_wrap = True
                        p = box.text_frame.paragraphs[0]
                        p.text = tr
                        p.font.size = Pt(18)
                        p.alignment = PP_ALIGN.RIGHT
            if shape.shape_type == 13:
                try:
                    image_bytes = shape.image.blob
                    ocr_text = ocr_and_translate_image(image_bytes)
                    if ocr_text:
                        box = slide.shapes.add_textbox(shape.left, shape.top + shape.height + Inches(0.1), shape.width, Inches(1))
                        box.text_frame.word_wrap = True
                        p = box.text_frame.paragraphs[0]
                        p.text = ocr_text
                        p.font.size = Pt(14)
                        p.alignment = PP_ALIGN.RIGHT
                except Exception as e:
                    print(e)
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

def translate_pdf_bytes(data):
    doc = fitz.open(stream=data, filetype="pdf")
    out_bio = io.BytesIO()
    templ = SimpleDocTemplate(out_bio, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    s1 = ParagraphStyle('o', fontName='Helvetica', fontSize=11, leading=14)
    s2 = ParagraphStyle('a', fontName='Amiri' if os.path.exists(FONT_PATH) else 'Helvetica', fontSize=12, leading=16, alignment=2)
    story = []
    for page in doc:
        txt = page.get_text("text").strip()
        if txt:
            for line in txt.split('\n'):
                if line.strip():
                    tr = translate_text(line)
                    story.append(Paragraph(line, s1))
                    story.append(Paragraph(fix_ar(tr), s2))
                    story.append(Spacer(1, 8))
    templ.build(story)
    out_bio.seek(0)
    return out_bio

HTML_PAGE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no">
<title>الريان للترجمة</title>
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800;900&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box;font-family:'Tajawal',sans-serif}
body{min-height:100vh;background:radial-gradient(1200px 600px at 20% 0%, #eef2ff 0%, #d8e2ff 20%, #c9d5ff 40%, #c7b8ff 85%, #b8a6ff 100%);color:#2f365f;padding:20px}
.app{max-width:440px;margin:0 auto;min-height:100vh;padding:12px 18px 100px}
.header{display:flex;justify-content:space-between;align-items:center;padding:6px 4px;margin-bottom:10px}
.logo{font-size:30px;font-weight:900;color:#5b5bd6}
.card{background:rgba(255,255,255,0.85);border-radius:20px;padding:16px;margin-bottom:12px;box-shadow:0 8px 20px rgba(120,130,220,0.1)}
.textarea{width:100%;height:100px;border:1.5px solid #ddd;border-radius:16px;padding:12px;font-family:'Tajawal';font-size:16px;resize:none}
.btn{width:100%;height:48px;border-radius:16px;border:none;background:linear-gradient(135deg,#7c8aff,#8b5cf6);color:#fff;font-weight:800;font-size:16px;cursor:pointer;margin-top:10px}
.result{background:#f8f8ff;border-radius:16px;padding:12px;margin-top:10px;border:1.5px solid #ddd;font-size:16px;min-height:50px}
</style>
</head>
<body>
<div class="app">
  <div class="header"><div class="logo">الريان للترجمة</div></div>
  <div class="card">
    <h3 style="margin-bottom:10px">ترجمة نص</h3>
    <textarea id="inputText" class="textarea" placeholder="اكتب النص هنا..."></textarea>
    <button class="btn" onclick="translateText()">ترجمة</button>
    <div id="resultText" class="result" style="display:none"></div>
  </div>
  <div class="card">
    <h3 style="margin-bottom:10px">رفع ملف (PDF/PPTX)</h3>
    <form id="uploadForm" action="/upload" method="post" enctype="multipart/form-data">
      <input type="file" name="file" accept=".pdf,.pptx" required style="width:100%;margin-bottom:10px">
      <button type="submit" class="btn">ترجمة الملف</button>
    </form>
    <div id="fileStatus" class="result" style="display:none"></div>
  </div>
</div>

<script>
window.translateText = async () => {
  const text = document.getElementById('inputText').value.trim();
  if(!text) return alert('اكتب نصاً للترجمة');
  const resultDiv = document.getElementById('resultText');
  resultDiv.style.display = 'block';
  resultDiv.innerText = 'جاري الترجمة...';
  try {
    const res = await fetch('/translate', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({text})
    });
    const data = await res.json();
    resultDiv.innerText = data.translation;
  } catch(e) { resultDiv.innerText = 'فشل الاتصال بالخادم'; }
};
</script>
</body>
</html>
"""

@app.route('/')
def home(): return render_template_string(HTML_PAGE)

@app.route('/translate', methods=['POST'])
def translate():
    data = request.get_json()
    text = data.get('text', '')
    if not text: return jsonify({'translation': ''})
    translated = translate_text(text)
    return jsonify({'translation': translated})

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files: return "لم يتم رفع ملف"
    file = request.files['file']
    if file.filename == '': return "اسم الملف فارغ"
    if file.filename.lower().endswith('.pptx'):
        data = file.read()
        bio = translate_pptx_bytes(data)
        return send_file(bio, as_attachment=True, download_name="translated.pptx")
    elif file.filename.lower().endswith('.pdf'):
        data = file.read()
        bio = translate_pdf_bytes(data)
        return send_file(bio, as_attachment=True, download_name="translated.pdf")
    return "الملف يجب أن يكون PDF أو PPTX"

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
