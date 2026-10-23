import os, io, uuid
from flask import Flask, request, jsonify, send_file, render_template_string
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
import requests

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 120 * 1024 * 1024

DEEPL_KEY = os.environ.get('DEEPL_API_KEY')

# === دالة الترجمة باستخدام DeepL فقط ===
def translate_text(text):
    if not DEEPL_KEY or not text.strip(): return text
    try:
        r = requests.post("https://api-free.deepl.com/v2/translate",
            headers={"Authorization": f"DeepL-Auth-Key {DEEPL_KEY}"},
            json={"text": [text], "target_lang": "AR"}, timeout=15)
        if r.status_code == 200:
            return r.json()['translations'][0]['text']
    except: pass
    return text

# === دالة ترجمة PPTX (نصوص فقط) ===
def translate_pptx(data):
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
                        # إضافة الترجمة كفقرة جديدة تحت النص الأصلي
                        new_p = shape.text_frame.add_paragraph()
                        new_p.text = tr
                        new_p.font.size = Pt(18) # حجم الخط
                        new_p.font.name = 'Arial'
                        new_p.alignment = PP_ALIGN.RIGHT
                        new_p.space_before = Pt(10)
        # تحديث التقدم (للتوضيح فقط، لكننا لن نستخدمه في الواجهة)
        print(f"Slide {i+1}/{total} done")
    
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out

# === الواجهة (HTML) ===
HTML_PAGE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1,user-scalable=no">
<title>الريان للترجمة - PPTX فقط</title>
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800;900&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box;font-family:'Tajawal',sans-serif}
body{min-height:100vh;background:radial-gradient(1200px 600px at 20% 0%, #eef2ff 0%, #d8e2ff 20%, #c9d5ff 40%, #c7b8ff 85%, #b8a6ff 100%);color:#2f365f;padding:20px}
.app{max-width:440px;margin:0 auto;min-height:100vh;padding:12px 18px 100px}
.header{display:flex;justify-content:space-between;align-items:center;padding:6px 4px;margin-bottom:10px}
.logo{font-size:30px;font-weight:900;color:#5b5bd6}
.card{background:rgba(255,255,255,0.85);border-radius:20px;padding:16px;margin-bottom:12px;box-shadow:0 8px 20px rgba(120,130,220,0.1)}
.btn{width:100%;height:48px;border-radius:16px;border:none;background:linear-gradient(135deg,#7c8aff,#8b5cf6);color:#fff;font-weight:800;font-size:16px;cursor:pointer;margin-top:10px}
</style>
</head>
<body>
<div class="app">
  <div class="header"><div class="logo">الريان للترجمة</div></div>
  <div class="card">
    <h3 style="margin-bottom:10px">رفع ملف PPTX (ترجمة النصوص فقط)</h3>
    <form id="uploadForm" action="/upload" method="post" enctype="multipart/form-data">
      <input type="file" name="file" accept=".pptx" required style="width:100%;margin-bottom:10px">
      <button type="submit" class="btn">ترجمة الملف</button>
    </form>
    <div id="status" style="margin-top:10px;color:#6b6b9a"></div>
  </div>
</div>
</body>
</html>
"""

@app.route('/')
def home(): return render_template_string(HTML_PAGE)

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files: return "لم يتم رفع ملف"
    file = request.files['file']
    if file.filename == '': return "اسم الملف فارغ"
    if not file.filename.lower().endswith('.pptx'): return "الملف يجب أن يكون PPTX"
    
    data = file.read()
    bio = translate_pptx(data)
    if bio:
        return send_file(bio, as_attachment=True, download_name="translated.pptx", mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    return "حدث خطأ أثناء الترجمة"

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
