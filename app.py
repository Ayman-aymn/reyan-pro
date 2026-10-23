import os, io, uuid, json, threading, time
from flask import Flask, request, jsonify, send_file, render_template_string
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import fitz
import requests
import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from functools import wraps
from collections import defaultdict

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

DEEPL_KEY = os.environ.get('DEEPL_API_KEY')
GEMINI_KEY = os.environ.get('GEMINI_API_KEY')

if GEMINI_KEY:
    import google.generativeai as genai
    genai.configure(api_key=GEMINI_KEY)
    gemini_model = genai.GenerativeModel('gemini-1.5-flash')

JOBS = {}
FONT_PATH = "/tmp/Amiri-Regular.ttf"

def ensure_font():
    if os.path.exists(FONT_PATH): return True
    urls = [
        "https://github.com/aliftype/amiri/raw/main/Amiri-Regular.ttf",
        "https://raw.githubusercontent.com/aliftype/amiri/main/Amiri-Regular.ttf"
    ]
    for u in urls:
        try:
            import urllib.request
            urllib.request.urlretrieve(u, FONT_PATH)
            if os.path.exists(FONT_PATH): return True
        except: pass
    return False

ensure_font()
try:
    pdfmetrics.registerFont(TTFont('Amiri', FONT_PATH))
    ARABIC_FONT = 'Amiri'
except:
    ARABIC_FONT = 'Helvetica'

def fix_ar(t):
    try: return get_display(arabic_reshaper.reshape(t))
    except: return t

def translate_text(text):
    if not text.strip(): return text
    if DEEPL_KEY:
        try:
            r = requests.post("https://api-free.deepl.com/v2/translate",
                headers={"Authorization": f"DeepL-Auth-Key {DEEPL_KEY}"},
                json={"text": [text], "target_lang": "AR"}, timeout=10)
            if r.status_code == 200: return r.json()['translations'][0]['text']
        except: pass
    if GEMINI_KEY:
        try:
            resp = gemini_model.generate_content(f"ترجم للعربية فقط:\n{text}", request_options={"timeout": 10})
            return resp.text.strip()
        except: pass
    return text

def process_pptx(job_id, data):
    try:
        prs = Presentation(io.BytesIO(data))
        total = len(prs.slides)
        for i, slide in enumerate(prs.slides):
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text.strip():
                    orig = shape.text.strip()
                    if len(orig) > 1:
                        tr = translate_text(orig)
                        if tr and tr != orig:
                            p = shape.text_frame.add_paragraph()
                            p.text = tr
                            p.font.size = Pt(18)
                            p.font.name = 'Arial'
                            p.alignment = PP_ALIGN.RIGHT
                            p.space_before = Pt(10)
            JOBS[job_id]['progress'] = int(((i+1)/total)*100)
        out = io.BytesIO()
        prs.save(out)
        out.seek(0)
        JOBS[job_id]['file'] = out
        JOBS[job_id]['status'] = 'done'
        JOBS[job_id]['progress'] = 100
    except Exception as e:
        JOBS[job_id]['status'] = f'error: {e}'

def process_pdf(job_id, data):
    try:
        doc = fitz.open(stream=data, filetype="pdf")
        total = len(doc)
        out_bio = io.BytesIO()
        templ = SimpleDocTemplate(out_bio, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
        s1 = ParagraphStyle('o', fontName='Helvetica', fontSize=11, leading=14)
        s2 = ParagraphStyle('a', fontName=ARABIC_FONT, fontSize=12, leading=16, alignment=2)
        story = []
        for i, page in enumerate(doc):
            txt = page.get_text("text").strip()
            if txt:
                for line in txt.split('\n'):
                    if line.strip():
                        tr = translate_text(line)
                        story.append(Paragraph(line, s1))
                        story.append(Paragraph(fix_ar(tr), s2))
                        story.append(Spacer(1, 8))
            JOBS[job_id]['progress'] = int(((i+1)/total)*100)
        templ.build(story)
        out_bio.seek(0)
        JOBS[job_id]['file'] = out_bio
        JOBS[job_id]['status'] = 'done'
        JOBS[job_id]['progress'] = 100
    except Exception as e:
        JOBS[job_id]['status'] = f'error: {e}'

HTML_PAGE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>الريان للترجمة</title>
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box;font-family:'Tajawal',sans-serif}
body{min-height:100vh;background:linear-gradient(135deg, #eef2ff 0%, #c7b8ff 100%);color:#2f365f;padding:20px;display:flex;justify-content:center;align-items:center}
.app{max-width:440px;width:100%;padding:20px}
.card{background:rgba(255,255,255,0.9);border-radius:20px;padding:20px;box-shadow:0 8px 20px rgba(120,130,220,0.1);text-align:center}
.btn{width:100%;height:48px;border-radius:16px;border:none;background:linear-gradient(135deg,#7c8aff,#8b5cf6);color:#fff;font-weight:800;font-size:16px;cursor:pointer;margin-top:10px}
.btn:disabled{opacity:0.5;cursor:not-allowed}
.btn-download{background:linear-gradient(135deg,#22c55e,#16a34a)}
.progress-container{display:flex;justify-content:center;margin:20px 0}
.progress-circle{width:120px;height:120px;border-radius:50%;background:conic-gradient(#7c8aff 0deg, #ddd 0deg);display:flex;justify-content:center;align-items:center;transition:background 0.3s}
.progress-inner{width:100px;height:100px;border-radius:50%;background:#fff;display:flex;justify-content:center;align-items:center;font-size:24px;font-weight:900;color:#5b5bd6}
</style>
</head>
<body>
<div class="app">
  <div class="card">
    <h2 style="color:#5b5bd6;margin-bottom:10px">الريان للترجمة</h2>
    <p style="color:#6b6b9a;margin-bottom:15px">رفع ملف PPTX أو PDF لترجمة النصوص</p>
    <form id="uploadForm" enctype="multipart/form-data">
      <input type="file" name="file" accept=".pptx,.pdf" required style="width:100%;margin-bottom:10px">
      <button type="submit" class="btn" id="submitBtn">ترجمة الملف</button>
    </form>
    <div id="progressContainer" style="display:none">
      <div class="progress-container">
        <div class="progress-circle" id="progressCircle">
          <div class="progress-inner" id="progressText">0%</div>
        </div>
      </div>
      <p id="statusText" style="color:#6b6b9a"></p>
    </div>
    <div id="downloadContainer" style="display:none">
      <button class="btn btn-download" id="downloadBtn">📥 تحميل الملف المترجم</button>
      <div style="margin-top:15px;padding:15px;background:#fff7e6;border-radius:12px;border:1px solid #f0e0b0;color:#4a4a7a;font-size:14px;line-height:1.8;text-align:center">
<strong>🤲 دعاء لأيمن:</strong><br>
اللهم احفظ أيمن بعينك التي لا تنام، وارزقه الصحة والعافية وسعة الرزق، ووفقه لما تحب وترضى، واجعل ختام أمره خيرًا. آمين.
</div>
    </div>
  </div>
</div>
<script>
let jobId = null, pollInterval = null;
document.getElementById('uploadForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const formData = new FormData(e.target);
  document.getElementById('submitBtn').disabled = true;
  document.getElementById('progressContainer').style.display = 'block';
  document.getElementById('downloadContainer').style.display = 'none';
  document.getElementById('statusText').innerText = 'جاري رفع الملف...';
  try {
    const res = await fetch('/upload', { method: 'POST', body: formData });
    const data = await res.json();
    if (data.error) {
      document.getElementById('statusText').innerText = '❌ ' + data.error;
      document.getElementById('submitBtn').disabled = false;
      return;
    }
    jobId = data.job_id;
    startPolling();
  } catch (err) {
    document.getElementById('statusText').innerText = '❌ فشل الاتصال';
    document.getElementById('submitBtn').disabled = false;
  }
});
function startPolling() {
  if (pollInterval) clearInterval(pollInterval);
  pollInterval = setInterval(async () => {
    try {
      const res = await fetch('/progress/' + jobId);
      const data = await res.json();
      const progress = data.progress || 0;
      const status = data.status;
      const circle = document.getElementById('progressCircle');
      circle.style.background = `conic-gradient(#7c8aff ${progress * 3.6}deg, #ddd ${progress * 3.6}deg)`;
      document.getElementById('progressText').innerText = progress + '%';
      if (status === 'done') {
        clearInterval(pollInterval);
        document.getElementById('statusText').innerText = '✅ اكتملت الترجمة!';
        document.getElementById('downloadContainer').style.display = 'block';
        document.getElementById('downloadBtn').onclick = () => {
          window.location.href = '/download/' + jobId;
        };
        document.getElementById('submitBtn').disabled = false;
      } else if (status && status.includes('error')) {
        clearInterval(pollInterval);
        document.getElementById('statusText').innerText = '❌ ' + status;
        document.getElementById('submitBtn').disabled = false;
      } else {
        document.getElementById('statusText').innerText = 'جاري الترجمة...';
      }
    } catch (err) {}
  }, 1000);
}
</script>
</body>
</html>
"""

@app.route('/')
def home(): return render_template_string(HTML_PAGE)

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files: return jsonify({"error": "no file"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "empty name"}), 400
    filename_lower = file.filename.lower()
    if not (filename_lower.endswith('.pptx') or filename_lower.endswith('.pdf')):
        return jsonify({"error": "الملف يجب أن يكون PPTX أو PDF"}), 400
    job_id = str(uuid.uuid4())[:8]
    JOBS[job_id] = {'progress': 0, 'status': 'working', 'file': None, 'type': 'pptx' if filename_lower.endswith('.pptx') else 'pdf'}
    data = file.read()
    if filename_lower.endswith('.pptx'):
        threading.Thread(target=process_pptx, args=(job_id, data), daemon=True).start()
    else:
        threading.Thread(target=process_pdf, args=(job_id, data), daemon=True).start()
    return jsonify({"job_id": job_id})

@app.route('/progress/<job_id>')
def progress(job_id):
    job = JOBS.get(job_id)
    if not job: return jsonify({"progress": 0, "status": "not found"})
    return jsonify({"progress": job['progress'], "status": job['status']})

@app.route('/download/<job_id>')
def download(job_id):
    job = JOBS.get(job_id)
    if not job or job['status'] != 'done' or not job['file']:
        return "الملف غير جاهز", 404
    ext = 'pptx' if job['type'] == 'pptx' else 'pdf'
    mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' if ext == 'pptx' else 'application/pdf'
    return send_file(job['file'], as_attachment=True, download_name=f"translated.{ext}", mimetype=mimetype)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
