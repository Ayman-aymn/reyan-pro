import os, io, uuid, json, threading, time, re, base64, tempfile
from flask import Flask, request, jsonify, send_file, render_template_string
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import fitz
import requests
import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from functools import wraps
from collections import defaultdict
from supabase import create_client
import google.generativeai as genai

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# ===== إعدادات =====
ADMIN_USER = 'admin'
ADMIN_PASS = 'Rayan_Secret_2026_!@#'
MAX_CHAT_LEN = 1000
MAX_PAGES = 100
MAX_HISTORY_LEN = 44
CHAT_COOLDOWN = 2
MAX_LOGIN_ATTEMPTS = 3
LOGIN_BLOCK_TIME = 900
MODELS_TO_TRY = ['gemini-3.5-flash', 'gemini-flash-latest', 'gemini-3.1-flash-lite-preview', 'gemini-3-flash-preview']

login_attempts = defaultdict(lambda: {'count': 0, 'blocked_until': 0})
last_chat_time = defaultdict(float)

@app.after_request
def add_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'DENY'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    response.headers['Content-Security-Policy'] = "default-src 'self'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com; script-src 'self' 'unsafe-inline'; img-src 'self' data:;"
    response.headers['Referrer-Policy'] = 'no-referrer'
    return response

def check_auth(username, password):
    ip = request.remote_addr
    now = time.time()
    attempt = login_attempts[ip]
    if attempt['blocked_until'] > now: return False
    if username == ADMIN_USER and password == ADMIN_PASS:
        attempt['count'] = 0
        return True
    else:
        attempt['count'] += 1
        if attempt['count'] >= MAX_LOGIN_ATTEMPTS:
            attempt['blocked_until'] = now + LOGIN_BLOCK_TIME
            attempt['count'] = 0
        return False

def authenticate():
    time.sleep(3)
    return app.response_class('ممنوع الدخول.', 401, {'WWW-Authenticate': 'Basic realm="Login Required"'})

DEEPL_KEY = os.environ.get('DEEPL_API_KEY')
GEMINI_KEY = os.environ.get('GEMINI_API_KEY')
SUPABASE_URL = 'https://lmgshabrtmvbnpdhkkst.supabase.co'
SUPABASE_SERVICE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImxtZ3NoYWJydG12Ym5wZGhra3N0Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc5MjYxNzgyMSwiZXhwIjoyMTA4MTkzODIxfQ.H8-yzWt9XlQEpKFfzSJxXey48u222UGcl5NuvCtATcA'

if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)

supabase = None
try: supabase = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)
except: pass

def increment_stat(field):
    if not supabase: return
    try:
        res = supabase.table('stats').select(field).eq('id', 1).execute()
        if res.data:
            current = res.data[0].get(field, 0)
            supabase.table('stats').update({field: current + 1}).eq('id', 1).execute()
    except: pass

RATE_LIMIT = defaultdict(list)
def rate_limit(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        ip = request.remote_addr
        now = time.time()
        RATE_LIMIT[ip] = [t for t in RATE_LIMIT[ip] if now - t < 3600]
        if len(RATE_LIMIT[ip]) >= 10:
            return jsonify({"error": "تجاوزت الحد المسموح"}), 429
        RATE_LIMIT[ip].append(now)
        return f(*args, **kwargs)
    return decorated

JOBS = {}

# ===== الخط العربي =====
try:
    with open('font_b64.txt', 'r') as f:
        font_b64 = f.read().strip()
    font_bytes = base64.b64decode(font_b64)
    FONT_PATH = tempfile.mktemp(suffix='.ttf')
    with open(FONT_PATH, 'wb') as f:
        f.write(font_bytes)
    pdfmetrics.registerFont(TTFont('Amiri', FONT_PATH))
    ARABIC_FONT = 'Amiri'
except Exception as e:
    print(f"Font error: {e}")
    ARABIC_FONT = 'Helvetica'

def fix_ar(t):
    try: return get_display(arabic_reshaper.reshape(t))
    except: return t

def translate_text(text):
    if not text.strip(): return text
    if DEEPL_KEY:
        try:
            r = requests.post("https://api-free.deepl.com/v2/translate", headers={"Authorization": f"DeepL-Auth-Key {DEEPL_KEY}"}, json={"text": [text], "target_lang": "AR"}, timeout=10)
            if r.status_code == 200: return r.json()['translations'][0]['text']
        except: pass
    for model_name in MODELS_TO_TRY:
        try:
            model = genai.GenerativeModel(model_name)
            resp = model.generate_content(f"ترجم للعربية فقط:\n{text}", request_options={"timeout": 10})
            if resp.text: return resp.text.strip()
        except: continue
    return text

def translate_batch(texts):
    if not texts: return []
    if DEEPL_KEY:
        try:
            r = requests.post("https://api-free.deepl.com/v2/translate", 
                headers={"Authorization": f"DeepL-Auth-Key {DEEPL_KEY}"}, 
                json={"text": texts, "target_lang": "AR"}, timeout=30)
            if r.status_code == 200: 
                return [t['text'] for t in r.json()['translations']]
        except: pass
    results = []
    for text in texts:
        results.append(translate_text(text))
    return results

def translate_image_with_gemini(image_bytes):
    try:
        img_b64 = base64.b64encode(image_bytes).decode('utf-8')
        for model_name in MODELS_TO_TRY:
            try:
                model = genai.GenerativeModel(model_name)
                prompt = "استخرج النص من هذه الصورة وترجمه إلى العربية. أعد الترجمة فقط."
                resp = model.generate_content(
                    [{'mime_type': 'image/jpeg', 'data': img_b64}, prompt],
                    request_options={"timeout": 30}
                )
                if resp.text: return resp.text.strip()
            except: continue
    except: pass
    return None

# ===== PPTX: الترجمة أسفل كل صورة مع مسافة =====
def process_pptx(job_id, data, translate_images=False):
    try:
        prs = Presentation(io.BytesIO(data))
        total = len(prs.slides)
        if total > MAX_PAGES: JOBS[job_id]['status'] = 'error: الحد الأقصى 100 شريحة'; return
        
        for i, slide in enumerate(prs.slides):
            # 1. ترجمة النصوص العادية
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text.strip():
                    orig = shape.text.strip()
                    if len(orig) > 1:
                        tr = translate_text(orig)
                        if tr and tr != orig:
                            p = shape.text_frame.add_paragraph()
                            p.text = tr
                            p.font.size = Pt(20)
                            p.font.name = 'Arial'
                            p.alignment = PP_ALIGN.RIGHT
                            p.space_before = Pt(10)
            
            # 2. ترجمة الصور (كل صورة على حدة)
            if translate_images:
                for shape in slide.shapes:
                    if shape.shape_type == 13: # صورة
                        try:
                            image_bytes = shape.image.blob
                            translated_text = translate_image_with_gemini(image_bytes)
                            if translated_text:
                                left = shape.left
                                top = shape.top
                                width = shape.width
                                height = shape.height
                                
                                # مسافة أمان 30 نقطة
                                textbox_top = top + height + Pt(30)
                                textbox_left = left
                                textbox_width = width
                                textbox_height = Pt(80) # ارتفاع أكبر
                                
                                txBox = slide.shapes.add_textbox(textbox_left, textbox_top, textbox_width, textbox_height)
                                tf = txBox.text_frame
                                tf.word_wrap = True
                                tf.text = "ترجمة الصورة:"
                                tf.paragraphs[0].font.size = Pt(16)
                                tf.paragraphs[0].font.name = 'Arial'
                                tf.paragraphs[0].alignment = PP_ALIGN.RIGHT
                                
                                p = tf.add_paragraph()
                                p.text = translated_text
                                p.font.size = Pt(18)
                                p.font.name = 'Arial'
                                p.alignment = PP_ALIGN.RIGHT
                                
                                # إضافة خلفية بيضاء
                                fill = txBox.fill
                                fill.solid()
                                fill.fore_color.rgb = RGBColor(255, 255, 255)
                        except Exception as e:
                            print(f"Error processing image: {e}")
            
            JOBS[job_id]['progress'] = int(((i+1)/total)*100)
            JOBS[job_id]['stage'] = f'ترجمة الشريحة {i+1} من {total}'
        
        out = io.BytesIO()
        prs.save(out)
        out.seek(0)
        JOBS[job_id]['file'] = out
        JOBS[job_id]['status'] = 'done'
        JOBS[job_id]['progress'] = 100
        JOBS[job_id]['type'] = 'pptx'
    except Exception as e:
        JOBS[job_id]['status'] = f'error: {str(e)[:100]}'

# ===== PDF: إنشاء ملف جديد =====
def process_pdf(job_id, data, translate_images=False):
    try:
        doc = fitz.open(stream=data, filetype="pdf")
        total = len(doc)
        if total > MAX_PAGES: JOBS[job_id]['status'] = 'error: الحد الأقصى 100 صفحة'; return
        
        out_bio = io.BytesIO()
        templ = SimpleDocTemplate(out_bio, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
        s_en = ParagraphStyle('en', fontName='Helvetica', fontSize=11, leading=14)
        s_ar = ParagraphStyle('ar', fontName=ARABIC_FONT, fontSize=12, leading=16, alignment=2)
        story = []
        
        for i, page in enumerate(doc):
            txt = page.get_text("text").strip()
            if txt:
                lines = [l for l in txt.split('\n') if l.strip()]
                if lines:
                    translated_lines = translate_batch(lines)
                    for line, tr in zip(lines, translated_lines):
                        story.append(Paragraph(line, s_en))
                        story.append(Paragraph(fix_ar(tr), s_ar))
                        story.append(Spacer(1, 8))
            
            if translate_images:
                for img_index, img in enumerate(page.get_images(full=True)):
                    try:
                        xref = img[0]
                        base_image = doc.extract_image(xref)
                        image_bytes = base_image["image"]
                        translated_text = translate_image_with_gemini(image_bytes)
                        if translated_text:
                            story.append(Paragraph(f"[ترجمة الصورة {img_index+1}]: {fix_ar(translated_text)}", s_ar))
                            story.append(Spacer(1, 8))
                    except Exception as e:
                        print(f"Error processing PDF image: {e}")
            
            JOBS[job_id]['progress'] = int(((i+1)/total)*100)
            JOBS[job_id]['stage'] = f'ترجمة الصفحة {i+1} من {total}'
        
        templ.build(story)
        out_bio.seek(0)
        JOBS[job_id]['file'] = out_bio; JOBS[job_id]['status'] = 'done'; JOBS[job_id]['progress'] = 100
        JOBS[job_id]['type'] = 'pdf'
    except Exception as e:
        JOBS[job_id]['status'] = f'error: {str(e)[:100]}'

# ===== الواجهة الرئيسية =====
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
.card{background:rgba(255,255,255,0.9);border-radius:20px;padding:20px;box-shadow:0 8px 20px rgba(120,130,220,0.1);text-align:center;margin-bottom:15px}
.btn{width:100%;height:48px;border-radius:16px;border:none;background:linear-gradient(135deg,#7c8aff,#8b5cf6);color:#fff;font-weight:800;font-size:16px;cursor:pointer;margin-top:10px}
.btn:disabled{opacity:0.5;cursor:not-allowed}
.btn-download{background:linear-gradient(135deg,#22c55e,#16a34a)}
.progress-container{display:flex;justify-content:center;margin:20px 0}
.progress-circle{width:120px;height:120px;border-radius:50%;background:conic-gradient(#7c8aff 0deg, #ddd 0deg);display:flex;justify-content:center;align-items:center;transition:background 0.3s}
.progress-inner{width:100px;height:100px;border-radius:50%;background:#fff;display:flex;justify-content:center;align-items:center;font-size:24px;font-weight:900;color:#5b5bd6}
.dua{margin-top:15px;padding:15px;background:#fff7e6;border-radius:12px;border:1px solid #f0e0b0;color:#4a4a7a;font-size:14px;line-height:1.8;text-align:center}
.checkbox{display:flex;align-items:center;gap:10px;margin:10px 0;font-size:14px}
.quick-translate{display:flex;gap:5px;margin-top:10px}
.quick-translate textarea{flex:1;padding:10px;border:1px solid #ddd;border-radius:10px;resize:none;height:60px;font-family:inherit}
.quick-translate button{background:#d4af37;color:white;border:none;padding:10px 15px;border-radius:10px;cursor:pointer;font-weight:bold}
</style>
</head>
<body>
<div class="app">
  <div class="card">
    <h2 style="color:#5b5bd6;margin-bottom:10px">الريان للترجمة</h2>
    <p style="color:#6b6b9a;margin-bottom:15px">رفع ملف PPTX أو PDF لترجمة النصوص (سينتج ملف PDF)</p>
    <form id="uploadForm" enctype="multipart/form-data">
      <input type="file" name="file" accept=".pptx,.pdf" required style="width:100%;margin-bottom:10px">
      <div class="checkbox">
        <input type="checkbox" id="translateImages">
        <label for="translateImages">ترجمة الصور أيضاً (أبطأ)</label>
      </div>
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
      <div class="dua">
        <strong>🤲 دعاء لأيمن:</strong><br>
        اللهم احفظ أيمن بعينك التي لا تنام، وارزقه الصحة والعافية وسعة الرزق، ووفقه لما تحب وترضى، واجعل ختام أمره خيرًا. آمين.
      </div>
    </div>
  </div>
  
  <!-- الترجمة السريعة -->
  <div class="card">
    <h3 style="color:#5b5bd6;margin-bottom:10px;">⚡ ترجمة سريعة</h3>
    <div class="quick-translate">
      <textarea id="quickText" placeholder="ألصق النص الإنجليزي هنا..."></textarea>
      <button onclick="quickTranslate()">ترجم</button>
    </div>
    <p id="quickResult" style="margin-top:10px;text-align:right;font-size:14px;color:#2f365f;display:none;padding:10px;background:#f8f9fa;border-radius:10px;"></p>
  </div>

  <!-- المساعد الذكي -->
  <div class="card" style="text-align:right;">
    <h3 style="color:#5b5bd6;margin-bottom:10px;">🧠 مساعد أيمن الذكي</h3>
    <p style="font-size:14px;color:#666;margin-bottom:10px;">اسأل عن أي شيء، أو ارفع ملفاً لتحليله.</p>
    <div style="display:flex;gap:5px;">
      <input type="text" id="quickChat" maxlength="1000" placeholder="اكتب سؤالك..." style="flex:1;padding:10px;border:1px solid #ddd;border-radius:10px;">
      <button onclick="goChat()" style="background:#8b5cf6;color:white;border:none;padding:10px 20px;border-radius:10px;cursor:pointer;">إرسال</button>
    </div>
  </div>
</div>
<script>
document.addEventListener('DOMContentLoaded', () => {
    const savedPref = localStorage.getItem('translateImages');
    if(savedPref === 'true') document.getElementById('translateImages').checked = true;
});
document.getElementById('translateImages').addEventListener('change', (e) => {
    localStorage.setItem('translateImages', e.target.checked);
});

let jobId = null, pollInterval = null;
document.getElementById('uploadForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  const formData = new FormData(e.target);
  const translateImages = document.getElementById('translateImages').checked;
  formData.append('translate_images', translateImages ? 'true' : 'false');
  
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
      const stage = data.stage || '';
      const circle = document.getElementById('progressCircle');
      circle.style.background = `conic-gradient(#7c8aff ${progress * 3.6}deg, #ddd ${progress * 3.6}deg)`;
      document.getElementById('progressText').innerText = progress + '%';
      if (status === 'done') {
        clearInterval(pollInterval);
        document.getElementById('statusText').innerText = '✅ اكتملت الترجمة!';
        document.getElementById('downloadContainer').style.display = 'block';
        document.getElementById('downloadBtn').onclick = () => { window.location.href = '/download/' + jobId; };
        document.getElementById('submitBtn').disabled = false;
      } else if (status && status.includes('error')) {
        clearInterval(pollInterval);
        document.getElementById('statusText').innerText = '❌ ' + status;
        document.getElementById('submitBtn').disabled = false;
      } else {
        document.getElementById('statusText').innerText = stage || 'جاري الترجمة...';
      }
    } catch (err) {}
  }, 1000);
}
function goChat() {
    const q = document.getElementById('quickChat').value.trim();
    if(q) window.location.href = '/chat_page?q=' + encodeURIComponent(q);
    else window.location.href = '/chat_page';
}
async function quickTranslate() {
    const text = document.getElementById('quickText').value.trim();
    if(!text) return;
    const resultDiv = document.getElementById('quickResult');
    resultDiv.style.display = 'block';
    resultDiv.textContent = 'جاري الترجمة...';
    try {
        const res = await fetch('/quick_translate', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({text: text})
        });
        const data = await res.json();
        resultDiv.textContent = data.translation || data.error || 'خطأ';
    } catch(e) {
        resultDiv.textContent = 'فشل الاتصال';
    }
}
</script>
</body>
</html>
"""

# ===== صفحة المساعد =====
CHAT_PAGE = """
<!DOCTYPE html>
<html dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>مساعد أيمن الذكي</title>
<style>
    body { font-family: sans-serif; background: #f0f2f5; margin: 0; padding: 0; display: flex; flex-direction: column; height: 100vh; }
    .header { background: #5b5bd6; color: white; padding: 15px; text-align: center; font-weight: bold; font-size: 18px; }
    .chat-container { flex: 1; overflow-y: auto; padding: 20px; display: flex; flex-direction: column; gap: 10px; }
    .msg { max-width: 80%; padding: 12px 15px; border-radius: 15px; line-height: 1.6; font-size: 15px; word-wrap: break-word; position: relative; }
    .msg.user { background: #5b5bd6; color: white; align-self: flex-start; border-bottom-left-radius: 0; }
    .msg.ai { background: white; color: #333; align-self: flex-end; border-bottom-right-radius: 0; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    .msg .copy-btn { position: absolute; top: 5px; left: 5px; background: #eee; border: none; border-radius: 5px; cursor: pointer; font-size: 10px; padding: 2px 5px; opacity: 0; transition: 0.3s; }
    .msg:hover .copy-btn { opacity: 1; }
    .input-area { background: white; padding: 10px; display: flex; flex-direction: column; gap: 10px; border-top: 1px solid #ddd; }
    .input-row { display: flex; gap: 10px; }
    .input-area input[type="text"] { flex: 1; padding: 12px; border: 1px solid #ddd; border-radius: 20px; outline: none; font-size: 15px; }
    .input-area button { background: #5b5bd6; color: white; border: none; padding: 12px 20px; border-radius: 20px; cursor: pointer; font-weight: bold; }
    .file-input-row { display: flex; gap: 10px; align-items: center; font-size: 12px; }
    .file-input-row input[type="file"] { flex: 1; }
    .file-input-row button { background: #d4af37; font-size: 12px; padding: 8px 15px; }
    .clear-btn { background: #e74c3c; color: white; border: none; padding: 5px 15px; border-radius: 15px; cursor: pointer; font-size: 12px; margin-bottom: 10px; }
</style>
</head>
<body>
<div class="header">🧠 مساعد أيمن الذكي</div>
<div class="chat-container" id="chatContainer"></div>
<button class="clear-btn" onclick="clearChat()">مسح المحادثة</button>
<div class="input-area">
    <div class="input-row">
        <input type="text" id="msgInput" maxlength="1000" placeholder="اكتب رسالتك هنا..." onkeypress="if(event.key==='Enter') sendMsg()">
        <button onclick="sendMsg()">إرسال</button>
    </div>
    <div class="file-input-row">
        <input type="file" id="fileInput" accept=".pdf,.pptx">
        <button onclick="uploadFileAndAsk()" style="background:#8b5cf6;">تحليل الملف</button>
    </div>
</div>
<script>
let history = [];
let lastSendTime = 0;
function addMsg(sender, text) {
    const c = document.getElementById('chatContainer');
    const d = document.createElement('div');
    d.className = 'msg ' + sender;
    d.textContent = text;
    if(sender === 'ai') {
        const btn = document.createElement('button');
        btn.className = 'copy-btn';
        btn.textContent = 'نسخ';
        btn.onclick = () => { navigator.clipboard.writeText(text); btn.textContent = 'تم!'; setTimeout(() => btn.textContent = 'نسخ', 2000); };
        d.appendChild(btn);
    }
    c.appendChild(d);
    c.scrollTop = c.scrollHeight;
    return d;
}
async function sendMsg() {
    const now = Date.now();
    if(now - lastSendTime < 2000) { alert('الرجاء الانتظار ثانيتين'); return; }
    const input = document.getElementById('msgInput');
    const text = input.value.trim();
    if(!text) return;
    if(text.length > 1000) { alert('الرسالة طويلة'); return; }
    if(history.length >= 44) { alert('انتهت المحادثة'); return; }
    lastSendTime = now;
    addMsg('user', text);
    input.value = '';
    const loadingDiv = addMsg('ai', 'جاري التفكير...');
    
    try {
        const res = await fetch('/chat', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({message: text, history: history})
        });
        const data = await res.json();
        if(data.reply) {
            loadingDiv.textContent = data.reply;
            history.push({role: 'user', parts: text});
            history.push({role: 'model', parts: data.reply});
        } else {
            loadingDiv.textContent = '❌ ' + (data.error || 'خطأ');
        }
    } catch(e) {
        loadingDiv.textContent = '❌ فشل الاتصال.';
    }
}
async function uploadFileAndAsk() {
    const file = document.getElementById('fileInput').files[0];
    if(!file) { alert('اختر ملفاً أولاً'); return; }
    const question = document.getElementById('msgInput').value.trim();
    
    addMsg('user', 'رفعت الملف: ' + file.name + (question ? ' | سؤالي: ' + question : ''));
    const loadingDiv = addMsg('ai', 'جاري قراءة الملف والتفكير...');
    
    const fd = new FormData();
    fd.append('file', file);
    fd.append('question', question);
    
    try {
        const res = await fetch('/chat_file', { method: 'POST', body: fd });
        const data = await res.json();
        if(data.reply) {
            loadingDiv.textContent = data.reply;
            history.push({role: 'user', parts: 'ملف: ' + file.name + ' | سؤال: ' + question});
            history.push({role: 'model', parts: data.reply});
        } else {
            loadingDiv.textContent = '❌ ' + (data.error || 'خطأ');
        }
    } catch(e) {
        loadingDiv.textContent = '❌ فشل الاتصال.';
    }
}
function clearChat() {
    document.getElementById('chatContainer').innerHTML = '';
    history = [];
}
window.onload = () => {
    const q = new URLSearchParams(window.location.search).get('q');
    if(q) { document.getElementById('msgInput').value = q; sendMsg(); }
};
</script>
</body>
</html>
"""

@app.route('/')
def home():
    increment_stat('visitors')
    return render_template_string(HTML_PAGE)

@app.route('/upload', methods=['POST'])
@rate_limit
def upload():
    if 'file' not in request.files: return jsonify({"error": "no file"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "empty name"}), 400
    filename_lower = file.filename.lower().strip()
    if not ('.pptx' in filename_lower or '.pdf' in filename_lower):
        return jsonify({"error": "الملف يجب أن يكون PPTX أو PDF"}), 400
    translate_images = request.form.get('translate_images') == 'true'
    increment_stat('translated_files')
    job_id = str(uuid.uuid4())[:8]
    is_pptx = '.pptx' in filename_lower
    JOBS[job_id] = {'progress': 0, 'status': 'working', 'file': None, 'type': 'pptx' if is_pptx else 'pdf', 'original_filename': file.filename, 'stage': 'جاري التحضير...'}
    data = file.read()
    if is_pptx: threading.Thread(target=process_pptx, args=(job_id, data, translate_images), daemon=True).start()
    else: threading.Thread(target=process_pdf, args=(job_id, data, translate_images), daemon=True).start()
    return jsonify({"job_id": job_id})

@app.route('/progress/<job_id>')
def progress(job_id):
    job = JOBS.get(job_id)
    if not job: return jsonify({"progress": 0, "status": "not found"})
    return jsonify({"progress": job['progress'], "status": job['status'], "stage": job.get('stage', '')})

@app.route('/download/<job_id>')
def download(job_id):
    job = JOBS.get(job_id)
    if not job or job['status'] != 'done' or not job['file']: return "الملف غير جاهز", 404
    increment_stat('downloads')
    ext = 'pptx' if job['type'] == 'pptx' else 'pdf'
    mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' if ext == 'pptx' else 'application/pdf'
    original_name = job.get('original_filename', f"translated.{ext}")
    base_name = original_name.rsplit('.', 1)[0]
    download_name = f"{base_name}_ar.{ext}"
    return send_file(job['file'], as_attachment=True, download_name=download_name, mimetype=mimetype)

@app.route('/quick_translate', methods=['POST'])
@rate_limit
def quick_translate():
    data = request.json
    text = data.get('text', '')
    if not text: return jsonify({'error': 'نص فارغ'})
    translation = translate_text(text)
    return jsonify({'translation': translation})

@app.route('/chat_page')
def chat_page_route():
    return render_template_string(CHAT_PAGE)

@app.route('/chat', methods=['POST'])
@rate_limit
def chat():
    ip = request.remote_addr
    now = time.time()
    if now - last_chat_time[ip] < CHAT_COOLDOWN: return jsonify({'error': 'انتظر ثانيتين'}), 429
    last_chat_time[ip] = now
    data = request.json
    message = data.get('message', '')
    history = data.get('history', [])
    if not message: return jsonify({'error': 'رسالة فارغة'})
    if len(message) > MAX_CHAT_LEN: return jsonify({'error': 'الرسالة طويلة جداً'})
    if len(history) >= MAX_HISTORY_LEN: return jsonify({'error': 'انتهت المحادثة'})
    if not GEMINI_KEY: return jsonify({'error': 'مفتاح Gemini غير مهيأ'})
    formatted_history = []
    for msg in history:
        if msg.get('role') == 'user': formatted_history.append({'role': 'user', 'parts': [msg.get('parts')]})
        elif msg.get('role') == 'model': formatted_history.append({'role': 'model', 'parts': [msg.get('parts')]})
    last_error = ''
    for model_name in MODELS_TO_TRY:
        try:
            model = genai.GenerativeModel(model_name)
            chat_session = model.start_chat(history=formatted_history)
            prompt = "أنت مساعد ذكي مفيد جداً اسمه أيمن. أجب باللغة العربية بوضوح وإيجاز. " + message
            resp = chat_session.send_message(prompt, request_options={'timeout': 30})
            return jsonify({'reply': resp.text})
        except Exception as e:
            last_error = str(e)
            continue
    return jsonify({'error': f'فشلت جميع النماذج: {last_error[:100]}'})

@app.route('/chat_file', methods=['POST'])
def chat_file():
    if 'file' not in request.files: return jsonify({'error': 'لا يوجد ملف'})
    file = request.files['file']
    if not GEMINI_KEY: return jsonify({'error': 'مفتاح Gemini غير مهيأ'})
    try:
        text = ''
        filename_lower = file.filename.lower()
        if '.pdf' in filename_lower:
            doc = fitz.open(stream=file.read(), filetype='pdf')
            if len(doc) > MAX_PAGES: return jsonify({'error': f'الملف كبير جداً'})
            for page in doc: text += page.get_text()
        elif '.pptx' in filename_lower:
            prs = Presentation(io.BytesIO(file.read()))
            if len(prs.slides) > MAX_PAGES: return jsonify({'error': f'الملف كبير جداً'})
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame: text += shape.text + '\n'
        else: return jsonify({'error': 'نوع الملف غير مدعوم'})
        if not text.strip(): return jsonify({'error': 'لم يتم العثور على نص'})
        
        question = request.form.get('question', '')
        if question:
            prompt = f"أنت مساعد ذكي. المستخدم يسأل: {question}\nبناءً على النص التالي:\n{text[:5000]}\nأجب على السؤال باللغة العربية."
        else:
            prompt = f"أنت مساعد ذكي. قم بتلخيص النص التالي، أو إعداد اختبار منه إذا طلب منك. النص:\n{text[:5000]}"
        
        last_error = ''
        for model_name in MODELS_TO_TRY:
            try:
                model = genai.GenerativeModel(model_name)
                resp = model.generate_content(prompt, request_options={'timeout': 30})
                return jsonify({'reply': resp.text})
            except Exception as e:
                last_error = str(e)
                continue
        return jsonify({'error': f'فشلت جميع النماذج: {last_error[:100]}'})
    except Exception as e:
        return jsonify({'error': f'خطأ تقني: {str(e)[:100]}'})

@app.route('/admin-dashboard-secret')
def admin_dashboard():
    auth = request.authorization
    if not auth or not check_auth(auth.username, auth.password): return authenticate()
    visitors = 0; translated = 0; downloads = 0
    if supabase:
        try:
            res = supabase.table('stats').select('*').eq('id', 1).execute()
            if res.data:
                visitors = res.data[0].get('visitors', 0); translated = res.data[0].get('translated_files', 0); downloads = res.data[0].get('downloads', 0)
        except: pass
    return f"<html dir='rtl'><body style='background:#1a1a2e;color:#fff;text-align:center;padding:40px;'><h1>📊 إحصائيات الريان</h1><div style='background:#2a2a4e;padding:20px;margin:20px;border-radius:10px;display:inline-block;width:200px;'>👁️ الزوار<br><h2 style='color:#d4af37'>{visitors}</h2></div><div style='background:#2a2a4e;padding:20px;margin:20px;border-radius:10px;display:inline-block;width:200px;'>📄 الملفات<br><h2 style='color:#d4af37'>{translated}</h2></div><div style='background:#2a2a4e;padding:20px;margin:20px;border-radius:10px;display:inline-block;width:200px;'>📥 التحميلات<br><h2 style='color:#d4af37'>{downloads}</h2></div></body></html>"

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
