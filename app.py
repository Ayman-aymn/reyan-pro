import os, io, uuid, json, threading, time, re
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
from supabase import create_client
import google.generativeai as genai

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

# ===== إعدادات الحماية =====
ADMIN_USER = 'admin'
ADMIN_PASS = 'Rayan_Secret_2026_!@#' # كلمة مرور لوحة التحكم (احتفظ بها)
MAX_CHAT_LEN = 1000 # أقصى طول لرسالة المحادثة

# ===== حماية الترويسات (Headers) =====
@app.after_request
def add_security_headers(response):
    response.headers['X-Content-Type-Options'] = 'nosniff'
    response.headers['X-Frame-Options'] = 'DENY'
    response.headers['X-XSS-Protection'] = '1; mode=block'
    response.headers['Strict-Transport-Security'] = 'max-age=31536000; includeSubDomains'
    response.headers['Content-Security-Policy'] = "default-src 'self'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com; script-src 'self' 'unsafe-inline'; img-src 'self' data:;"
    return response

# ===== حماية لوحة التحكم (Basic Auth) =====
def check_auth(username, password):
    return username == ADMIN_USER and password == ADMIN_PASS

def authenticate():
    return app.response_class(
        'ممنوع الدخول. هذه المنطقة محمية.', 401,
        {'WWW-Authenticate': 'Basic realm="Login Required"'}
    )

# ===== الإعدادات =====
DEEPL_KEY = os.environ.get('DEEPL_API_KEY')
GEMINI_KEY = os.environ.get('GEMINI_API_KEY')
SUPABASE_URL = 'https://lmgshabrtmvbnpdhkkst.supabase.co'
SUPABASE_SERVICE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImxtZ3NoYWJydG12Ym5wZGhra3N0Iiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc5MjYxNzgyMSwiZXhwIjoyMTA4MTkzODIxfQ.H8-yzWt9XlQEpKFfzSJxXey48u222UGcl5NuvCtATcA'

if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)
    gemini_model = genai.GenerativeModel('gemini-1.5-flash')

supabase = None
try:
    supabase = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)
except Exception as e:
    print('Supabase error:', e)

def increment_stat(field):
    if not supabase: return
    try:
        res = supabase.table('stats').select(field).eq('id', 1).execute()
        if res.data:
            current = res.data[0].get(field, 0)
            supabase.table('stats').update({field: current + 1}).eq('id', 1).execute()
    except Exception as e:
        print('Stat error:', e)

# ===== الحماية من إغراق الطلبات (Rate Limiting) =====
RATE_LIMIT = defaultdict(list)
MAX_REQUESTS_PER_HOUR = 10

def rate_limit(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        ip = request.remote_addr
        now = time.time()
        RATE_LIMIT[ip] = [t for t in RATE_LIMIT[ip] if now - t < 3600]
        if len(RATE_LIMIT[ip]) >= MAX_REQUESTS_PER_HOUR:
            return jsonify({"error": "لقد تجاوزت الحد المسموح. حاول لاحقاً."}), 429
        RATE_LIMIT[ip].append(now)
        return f(*args, **kwargs)
    return decorated

JOBS = {}

# ===== الخط العربي =====
FONT_PATH = "/tmp/Amiri-Regular.ttf"
if not os.path.exists(FONT_PATH):
    try:
        import urllib.request
        urllib.request.urlretrieve("https://github.com/aliftype/amiri/raw/main/Amiri-Regular.ttf", FONT_PATH)
    except Exception as e:
        print(f"Font download error: {e}")

try:
    pdfmetrics.registerFont(TTFont('Amiri', FONT_PATH))
    ARABIC_FONT = 'Amiri'
except:
    ARABIC_FONT = 'Helvetica'

def fix_ar(t):
    try: return get_display(arabic_reshaper.reshape(t))
    except: return t

# ===== دالة الترجمة =====
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

# ===== معالجة الملفات =====
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
        JOBS[job_id]['status'] = f'error: حدث خطأ في المعالجة'

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
        JOBS[job_id]['status'] = f'error: حدث خطأ في المعالجة'

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
      <div class="dua">
        <strong>🤲 دعاء لأيمن:</strong><br>
        اللهم احفظ أيمن بعينك التي لا تنام، وارزقه الصحة والعافية وسعة الرزق، ووفقه لما تحب وترضى، واجعل ختام أمره خيرًا. آمين.
      </div>
    </div>
  </div>
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
        document.getElementById('downloadBtn').onclick = () => { window.location.href = '/download/' + jobId; };
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
function goChat() {
    const q = document.getElementById('quickChat').value.trim();
    if(q) window.location.href = '/chat_page?q=' + encodeURIComponent(q);
    else window.location.href = '/chat_page';
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
    .msg { max-width: 80%; padding: 12px 15px; border-radius: 15px; line-height: 1.6; font-size: 15px; word-wrap: break-word; }
    .msg.user { background: #5b5bd6; color: white; align-self: flex-start; border-bottom-left-radius: 0; }
    .msg.ai { background: white; color: #333; align-self: flex-end; border-bottom-right-radius: 0; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
    .input-area { background: white; padding: 10px; display: flex; gap: 10px; border-top: 1px solid #ddd; }
    .input-area input { flex: 1; padding: 12px; border: 1px solid #ddd; border-radius: 20px; outline: none; font-size: 15px; }
    .input-area button { background: #5b5bd6; color: white; border: none; padding: 12px 20px; border-radius: 20px; cursor: pointer; font-weight: bold; }
</style>
</head>
<body>
<div class="header">🧠 مساعد أيمن الذكي</div>
<div class="chat-container" id="chatContainer"></div>
<div class="input-area">
    <input type="text" id="msgInput" maxlength="1000" placeholder="اكتب رسالتك هنا..." onkeypress="if(event.key==='Enter') sendMsg()">
    <button onclick="sendMsg()">إرسال</button>
</div>
<script>
let history = [];
function addMsg(sender, text) {
    const c = document.getElementById('chatContainer');
    const d = document.createElement('div');
    d.className = 'msg ' + sender;
    // تأمين ضد XSS: استخدام textContent بدلاً من innerHTML للردود
    if(sender === 'ai') {
        d.textContent = text;
    } else {
        d.textContent = text;
    }
    c.appendChild(d);
    c.scrollTop = c.scrollHeight;
    return d;
}
async function sendMsg() {
    const input = document.getElementById('msgInput');
    const text = input.value.trim();
    if(!text) return;
    if(text.length > 1000) { alert('الرسالة طويلة جداً (الحد 1000 حرف)'); return; }
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
window.onload = () => {
    const q = new URLSearchParams(window.location.search).get('q');
    if(q) { document.getElementById('msgInput').value = q; sendMsg(); }
};
</script>
</body>
</html>
"""

# ===== المسارات =====
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
    
    increment_stat('translated_files')
    job_id = str(uuid.uuid4())[:8]
    is_pptx = '.pptx' in filename_lower
    JOBS[job_id] = {'progress': 0, 'status': 'working', 'file': None, 'type': 'pptx' if is_pptx else 'pdf', 'original_filename': file.filename}
    
    data = file.read()
    if is_pptx:
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
    increment_stat('downloads')
    ext = 'pptx' if job['type'] == 'pptx' else 'pdf'
    mimetype = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' if ext == 'pptx' else 'application/pdf'
    original_name = job.get('original_filename', f"translated.{ext}")
    base_name = original_name.rsplit('.', 1)[0]
    download_name = f"{base_name}_ar.{ext}"
    return send_file(job['file'], as_attachment=True, download_name=download_name, mimetype=mimetype)

@app.route('/chat_page')
def chat_page_route():
    return render_template_string(CHAT_PAGE)

@app.route('/chat', methods=['POST'])
@rate_limit
def chat():
    data = request.json
    message = data.get('message', '')
    history = data.get('history', [])
    if not message: return jsonify({'error': 'رسالة فارغة'})
    if len(message) > MAX_CHAT_LEN: return jsonify({'error': 'الرسالة طويلة جداً'})
    if not GEMINI_KEY: return jsonify({'error': 'مفتاح Gemini غير مهيأ'})
    
    try:
        chat_session = gemini_model.start_chat(history=history)
        prompt = "أنت مساعد ذكي مفيد جداً اسمه أيمن. أجب باللغة العربية بوضوح وإيجاز. " + message
        resp = chat_session.send_message(prompt, request_options={'timeout': 30})
        return jsonify({'reply': resp.text})
    except Exception as e:
        return jsonify({'error': 'حدث خطأ ما. حاول لاحقاً.'})

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
            for page in doc:
                text += page.get_text()
        elif '.pptx' in filename_lower:
            prs = Presentation(io.BytesIO(file.read()))
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        text += shape.text + '\n'
        else:
            return jsonify({'error': 'نوع الملف غير مدعوم'})
            
        if not text.strip():
            return jsonify({'error': 'لم يتم العثور على نص'})
            
        prompt = f"أنت مساعد ذكي. قم بتلخيص النص التالي، أو إعداد اختبار منه إذا طلب منك. النص:\n{text[:5000]}"
        resp = gemini_model.generate_content(prompt, request_options={'timeout': 30})
        return jsonify({'reply': resp.text})
    except Exception as e:
        return jsonify({'error': 'حدث خطأ ما في تحليل الملف.'})

@app.route('/admin-dashboard-secret')
def admin_dashboard():
    auth = request.authorization
    if not auth or not check_auth(auth.username, auth.password):
        return authenticate()
    
    visitors = 0; translated = 0; downloads = 0
    if supabase:
        try:
            res = supabase.table('stats').select('*').eq('id', 1).execute()
            if res.data:
                visitors = res.data[0].get('visitors', 0)
                translated = res.data[0].get('translated_files', 0)
                downloads = res.data[0].get('downloads', 0)
        except Exception as e:
            print('Dashboard error:', e)
    html = f"""
    <html dir='rtl'><head><title>لوحة التحكم</title>
    <style>
    body {{ font-family: sans-serif; background: #1a1a2e; color: #fff; padding: 40px; text-align: center; }}
    .card {{ background: #2a2a4e; padding: 30px; margin: 20px; border-radius: 15px; display: inline-block; width: 250px; box-shadow: 0 4px 15px rgba(0,0,0,0.3); }}
    .num {{ font-size: 48px; color: #d4af37; font-weight: bold; margin-top: 10px; }}
    h1 {{ color: #d4af37; }}
    </style></head><body>
    <h1>📊 إحصائيات الريان للترجمة</h1>
    <div class='card'><h3>👁️ الزوار</h3><div class='num'>{visitors}</div></div>
    <div class='card'><h3>📄 الملفات المترجمة</h3><div class='num'>{translated}</div></div>
    <div class='card'><h3>📥 التحميلات</h3><div class='num'>{downloads}</div></div>
    </body></html>
    """
    return html

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
