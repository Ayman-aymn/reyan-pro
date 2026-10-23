import os, io, uuid, json, threading, time
from flask import Flask, request, jsonify, send_file, render_template_string
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
import requests
from functools import wraps
from collections import defaultdict

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 120 * 1024 * 1024

DEEPL_KEY = os.environ.get('DEEPL_API_KEY')

# === حماية Rate Limiting ===
RATE_LIMIT = defaultdict(list)  # IP -> [timestamps]
MAX_REQUESTS_PER_HOUR = 5  # 5 طلبات في الساعة لكل IP

def rate_limit(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        ip = request.remote_addr
        now = time.time()
        # تنظيف الطلبات القديمة (أكثر من ساعة)
        RATE_LIMIT[ip] = [t for t in RATE_LIMIT[ip] if now - t < 3600]
        if len(RATE_LIMIT[ip]) >= MAX_REQUESTS_PER_HOUR:
            return jsonify({"error": "لقد تجاوزت الحد المسموح (5 طلبات في الساعة)"}), 429
        RATE_LIMIT[ip].append(now)
        return f(*args, **kwargs)
    return decorated

# === تخزين الملفات ===
JOBS = {}

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

def process_pptx(job_id, data):
    try:
        prs = Presentation(io.BytesIO(data))
        slides = list(prs.slides)
        total = len(slides)
        if total == 0:
            JOBS[job_id]['progress'] = 100
            JOBS[job_id]['status'] = 'done'
            JOBS[job_id]['file'] = None
            return
        
        for i, slide in enumerate(slides):
            for shape in slide.shapes:
                if shape.has_text_frame and shape.text.strip():
                    orig = shape.text.strip()
                    if len(orig) > 1:
                        tr = translate_text(orig)
                        if tr and tr != orig:
                            new_p = shape.text_frame.add_paragraph()
                            new_p.text = tr
                            new_p.font.size = Pt(18)
                            new_p.font.name = 'Arial'
                            new_p.alignment = PP_ALIGN.RIGHT
                            new_p.space_before = Pt(10)
            JOBS[job_id]['progress'] = int(((i + 1) / total) * 100)
        
        out = io.BytesIO()
        prs.save(out)
        out.seek(0)
        JOBS[job_id]['file'] = out
        JOBS[job_id]['status'] = 'done'
        JOBS[job_id]['progress'] = 100
    except Exception as e:
        print(f"Error: {e}")
        JOBS[job_id]['status'] = f'error: {str(e)}'

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
body{min-height:100vh;background:radial-gradient(1200px 600px at 20% 0%, #eef2ff 0%, #d8e2ff 20%, #c9d5ff 40%, #c7b8ff 85%, #b8a6ff 100%);color:#2f365f;padding:20px;display:flex;justify-content:center;align-items:center}
.app{max-width:440px;width:100%;min-height:100vh;padding:20px}
.card{background:rgba(255,255,255,0.85);border-radius:20px;padding:20px;margin-bottom:15px;box-shadow:0 8px 20px rgba(120,130,220,0.1);text-align:center}
.btn{width:100%;height:48px;border-radius:16px;border:none;background:linear-gradient(135deg,#7c8aff,#8b5cf6);color:#fff;font-weight:800;font-size:16px;cursor:pointer;margin-top:10px}
.btn:disabled{opacity:0.5;cursor:not-allowed}
.btn-download{background:linear-gradient(135deg,#22c55e,#16a34a)}
.progress-container{display:flex;justify-content:center;align-items:center;margin:20px 0}
.progress-circle{width:120px;height:120px;border-radius:50%;background:conic-gradient(#7c8aff 0deg, #ddd 0deg);display:flex;justify-content:center;align-items:center;transition:background 0.3s}
.progress-inner{width:100px;height:100px;border-radius:50%;background:#fff;display:flex;justify-content:center;align-items:center;font-size:24px;font-weight:900;color:#5b5bd6}
.dua{background:#fff7e6;border-radius:16px;padding:15px;margin-top:15px;font-size:15px;line-height:1.8;color:#4a4a7a;border:1px solid #f0e0b0;display:none}
</style>
</head>
<body>
<div class="app">
  <div class="card">
    <h2 style="color:#5b5bd6;margin-bottom:10px">الريان للترجمة</h2>
    <p style="color:#6b6b9a;margin-bottom:15px">رفع ملف PPTX لترجمة النصوص إلى العربية</p>
    <form id="uploadForm" enctype="multipart/form-data">
      <input type="file" name="file" accept=".pptx" required style="width:100%;margin-bottom:10px">
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
      <div class="dua" id="duaBox"></div>
    </div>
  </div>
</div>

<script>
let jobId = null;
let pollInterval = null;

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
    document.getElementById('statusText').innerText = '❌ فشل الاتصال بالخادم';
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
        document.getElementById('duaBox').style.display = 'block';
        document.getElementById('duaBox').innerHTML = `
          <strong>🤲 دعاء لأيمن:</strong><br>
          اللهم احفظ أيمن بعينك التي لا تنام، وارزقه الصحة والعافية وسعة الرزق، ووفقه لما تحب وترضى، واجعل ختام أمره خيرًا. آمين.
        `;
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
    } catch (err) {
      console.error(err);
    }
  }, 1000);
}
</script>
</body>
</html>
"""

@app.route('/')
def home(): return render_template_string(HTML_PAGE)

@app.route('/upload', methods=['POST'])
@rate_limit
def upload():
    if 'file' not in request.files: return jsonify({"error": "no file"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "empty name"}), 400
    if not file.filename.lower().endswith('.pptx'): return jsonify({"error": "not pptx"}), 400
    
    # التحقق من أن الملف هو PPTX حقيقي (وليس ملفاً مموهاً)
    header = file.read(4)
    file.seek(0)
    if header != b'PK\x03\x04':
        return jsonify({"error": "الملف ليس PPTX صالح"}), 400
    
    job_id = str(uuid.uuid4())[:8]
    JOBS[job_id] = {'progress': 0, 'status': 'working', 'file': None}
    
    data = file.read()
    threading.Thread(target=process_pptx, args=(job_id, data), daemon=True).start()
    
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
    return send_file(job['file'], as_attachment=True, download_name="translated.pptx", mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 8080))
    app.run(host='0.0.0.0', port=port)
