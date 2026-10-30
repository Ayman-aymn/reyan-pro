import os, io, uuid, json, threading, time, re, base64, tempfile
from flask import Flask, request, jsonify, send_file, render_template_string, send_from_directory, redirect
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
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
MODELS_TO_TRY = ['gemini-3.5-flash']
WHATSAPP_NUMBER = '249925284869'
PHONE_NUMBER = '+249925284869'
WHATSAPP_CHANNEL = 'https://whatsapp.com/channel/0029VbDbf1V9MF8xMszlN61S'
SITE_URL = 'https://reyan-pro.onrender.com'

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
SUPABASE_URL = os.environ.get('SUPABASE_URL')
SUPABASE_SERVICE_KEY = os.environ.get('SUPABASE_SERVICE_KEY')

if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)

supabase = None
if SUPABASE_URL and SUPABASE_SERVICE_KEY:
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
        from google import genai
        from google.genai import types
        
        client = genai.Client(api_key=GEMINI_KEY)
        image_part = types.Part.from_bytes(data=image_bytes, mime_type='image/jpeg')
        prompt = "استخرج النص من هذه الصورة وترجمه إلى العربية. أعد الترجمة فقط."
        
        for model_name in MODELS_TO_TRY:
            try:
                resp = client.models.generate_content(model=model_name, contents=[prompt, image_part])
                if resp.text: return resp.text.strip()
            except Exception as e:
                print(f"Gemini image error with {model_name}: {e}")
                continue
    except Exception as e:
        print(f"Gemini client error: {e}")
    return None

# ===== PPTX =====

CHAT_PAGE = """
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


def process_pptx(job_id, data, translate_images=False, to_pdf=False):
    print(f"=== process_pptx: translate_images={translate_images}, to_pdf={to_pdf} ===")
    try:
        prs = Presentation(io.BytesIO(data))
        total = len(prs.slides)
        if total > MAX_PAGES: JOBS[job_id]['status'] = 'error: الحد الأقصى 100 شريحة'; return
        
        if to_pdf:
            # ... (نفس كود PDF السابق، لا تغيير) ...
            out_bio = io.BytesIO()
            templ = SimpleDocTemplate(out_bio, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
            s_en = ParagraphStyle('en', fontName='Helvetica', fontSize=12, leading=16)
            s_ar = ParagraphStyle('ar', fontName=ARABIC_FONT, fontSize=14, leading=18, alignment=2)
            story = []
            
            for i, slide in enumerate(prs.slides):
                story.append(Paragraph(f"--- شريحة {i+1} ---", s_en))
                story.append(Spacer(1, 10))
                english_texts = []
                for shape in slide.shapes:
                    if shape.has_text_frame and shape.text.strip():
                        orig = shape.text.strip()
                        if len(orig) > 1: english_texts.append(orig)
                
                if english_texts:
                    for eng in english_texts:
                        story.append(Paragraph(eng, s_en))
                        story.append(Spacer(1, 5))
                    story.append(Spacer(1, 20))
                    translations = translate_batch(english_texts)
                    for tr in translations:
                        story.append(Paragraph(fix_ar(tr), s_ar))
                        story.append(Spacer(1, 10))
                
                if translate_images:
                    for shape in slide.shapes:
                        if shape.shape_type == 13:
                            try:
                                image_bytes = shape.image.blob
                                translated_text = translate_image_with_gemini(image_bytes)
                                if translated_text:
                                    story.append(Paragraph(f"[ترجمة الصورة]: {fix_ar(translated_text)}", s_ar))
                                    story.append(Spacer(1, 10))
                            except: pass
                story.append(Spacer(1, 30))
                JOBS[job_id]['progress'] = int(((i+1)/total)*100)
                JOBS[job_id]['stage'] = f'ترجمة الشريحة {i+1} من {total}'
            
            templ.build(story)
            out_bio.seek(0)
            JOBS[job_id]['file'] = out_bio
            JOBS[job_id]['status'] = 'done'
            JOBS[job_id]['progress'] = 100
            JOBS[job_id]['type'] = 'pdf'
        else:
            # === العودة إلى الملاحظات (Notes) ===
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
                
                # 2. ترجمة الصور وإضافتها إلى ملاحظات الشريحة
                if translate_images:
                    notes_text = []
                    for shape in slide.shapes:
                        if shape.shape_type == 13:
                            try:
                                image_bytes = shape.image.blob
                                translated_text = translate_image_with_gemini(image_bytes)
                                if translated_text:
                                    notes_text.append(translated_text)
                            except: pass
                    
                    if notes_text:
                        try:
                            notes_slide = slide.notes_slide
                            tf = notes_slide.notes_text_frame
                            tf.text = "ترجمة الصور:" + chr(10)
                            for tr in notes_text:
                                p = tf.add_paragraph()
                                p.text = tr
                                p.font.size = Pt(20)
                                p.font.name = 'Arial'
                                p.alignment = PP_ALIGN.RIGHT
                        except Exception as e:
                            print(f"Error adding notes: {e}")
                
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
                    except: pass
            JOBS[job_id]['progress'] = int(((i+1)/total)*100)
            JOBS[job_id]['stage'] = f'ترجمة الصفحة {i+1} من {total}'
        
        templ.build(story)
        out_bio.seek(0)
        JOBS[job_id]['file'] = out_bio; JOBS[job_id]['status'] = 'done'; JOBS[job_id]['progress'] = 100
        JOBS[job_id]['type'] = 'pdf'
    except Exception as e:
        JOBS[job_id]['status'] = f'error: {str(e)[:100]}'

# ===== SEO =====
SITEMAP_XML = f"""<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
  <url><loc>{SITE_URL}/</loc><changefreq>daily</changefreq><priority>1.0</priority></url>
  <url><loc>{SITE_URL}/about</loc><changefreq>monthly</changefreq><priority>0.8</priority></url>
  <url><loc>{SITE_URL}/contact</loc><changefreq>monthly</changefreq><priority>0.8</priority></url>
  <url><loc>{SITE_URL}/privacy</loc><changefreq>monthly</changefreq><priority>0.5</priority></url>
  <url><loc>{SITE_URL}/terms</loc><changefreq>monthly</changefreq><priority>0.5</priority></url>
  <url><loc>{SITE_URL}/faq</loc><changefreq>monthly</changefreq><priority>0.7</priority></url>
  <url><loc>{SITE_URL}/feedback</loc><changefreq>daily</changefreq><priority>0.8</priority></url>
  <url><loc>{SITE_URL}/community</loc><changefreq>daily</changefreq><priority>0.9</priority></url>
</urlset>"""

ROBOTS_TXT = f"""User-agent: *
Allow: /
Sitemap: {SITE_URL}/sitemap.xml
"""

# ===== الواجهة الرئيسية =====
HTML_PAGE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>الريان للترجمة | ترجمة ملفات PDF و PPTX</title>
<meta name="description" content="منصة الريان للترجمة: ترجمة ملفات PDF و PPTX والنصوص من الإنجليزية إلى العربية مجاناً. ترجمة دقيقة وسريعة مع الحفاظ على التنسيق.">
<meta name="keywords" content="ترجمة, ترجمة ملفات, ترجمة PDF, ترجمة PPTX, الريان للترجمة">
<link rel="icon" href="/static/favicon.jpg" type="image/jpeg">
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;500;700;800&display=swap" rel="stylesheet">
<style>
*{margin:0;padding:0;box-sizing:border-box;font-family:'Tajawal',sans-serif}
body{min-height:100vh;background:linear-gradient(135deg, #eef2ff 0%, #c7b8ff 100%);color:#2f365f;padding:20px;display:flex;flex-direction:column;justify-content:center;align-items:center}
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
.footer{max-width:440px;width:100%;padding:20px;text-align:center;font-size:14px;color:#6b6b9a}
.footer a{color:#5b5bd6;text-decoration:none;margin:0 5px;}
</style>
</head>
<body>
<div class="app">
  <div class="card">
    <h2 style="color:#5b5bd6;margin-bottom:10px">الريان للترجمة</h2>
    <p style="color:#6b6b9a;margin-bottom:15px">رفع ملف PPTX أو PDF لترجمة النصوص</p>
    <form id="uploadForm" enctype="multipart/form-data">
      <input type="file" name="file" accept=".pptx,.pdf" required style="width:100%;margin-bottom:10px">
      <div class="checkbox">
        <input type="checkbox" id="translateImages">
        <label for="translateImages">ترجمة الصور أيضاً (أبطأ)</label>
      </div>
      <div class="checkbox">
        <input type="checkbox" id="toPdf">
        <label for="toPdf">تحويل إلى PDF (ترتيب: نص أصلي ثم ترجمة)</label>
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
  
  <div class="card">
    <h3 style="color:#5b5bd6;margin-bottom:10px;">⚡ ترجمة سريعة</h3>
    <div class="quick-translate">
      <textarea id="quickText" placeholder="ألصق النص الإنجليزي هنا..."></textarea>
      <button onclick="quickTranslate()">ترجم</button>
    </div>
    <p id="quickResult" style="margin-top:10px;text-align:right;font-size:14px;color:#2f365f;display:none;padding:10px;background:#f8f9fa;border-radius:10px;"></p>
  </div>

  <div class="card" style="text-align:right;">
    <h3 style="color:#5b5bd6;margin-bottom:10px;">🧠 مساعد أيمن الذكي</h3>
    <p style="font-size:14px;color:#666;margin-bottom:10px;">اسأل عن أي شيء، أو ارفع ملفاً لتحليله.</p>
    <div style="display:flex;gap:5px;">
      <input type="text" id="quickChat" maxlength="1000" placeholder="اكتب سؤالك..." style="flex:1;padding:10px;border:1px solid #ddd;border-radius:10px;">
      <button onclick="goChat()" style="background:#8b5cf6;color:white;border:none;padding:10px 20px;border-radius:10px;cursor:pointer;">إرسال</button>
    </div>
  </div>

  <div class="card" style="display:flex;gap:10px;">
    <a href="/feedback" style="flex:1;text-decoration:none;">
      <button class="btn" style="background:linear-gradient(135deg,#f59e0b,#d97706);margin:0;">💬 تعليقات واقتراحات</button>
    </a>
    <a href="/community" style="flex:1;text-decoration:none;">
      <button class="btn" style="background:linear-gradient(135deg,#10b981,#059669);margin:0;">🌐 مجتمع الريان</button>
    </a>
  </div>
</div>
<div class="footer">
    <a href="/about">من نحن</a> | 
    <a href="/contact">اتصل بنا</a> | 
    <a href="/privacy">سياسة الخصوصية</a> | 
    <a href="/terms">شروط الاستخدام</a> | 
    <a href="/faq">الأسئلة الشائعة</a>
    <p style="margin-top:10px;">تابعنا على <a href="{{WHATSAPP_CHANNEL}}" target="_blank" style="color:#25D366;font-weight:bold;">قناة الواتساب</a></p>
    <p style="margin-top:5px;">© 2026 الريان للترجمة. جميع الحقوق محفوظة.</p>
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
  const toPdf = document.getElementById('toPdf').checked;
  formData.append('translate_images', translateImages ? 'true' : 'false');
  formData.append('to_pdf', toPdf ? 'true' : 'false');
  
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

# ===== الصفحات =====
ABOUT_PAGE = """
<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>من نحن | الريان للترجمة</title><meta name="description" content="تعرف على منصة الريان للترجمة، أهدافها، وكيف تساعد الطلاب والباحثين في ترجمة الملفات."><link rel="icon" href="/static/favicon.jpg" type="image/jpeg"><link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet"><style>body{font-family:'Tajawal',sans-serif;background:#f0f2f5;color:#2f365f;padding:40px;text-align:center;} .container{max-width:600px;margin:auto;background:#fff;padding:30px;border-radius:20px;box-shadow:0 8px 20px rgba(0,0,0,0.1);} h1{color:#5b5bd6;} p{line-height:1.8;font-size:16px;text-align:right;}</style></head><body><div class="container"><h1>من نحن</h1><p>الريان للترجمة هي منصة رقمية مجانية تهدف إلى تسهيل ترجمة الملفات الأكاديمية (PDF و PPTX) للطلاب والباحثين في العالم العربي. انطلقت المنصة من السودان، وتستخدم أحدث تقنيات الذكاء الاصطناعي (مثل Gemini و DeepL) لتقديم ترجمات دقيقة وسريعة، مع الحفاظ على تنسيق الملفات الأصلية. نؤمن بأن المعرفة لا تحدها حدود، ونسعى لجعل المحتوى العلمي متاحاً للجميع بأسهل طريقة.</p><br><a href="/" style="color:#5b5bd6;text-decoration:none;font-weight:bold;">← العودة للرئيسية</a></div></body></html>
"""

CONTACT_PAGE = """
<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>اتصل بنا | الريان للترجمة</title><meta name="description" content="تواصل مع فريق الريان للترجمة عبر البريد الإلكتروني أو الواتساب لأي استفسار أو اقتراح."><link rel="icon" href="/static/favicon.jpg" type="image/jpeg"><link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet"><style>body{font-family:'Tajawal',sans-serif;background:#f0f2f5;color:#2f365f;padding:40px;text-align:center;} .container{max-width:600px;margin:auto;background:#fff;padding:30px;border-radius:20px;box-shadow:0 8px 20px rgba(0,0,0,0.1);} h1{color:#5b5bd6;} p{line-height:1.8;font-size:16px;text-align:right;}</style></head><body><div class="container"><h1>اتصل بنا</h1><p>هل لديك سؤال، اقتراح، أو واجهت مشكلة أثناء استخدام المنصة؟ يسعدنا تواصلك معنا.</p><p style="text-align:center;font-size:18px;margin:20px 0;"><strong>البريد الإلكتروني:</strong> <a href="mailto:Abureyan810@gmail.com" style="color:#5b5bd6;">Abureyan810@gmail.com</a></p><p style="text-align:center;font-size:18px;margin:20px 0;"><strong>واتساب:</strong> <a href="https://wa.me/249925284869" style="color:#5b5bd6;" target="_blank">اضغط هنا</a></p><p style="text-align:center;font-size:18px;"><strong>اتصال هاتفي:</strong> <a href="tel:+249925284869" style="color:#5b5bd6;">+249925284869</a></p><br><a href="/" style="color:#5b5bd6;text-decoration:none;font-weight:bold;">← العودة للرئيسية</a></div></body></html>
"""

PRIVACY_PAGE = """
<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>سياسة الخصوصية | الريان للترجمة</title><meta name="description" content="سياسة الخصوصية لمنصة الريان للترجمة: كيف نتعامل مع ملفاتك وبياناتك."><link rel="icon" href="/static/favicon.jpg" type="image/jpeg"><link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet"><style>body{font-family:'Tajawal',sans-serif;background:#f0f2f5;color:#2f365f;padding:40px;text-align:center;} .container{max-width:600px;margin:auto;background:#fff;padding:30px;border-radius:20px;box-shadow:0 8px 20px rgba(0,0,0,0.1);} h1{color:#5b5bd6;} p,li{line-height:1.8;font-size:16px;text-align:right;}</style></head><body><div class="container"><h1>سياسة الخصوصية</h1><p>نحن في الريان للترجمة نحترم خصوصيتك. نوضح هنا كيفية تعاملنا مع بياناتك:</p><ul style="text-align:right;line-height:1.8;"><li><strong>الملفات المرفوعة:</strong> يتم معالجة الملفات مؤقتاً في الذاكرة، ولا نقوم بتخزينها أو مشاركتها مع أي طرف ثالث.</li><li><strong>ملفات تعريف الارتباط (Cookies):</strong> قد نستخدم ملفات تعريف الارتباط لتحسين تجربتك، ولعرض إعلانات مناسبة (إذا تم تفعيل الإعلانات لاحقاً).</li><li><strong>خدمات الطرف الثالث:</strong> نستخدم واجهات برمجة تطبيقات (APIs) من Google و DeepL للترجمة.</li><li><strong>لا نبيع بياناتك:</strong> نحن لا نبيع أو نشارك بياناتك الشخصية مع أي جهة خارجية.</li></ul><br><a href="/" style="color:#5b5bd6;text-decoration:none;font-weight:bold;">← العودة للرئيسية</a></div></body></html>
"""

TERMS_PAGE = """
<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>شروط الاستخدام | الريان للترجمة</title><meta name="description" content="شروط استخدام منصة الريان للترجمة."><link rel="icon" href="/static/favicon.jpg" type="image/jpeg"><link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet"><style>body{font-family:'Tajawal',sans-serif;background:#f0f2f5;color:#2f365f;padding:40px;text-align:center;} .container{max-width:600px;margin:auto;background:#fff;padding:30px;border-radius:20px;box-shadow:0 8px 20px rgba(0,0,0,0.1);} h1{color:#5b5bd6;} p,li{line-height:1.8;font-size:16px;text-align:right;}</style></head><body><div class="container"><h1>شروط الاستخدام</h1><p>باستخدامك للمنصة، فإنك توافق على:</p><ul style="text-align:right;line-height:1.8;"><li>عدم استخدام المنصة لترجمة محتوى غير قانوني أو مسيء.</li><li>أنت المسؤول الوحيد عن المحتوى الذي ترفعه.</li><li>نحن غير مسؤولين عن أي أخطاء في الترجمة قد تنتج عن استخدام التقنيات الآلية.</li></ul><br><a href="/" style="color:#5b5bd6;text-decoration:none;font-weight:bold;">← العودة للرئيسية</a></div></body></html>
"""

FAQ_PAGE = """
<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>الأسئلة الشائعة | الريان للترجمة</title><meta name="description" content="الأسئلة الشائعة حول منصة الريان للترجمة: هل الخدمة مجانية؟ هل تحفظون الملفات؟"><link rel="icon" href="/static/favicon.jpg" type="image/jpeg"><link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet"><style>body{font-family:'Tajawal',sans-serif;background:#f0f2f5;color:#2f365f;padding:40px;text-align:center;} .container{max-width:600px;margin:auto;background:#fff;padding:30px;border-radius:20px;box-shadow:0 8px 20px rgba(0,0,0,0.1);} h1{color:#5b5bd6;} p{line-height:1.8;font-size:16px;text-align:right;} strong{color:#5b5bd6;}</style></head><body><div class="container"><h1>الأسئلة الشائعة</h1><p><strong>هل الخدمة مجانية؟</strong><br>نعم، الخدمة مجانية بالكامل.</p><p><strong>هل تحفظون ملفاتي؟</strong><br>لا، يتم حذف الملفات فور إتمام الترجمة.</p><p><strong>لماذا تظهر المربعات في ملفات PDF؟</strong><br>لأننا نستخدم خطاً عربياً مضمناً، وقد يحتاج المتصفح للتحديث. جرب فتح الملف في برنامج آخر.</p><p><strong>هل يمكنني ترجمة ملفات كبيرة؟</strong><br>نعم، بحد أقصى 100 صفحة/شريحة لكل ملف.</p><br><a href="/" style="color:#5b5bd6;text-decoration:none;font-weight:bold;">← العودة للرئيسية</a></div></body></html>
"""

# ===== صفحة التعليقات =====
FEEDBACK_PAGE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>التعليقات والاقتراحات | الريان للترجمة</title>
<meta name="description" content="شاركنا رأيك واقتراحاتك حول منصة الريان للترجمة.">
<link rel="icon" href="/static/favicon.jpg" type="image/jpeg">
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
<style>
body{font-family:'Tajawal',sans-serif;background:#f0f2f5;color:#2f365f;padding:20px;text-align:center;}
.container{max-width:600px;margin:auto;background:#fff;padding:30px;border-radius:20px;box-shadow:0 8px 20px rgba(0,0,0,0.1);}
h1{color:#5b5bd6;margin-bottom:20px;}
input,textarea{width:100%;padding:10px;margin-bottom:10px;border:1px solid #ddd;border-radius:10px;font-family:inherit;box-sizing:border-box;}
textarea{height:100px;resize:none;}
button{background:linear-gradient(135deg,#f59e0b,#d97706);color:#fff;border:none;padding:12px 20px;border-radius:10px;cursor:pointer;font-weight:bold;width:100%;}
.comment{background:#f8f9fa;border-radius:10px;padding:15px;margin-bottom:10px;text-align:right;}
.comment strong{color:#5b5bd6;}
.comment small{color:#888;font-size:12px;}
</style>
</head>
<body>
<div class="container">
<h1>💬 التعليقات والاقتراحات</h1>
<form action="/submit_feedback" method="POST">
<input type="text" name="name" placeholder="اسمك (اختياري)" maxlength="50">
<textarea name="message" placeholder="اكتب تعليقك أو اقتراحك هنا..." required maxlength="500"></textarea>
<button type="submit">إرسال</button>
</form>
<hr style="margin:20px 0;border:1px solid #eee;">
<h2 style="color:#5b5bd6;margin-bottom:15px;">آخر التعليقات</h2>
<div id="comments_list">
{{COMMENTS|safe}}
</div>
<br><a href="/" style="color:#5b5bd6;text-decoration:none;font-weight:bold;">← العودة للرئيسية</a>
</div>
</body>
</html>
"""

# ===== صفحة المجتمع =====
COMMUNITY_PAGE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>مجتمع الريان | الريان للترجمة</title>
<meta name="description" content="مجتمع الريان الأكاديمي: تواصل مع طلاب من جامعتك، واطلب المساعدة.">
<link rel="icon" href="/static/favicon.jpg" type="image/jpeg">
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
<style>
body{font-family:'Tajawal',sans-serif;background:#f0f2f5;color:#2f365f;padding:20px;text-align:center;}
.container{max-width:600px;margin:auto;background:#fff;padding:30px;border-radius:20px;box-shadow:0 8px 20px rgba(0,0,0,0.1);}
h1{color:#5b5bd6;margin-bottom:20px;}
input,textarea{width:100%;padding:10px;margin-bottom:10px;border:1px solid #ddd;border-radius:10px;font-family:inherit;box-sizing:border-box;}
textarea{height:80px;resize:none;}
button{background:linear-gradient(135deg,#10b981,#059669);color:#fff;border:none;padding:12px 20px;border-radius:10px;cursor:pointer;font-weight:bold;width:100%;}
.post{background:#f8f9fa;border-radius:10px;padding:15px;margin-bottom:10px;text-align:right;border-left:4px solid #10b981;}
.post strong{color:#5b5bd6;}
.post small{color:#888;font-size:12px;}
.post .whatsapp-btn{display:inline-block;background:#25D366;color:white;padding:8px 15px;border-radius:8px;text-decoration:none;margin-top:8px;font-size:14px;}
</style>
</head>
<body>
<div class="container">
<h1>🌐 مجتمع الريان</h1>
<p style="margin-bottom:20px;">تواصل مع طلاب من جامعتك، واطلب المساعدة.</p>
<form action="/submit_post" method="POST">
<input type="text" name="name" placeholder="اسمك" required maxlength="50">
<input type="text" name="university" placeholder="جامعتك (مثال: جامعة أم درمان الإسلامية)" required maxlength="100">
<textarea name="message" placeholder="اكتب طلبك أو سؤالك هنا..." required maxlength="500"></textarea>
<input type="text" name="whatsapp" placeholder="رقم واتساب (اختياري - للتواصل)" maxlength="20">
<button type="submit">نشر</button>
</form>
<hr style="margin:20px 0;border:1px solid #eee;">
<h2 style="color:#5b5bd6;margin-bottom:15px;">آخر المنشورات</h2>
<div id="posts_list">
{{POSTS|safe}}
</div>
<br><a href="/" style="color:#5b5bd6;text-decoration:none;font-weight:bold;">← العودة للرئيسية</a>
</div>
</body>
</html>
"""

# ===== المسارات =====
@app.route('/')
def home():
    increment_stat('visitors')
    return render_template_string(HTML_PAGE, WHATSAPP_CHANNEL=WHATSAPP_CHANNEL)

@app.route('/sitemap.xml')
def sitemap():
    return app.response_class(SITEMAP_XML, mimetype='application/xml')

@app.route('/robots.txt')
def robots():
    return app.response_class(ROBOTS_TXT, mimetype='text/plain')

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
    to_pdf = request.form.get('to_pdf') == 'true'
    increment_stat('translated_files')
    job_id = str(uuid.uuid4())[:8]
    is_pptx = '.pptx' in filename_lower
    JOBS[job_id] = {'progress': 0, 'status': 'working', 'file': None, 'type': 'pptx' if is_pptx else 'pdf', 'original_filename': file.filename, 'stage': 'جاري التحضير...'}
    data = file.read()
    if is_pptx: threading.Thread(target=process_pptx, args=(job_id, data, translate_images, to_pdf), daemon=True).start()
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
            # إزالة كلمة models/ إذا كانت موجودة
            clean_name = model_name.replace('models/', '')
            model = genai.GenerativeModel(clean_name)
            chat_session = model.start_chat(history=formatted_history)
            prompt = "أنت مساعد ذكي مفيد جداً اسمه أيمن. أجب باللغة العربية بوضوح وإيجاز. " + message
            resp = chat_session.send_message(prompt, request_options={'timeout': 90})
            return jsonify({'reply': resp.text})
        except Exception as e:
            last_error = str(e)
            print(f'فشل النموذج {model_name}: {e}')
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
                resp = model.generate_content(prompt, request_options={'timeout': 90})
                return jsonify({'reply': resp.text})
            except Exception as e:
                last_error = str(e)
                continue
        return jsonify({'error': f'فشلت جميع النماذج: {last_error[:100]}'})
    except Exception as e:
        return jsonify({'error': f'خطأ تقني: {str(e)[:100]}'})

@app.route('/feedback')
def feedback():
    comments = ""
    if supabase:
        try:
            res = supabase.table('feedback').select('*').order('created_at', desc=True).limit(20).execute()
            if res.data:
                for c in res.data:
                    comments += f"<div class='comment'><strong>{c['name']}</strong><br>{c['message']}<br><small>{c['created_at'][:16]}</small></div>"
            else:
                comments = "<p style='color:#888;'>لا توجد تعليقات بعد. كن أول من يشارك رأيه!</p>"
        except: pass
    if not comments: comments = "<p style='color:#888;'>لا توجد تعليقات بعد. كن أول من يشارك رأيه!</p>"
    return render_template_string(FEEDBACK_PAGE, COMMENTS=comments)

@app.route('/submit_feedback', methods=['POST'])
def submit_feedback():
    print('=== بدء إرسال التعليق ===')
    name = request.form.get('name', 'زائر') or 'زائر'
    message = request.form.get('message', '')
    print(f'الاسم: {name}')
    print(f'الرسالة: {message}')
    
    if not supabase:
        print('خطأ: Supabase غير متصل')
        return 'Supabase غير متصل', 500
        
    if not message:
        print('خطأ: الرسالة فارغة')
        return 'الرسالة فارغة', 400
        
    try:
        print('محاولة الإدراج في Supabase...')
        res = supabase.table('feedback').insert({'name': name[:50], 'message': message[:500]}).execute()
        print(f'نتيجة الإدراج: {res}')
    except Exception as e:
        print(f'خطأ في الإدراج: {e}')
        return f'خطأ: {e}', 500
        
    print('=== تم الإرسال بنجاح ===')
    return redirect('/feedback')

@app.route('/community')
def community():
    posts = ""
    if supabase:
        try:
            res = supabase.table('community_posts').select('*').order('created_at', desc=True).limit(20).execute()
            if res.data:
                for p in res.data:
                    whatsapp_btn = ""
                    if p.get('whatsapp'):
                        whatsapp_btn = f"<a href='https://wa.me/{p['whatsapp']}' class='whatsapp-btn' target='_blank'>📱 تواصل عبر واتساب</a>"
                    posts += f"<div class='post'><strong>{p['name']}</strong> <small>({p['university']})</small><br>{p['message']}<br>{whatsapp_btn}<br><small>{p['created_at'][:16]}</small></div>"
            else:
                posts = "<p style='color:#888;'>لا توجد منشورات بعد. كن أول من ينشر!</p>"
        except: pass
    if not posts: posts = "<p style='color:#888;'>لا توجد منشورات بعد. كن أول من ينشر!</p>"
    return render_template_string(COMMUNITY_PAGE, POSTS=posts)

@app.route('/submit_post', methods=['POST'])
def submit_post():
    print('=== بدء إرسال المنشور ===')
    name = request.form.get('name', '')
    university = request.form.get('university', '')
    message = request.form.get('message', '')
    whatsapp = request.form.get('whatsapp', '')
    print(f'الاسم: {name}')
    print(f'الجامعة: {university}')
    print(f'الرسالة: {message}')
    
    if not supabase:
        print('خطأ: Supabase غير متصل')
        return 'Supabase غير متصل', 500
        
    if not name or not university or not message:
        print('خطأ: بيانات ناقصة')
        return 'بيانات ناقصة', 400
        
    try:
        print('محاولة الإدراج في Supabase...')
        res = supabase.table('community_posts').insert({
            'name': name[:50], 'university': university[:100], 'message': message[:500], 'whatsapp': whatsapp[:20]
        }).execute()
        print(f'نتيجة الإدراج: {res}')
    except Exception as e:
        print(f'خطأ في الإدراج: {e}')
        return f'خطأ: {e}', 500
        
    print('=== تم الإرسال بنجاح ===')
    return redirect('/community')

@app.route('/about')
def about():
    return render_template_string(ABOUT_PAGE)

@app.route('/contact')
def contact():
    return render_template_string(CONTACT_PAGE)

@app.route('/privacy')
def privacy():
    return render_template_string(PRIVACY_PAGE)

@app.route('/terms')
def terms():
    return render_template_string(TERMS_PAGE)

@app.route('/faq')
def faq():
    return render_template_string(FAQ_PAGE)

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
