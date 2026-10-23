import os, io, uuid, threading, time, tempfile
from flask import Flask, request, jsonify, send_file, render_template_string
from collections import OrderedDict
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
import fitz, requests, arabic_reshaper
from bidi.algorithm import get_display
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import letter
import google.generativeai as genai

app = Flask(__name__)
# === حل مشكلة 100MB ===
app.config['MAX_CONTENT_LENGTH'] = 120 * 1024 * 1024  # 120MB مسموح

DEEPL_KEY = os.environ.get('DEEPL_API_KEY')
GEMINI_KEY = os.environ.get('GEMINI_API_KEY')
if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)
    gemini_model = genai.GenerativeModel('gemini-1.5-flash')

JOBS = {}

def fix_ar(t):
    try: return get_display(arabic_reshaper.reshape(t))
    except: return t

def translate_smart(text):
    if not text.strip(): return text
    try:
        prompt = f"Translate to formal Arabic only, auto-detect source language:\n\n{text}"
        return genai.GenerativeModel('gemini-1.5-flash').generate_content(prompt).text.strip()
    except: return text

def process_job(job_id, tmp_path, filename):
    try:
        JOBS[job_id]['status']='working'
        ext = filename.lower()
        if ext.endswith('.pptx'):
            # نقرا من الملف في الهارد مش الرام
            prs = Presentation(tmp_path)
            total = sum(1 for s in prs.slides for sh in s.shapes if sh.has_text_frame and sh.text.strip())
            done=0
            for slide in prs.slides:
                for shape in list(slide.shapes):
                    if shape.has_text_frame and shape.text.strip():
                        orig=shape.text.strip()
                        if len(orig)>1 and len(orig)<500: # ما نترجم نص طويل شديد مرة واحدة
                            tr=translate_smart(orig)
                            if tr!=orig:
                                try:
                                    box=slide.shapes.add_textbox(shape.left, shape.top+shape.height+Pt(5), shape.width, Inches(0.6))
                                    box.text_frame.word_wrap=True
                                    p=box.text_frame.paragraphs[0]
                                    p.text=tr; p.font.size=Pt(14); p.alignment=PP_ALIGN.RIGHT
                                except: pass
                            done+=1
                            JOBS[job_id]['progress']=int(done/max(total,1)*100)
            out_path = tmp_path + "_out.pptx"
            prs.save(out_path)
            JOBS[job_id]['file_path']=out_path
            JOBS[job_id]['filename']='translated.pptx'
        else:
            doc=fitz.open(tmp_path)
            out_bio=io.BytesIO()
            templ=SimpleDocTemplate(out_bio, pagesize=letter)
            s1=ParagraphStyle('o', fontSize=11); s2=ParagraphStyle('a', fontSize=12, alignment=2)
            story=[]; pages=len(doc)
            for i, page in enumerate(doc):
                txt=page.get_text("text")
                for line in txt.split('\n'):
                    if line.strip():
                        story.append(Paragraph(line[:500], s1))
                        story.append(Paragraph(fix_ar(translate_smart(line[:500])), s2))
                JOBS[job_id]['progress']=int((i+1)/pages*100)
            templ.build(story); out_bio.seek(0)
            out_path = tmp_path + "_out.pdf"
            open(out_path,'wb').write(out_bio.read())
            JOBS[job_id]['file_path']=out_path
            JOBS[job_id]['filename']='translated.pdf'
        JOBS[job_id]['progress']=100; JOBS[job_id]['status']='done'
        # امسح الملف الأصلي عشان نوفر مساحة
        try: os.remove(tmp_path)
        except: pass
    except Exception as e:
        JOBS[job_id]['status']=f'error: {e}'; JOBS[job_id]['progress']=0

HTML_PAGE = """
<!DOCTYPE html><html dir=rtl lang=ar><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>الريان - 120MB</title>
<style>body{font-family:sans-serif;background:#eef2ff;padding:20px}
.card{max-width:600px;margin:auto;background:#fff;padding:20px;border-radius:16px}
button{width:100%;height:50px;background:#6c5cff;color:#fff;border:none;border-radius:12px;font-size:18px}
#bar{height:20px;background:#eee;border-radius:10px;overflow:hidden;margin-top:15px;display:none}
#fill{height:100%;width:0%;background:#6c5cff;transition:0.3s}</style></head><body>
<div class=card>
<h2>الريان - يدعم حتى 100MB مع الصوت</h2>
<input type=file id=file accept=.pdf,.pptx><br><br>
<button onclick=upload()>رفع وترجمة</button>
<div id=bar><div id=fill></div></div><p id=status></p>
<div id=actions style=display:none>
<button id=dl onclick=download()>تحميل الملف المترجم</button>
<div id=dua style=margin-top:20px;background:#fff7e6;padding:15px;border-radius:10px;display:none"></div>
</div></div>
<script>
let jobId=null;
async function upload(){
 let f=document.getElementById('file').files[0]; if(!f) return alert('اختار ملف');
 if(f.size>120*1024*1024) return alert('الملف أكبر من 120MB');
 let fd=new FormData(); fd.append('file',f);
 let r=await fetch('/upload',{method:'POST',body:fd}); let j=await r.json(); jobId=j.job_id;
 document.getElementById('bar').style.display='block'; poll();
}
async function poll(){
 let r=await fetch('/progress/'+jobId); let j=await r.json();
 document.getElementById('fill').style.width=j.progress+'%';
 document.getElementById('status').innerText='التقدم: '+j.progress+'% - '+j.status;
 if(j.status=='done'){ document.getElementById('actions').style.display='block';
   document.getElementById('dua').style.display='block';
   document.getElementById('dua').innerHTML="اللهم وفق عبدك <b>أيمن</b> لما تحب وترضى، اللهم سهل له كل عسير واجعله من الناجحين<br><br>ولا تنس أن <b>تصلي على النبي ﷺ</b>"; 
 } else if(j.status.includes('error')){ document.getElementById('status').innerText=j.status }
 else setTimeout(poll,1000);
}
function download(){ window.location='/download/'+jobId; }
</script></body></html>
"""

@app.route('/')
def home(): return render_template_string(HTML_PAGE)

@app.route('/upload', methods=['POST'])
def upload():
    f=request.files.get('file')
    if not f: return jsonify({"error":"no file"}),400
    # حماية 2: نحفظ في /tmp مش في الرام
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f.filename)[1])
    f.save(tmp.name)
    tmp.close()
    job_id=str(uuid.uuid4())[:8]
    JOBS[job_id]={'progress':0,'status':'queued','file_path':None,'filename':'','text':''}
    threading.Thread(target=process_job, args=(job_id, tmp.name, f.filename), daemon=True).start()
    return jsonify({"job_id":job_id})

@app.route('/progress/<job_id>')
def progress(job_id):
    j=JOBS.get(job_id)
    if not j: return jsonify({"error":"not found"}),404
    return jsonify({"progress":j['progress'],"status":j['status']})

@app.route('/download/<job_id>')
def download_file(job_id):
    j=JOBS.get(job_id)
    if not j or j['status']!='done': return "not ready",404
    return send_file(j['file_path'], as_attachment=True, download_name=j['filename'])

@app.route('/ping')
def ping(): return jsonify({"ok":True})

if __name__=='__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT',10000)))
