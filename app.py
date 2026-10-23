import os, io, uuid, threading, tempfile
from flask import Flask, request, jsonify, send_file, render_template_string
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
import fitz, arabic_reshaper
from bidi.algorithm import get_display
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.pagesizes import letter
from PIL import Image
import google.generativeai as genai

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 120 * 1024 * 1024

GEMINI_KEY = os.environ.get('GEMINI_API_KEY')
if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)
    gemini_model = genai.GenerativeModel('gemini-1.5-flash')

JOBS = {}

def fix_ar(t):
    try: return get_display(arabic_reshaper.reshape(t))
    except: return t

def translate_text_only(text):
    if not text.strip() or len(text.strip())<2: return text
    try:
        prompt = f"Translate this to formal Arabic only, auto-detect source language, no explanation:\n\n{text}"
        return gemini_model.generate_content(prompt).text.strip()
    except: return text

def ocr_and_translate_image(image_bytes):
    if not GEMINI_KEY: return ""
    try:
        img = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        prompt = "Extract any text in this image, translate it to formal Arabic. Return in format:\nORIGINAL: [text]\nARABIC: [translation]\nIf no text, return empty."
        resp = gemini_model.generate_content([prompt, img])
        txt = resp.text.strip()
        return txt if txt and "empty" not in txt.lower() else ""
    except Exception as e:
        print("OCR error:", e)
        return ""

def process_job(job_id, tmp_path, filename):
    try:
        JOBS[job_id]['status']='working'
        if filename.lower().endswith('.pptx'):
            prs = Presentation(tmp_path)
            total_shapes = sum(len(list(s.shapes)) for s in prs.slides)
            done=0
            for slide in prs.slides:
                for shape in list(slide.shapes):
                    # نص
                    if shape.has_text_frame and shape.text.strip():
                        orig = shape.text.strip()
                        if 1 < len(orig) < 1000:
                            tr = translate_text_only(orig)
                            if tr != orig:
                                try:
                                    box = slide.shapes.add_textbox(shape.left, shape.top + shape.height + Pt(5), shape.width, Inches(0.8))
                                    box.text_frame.word_wrap=True
                                    p=box.text_frame.paragraphs[0]
                                    p.text=tr; p.font.size=Pt(16); p.alignment=PP_ALIGN.RIGHT
                                except: pass
                    # صورة - الميزة الجديدة
                    if shape.shape_type == 13 or hasattr(shape, "image"):
                        try:
                            image_bytes = shape.image.blob
                            ocr_result = ocr_and_translate_image(image_bytes)
                            if ocr_result:
                                box = slide.shapes.add_textbox(shape.left, shape.top + shape.height + Inches(0.1), shape.width, Inches(1.2))
                                box.text_frame.word_wrap=True
                                p=box.text_frame.paragraphs[0]
                                p.text=ocr_result; p.font.size=Pt(12); p.alignment=PP_ALIGN.RIGHT
                        except: pass
                    done+=1
                    if total_shapes>0:
                        JOBS[job_id]['progress']=int(done/total_shapes*90)
            out_path = tmp_path + "_out.pptx"
            prs.save(out_path)
            JOBS[job_id]['file_path']=out_path
            JOBS[job_id]['filename']='translated_with_images.pptx'
        else:
            doc=fitz.open(tmp_path)
            out_bio=io.BytesIO()
            templ=SimpleDocTemplate(out_bio, pagesize=letter)
            s1=ParagraphStyle('o', fontSize=11); s2=ParagraphStyle('a', fontSize=12, alignment=2)
            story=[]
            for i, page in enumerate(doc):
                txt=page.get_text("text")
                for line in txt.split('\n'):
                    if line.strip():
                        story.append(Paragraph(line[:500], s1))
                        story.append(Paragraph(fix_ar(translate_text_only(line[:500])), s2))
                JOBS[job_id]['progress']=int((i+1)/len(doc)*90)
            templ.build(story); out_bio.seek(0)
            out_path = tmp_path + "_out.pdf"
            open(out_path,'wb').write(out_bio.read())
            JOBS[job_id]['file_path']=out_path
            JOBS[job_id]['filename']='translated.pdf'
        JOBS[job_id]['progress']=100; JOBS[job_id]['status']='done'
        try: os.remove(tmp_path)
        except: pass
    except Exception as e:
        JOBS[job_id]['status']=f'error: {e}'; print(e)

HTML = """
<!DOCTYPE html><html dir=rtl lang=ar><head><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1">
<title>الريان - Gemini Vision</title>
<style>body{font-family:sans-serif;background:#eef2ff;padding:20px}
.card{max-width:600px;margin:auto;background:#fff;padding:20px;border-radius:16px;box-shadow:0 8px 20px #0001}
button{width:100%;height:50px;background:#6c5cff;color:#fff;border:none;border-radius:12px;font-size:18px}
#bar{height:20px;background:#eee;border-radius:10px;overflow:hidden;margin-top:15px;display:none}
#fill{height:100%;width:0%;background:#6c5cff;transition:0.3s}
</style></head><body>
<div class=card>
<h2>الريان - يترجم النصوص والصور</h2>
<input type=file id=file accept=.pdf,.pptx><br><br>
<button onclick=upload()>رفع وترجمة</button>
<div id=bar><div id=fill></div></div><p id=status></p>
<div id=actions style=display:none>
<button onclick=download()>تحميل الملف</button>
<div id=dua style=margin-top:20px;background:#fff7e6;padding:15px;border-radius:10px"></div>
</div></div>
<script>
let jobId=null;
async function upload(){
 let f=document.getElementById('file').files[0]; if(!f) return alert('اختار ملف');
 let fd=new FormData(); fd.append('file',f);
 let r=await fetch('/upload',{method:'POST',body:fd}); let j=await r.json(); jobId=j.job_id;
 document.getElementById('bar').style.display='block'; poll();
}
async function poll(){
 let r=await fetch('/progress/'+jobId); let j=await r.json();
 document.getElementById('fill').style.width=j.progress+'%';
 document.getElementById('status').innerText=j.progress+'% - '+j.status;
 if(j.status=='done'){ document.getElementById('actions').style.display='block';
  document.getElementById('dua').innerHTML="اللهم وفق عبدك <b>أيمن</b> لما تحب وترضى، اللهم سهل له كل عسير واجعله من الناجحين<br><br>ولا تنس أن <b>تصلي على النبي ﷺ</b>"; 
 } else if(j.status.includes('error')){ document.getElementById('status').innerText=j.status }
 else setTimeout(poll,1000);
}
function download(){ window.location='/download/'+jobId; }
</script></body></html>
"""

@app.route('/')
def home(): return render_template_string(HTML)
@app.route('/upload', methods=['POST'])
def up():
    f=request.files.get('file')
    if not f: return jsonify({"error":"no file"}),400
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f.filename)[1])
    f.save(tmp.name); tmp.close()
    job_id=str(uuid.uuid4())[:8]
    JOBS[job_id]={'progress':0,'status':'queued','file_path':None}
    threading.Thread(target=process_job, args=(job_id, tmp.name, f.filename), daemon=True).start()
    return jsonify({"job_id":job_id})
@app.route('/progress/<jid>')
def prog(jid): return jsonify(JOBS.get(jid,{"progress":0,"status":"not found"}))
@app.route('/download/<jid>')
def dl(jid):
    j=JOBS.get(jid)
    if not j or j['status']!='done': return "not ready",404
    return send_file(j['file_path'], as_attachment=True, download_name=j['filename'])
@app.route('/ping')
def ping(): return jsonify({"ok":True})

if __name__=='__main__': app.run(host='0.0.0.0', port=int(os.environ.get('PORT',10000)))
