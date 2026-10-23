import os, uuid, io
from flask import Flask, request, jsonify, send_file, render_template_string
from pptx import Presentation
import requests

app = Flask(__name__)

# ذاكرة مؤقتة للمهام - تعمل على Render المجاني
JOBS = {}

DEEPL_KEY = os.environ.get("DEEPL_API_KEY")
SUPABASE_URL = os.environ.get("SUPABASE_URL") or "https://lmgshabrtmvbnpdhkkst.supabase.co"
SUPABASE_KEY = os.environ.get("SUPABASE_KEY") or os.environ.get("SUPABASE_ANON_KEY")

def translate_text(text):
    if not text.strip(): return text
    # 1. حاول من Supabase (القاموس) لتوفير الحصة
    # 2. ترجم بـ DeepL
    try:
        url = "https://api-free.deepl.com/v2/translate"
        data = {"auth_key": DEEPL_KEY, "text": text, "target_lang": "AR"}
        r = requests.post(url, data=data, timeout=10)
        if r.status_code == 200:
            return r.json()["translations"][0]["text"]
    except:
        pass
    return text

HTML_PAGE = """
<!DOCTYPE html><html dir="rtl"><head><meta charset="utf-8"><title>الريان للترجمة</title>
<style>body{font-family:Tajawal;padding:20px;background:#f5f7ff} .card{background:#fff;padding:20px;border-radius:16px;box-shadow:0 8px 24px rgba(0,0,0,0.08);max-width:600px;margin:auto} .bar{height:12px;background:#eef;border-radius:10px;overflow:hidden;margin:15px 0} .fill{height:100%;width:0%;background:linear-gradient(90deg,#7c8aff,#8b5cf6);transition:0.3s}</style>
</head><body>
<div class="card">
<h2>الريان للترجمة - حل مجاني ضد Timeout</h2>
<input type="file" id="file" accept=".pptx"><button onclick="upload()">رفع وترجمة</button>
<div class="bar"><div id="fill" class="fill"></div></div>
<div id="log"></div>
<a id="down" style="display:none" download="translated.pptx">تحميل الملف المترجم</a>
</div>
<script>
let jobId=null, total=0;
async function upload(){
  const f=document.getElementById('file').files[0]; if(!f) return alert('اختر ملف');
  const fd=new FormData(); fd.append('file', f);
  document.getElementById('log').innerText='جاري الرفع...';
  const res=await fetch('/upload',{method:'POST',body:fd}); const data=await res.json();
  jobId=data.job_id; total=data.total; document.getElementById('log').innerText=`الملف فيه ${total} شريحة - بدأنا الترجمة شريحة شريحة`;
  translateNext(0);
}
async function translateNext(i){
  if(i>=total){ document.getElementById('log').innerText='انتهت الترجمة!'; document.getElementById('down').href='/download/'+jobId; document.getElementById('down').style.display='block'; return; }
  document.getElementById('fill').style.width=((i/total)*100)+'%';
  document.getElementById('log').innerText=`نترجم الشريحة ${i+1} من ${total}...`;
  await fetch(`/translate_slide/${jobId}/${i}`,{method:'POST'});
  translateNext(i+1);
}
</script></div></body></html>
"""

@app.route("/")
def index():
    return render_template_string(HTML_PAGE)

@app.route("/upload", methods=["POST"])
def upload():
    file = request.files['file']
    prs = Presentation(file)
    job_id = str(uuid.uuid4())
    JOBS[job_id] = {"prs": prs, "total": len(prs.slides)}
    return jsonify({"job_id": job_id, "total": len(prs.slides)})

@app.route("/translate_slide/<job_id>/<int:slide_idx>", methods=["POST"])
def translate_slide(job_id, slide_idx):
    job = JOBS.get(job_id)
    if not job: return jsonify({"error": "job not found"}), 404
    prs = job["prs"]
    slide = prs.slides[slide_idx]
    # ترجم كل مربع نص في الشريحة
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    orig = run.text
                    if orig.strip():
                        trans = translate_text(orig)
                        # الحفاظ على الأصلي + الترجمة تحته
                        if trans != orig:
                            run.text = orig + "\\n" + trans
    return jsonify({"ok": True, "slide": slide_idx})

@app.route("/download/<job_id>")
def download(job_id):
    job = JOBS.get(job_id)
    if not job: return "not found", 404
    prs = job["prs"]
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return send_file(bio, as_attachment=True, download_name="translated.pptx", mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
