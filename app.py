import os, uuid, io, json
from flask import Flask, request, jsonify, send_file, render_template_string
from pptx import Presentation
import google.generativeai as genai

app = Flask(__name__)

GEMINI_KEY = os.environ.get("GEMINI_API_KEY")
if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)
    model = genai.GenerativeModel('gemini-1.5-flash')

JOBS = {}

def translate_texts_batch(texts):
    if not texts:
        return []
    prompt = (
        "ترجم النصوص التالية من الإنجليزية إلى العربية. "
        "أعد النتيجة كـ JSON array فقط بنفس الترتيب وبنفس عدد العناصر، بدون شرح.\n"
        f"النصوص: {json.dumps(texts, ensure_ascii=False)}"
    )
    try:
        resp = model.generate_content(prompt, request_options={"timeout": 30})
        cleaned = resp.text.strip()
        if cleaned.startswith("```"):
            cleaned = cleaned.split("```")[1]
            if cleaned.startswith("json"):
                cleaned = cleaned[4:]
            cleaned = cleaned.strip()
        result = json.loads(cleaned)
        if isinstance(result, list) and len(result) == len(texts):
            return result
    except Exception as e:
        print("Gemini error:", e)
    return texts

HTML_PAGE = """<!DOCTYPE html><html dir="rtl"><head><meta charset="utf-8">
<title>الريان للترجمة</title>
<style>
body{font-family:sans-serif;background:#1a1a2e;color:#fff;text-align:center;padding:40px}
button,input{font-size:16px;padding:10px;margin:10px;border-radius:8px;border:none}
button{background:#d4af37;color:#000;cursor:pointer;font-weight:bold}
button:disabled{opacity:0.5}
#log{margin-top:20px;font-size:18px}
.progress{width:80%;margin:20px auto;background:#333;border-radius:10px;height:30px;overflow:hidden}
.bar{height:100%;background:#d4af37;width:0%;transition:width 0.5s}
a{color:#d4af37;font-size:20px;display:inline-block;margin-top:20px}
.warn{color:#ff6b6b;font-size:14px;margin-top:10px}
</style></head><body>
<h2>الريان للترجمة</h2>
<p class="warn">⚠️ لا تغلق هذه الصفحة أثناء الترجمة، وإلا ستفقد التقدم.</p>
<input type="file" id="f" accept=".pptx"><br>
<button id="btn" onclick="up()">ابدأ الترجمة</button>
<div id="log"></div>
<div class="progress"><div class="bar" id="bar"></div></div>
<a id="d" style="display:none">⬇️ تحميل الملف المترجم</a>
<script>
let jobId, total;
async function up(){
  const file = document.getElementById('f').files[0];
  if(!file){alert('اختر ملفاً');return;}
  document.getElementById('btn').disabled = true;
  document.getElementById('log').innerText = 'جاري رفع الملف...';
  const fd = new FormData(); fd.append('file', file);
  const r = await fetch('/upload', {method:'POST', body:fd});
  const d = await r.json();
  if(d.error){document.getElementById('log').innerText='خطأ: '+d.error; document.getElementById('btn').disabled=false; return;}
  jobId = d.job_id; total = d.total;
  await translateNext(0);
}
async function translateNext(i){
  if(i >= total){
    document.getElementById('log').innerText = '✅ اكتملت الترجمة';
    const a = document.getElementById('d');
    a.href = '/download/' + jobId; a.style.display = 'inline-block';
    document.getElementById('btn').disabled = false;
    return;
  }
  document.getElementById('log').innerText = `ترجمة الشريحة ${i+1} من ${total}...`;
  const r = await fetch(`/translate_slide/${jobId}/${i}`, {method:'POST'});
  const d = await r.json();
  if(d.error){
    document.getElementById('log').innerText = `خطأ في الشريحة ${i+1}: ${d.error}`;
    document.getElementById('btn').disabled = false;
    return;
  }
  document.getElementById('bar').style.width = Math.round(((i+1)/total)*100) + '%';
  await translateNext(i+1);
}
</script></body></html>"""

@app.route("/")
def index():
    return render_template_string(HTML_PAGE)

@app.route("/upload", methods=["POST"])
def upload():
    if 'file' not in request.files:
        return jsonify({"error": "لا يوجد ملف"})
    f = request.files['file']
    try:
        data = f.read()
        prs = Presentation(io.BytesIO(data))
        total = len(prs.slides)
        jid = str(uuid.uuid4())
        JOBS[jid] = {"data": data, "total": total}
        return jsonify({"job_id": jid, "total": total})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/translate_slide/<job_id>/<int:idx>", methods=["POST"])
def translate_slide(job_id, idx):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "انتهت الجلسة، ارفع الملف من جديد"})
    try:
        prs = Presentation(io.BytesIO(job["data"]))
        slide = prs.slides[idx]
        
        paragraphs = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    txt = "".join(r.text for r in para.runs).strip()
                    if txt:
                        paragraphs.append((para, txt))
        
        texts = [p[1] for p in paragraphs]
        if texts:
            translations = translate_texts_batch(texts)
            for (para, original), trans in zip(paragraphs, translations):
                if trans and trans != original:
                    from pptx.oxml.ns import qn
                    from lxml import etree
                    import copy
                    ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
                    new_p = copy.deepcopy(para._p)
                    para._p.addnext(new_p)
                    for r in new_p.findall(qn('a:r')):
                        new_p.remove(r)
                    r = etree.SubElement(new_p, f'{{{ns}}}r')
                    t = etree.SubElement(r, f'{{{ns}}}t')
                    t.text = trans
        
        bio = io.BytesIO()
        prs.save(bio)
        job["data"] = bio.getvalue()
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/download/<job_id>")
def download(job_id):
    job = JOBS.get(job_id)
    if not job:
        return "انتهت الجلسة", 404
    return send_file(io.BytesIO(job["data"]), as_attachment=True,
                     download_name="translated.pptx",
                     mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 10000)))
