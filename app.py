import os, uuid, io, json
from flask import Flask, request, jsonify, send_file, render_template_string
from pptx import Presentation
import requests
import google.generativeai as genai

app = Flask(__name__)
JOBS = {}

DEEPL_KEY = os.environ.get("DEEPL_API_KEY")
GEMINI_KEY = os.environ.get("GEMINI_API_KEY")
if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)
    model = genai.GenerativeModel('gemini-1.5-flash')

def translate_with_deepl(text):
    try:
        r = requests.post("https://api-free.deepl.com/v2/translate", data={"auth_key": DEEPL_KEY, "text": text, "target_lang": "AR"}, timeout=8)
        if r.status_code==200: return r.json()["translations"][0]["text"]
    except: pass
    return None

def translate_slide_with_gemini(texts):
    # texts = قائمة بكل نصوص الشريحة
    if not texts: return []
    prompt = f"""ترجم النصوص التالية من الإنجليزية إلى العربية. 
    أعدها كـ JSON array فقط بنفس الترتيب بدون شرح.
    النصوص:
    {json.dumps(texts, ensure_ascii=False)}
    """
    try:
        resp = model.generate_content(prompt)
        # حاول استخراج JSON
        cleaned = resp.text.replace("```json","").replace("```","").strip()
        return json.loads(cleaned)
    except Exception as e:
        print("Gemini error", e)
        # fallback لـ DeepL جملة جملة
        return [translate_with_deepl(t) or t for t in texts]

HTML_PAGE = """<html dir="rtl"><head><meta charset="utf-8"><title>الريان Gemini</title></head><body>
<h3>الترجمة بـ Gemini - شريحة شريحة (أسرع)</h3>
<input type="file" id="f"><button onclick="up()">رفع</button><div id="log"></div><a id="d" style="display:none">تحميل</a>
<script>
let jobId,total;
async function up(){
 const fd=new FormData(); fd.append('file', document.getElementById('f').files[0]);
 const r=await fetch('/upload',{method:'POST',body:fd}); const d=await r.json(); jobId=d.job_id; total=d.total; next(0);
}
async function next(i){
 if(i>=total){ document.getElementById('log').innerText='انتهى'; document.getElementById('d').href='/download/'+jobId; document.getElementById('d').style.display='block'; return; }
 document.getElementById('log').innerText=`ترجمة شريحة ${i+1}/${total} بـ Gemini...`;
 await fetch(`/translate_slide/${jobId}/${i}`,{method:'POST'});
 next(i+1);
}
</script></body></html>"""

@app.route("/")
def index(): return render_template_string(HTML_PAGE)

@app.route("/upload", methods=["POST"])
def upload():
    prs = Presentation(request.files['file'])
    jid = str(uuid.uuid4())
    JOBS[jid] = {"prs": prs, "total": len(prs.slides)}
    return jsonify({"job_id": jid, "total": len(prs.slides)})

@app.route("/translate_slide/<job_id>/<int:idx>", methods=["POST"])
def translate_slide(job_id, idx):
    job = JOBS.get(job_id); prs=job["prs"]; slide=prs.slides[idx]
    # جمع كل النصوص في الشريحة
    runs_list = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    if run.text.strip():
                        runs_list.append(run)
    
    texts = [r.text for r in runs_list]
    if texts:
        translated = translate_slide_with_gemini(texts)
        # الصق الترجمة تحت الأصلي
        for run, trans in zip(runs_list, translated):
            if trans and trans != run.text:
                run.text = run.text + "\n" + trans
    return jsonify({"ok": True})

@app.route("/download/<job_id>")
def download(job_id):
    bio=io.BytesIO(); JOBS[job_id]["prs"].save(bio); bio.seek(0)
    return send_file(bio, as_attachment=True, download_name="translated_gemini.pptx")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
