import os
from flask import Flask, request, jsonify, render_template_string
import google.generativeai as genai

app = Flask(__name__)

GEMINI_KEY = os.environ.get("GEMINI_API_KEY")
if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)
    model = genai.GenerativeModel('gemini-1.5-flash')

HTML_PAGE = """<!DOCTYPE html>
<html dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>الريان للترجمة - النصوص</title>
<style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
        font-family: 'Tajawal', sans-serif;
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        min-height: 100vh; display: flex; justify-content: center; align-items: center; padding: 20px;
    }
    .container {
        background: rgba(255, 255, 255, 0.95); border-radius: 24px; padding: 40px;
        max-width: 600px; width: 100%; text-align: center; box-shadow: 0 20px 60px rgba(0, 0, 0, 0.4);
        border: 1px solid rgba(255, 215, 0, 0.3);
    }
    .logo { font-size: 36px; font-weight: 700; color: #1a1a2e; margin-bottom: 10px; }
    .logo span { color: #d4af37; }
    .subtitle { font-size: 16px; color: #666; margin-bottom: 30px; }
    textarea {
        width: 100%; height: 150px; padding: 15px; border: 2px solid #ddd; border-radius: 12px;
        font-family: inherit; font-size: 16px; resize: vertical; margin-bottom: 20px;
    }
    textarea:focus { outline: none; border-color: #d4af37; }
    .btn {
        background: linear-gradient(135deg, #d4af37 0%, #b8960c 100%); color: white;
        padding: 15px 30px; border: none; border-radius: 12px; font-size: 18px; font-weight: 700;
        cursor: pointer; width: 100%; transition: 0.3s;
    }
    .btn:disabled { opacity: 0.5; cursor: not-allowed; }
    .result-box {
        margin-top: 25px; padding: 20px; background: #f8f9fa; border-radius: 12px;
        text-align: right; font-size: 18px; line-height: 1.6; display: none; border: 1px solid #d4af37;
    }
</style>
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@400;700&display=swap" rel="stylesheet">
</head>
<body>
    <div class="container">
        <div class="logo">الريان <span>للترجمة</span></div>
        <p class="subtitle">ترجمة نصوص دقيقة وسريعة</p>
        
        <textarea id="inputText" placeholder="ألصق النص الإنجليزي هنا..."></textarea>
        <button class="btn" id="translateBtn" onclick="translateText()">ترجم الآن</button>
        
        <div id="resultBox" class="result-box"></div>
        <div id="errorMsg" style="color: red; margin-top: 15px;"></div>
    </div>

    <script>
        async function translateText() {
            const text = document.getElementById('inputText').value.trim();
            const btn = document.getElementById('translateBtn');
            const resultBox = document.getElementById('resultBox');
            const errorMsg = document.getElementById('errorMsg');

            if (!text) {
                alert('الرجاء إدخال نص للترجمة');
                return;
            }

            btn.disabled = true;
            btn.innerText = 'جاري الترجمة...';
            resultBox.style.display = 'none';
            errorMsg.innerText = '';

            try {
                const response = await fetch('/translate', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ text: text })
                });
                
                const data = await response.json();
                
                if (data.error) {
                    errorMsg.innerText = 'خطأ: ' + data.error;
                } else {
                    resultBox.innerText = data.translation;
                    resultBox.style.display = 'block';
                }
            } catch (error) {
                errorMsg.innerText = 'حدث خطأ في الاتصال بالخادم.';
            } finally {
                btn.disabled = false;
                btn.innerText = 'ترجم الآن';
            }
        }
    </script>
</body>
</html>"""

@app.route('/')
def index():
    return render_template_string(HTML_PAGE)

@app.route('/translate', methods=['POST'])
def translate():
    data = request.json
    text = data.get('text', '')
    
    if not text:
        return jsonify({"error": "لم يتم إرسال نص"})

    try:
        prompt = f"ترجم النص التالي من الإنجليزية إلى العربية بدقة عالية. أعد الترجمة فقط بدون أي مقدمات أو شروحات:\n\n{text}"
        response = model.generate_content(prompt, request_options={"timeout": 30})
        
        translation = response.text.strip()
        return jsonify({"translation": translation})
    except Exception as e:
        return jsonify({"error": str(e)})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get("PORT", 10000)))
