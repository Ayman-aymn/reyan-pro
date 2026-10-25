import gradio as gr
import easyocr
import requests

reader = easyocr.Reader(['en', 'ar'])

def translate_image(image):
    results = reader.readtext(image)
    text = " ".join([res[1] for res in results])
    
    if not text.strip():
        return "لم يتم العثور على نص في الصورة"
    
    try:
        r = requests.post(
            "https://api-free.deepl.com/v2/translate",
            headers={"Authorization": "DeepL-Auth-Key YOUR_DEEPL_KEY"},
            json={"text": [text], "target_lang": "AR"},
            timeout=10
        )
        if r.status_code == 200:
            translated = r.json()['translations'][0]['text']
            return f"النص الأصلي: {text}\n\nالترجمة: {translated}"
    except:
        pass
    
    return f"النص الأصلي: {text}\n\n(لم تتم الترجمة، تحقق من مفتاح DeepL)"

with gr.Blocks() as demo:
    gr.Markdown("# 🖼️ مترجم الصور - الريان")
    gr.Markdown("ارفع صورة، وسيتم استخراج النص منها وترجمته إلى العربية.")
    with gr.Row():
        image_input = gr.Image(type="filepath", label="ارفع الصورة")
        text_output = gr.Textbox(label="النتيجة", lines=5)
    btn = gr.Button("ترجم الصورة")
    btn.click(fn=translate_image, inputs=image_input, outputs=text_output)

if __name__ == "__main__":
    demo.launch()
