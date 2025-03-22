from flask import Blueprint, request, jsonify
import openai
import os

translate_bp = Blueprint('translate', __name__)

openai.api_key = os.getenv("OPENAI_API_KEY")

@translate_bp.route('/translate', methods=['POST'])
def translate_text():
    data = request.json
    text = data.get("text")
    source_lang = data.get("source_lang")
    target_lang = data.get("target_lang")

    if not text or not source_lang or not target_lang:
        return jsonify({"error": "Missing parameters"}), 400

    prompt = f"Translate this from {source_lang} to {target_lang}: {text}"

    try:
        client = openai.OpenAI() 

        response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}]
        )

        translation = response.choices[0].message.content

        return jsonify({"translated_text": translation})
    except Exception as e:
        return jsonify({"error": str(e)}), 500
