import os
import uuid
import json
import re
import traceback

from flask import Flask, render_template, request, send_file, flash, redirect, jsonify
from werkzeug.utils import secure_filename
from PyPDF2 import PdfReader
from docx import Document
import language_tool_python
import google.generativeai as genai
from playwright.sync_api import sync_playwright

# === Flask Setup ===
app = Flask(__name__)
app.secret_key = 'supersecret'

UPLOAD_FOLDER = 'uploads'
GENERATED_FOLDER = 'generated'
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'txt'}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(GENERATED_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# === Gemini API Key Rotation Setup ===
GENAI_API_KEYS = [
    "AIzaSyBWHfbDFcuhKwEL-uX0j-SRyUUSidgBhaE",
    "AIzaSyDzzIncgM-mfsad8QYWyYG5PvQtXYdlpbs"
]
API_USE_LIMIT = 50
api_call_count = 0
current_api_index = 0

def configure_genai():
    genai.configure(api_key=GENAI_API_KEYS[current_api_index])

configure_genai()

def increment_api_usage():
    global api_call_count, current_api_index
    api_call_count += 1
    if api_call_count >= API_USE_LIMIT:
        api_call_count = 0
        current_api_index = (current_api_index + 1) % len(GENAI_API_KEYS)
        configure_genai()

# === LanguageTool Setup ===
tool = language_tool_python.LanguageTool('en-US')

# === Helper Functions ===
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def extract_text(filepath, ext):
    if ext == 'pdf':
        with open(filepath, 'rb') as f:
            return ' '.join([p.extract_text() for p in PdfReader(f).pages if p.extract_text()])
    elif ext == 'docx':
        return '\n'.join([p.text for p in Document(filepath).paragraphs])
    elif ext == 'txt':
        with open(filepath, 'r', encoding='utf-8') as f:
            return f.read()
    return ""

def clean_formatting(text):
    text = re.sub(r'\r\n|\r', '\n', text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    text = re.sub(r'[ \t]{2,}', ' ', text)
    text = re.sub(r'•', '-', text)
    text = re.sub(r'\n\s*-\s*', '\n- ', text)
    return text.strip()

def get_grammar_suggestions(text):
    matches = tool.check(text)
    suggestions = []
    for i, match in enumerate(matches):
        suggestion = {
            'id': i,
            'offset': match.offset,
            'length': match.errorLength,
            'message': match.message,
            'original': text[match.offset:match.offset + match.errorLength],
            'suggestions': match.replacements[:3] if match.replacements else [],
            'rule_id': match.ruleId,
            'category': match.category
        }
        suggestions.append(suggestion)
    return suggestions

def apply_grammar_corrections(text, accepted_corrections):
    # Sort in reverse by offset so that replacements don't affect subsequent offsets
    accepted_corrections.sort(key=lambda x: x['offset'], reverse=True)
    corrected_text = text
    for correction in accepted_corrections:
        start = correction['offset']
        end = start + correction['length']
        corrected_text = corrected_text[:start] + correction['replacement'] + corrected_text[end:]
    return corrected_text

def auto_correct_text(text):
    matches = tool.check(text)
    matches.sort(key=lambda x: x.offset, reverse=True)
    corrected_text = text
    for match in matches:
        if match.replacements:
            start = match.offset
            end = start + match.errorLength
            corrected_text = corrected_text[:start] + match.replacements[0] + corrected_text[end:]
    return corrected_text

def extract_json(text):
    if "```json" in text:
        match = re.findall(r"```json(.*?)```", text, re.DOTALL)
        if match:
            try:
                return json.loads(match[0].strip())
            except Exception:
                return {}
    try:
        return json.loads(text)
    except Exception:
        return {}

def generate_structured_data(text):
    increment_api_usage()
    prompt = f"""
You are an expert HR resume parser. Convert the following text into structured JSON with these fields:

{{
  "name": "",
  "education_training_certifications": [],
  "total_experience": "",
  "professional_summary": "",
  "netweb_projects": [{{"title": "", "description": ""}}],
  "past_projects": [{{"title": "", "description": ""}}],
  "roles_responsibilities": "",
  "technical_skills": {{
    "web_technologies": [],
    "scripting_languages": [],
    "frameworks": [],
    "databases": [],
    "web_servers": [],
    "tools": []
  }},
  "personal_details": {{
    "employee_id": "",
    "permanent_address": "",
    "local_address": "",
    "contact_number": "",
    "date_of_joining": "",
    "designation": "",
    "overall_experience": "",
    "date_of_birth": "",
    "passport_details": ""
  }}
}}

Return ONLY valid JSON. Do not return markdown or explanation.

Resume:

{text}
"""
    model = genai.GenerativeModel("gemini-2.5-flash")
    response = model.generate_content(prompt)
    return extract_json(response.text)

def render_html_to_pdf(html_string, output_path):
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page()
        page.set_content(html_string)
        page.pdf(path=output_path, format="A4", print_background=True)
        browser.close()

# === Routes ===
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('file_input')
        prompt_text = request.form.get('text_input')
        text = ""

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(path)
            ext = filename.rsplit('.', 1)[1].lower()
            text = extract_text(path, ext)
        elif prompt_text and prompt_text.strip():
            text = prompt_text.strip()

        if not text:
            flash("Please upload a file or enter resume text.")
            return redirect('/')

        formatted_text = clean_formatting(text)
        grammar_suggestions = get_grammar_suggestions(formatted_text)

        return render_template("grammar_check.html", 
                               original_text=formatted_text, 
                               suggestions=grammar_suggestions)

    return render_template("index.html")

@app.route('/process_grammar', methods=['POST'])
def process_grammar():
    original_text = request.form.get('original_text')
    action = request.form.get('action')
    
    if action == 'skip':
        profile = generate_structured_data(original_text)
        return render_template("profile.html", profile=profile)
    elif action == 'auto_correct':
        corrected_text = auto_correct_text(original_text)
        profile = generate_structured_data(corrected_text)
        return render_template("profile.html", profile=profile)

    corrections_data = request.form.get('corrections')

    try:
        accepted_corrections = json.loads(corrections_data) if corrections_data else []
    except json.JSONDecodeError:
        flash("Invalid correction data submitted.")
        return redirect('/')

    corrected_text = apply_grammar_corrections(original_text, accepted_corrections)
    profile = generate_structured_data(corrected_text)
    return render_template("profile.html", profile=profile)

@app.route('/download', methods=['POST'])
def download():
    profile_data = request.form.get("profile_data")
    try:
        profile = json.loads(profile_data)
    except Exception:
        flash("Invalid profile data.")
        return redirect('/')

    try:
        html = render_template("profile.html", profile=profile)
        output_path = os.path.join(GENERATED_FOLDER, f"profile_{uuid.uuid4().hex}.pdf")
        render_html_to_pdf(html, output_path)
        return send_file(output_path, as_attachment=True, download_name="Employee_Profile.pdf")
    except Exception as e:
        print("❌ PDF Generation Error:", e)
        print(traceback.format_exc())
        flash("PDF generation failed.")
        return redirect('/')

# === Run Server ===
if __name__ == '__main__':
    app.run(debug=True)