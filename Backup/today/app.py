import os
 
import uuid
 
import json
 
import re
 
import traceback
 
from flask import Flask, render_template, request, send_file, flash, redirect
 
from werkzeug.utils import secure_filename
 
from PyPDF2 import PdfReader
 
from docx import Document
 
import language_tool_python
 
import google.generativeai as genai
 
from playwright.sync_api import sync_playwright
 
# ✅ Gemini API Key
 
genai.configure(api_key="AIzaSyBWHfbDFcuhKwEL-uX0j-SRyUUSidgBhaE")
 
app = Flask(__name__)
 
app.secret_key = 'supersecret'
 
UPLOAD_FOLDER = 'uploads'
 
GENERATED_FOLDER = 'generated'
 
ALLOWED_EXTENSIONS = {'pdf', 'docx', 'txt'}
 
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
 
os.makedirs(GENERATED_FOLDER, exist_ok=True)
 
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
 
model = genai.GenerativeModel("gemini-2.5-flash")
 
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
 
def correct_grammar(text):
 
    matches = tool.check(text)
 
    return language_tool_python.utils.correct(text, matches)
 
def extract_json(text):
 
    if "```json" in text:
 
        match = re.findall(r"```json(.*?)```", text, re.DOTALL)
 
        if match:
 
            try:
 
                return json.loads(match[0].strip())
 
            except:
 
                return {}
 
    try:
 
        return json.loads(text)
 
    except:
 
        return {}
 
def generate_structured_data(text):
 
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
 
        cleaned_text = correct_grammar(formatted_text)
 
        profile = generate_structured_data(cleaned_text)
 
        return render_template("profile.html", profile=profile)
 
    return render_template("index.html")
 
 
@app.route('/download', methods=['POST'])
 
def download():
 
    profile_data = request.form.get("profile_data")
 
    try:
 
        profile = json.loads(profile_data)
 
    except:
 
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
 
 
# === Run ===
 
if __name__ == '__main__':
 
    app.run(debug=True)
 
 
 