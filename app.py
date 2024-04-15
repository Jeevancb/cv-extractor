import os
import zipfile
import tempfile
from PyPDF2 import PdfReader # type: ignore
from docx import Document # type: ignore
import re
from openpyxl import Workbook # type: ignore
from flask import Flask, request, send_file,render_template,Response,redirect,url_for
from io import BytesIO
app = Flask(__name__)


def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        pdf = PdfReader(file)
        text = ""
        for page in range(len(pdf.pages)):
            text += pdf.pages[page].extract_text()
    return text

def extract_text_from_docx(docx_path):
    doc = Document(docx_path)
    text = " ".join([paragraph.text for paragraph in doc.paragraphs])
    return text

def extract_info_from_text(text):
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    contact_pattern = r'\b\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}\b'
    
    emails = re.findall(email_pattern, text)
    contacts = re.findall(contact_pattern, text)
    
   
    return  emails, contacts, text


def process_cv(file_path):
    if file_path.endswith('.pdf'):
        return extract_text_from_pdf(file_path)
    elif file_path.endswith('.docx'):
        return extract_text_from_docx(file_path)
    else:
        return None
    

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file and allowed_file(file.filename):
        # Process the uploaded file
        wb = Workbook()
        ws = wb.active
        ws.append(['Email ID', 'Contact No.', 'Overall Text'])
        # Assuming the file is a ZIP file containing CVs
        with zipfile.ZipFile(file, 'r') as zip_ref:
            with tempfile.TemporaryDirectory() as temp_dir:
                zip_ref.extractall(temp_dir)
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        text = process_cv(file_path)
                        if text:
                            emails, contacts, overall_text = extract_info_from_text(text)
                            for email in emails:
                                for contact in contacts:
                                    ws.append([email, contact, overall_text])
        # Save the workbook to a temporary file
        wb.save('cv_data.xlsx')
        return redirect(url_for('download_page'))
    return 'File type not allowed'

@app.route('/download_page')
def download_page():
    return render_template('download.html')

@app.route('/download')
def download():
    return send_file('cv_data.xlsx', as_attachment=True, download_name='cv_data.xlsx')


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ['zip']

@app.route('/')
def upload_form():
    return render_template('upload.html')


# Example usage
if __name__ == '__main__':
    app.run(host='0.0.0.0',port=5000)





# Assuming you have a list of email addresses extracted from text


    

# # Save the workbook
