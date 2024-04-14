import os
import zipfile
import tempfile
from PyPDF2 import PdfReader
from docx import Document
import re
from openpyxl import Workbook


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

def extract_and_process_cvs_from_zip(zip_path):
    wb = Workbook()
    ws = wb.active
    ws.append(['Email ID', 'Contact No.', 'Overall Text'])
    myl=[]

    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_ref.extractall(temp_dir)
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    text = process_cv(file_path)
                    if text:
                        emails, contacts, overall_text=extract_info_from_text(text)
                        
                        for email in emails:
                            for contact in contacts:
                                ws.append([email, contact, overall_text])
                            

    wb.save('output.xlsx')

                       
                        # ws.append([emails, contacts, overall_text])
                        # # Add your logic here to extract data from the CVs
                        # # For example:
                        # # ws.append([email, contact_no, text])
                        # wb.save('output.xlsx')
                     #print(f"Processed {file}: {text[:100]}...") # Print the first 100 characters of the text

# Example usage
zip_path = 'Sample2-20240406T093029Z-001.zip'
extract_and_process_cvs_from_zip(zip_path)



# Assuming you have a list of email addresses extracted from text


    

# # Save the workbook
