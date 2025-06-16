from PyPDF2 import PdfReader
import docx

def extract_text_from_pdf(filepath):
    reader = PdfReader(filepath)
    text = ""
    for page in reader.pages:
        text += page.extract_text()
    return text

def extract_text_from_docx(filepath):
    doc = docx.Document(filepath)
    text = " ".join([para.text for para in doc.paragraphs])
    return text

def extract_text_from_txt(filepath):
    with open(filepath, "r", encoding="utf-8") as file:
        return file.read()

def extract_text_from_files(filepaths):
    extracted_texts = []
    for filepath in filepaths:
        if filepath.endswith(".pdf"):
            extracted_texts.append(extract_text_from_pdf(filepath))
        elif filepath.endswith(".docx"):
            extracted_texts.append(extract_text_from_docx(filepath))
        elif filepath.endswith(".txt"):
            extracted_texts.append(extract_text_from_txt(filepath))
    return extracted_texts
