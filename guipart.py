import os
import re
import docx2txt
import pandas as pd
from PyPDF2 import PdfReader
import subprocess
from win32com import client as wc
import streamlit as st

def doc_to_docx(doc_file, docx_file):
    try:
        word = wc.Dispatch('Word.Application')
        doc = word.Documents.Open(doc_file)
        doc.SaveAs(docx_file, 12)  # FileFormat for .docx
        doc.Close()
        word.Quit()
        print(f"Converted {doc_file} to {docx_file}")
        return True
    except Exception as e:
        print(f"Error converting {doc_file} to {docx_file}: {e}")
        return False

def extract_mobile_numbers(text):
    # Regular expression to extract mobile numbers
    mobile_numbers = re.findall(r'\b(?:[0-9][\s-]*){9,}\b', text)
    return mobile_numbers

def extract_email_addresses(text):
    # Regular expression to extract email addresses
    email_addresses = re.findall(r'[\w\.-]+@[\w\.-]+', text)
    return email_addresses

def extract_other_text(text):
    # Remove mobile numbers and email addresses
    text = re.sub(r'\b(?:[0-9][\s-]*){9,}\b', '', text)
    text = re.sub(r'[\w\.-]+@[\w\.-]+', '', text)
    # Remove extra spaces and newlines
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_information_from_docx(cv_file):
    # Load the CV text
    text = docx2txt.process(cv_file)
    mobile_numbers = extract_mobile_numbers(text)
    email_addresses = extract_email_addresses(text)
    other_text = extract_other_text(text)
    return {'Text': other_text, 'Mobile Numbers': mobile_numbers, 'Email Addresses': email_addresses}

def extract_information_from_pdf(cv_file):
    # Extract text from PDF
    with open(cv_file, 'rb') as f:
        reader = PdfReader(f)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    mobile_numbers = extract_mobile_numbers(text)
    email_addresses = extract_email_addresses(text)
    other_text = extract_other_text(text)
    return {'Text': other_text, 'Mobile Numbers': mobile_numbers, 'Email Addresses': email_addresses}

def extract_information_from_cv(cv_file):
    # Extracting text based on file type
    if cv_file.endswith('.docx'):
        return extract_information_from_docx(cv_file)
    elif cv_file.endswith('.pdf'):
        return extract_information_from_pdf(cv_file)
    else:
        return None  # Unsupported file format

def save_to_excel(data, output_file):
    # Create a DataFrame from the extracted data
    df = pd.DataFrame(data)

    # Save to Excel
    df.to_excel(output_file, index=False)

def main():
    st.title("CV Information Extractor")

    uploaded_files = st.file_uploader("Upload CV files", accept_multiple_files=True)
    if st.button("Extract"):
        if uploaded_files:
            output_folder = "output_folder"
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            all_cv_data = []
            for uploaded_file in uploaded_files:
                with open(os.path.join(output_folder, uploaded_file.name), "wb") as f:
                    f.write(uploaded_file.getvalue())

                cv_data = extract_information_from_cv(os.path.join(output_folder, uploaded_file.name))
                if cv_data:
                    cv_data['Filename'] = uploaded_file.name
                    all_cv_data.append(cv_data)

            if all_cv_data:
                output_file = 'cv_data.xlsx'
                save_to_excel(all_cv_data, output_file)
                st.success(f"CV data extracted from {len(all_cv_data)} CVs and saved to {output_file}")
            else:
                st.warning("No CVs found or unsupported file formats.")
        else:
            st.warning("Please upload CV files.")

if __name__ == "__main__":
    main()
