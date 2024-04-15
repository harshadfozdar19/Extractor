import os
import re
import docx2txt
import pandas as pd
from PyPDF2 import PdfReader
import subprocess


def extract_text_from_doc(cv_file):
    try:
        result = subprocess.run(["antiword", cv_file], capture_output=True, text=True)
        text = result.stdout
        return text
    except Exception as e:
        print(f"Error extracting text: {e}")
        return None

# Example usage:
# file_path = "example.doc"  # Replace with your file path
# text = extract_text_from_doc(file_path)
# if text:
#     print(text)





def extract_information_from_docx(cv_file):
    # Load the CV text
    return docx2txt.process(cv_file)

def extract_information_from_pdf(cv_file):
    # Extract text from PDF
    with open(cv_file, 'rb') as f:
        reader = PdfReader(f)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    return text

def extract_information_from_cv(cv_file):
    # Extracting text based on file type
    if cv_file.endswith('.docx'):
        return extract_information_from_docx(cv_file)
    elif cv_file.endswith('.pdf'):
        return extract_information_from_pdf(cv_file)
    elif  cv_file.endswith('.doc'):
        return extract_text_from_doc(cv_file)

    else:
        return None  # Unsupported file format

def save_to_excel(data, output_file):
    # Create a DataFrame from the extracted data
    df = pd.DataFrame(data)

    # Save to Excel
    df.to_excel(output_file, index=False)

if __name__ == "__main__":
    # Specify the folder containing CV files
    cv_folder = r'C:\Users\harsh\Desktop\Python development question for internship\Sample2'  

    # Specify the output Excel file
    output_file = 'cv_data.xlsx'

    # Initialize list to store extracted data from all CVs
    all_cv_data = []

    # Process each CV file in the folder
    for filename in os.listdir(cv_folder):
        cv_file = os.path.join(cv_folder, filename)
        if os.path.isfile(cv_file):
            cv_text = extract_information_from_cv(cv_file)
            if cv_text:
                all_cv_data.append({'Filename': filename, 'Text': cv_text})

    if all_cv_data:
        # Save the data to Excel
        save_to_excel(all_cv_data, output_file)

        print(f"CV data extracted from {len(all_cv_data)} CVs and saved to {output_file}")
    else:
        print("No CVs found in the specified folder or unsupported file formats.")
