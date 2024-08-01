import PyPDF2
import pdfplumber

def unlock_pdf(input_path, output_path, password):
    with open(input_path, 'rb') as infile:
        reader = PyPDF2.PdfReader(infile)
        if reader.is_encrypted:
            reader.decrypt(password)
            writer = PyPDF2.PdfWriter()
            for page_num in range(len(reader.pages)):
                writer.add_page(reader.pages[page_num])

            with open(output_path, 'wb') as outfile:
                writer.write(outfile)

def extract_text_from_unlocked_pdf(pdf_path):
    text = ''
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text
    return text

# Provide the path to your password-protected PDF and the password
input_pdf_path = r"C:\Users\aspsh\OneDrive\Desktop\cc\sbi\bpcl\9654519415575153_16012024.pdf"
output_pdf_path = r"C:\Users\aspsh\OneDrive\Desktop\cc\unlocked_airtel_June24.pdf"
password = "XXXX"

# Unlock the PDF
unlock_pdf(input_pdf_path, output_pdf_path, password)

# Extract text from the unlocked PDF
text = extract_text_from_unlocked_pdf(output_pdf_path)
print(text)