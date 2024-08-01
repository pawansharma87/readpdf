import os
import re
import nltk
import pdfplumber
import pandas as pd
from datetime import datetime
import PyPDF2

# Download necessary NLTK data files (first-time setup)
nltk.download('punkt')

category_keywords = {
    "Utility Bill": ["UTILITIES", "PAYMENTS BANK", "ELECTRICITY", "GASCYLINDER", "RECHARGE"],
    "Travel": ["AUTO SERVICES", "BANGALORE", "TRANSPORT"],
    "Education": ["EDUCATION", "SCHOOL", "COLLEGE", "INSTITUTE"],
    "Food": ["FOOD PRODUCTS", "ZOMATO", "RESTAURANT", "CAFE", "HOSPITALITY", "SWIGGY"],
    "Shopping": ["DEPT STORES", "RETAIL", "SHOP", "AMAZON", "FLIPKART", "MOTO", "MARKET"],
    "Fees": ["JOINING FEE", "MEMBERSHIP FEE"],
    "Tax": ["GST", "TAX"],
    "Cashback": ["CASHBACK"],
    "Payment": ["ONLINE PAYMENT", "IMPS PAYMENT"],
    "Insurance": ["INSURANCE", "MAX BUPA", "COVERFOX"],
    "Medical": ["MEDICAL", "PHARMACY", "HOSPITAL"],
    "Home Furnishing": ["HOME FURNISHING", "NOBROKER"],
    "Miscellaneous": ["MISCELLANEOUS", "PAYZAPP", "OTHERS"],
    "Others": []
}

def categorize_description(description):
    tokens = nltk.word_tokenize(description.upper())
    for token in tokens:
        for category, keywords in category_keywords.items():
            if token in keywords:
                return category
    return "Others"

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

def extract_text_from_unlocked_pdf(pdf_path, credit_card_name):
    transactions = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                lines = page_text.strip().split('\n')
                for line in lines:
                    match = re.match(r'(\d{2}/\d{2}/\d{4})\s+(.*?)\s+([\d,]+\.\d{2})\s+(Dr|Cr)', line)
                    if not match:
                        match = re.match(r'(\d{2}/\d{2})\s+(\d+)\s+(.*?)\s+([\d,]+\.\d{2})(Dr|Cr)', line)
                    if match:
                        date, description, amount, trans_type = match.groups()
                        amount = float(amount.replace(',', ''))
                        type = 'expense' if trans_type == 'Dr' else 'income'
                        category = categorize_description(description)
                        transactions.append({
                            'Date': datetime.strptime(date, '%d/%m/%Y').date(),
                            'Description': description,
                            'Amount': amount,
                            'Type': type,
                            'CreditCard': credit_card_name,
                            'Category': category
                        })
    return transactions

# Specify the directory containing the PDF files
#pdf_directory = r"C:\Users\aspsh\OneDrive\Desktop\cc\airtel"
#password = "BITT0106"  # Add the password if required

pdf_directory = r"C:\Users\aspsh\OneDrive\Desktop\cc\rewards"
password = "PAWA27AUG"

# Aggregate all transactions
all_transactions = []

# Process each PDF in the directory
for filename in os.listdir(pdf_directory):
    if filename.endswith(".pdf"):
        pdf_path = os.path.join(pdf_directory, filename)
        output_pdf_path = os.path.join(pdf_directory, f"unlocked_{filename}")

        # Unlock the PDF if it's password protected
        try:
            unlock_pdf(pdf_path, output_pdf_path, password)
        except Exception as e:
            print(f"Could not unlock {filename}: {e}")
            continue

        # Extract text from the unlocked PDF
        transactions = extract_text_from_unlocked_pdf(output_pdf_path, "Axis Airtel")
        all_transactions.extend(transactions)

# Convert to DataFrame for easier manipulation
df = pd.DataFrame(all_transactions)

# Summarize expenses (debit) by description and category
expense_summary = df[df['Type'] == 'expense'].groupby(['Description', 'Category'])['Amount'].sum().reset_index()
expense_summary = expense_summary.rename(columns={'Amount': 'Total Amount (Expense)'})

# Summarize income (credit) by description and category
income_summary = df[df['Type'] == 'income'].groupby(['Description', 'Category'])['Amount'].sum().reset_index()
income_summary = income_summary.rename(columns={'Amount': 'Total Amount (Income)'})

# Save the data to an Excel file
output_excel_path = os.path.join(pdf_directory, "transaction_summary.xlsx")
with pd.ExcelWriter(output_excel_path) as writer:
    df.to_excel(writer, sheet_name='Transactions', index=False)
    expense_summary.to_excel(writer, sheet_name='Expense Summary', index=False)
    income_summary.to_excel(writer, sheet_name='Income Summary', index=False)

print(f"\nData has been exported to {output_excel_path}")
