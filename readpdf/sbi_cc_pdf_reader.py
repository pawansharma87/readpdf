import os
import re
import nltk
import pdfplumber
import pandas as pd
from datetime import datetime
import PyPDF2

# Download necessary NLTK data files (first-time setup)
nltk.download('punkt')

# Define category keywords
category_keywords = {
    "Utility Bill": ["UTILITIES", "PAYMENTS BANK", "ELECTRICITY", "GASCYLINDER", "RECHARGE", "PZRECHARGE"],
    "Travel": ["AUTO SERVICES", "TRANSPORT"],
    "Education": ["EDUCATION", "SCHOOL", "COLLEGE", "INSTITUTE"],
    "Food": ["FOOD PRODUCTS", "ZOMATO", "RESTAURANT", "CAFE", "HOSPITALITY", "SWIGGY"],
    "Shopping": ["DEPT STORES", "RETAIL", "SHOP", "AMAZON", "FLIPKART", "MOTO", "MARKET"],
    "Fees": ["JOINING FEE", "MEMBERSHIP FEE"],
    "Tax": ["GST", "TAX"],
    "Cashback": ["CASHBACK"],
    "Payment": ["ONLINE PAYMENT", "IMPS PAYMENT", "PAYMENT RECEIVED"],
    "Insurance": ["INSURANCE", "MAX BUPA", "COVERFOX", "PZINSURANCE"],
    "Medical": ["MEDICAL", "PHARMACY", "HOSPITAL", "DENTREE"],
    "Home Furnishing": ["HOME FURNISHING", "NOBROKER"],
    "Miscellaneous": ["MISCELLANEOUS", "PAYZAPP", "OTHERS"],
    "Fuel": ["FUEL SURCHARGE WAIVER EXCL TAX", "FUEL SURCHARGE", "FILLING", "DEEP AUTOMOBILES", "FUEL", "PATEL"],
    "Others": []
}

def categorize_description(description):
    tokens = nltk.word_tokenize(description.upper())
    for token in tokens:
        for category, keywords in category_keywords.items():
            if token in keywords:
                print(f"token : {token} {category} {description}")
                return category
    return "Others"

def unlock_pdf(input_path, output_path, password):
    """Unlock a password-protected PDF and save it as a new file."""
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
    """Extract transaction data from an unlocked PDF and categorize it."""
    # Regex patterns for different transaction formats
    patterns = [
        r'(\d{2} \w{3} \d{2})\s+(.*?)\s+IN\s+([\d,\.]+)\s*([DC])',
        r'(\d{2} \w{3} \d{2})\s+(.*?)\s+([\d,\.]+)\s+([DC])',
        r'(\d{2} \w{3} \d{2})\s+(.*?)\s+([\d,\.]+)\s*([DC])?',  # Optional transaction type
        r'(\d{2} \w{3} \d{2})\s+(.*?)\s+([\d,\.]+)\s*([D|C])?'  # Handle optional types and spaces
    ]

    transactions = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                lines = page_text.strip().split('\n')
                for line in lines:
                    matched = False
                    for pattern in patterns:
                        match = re.match(pattern, line)
                        if match:
                            matched = True
                            groups = match.groups()
                            if len(groups) == 4:
                                date_str, description, amount_str, trans_type = groups
                            elif len(groups) == 3:
                                date_str, description, amount_str = groups
                                trans_type = 'D'  # Default to debit if type is missing
                            else:
                                #print(f"Unmatched line: {line}")
                                continue

                            # Parse date
                            try:
                                date = datetime.strptime(date_str, '%d %b %y').date()
                            except ValueError:
                                date = datetime(2024, 1, 1).date()  # Default date if parsing fails

                            # Parse amount
                            amount = float(amount_str.replace(',', ''))
                            # Determine transaction type
                            trans_type = 'expense' if trans_type in ['D', 'D'] else 'income'
                            # Categorize description
                            category = categorize_description(description)

                            transactions.append({
                                'Date': date,
                                'Description': description,
                                'Amount': amount,
                                'Type': trans_type,
                                'CreditCard': credit_card_name,
                                'Category': category
                            })
                            break  # Exit pattern matching loop if match is found
                    # if not matched:
                    #     print(f"Unmatched line: {line}")
    return transactions

# Specify the directory containing the PDF files and password
pdf_directory = r"C:\Users\aspsh\OneDrive\Desktop\cc\sbi\bpcl"
password = "270819870503"

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
        transactions = extract_text_from_unlocked_pdf(output_pdf_path, "SBI BPCL")
        all_transactions.extend(transactions)

# Convert to DataFrame for easier manipulation
df = pd.DataFrame(all_transactions)

# Summarize expenses (debit) by description and category
expense_summary = df[df['Type'] == 'expense'].groupby(['Category'])['Amount'].sum().reset_index()
expense_summary = expense_summary.rename(columns={'Amount': 'Total Amount (Expense)'})

# Summarize income (credit) by description and category
income_summary = df[df['Type'] == 'income'].groupby(['Category'])['Amount'].sum().reset_index()
income_summary = income_summary.rename(columns={'Amount': 'Total Amount (Income)'})

# Save the data to an Excel file
output_excel_path = os.path.join(pdf_directory, "transaction_summary.xlsx")
with pd.ExcelWriter(output_excel_path) as writer:
    df.to_excel(writer, sheet_name='Transactions', index=False)
    expense_summary.to_excel(writer, sheet_name='Expense Summary', index=False)
    income_summary.to_excel(writer, sheet_name='Income Summary', index=False)

print(f"\nData has been exported to {output_excel_path}")
