import os
import re
import pdfplumber
import pandas as pd
from datetime import datetime
import PyPDF2
import logging

# Configure logging
logging.basicConfig(filename='transaction_processing.log', level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')

pdf_directory = r"C:\Users\aspsh\OneDrive\Desktop\cc\axis\both"

patterns = [
    r'(\d{2}/\d{2}/\d{4})\s+(.*?)\s+([\d,]+\.\d{2})\s+(Dr|Cr)',  # New pattern
    r'(\d{2}/\d{2})\s+(\d+)\s+(.*?)\s+([\d,]+\.\d{2})(Dr|Cr)',  # New pattern
    r'(\d{2} \w{3} \d{2})\s+(.*?)\s+IN\s+([\d,\.]+)\s*([DC])',
    r'(\d{2} \w{3} \d{2})\s+(.*?)\s+([\d,\.]+)\s+([DC])',
    r'(\d{2} \w{3} \d{2})\s+(.*?)\s+([\d,\.]+)\s*([DC])?',  # Optional transaction type
    r'(\d{2} \w{3} \d{2})\s+(.*?)\s+([\d,\.]+)\s*([D|C])?'  # Handle optional types and spaces
]
# Define category keywords
category_keywords = {
    "Utility Bill": ["UTILITIES", "PAYMENTS BANK", "ELECTRICITY", "GASCYLINDER", "RECHARGE", "PZRECHARGE"],
    "Travel": ["AUTO SERVICES", "TRANSPORT"],
    "Grocery": ["THE BIG MARKET", "RELIANCE", "LULU VALUE MART", "MARKET"],
    "Education": ["EDUCATION", "SCHOOL", "COLLEGE", "INSTITUTE", "NEW HORIZON", "ONIP"],
    "Food": ["FOOD PRODUCTS", "ZOMATO", "RESTAURANT", "CAFE", "HOSPITALITY", "SWIGGY"],
    "Shopping": ["DEPT STORES", "RETAIL", "SHOP", "AMAZON", "FLIPKART", "MOTO", "MYNTRA", "CASIO HYPER SHOPPEE"],
    "Fees": ["JOINING FEE", "MEMBERSHIP FEE"],
    "Tax": ["GST", "TAX"],
    "Cashback": ["CASHBACK"],
    "Payment": ["ONLINE PAYMENT", "IMPS PAYMENT", "PAYMENT RECEIVED"],
    "Insurance": ["INSURANCE", "MAX BUPA", "COVERFOX", "PZINSURANCE"],
    "Medical": ["MEDICAL", "PHARMACY", "HOSPITAL", "DENTREE", "TATA 1MG HEALTHCARE", "LKST931"],
    "Home Furnishing": ["HOME FURNISHING", "NOBROKER"],
    "Miscellaneous": ["MISCELLANEOUS", "PAYZAPP", "OTHERS", "NAIVEDYA", "SWADESHI ENTERPRISES", "Unity Enterprises", "PICTURE PEOPLE", "GURUMURTHY S"],
    "Fuel": ["FUEL SURCHARGE WAIVER EXCL TAX", "FUEL SURCHARGE", "FILLING", "DEEP AUTOMOBILES", "FUEL", "PATEL"],
    "Entertainment": ["HOTSTAR", "HOTSTAR CYBS SI"]
}

def categorize_description(description):
    description_upper = description.upper()
    for category, keywords in category_keywords.items():
        for keyword in keywords:
            if keyword in description_upper:
                logging.info(f"Matched keyword: '{keyword}' in description: '{description}' under category: {category}")
                return category
    return "Others"

def unlock_pdf(input_path, output_path, passwords):
    try:
        with open(input_path, 'rb') as infile:
            reader = PyPDF2.PdfReader(infile)
            if reader.is_encrypted:
                for password in passwords:
                    try:
                        reader.decrypt(password)
                        writer = PyPDF2.PdfWriter()
                        for page_num in range(len(reader.pages)):
                            writer.add_page(reader.pages[page_num])
                        with open(output_path, 'wb') as outfile:
                            writer.write(outfile)
                        return True  # Successfully unlocked
                    except Exception as e:
                        logging.warning(f"Password '{password}' failed for {input_path}: {e}")
                return False  # All passwords failed
    except Exception as e:
        logging.error(f"Error unlocking PDF {input_path}: {e}")
        return False

def extract_text_from_unlocked_pdf(pdf_path):

    transactions = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    lines = page_text.strip().split('\n')
                    for line in lines:
                        match = None
                        for pattern in patterns:
                            match = re.match(pattern, line)
                            if match:
                                break
                        if match:
                            groups = match.groups()
                            if len(groups) == 4:
                                date, description, amount, trans_type = groups
                            elif len(groups) == 5:
                                date, description, amount, trans_type = groups[0], groups[2], groups[3], groups[4]
                            else:
                                logging.warning(f"Unmatched line: {line}")
                                continue
                            try:
                                date = datetime.strptime(date, '%d %b %y').date()
                            except ValueError:
                                try:
                                    date = datetime.strptime(date, '%d/%m/%Y').date()
                                except ValueError:
                                    date = datetime(2024, 1, 1).date()  # Default date if parsing fails
                            amount = float(amount.replace(',', ''))

                            type = 'expense' if trans_type in ['D', 'Dr'] else 'income'
                            category = categorize_description(description)
                            transactions.append({
                                'Date': date,
                                'Description': description,
                                'Amount': amount,
                                'Type': type,
                                'Category': category
                            })
    except Exception as e:
        logging.error(f"Error processing PDF {pdf_path}: {e}")
    return transactions

# Main logic
def main():
    output_directory = os.path.join(pdf_directory, 'output')
    os.makedirs(output_directory, exist_ok=True)
    passwords = ["XXX", "XX", "XXX", "XXX"]

    all_transactions = []

    for filename in os.listdir(pdf_directory):
        if filename.endswith(".pdf"):
            pdf_path = os.path.join(pdf_directory, filename)
            output_pdf_path = os.path.join(output_directory, f"unlocked_{filename}")

            unlocked = unlock_pdf(pdf_path, output_pdf_path, passwords)
            if not unlocked:
                logging.warning(f"Could not unlock {filename} with provided passwords.")
                continue

            transactions = extract_text_from_unlocked_pdf(output_pdf_path)
            all_transactions.extend(transactions)

    if all_transactions:
        df = pd.DataFrame(all_transactions)

        expense_summary = df[df['Type'] == 'expense'].groupby('Category')['Amount'].sum().reset_index()
        expense_summary = expense_summary.rename(columns={'Amount': 'Total Amount (Expense)'})

        income_summary = df[df['Type'] == 'income'].groupby('Category')['Amount'].sum().reset_index()
        income_summary = income_summary.rename(columns={'Amount': 'Total Amount (Income)'})

        output_excel_path = os.path.join(pdf_directory, "transaction_summary.xlsx")
        with pd.ExcelWriter(output_excel_path) as writer:
            df.to_excel(writer, sheet_name='Transactions', index=False)
            expense_summary.to_excel(writer, sheet_name='Expense Summary', index=False)
            income_summary.to_excel(writer, sheet_name='Income Summary', index=False)

        logging.info(f"Data has been exported to {output_excel_path}")
    else:
        logging.info("No transactions were processed.")

if __name__ == "__main__":
    main()
