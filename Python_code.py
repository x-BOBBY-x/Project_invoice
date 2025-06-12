import pdfplumber
import pandas as pd
import re
import os

input_folder = "Files"  
output_file = "All_Invoice_Details.xlsx"

def extract_full_invoice_details_safe(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])

    patterns = {
        "Invoice Number": re.search(r"Invoice (?:Number|No)[^\n:]*[:\s]+([A-Z0-9\-#]+)", full_text, re.IGNORECASE),
        "Order ID": re.search(r"(?:Order ID|Order Number)[^\n:]*[:\s]+([A-Z0-9\-]+)", full_text, re.IGNORECASE),
        "Order Date": re.search(r"Order Date[^\n:]*[:\s]+([0-9]{2}[-/][0-9]{2}[-/][0-9]{4})", full_text),
        "Invoice Date": re.search(r"Invoice Date[^\n:]*[:\s]+([0-9]{2}[-/][0-9]{2}[-/][0-9]{4})", full_text),
        "Billing Name": re.search(r"Billing Address\s*[:\s]*([^\n]+)", full_text),
        "Shipping Name": re.search(r"Shipping Address\s*[:\s]*([^\n]+)", full_text),
        "PAN": re.search(r"PAN(?: No)?:\s*([A-Z0-9]+)", full_text),
        "GSTIN": re.search(r"GST(?:IN| Registration No)?(?: No)?[:\s]*([A-Z0-9]+)", full_text),
        "Seller": re.search(r"Sold By\s*:?(.+?)(?:\n|,)", full_text, re.DOTALL),
        "Total Amount": re.search(r"Grand Total.*?₹?\s*([0-9,]+\.\d{2})", full_text),
        "Shipping Charges": re.search(r"Shipping (?:and Handling)? Charges.*?₹?\s*([0-9,]+\.\d{2})", full_text),
        "Tax Amount": re.search(r"Tax Amount.*?₹?\s*([0-9,]+\.\d{2})", full_text),
        "Item Description": re.search(r"(?:Product Description|Description|Item)[^\n:]*[:\s]*([^\n]+)", full_text, re.IGNORECASE)
    }

    details = {"File": os.path.basename(pdf_path)}
    for key, match in patterns.items():
        if match:
            try:
                details[key] = match.group(1).strip()
            except IndexError:
                details[key] = "Not found"
        else:
            details[key] = "Not found"

    return details

pdf_files = [os.path.join(input_folder, f) for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]

invoice_data = [extract_full_invoice_details_safe(file) for file in pdf_files]

df = pd.DataFrame(invoice_data)
df.to_excel(output_file, index=False)

print(f"All invoice data extracted and saved to: {output_file}")