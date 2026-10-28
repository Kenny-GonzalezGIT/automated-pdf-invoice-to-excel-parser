import os
import glob
import pypdf
import pandas as pd
from datetime import datetime

def extract_invoice_data(pdf_path):
    """
    Scans a PDF invoice file and attempts to extract fundamental data points.
    """
    invoice_num = "N/A"
    invoice_date = "N/A"
    total_amount = "N/A"
    
    try:
        reader = pypdf.PdfReader(pdf_path)
        full_text = ""
        for page in reader.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"
        
        # Tokenize lines for semantic scanning
        lines = [line.strip() for line in full_text.split('\n') if line.strip()]
        
        for i, line in enumerate(lines):
            lowered = line.lower()
            
            # 1. Look for Invoice Number identifiers
            if "invoice number" in lowered or "invoice #" in lowered or "inv #" in lowered:
                parts = line.split(':') if ':' in line else line.split()
                invoice_num = parts[-1].strip()
                
            # 2. Look for Date identifiers
            if "date" in lowered and invoice_date == "N/A":
                parts = line.split(':') if ':' in line else line.split()
                invoice_date = parts[-1].strip()
                
            # 3. Look for Balance / Due Amount totals
            if "total due" in lowered or "amount due" in lowered or "total amount" in lowered:
                parts = line.split(':') if ':' in line else line.split()
                total_amount = parts[-1].strip()
                
    except Exception as e:
        print(f"Error parsing file {os.path.basename(pdf_path)}: {str(e)}")
        
    return invoice_num, invoice_date, total_amount

def batch_process_directory():
    print("="*60)
    print("Initializing Local PDF Invoice Processing Engine...")
    print("="*60)
    
    # Scan current directory for files
    pdf_files = glob.glob("*.pdf")
    if not pdf_files:
        print("ALERT: No PDF files found in the active workspace directory.")
        print("Please place this script inside the folder containing your invoices.")
        return

    dataset = []
    for pdf in pdf_files:
        print(f"Processing: {os.path.basename(pdf)}...")
        inv_num, inv_date, total = extract_invoice_data(pdf)
        dataset.append({
            "Source File": os.path.basename(pdf),
            "Invoice Number": inv_num,
            "Invoice Date": inv_date,
            "Total Amount Due": total,
            "Processed Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

    # Pack metrics matrix into structured framework
    df = pd.DataFrame(dataset)
    output_file = "Parsed_Invoice_Summary.xlsx"
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    print("="*60)
    print("SUCCESS: Data parsing sequence completed.")
    print(f"Universal asset generated: {os.path.abspath(output_file)}")
    print("="*60)

if __name__ == "__main__":
    batch_process_directory()
  
