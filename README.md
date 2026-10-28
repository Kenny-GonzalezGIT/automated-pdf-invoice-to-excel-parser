# automated-pdf-invoice-to-excel-parser[README.md](https://github.com/user-attachments/files/34102450/README.md)
> **WINDOWS EXECUTABLE — FLAT $25:** Non-technical users can purchase the ready-to-run Windows Executable (.exe) version by transferring **$25 directly via Cash App to `$KennyGonzalezR` or via Venmo to `@Kenny-Gonzalez-68`. Contact Cozypanda64@gmail.com with your payment reference to receive the download. No Python installation is needed for the packaged executable.
>
> **Publisher setup:** Replace both payment placeholders and the delivery contact before publishing. Build and test the executable on Windows before accepting orders. This source download does not include an executable.

# PDF Invoice Parser for Windows — Batch Invoice to Excel Automation

Stop copying invoice numbers, dates, and amounts into Excel one PDF at a time. This Python PDF invoice parser processes a folder of invoices and organizes the results into a clean Excel spreadsheet, helping office workers, bookkeepers, administrative assistants, and accounts payable teams reclaim time spent on repetitive data entry.

For offices handling recurring invoice batches, automated PDF-to-Excel extraction can save hours of manual copy-pasting. Actual savings depend on invoice volume, layout consistency, and the review required. Instead of opening each PDF and retyping its details, run one command and review a single workbook.

## Features

- Batch PDF invoice processing with optional subfolder scanning.
- Invoice Number, Date, and Total Amount Due extraction using pypdf.
- Pandas Excel export with styled headers, filters, frozen headings, and numeric totals.
- Source filenames, currency labels, and review notes for traceability.
- Per-file error handling so damaged or password-protected PDFs do not stop the batch.
- Local processing: the script does not upload invoice contents to a cloud service.
- Explicit overwrite protection and preservation of leading zeros in invoice numbers.

## Install on Windows

Install Python 3.10 or newer, then open PowerShell in this project folder:

```powershell
py -m pip install -r requirements.txt
```

## Convert PDF invoices to Excel

```powershell
py invoice_parser.py "C:\Invoices"
```

The workbook is saved as `C:\Invoices\invoices.xlsx`. For a custom destination and subfolders:

```powershell
py invoice_parser.py "C:\Invoices" --recursive --output "C:\Reports\invoices.xlsx"
```

Add `--overwrite` to replace an existing workbook. Close the workbook in Excel before replacing it. Use `py invoice_parser.py --help` for all options.

## Supported invoice layouts

Each PDF should contain one invoice with selectable text and clearly labeled fields, for example:

```text
Invoice Number: INV-00123
Date: 2026-10-28
Total Amount Due: $1,250.00
```

Labels and values may also appear on adjacent lines. The parser recognizes `Invoice No.` and `Invoice #`. Customize `LABELS` in `invoice_parser.py` to match additional vendor labels. It does not infer arbitrary table layouts or guarantee extraction from every invoice template.

Scanned image-only PDFs need OCR before processing. Password-protected files must be decrypted first. Dates remain literal text to avoid misinterpreting day/month order. Amounts support decimal points and comma thousands separators; comma-decimal amounts and other unsupported formats are flagged for review. Currency symbols are preserved as printed; `$` does not identify a specific country's dollar. Do not aggregate mixed currencies without conversion.

## Review the results

The worksheet includes Source File, Invoice Number, Date, Total Amount Due, Currency, Status, and Notes. Missing or conflicting fields stay blank and receive a review note. Always check extracted results against the originals before accounting use; `Extracted` means a label match was found, not that the invoice was independently verified.

Exit codes: `0` = workbook created with all rows extracted; `1` = processing/output failure or no PDFs; `2` = argument error, or workbook created with rows requiring review. Read the console message to distinguish these cases.

## Build the optional Windows executable

Run these commands on Windows after testing the Python script:

```powershell
py -m pip install pyinstaller
py -m PyInstaller --onefile --name invoice_parser invoice_parser.py
.\dist\invoice_parser.exe "C:\Invoices" --output "C:\Reports\invoices.xlsx"
```

The executable uses the same command-line options as the script. Test the packaged executable on a clean Windows machine before distributing it. A separate graphical interface is not included.
