# Pharmacy Quotation Automation

A Python tool that automates the generation of medical supply quotations from purchase requisition documents. It matches requested items against a historical price database, applies VAT, and produces a formatted Excel quotation ready for review or submission.

## Features

- **Multi-format input** — accepts PDF (text-based or scanned via OCR) and Excel requisition files
- **Smart price matching** — fuzzy matching with confidence scoring handles spelling variations and abbreviations
- **VAT tracking** — extracts and applies VAT rates from historical purchase data
- **Structured Excel output** — three-sheet workbook:
  - `Quotation` — clean sheet for sending to suppliers
  - `Internal Review` — confidence scores and match details for internal use
  - `Items to Price` — flagged items requiring manual pricing
- **Flexible parsing** — multiple strategies handle numbered lists, pipe-separated tables, and freeform text

## How It Works

```
Requisition file (PDF / Excel)
         │
         ▼
   Item extraction
  (text → OCR fallback)
         │
         ▼
  Fuzzy price lookup
  against price database
         │
         ▼
  Excel quotation output
  (clean + internal sheets)
```

## Usage

```bash
python pharmacy_automation_v3.py <request_file> <price_database.xlsx> [output.xlsx]
```

**Examples:**

```bash
# PDF requisition
python pharmacy_automation_v3.py requisition.pdf price_database.xlsx quotation.xlsx

# Excel requisition
python pharmacy_automation_v3.py sickbay_request.xlsx price_database.xlsx

# Programmatic use
from pharmacy_automation_v3 import process_requisition

process_requisition(
    request_path="requisition.pdf",
    price_db_path="price_database.xlsx",
    output_path="quotation_output.xlsx"
)
```

## Setup

```bash
pip install -r requirements.txt
```

For OCR support on scanned PDFs, also install:

```bash
pip install pdf2image pytesseract
# macOS: brew install tesseract poppler
# Ubuntu: apt install tesseract-ocr poppler-utils
# Windows: install Tesseract from https://github.com/UB-Mannheim/tesseract/wiki
```

## Price Database Format

The price database should be an Excel file with columns in any order containing:
- Item description / name
- Quantity
- Unit price
- Total amount

The loader automatically detects columns by verifying the `qty × unit_price ≈ total` relationship, so exact column headers are not required.

## Output Example

| No. | Description | Unit | Qty | Unit Price | Amount | VAT | Gross Amount |
|-----|-------------|------|-----|-----------|--------|-----|-------------|
| 1   | Paracetamol 500mg | PCS | 100 | 2.50 | 250.00 | 12.50 | 262.50 |
| 2   | Bandage 10cm | PCS | 50  | 4.00 | 200.00 | 10.00 | 210.00 |

## Project Structure

```
pharmacy_automation_v3.py   # Main script
requirements.txt            # Python dependencies
sample_data/                # Example input files
```

## Dependencies

| Package | Purpose |
|---------|---------|
| `pandas` | Excel/CSV parsing |
| `pdfplumber` | PDF text extraction |
| `fuzzywuzzy` | Fuzzy string matching |
| `openpyxl` | Excel output generation |
| `pdf2image` + `pytesseract` | OCR for scanned PDFs (optional) |
