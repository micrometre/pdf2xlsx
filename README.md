# Pdf to Xlsx examples

Convert tables from PDF files to Excel (XLSX) format.

## Installation

1. Clone or navigate to this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Ensure Java is installed (required for `tabula-py`):
   ```bash
   java -version
   ```

## Usage

Convert all PDFs in the `pdfs/` directory:
```bash
python main.py
```

Convert a single PDF file:
```bash
python main.py pdfs/example.pdf -o output.xlsx
```

Convert all PDFs to a custom output directory:
```bash
python main.py pdfs -o custom_output/
```

Output files will be named with `.xlsx` extension. Each table in a PDF is saved as a separate sheet in the Excel file.

## Requirements

- Python 3.7+
- Java (for PDF table extraction)
- Dependencies listed in `requirements.txt`
