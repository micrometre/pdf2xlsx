# PDF to Excel Converter

Convert tables from PDF files to Excel (XLSX) format with an optional filtering feature.

## Features

- **GUI Interface**: User-friendly tkinter-based graphical interface
- **Single File or Batch Conversion**: Convert one PDF or an entire directory
- **PDF Table Extraction**: Automatically extracts tables from PDFs using Tabula
- **Excel Output**: Each table is saved as a separate sheet in Excel files
- **Row Filtering**: Optional filtering to search for specific strings across converted Excel files
- **CLI Support**: Command-line interface for automation and scripting
- **Real-time Logging**: View conversion progress in the application log

## Installation

1. Clone or navigate to this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Ensure Java 8+ is installed (required by Tabula for PDF table extraction)

## Usage

### GUI Mode (Default)

Launch the application with a simple command:

```bash
python3 gui.py
```

The GUI provides:

#### Single File Conversion
1. Select **"Single File"** mode
2. Click **"Browse..."** under Input and select a PDF file
3. Click **"Browse..."** under Output and specify where to save the Excel file
4. Click **"Convert"**

#### Batch Directory Conversion
1. Select **"Directory (Batch)"** mode
2. Click **"Browse..."** under Input and select a folder containing PDF files
3. Click **"Browse..."** under Output and select an output directory
4. Click **"Convert"** - each PDF will be converted to its own Excel file

#### Optional Filtering

After conversion, you can automatically filter and consolidate results:

1. Enable the **"Enable Filtering after Conversion"** checkbox
2. Enter a search string in the **"Search String"** field (e.g., "Barrier gate opened")
3. Click **"Convert"**
4. After conversion completes, a `filtered_results.xlsx` file will be created containing only rows that match your search string

**Note**: The search is case-sensitive and requires an exact string match.

### CLI Mode

For command-line usage:

```bash
python3 gui.py --cli [input] -o [output]
```

Examples:

```bash
# Convert single PDF
python3 gui.py --cli example.pdf -o output.xlsx

# Convert all PDFs in a directory
python3 gui.py --cli pdfs/ -o xlsx_outputs/
```

## How It Works

### Conversion Process
1. Reads PDF file(s) using Tabula's table extraction
2. Extracts all tables from each page
3. Saves each table as a separate sheet in an Excel workbook
4. Automatically names sheets as `table_1`, `table_2`, etc.

### Filtering Process
1. Searches through all Excel files in the output directory
2. Scans each row for the specified search string
3. Consolidates all matching rows into a single Excel file
4. Provides a count of files processed and rows found

## Requirements

- Python 3.7+
- Java 8 or higher (required for Tabula)
- Dependencies listed in `requirements.txt`:
  - pandas: Data manipulation
  - tabula-py: PDF table extraction
  - openpyxl: Excel file creation and manipulation

## Output

- **Converted Files**: Named as `[original_filename].xlsx` containing extracted tables
- **Filtered Results**: Named as `filtered_results.xlsx` (if filtering is enabled)
- Each table from the PDF becomes a separate sheet in the Excel file
- Sheet names follow the pattern `table_1`, `table_2`, etc.

## Troubleshooting

- **Java not found**: Ensure Java 8+ is installed and in your system PATH
- **No tables found**: The PDF may not contain structured tables, or they may be in an image format
- **Filtering returns no results**: Check that the search string exactly matches the cell content (case-sensitive)

1. Clone or navigate to this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Convert all PDFs in the `test_pdfs/` directory:

Convert a single PDF file:

Convert all PDFs to a custom output directory:

Output files will be named with `.xlsx` extension. Each table in a PDF is saved as a separate sheet in the Excel file.

## Requirements

- Python 3.7+
- Dependencies listed in `requirements.txt` (pandas, pdfplumber, openpyxl)
- windows(java)