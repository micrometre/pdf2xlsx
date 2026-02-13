import argparse
import sys
from pathlib import Path

import pandas as pd

try:
	from tabula import read_pdf
except Exception as e:
	raise ImportError(
		"tabula-py is not installed or cannot be imported. Install it with `pip install tabula-py openpyxl` and ensure Java is available on your PATH."
	) from e


def convert_pdf_to_xlsx(pdf_path: Path, xlsx_path: Path) -> bool:
	try:
		tables = read_pdf(str(pdf_path), pages="all", multiple_tables=True)
	except Exception as e:
		print(f"Failed to read PDF {pdf_path}: {e}")
		return False

	if not tables:
		print(f"No tables found in {pdf_path}")
		return False

	try:
		with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
			if isinstance(tables, list):
				for i, df in enumerate(tables, start=1):
					sheet_name = f"table_{i}"
					df.to_excel(writer, sheet_name=sheet_name, index=False)
			else:
				tables.to_excel(writer, sheet_name="Sheet1", index=False)
	except Exception as e:
		print(f"Failed to write Excel {xlsx_path}: {e}")
		return False

	print(f"Wrote {xlsx_path}")
	return True


def main(argv=None):
	parser = argparse.ArgumentParser(description="Convert PDF tables to XLSX")
	parser.add_argument("input", nargs="?", default="pdfs", help="PDF file or directory (default: pdfs)")
	parser.add_argument("-o", "--output", help="Output XLSX file or directory")

	args = parser.parse_args(argv)
	input_path = Path(args.input)

	if input_path.is_dir():
		out_dir = Path(args.output) if args.output else Path("xlsx_outputs")
		out_dir.mkdir(parents=True, exist_ok=True)
		pdf_files = list(input_path.glob("*.pdf"))
		if not pdf_files:
			print(f"No PDF files found in {input_path}")
			return 1
		for pdf in pdf_files:
			out_file = out_dir / (pdf.stem + ".xlsx")
			convert_pdf_to_xlsx(pdf, out_file)
	else:
		pdf = input_path
		if not pdf.exists():
			print(f"Input file not found: {pdf}")
			return 1
		if args.output:
			out_file = Path(args.output)
		else:
			out_file = pdf.with_suffix(".xlsx")
		convert_pdf_to_xlsx(pdf, out_file)


if __name__ == "__main__":
	raise SystemExit(main())