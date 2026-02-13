import argparse
import sys
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading

import pandas as pd

def get_java_path():
    """Get Java executable path, checking bundled Java first"""
    
    # If running as PyInstaller bundle
    if hasattr(sys, '_MEIPASS'):
        base_path = Path(sys._MEIPASS)
        # Check for bundled Java in common locations
        java_candidates = [
            base_path / "java" / "bin" / "java.exe",
            base_path / "jre" / "bin" / "java.exe",
            base_path / "jdk" / "bin" / "java.exe",
            base_path / "runtime" / "bin" / "java.exe",
        ]
        
        for java_path in java_candidates:
            if java_path.exists():
                return str(java_path)
    
    # Fall back to system Java
    return "java"

import os
from pathlib import Path


java_path = get_java_path()
if java_path:
    os.environ['JAVA_HOME'] = str(Path(java_path).parent.parent)
    os.environ['PATH'] = str(Path(java_path).parent) + os.pathsep + os.environ['PATH']

try:
    from tabula import read_pdf
except Exception as e:
    raise ImportError(
        "Java is required but not found. Please install Java 8+ or bundle it with the application."
    ) from e



def convert_pdf_to_xlsx(pdf_path: Path, xlsx_path: Path, callback=None) -> bool:
    try:
        tables = read_pdf(str(pdf_path), pages="all", multiple_tables=True)
    except Exception as e:
        error_msg = f"Failed to read PDF {pdf_path}: {e}"
        print(error_msg)
        if callback:
            callback(error_msg)
        return False

    if not tables:
        error_msg = f"No tables found in {pdf_path}"
        print(error_msg)
        if callback:
            callback(error_msg)
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
        error_msg = f"Failed to write Excel {xlsx_path}: {e}"
        print(error_msg)
        if callback:
            callback(error_msg)
        return False

    success_msg = f"✓ Wrote {xlsx_path}"
    print(success_msg)
    if callback:
        callback(success_msg)
    return True


class PDFConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter")
        self.root.geometry("600x500")
        self.root.resizable(True, True)
        
        # Set icon if available (optional)
        try:
            self.root.iconbitmap(default='icon.ico')
        except:
            pass
        
        # Variables
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.is_directory_mode = tk.BooleanVar(value=False)
        self.conversion_running = False
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="PDF to Excel Converter", 
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Mode selection
        mode_frame = ttk.LabelFrame(main_frame, text="Mode", padding="10")
        mode_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Radiobutton(mode_frame, text="Single File", 
                       variable=self.is_directory_mode, 
                       value=False).grid(row=0, column=0, padx=(0, 20))
        ttk.Radiobutton(mode_frame, text="Directory (Batch)", 
                       variable=self.is_directory_mode, 
                       value=True).grid(row=0, column=1)
        
        # Input selection
        input_frame = ttk.LabelFrame(main_frame, text="Input", padding="10")
        input_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        input_frame.columnconfigure(1, weight=1)
        
        ttk.Entry(input_frame, textvariable=self.input_path).grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(input_frame, text="Browse...", command=self.browse_input).grid(row=0, column=2)
        
        # Output selection
        output_frame = ttk.LabelFrame(main_frame, text="Output", padding="10")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(1, weight=1)
        
        ttk.Entry(output_frame, textvariable=self.output_path).grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(output_frame, text="Browse...", command=self.browse_output).grid(row=0, column=2)
        
        # Control buttons
        control_frame = ttk.Frame(main_frame)
        control_frame.grid(row=4, column=0, columnspan=3, pady=(0, 10))
        
        self.convert_btn = ttk.Button(control_frame, text="Convert", command=self.start_conversion)
        self.convert_btn.grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(control_frame, text="Clear", command=self.clear_fields).grid(row=0, column=1)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Log area
        log_frame = ttk.LabelFrame(main_frame, text="Conversion Log", padding="10")
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        # Scrollbar for log
        self.log_text = tk.Text(log_frame, height=12, width=70, wrap=tk.WORD)
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
    def browse_input(self):
        if self.is_directory_mode.get():
            path = filedialog.askdirectory(title="Select Directory with PDF files")
        else:
            path = filedialog.askopenfilename(
                title="Select PDF file",
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
        if path:
            self.input_path.set(path)
            self.log(f"Selected input: {path}")
            
            # Auto-suggest output
            if not self.output_path.get():
                self.suggest_output_path()
    
    def browse_output(self):
        if self.is_directory_mode.get():
            path = filedialog.askdirectory(title="Select Output Directory")
        else:
            path = filedialog.asksaveasfilename(
                title="Save Excel file as",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
        if path:
            self.output_path.set(path)
            self.log(f"Selected output: {path}")
    
    def suggest_output_path(self):
        input_path = Path(self.input_path.get())
        if self.is_directory_mode.get():
            suggested = input_path.parent / f"{input_path.name}_excel"
        else:
            suggested = input_path.with_suffix(".xlsx")
        
        self.output_path.set(str(suggested))
        self.log(f"Auto-suggested output: {suggested}")
    
    def clear_fields(self):
        self.input_path.set("")
        self.output_path.set("")
        self.log_text.delete(1.0, tk.END)
        self.status_var.set("Ready")
    
    def log(self, message):
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
    
    def start_conversion(self):
        if self.conversion_running:
            return
        
        input_path = self.input_path.get().strip()
        output_path = self.output_path.get().strip()
        
        if not input_path:
            messagebox.showerror("Error", "Please select input file or directory")
            return
        
        if not output_path:
            messagebox.showerror("Error", "Please select output location")
            return
        
        # Disable controls during conversion
        self.conversion_running = True
        self.convert_btn.config(state='disabled')
        self.progress.start()
        self.status_var.set("Converting...")
        
        # Run conversion in separate thread
        thread = threading.Thread(target=self.convert, args=(input_path, output_path))
        thread.daemon = True
        thread.start()
    
    def convert(self, input_path, output_path):
        try:
            input_path = Path(input_path)
            output_path = Path(output_path)
            
            if input_path.is_dir():
                # Batch conversion
                out_dir = output_path if output_path.suffix == '' else output_path.parent
                out_dir.mkdir(parents=True, exist_ok=True)
                
                pdf_files = list(input_path.glob("*.pdf"))
                if not pdf_files:
                    self.log("❌ No PDF files found in the directory")
                    return
                
                self.log(f"Found {len(pdf_files)} PDF file(s)")
                success_count = 0
                
                for i, pdf in enumerate(pdf_files, 1):
                    self.log(f"\nProcessing ({i}/{len(pdf_files)}): {pdf.name}")
                    
                    if output_path.suffix == '':
                        out_file = out_dir / (pdf.stem + ".xlsx")
                    else:
                        out_file = output_path
                    
                    if convert_pdf_to_xlsx(pdf, out_file, self.log):
                        success_count += 1
                
                self.log(f"\n{'='*50}")
                self.log(f"✅ Batch conversion complete: {success_count}/{len(pdf_files)} successful")
                
            else:
                # Single file conversion
                if not input_path.exists():
                    self.log(f"❌ Input file not found: {input_path}")
                    return
                
                self.log(f"Converting: {input_path.name}")
                convert_pdf_to_xlsx(input_path, output_path, self.log)
                self.log(f"✅ Conversion complete")
        
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
        
        finally:
            # Re-enable controls
            self.root.after(0, self.conversion_finished)
    
    def conversion_finished(self):
        self.conversion_running = False
        self.convert_btn.config(state='normal')
        self.progress.stop()
        self.status_var.set("Ready")


def main(argv=None):
    # Check if GUI should be used
    if len(sys.argv) > 1 and sys.argv[1] in ['--cli', '-c', '--no-gui']:
        # Remove the CLI flag before parsing
        if sys.argv[1] in ['--cli', '-c', '--no-gui']:
            sys.argv.pop(1)
        return cli_main(argv)
    else:
        # Launch GUI
        root = tk.Tk()
        app = PDFConverterGUI(root)
        root.mainloop()
        return 0


def cli_main(argv=None):
    """Original command-line interface"""
    parser = argparse.ArgumentParser(description="Convert PDF tables to XLSX")
    parser.add_argument("input", nargs="?", default="pdfs", help="PDF file or directory (default: pdfs)")
    parser.add_argument("-o", "--output", help="Output XLSX file or directory")
    parser.add_argument("--cli", action="store_true", help="Force command-line mode")
    
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
    return 0


if __name__ == "__main__":
    raise SystemExit(main())