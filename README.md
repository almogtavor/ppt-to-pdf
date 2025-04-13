# PPT/PDF to Multi-Slide PDF Converter

A tool to convert PowerPoint presentations and PDFs into multi-slide PDFs with customizable layouts.

## Features

- Convert PowerPoint (.ppt, .pptx) and PDF files to multi-slide PDFs
- Customize layout with adjustable slides per row, gaps, and margins
- Combine multiple files into a single PDF
- Option to start each PDF's slides on a new page
- Support for right-to-left (RTL) layout for languages like Arabic and Hebrew
- Modern web interface with drag-and-drop support
- Layout page showing the order of PDFs in the combined output

## Requirements

- Python 3.7+
- Poppler (for PDF conversion)
- Ghostscript (gswin64c) (for PDF processing)
- Tesseract OCR (only required if using the OCR option)

### Installing Dependencies

1. Install Python packages:
```bash
pip install -r requirements.txt
```

2. Install Poppler:
   - Windows: Download from [poppler releases](https://github.com/oschwartz10612/poppler-windows/releases/)
   - Extract to a location (e.g., `C:\Program Files\poppler-23.11.0`)
   - Add the bin folder to your PATH (e.g., `C:\Program Files\poppler-23.11.0\Library\bin`)

3. Install Ghostscript:
   - Download from [Ghostscript releases](https://github.com/ArtifexSoftware/ghostpdl-downloads/releases)
   - Run the installer
   - Make sure to check "Add to PATH" during installation
   - Restart your terminal/PowerShell window for the PATH changes to take effect

4. Install Tesseract OCR (only needed if using OCR option):
   - Visit [Tesseract OCR GitHub](https://github.com/UB-Mannheim/tesseract/wiki)
   - Download the installer for your system (64-bit Windows: `tesseract-ocr-w64-setup-v5.3.3.20231005.exe`)
   - Run the installer and follow these steps:
     - Accept the license agreement
     - Choose "Install for all users" if you have admin rights
     - Make sure to check the box that says "Add to PATH" during installation
     - Complete the installation
   - Restart your terminal/PowerShell window for the PATH changes to take effect

## Usage

### Web Interface

1. Start the Streamlit app:
```bash
streamlit run app.py
```

2. Open your browser and go to `http://localhost:8501`

3. Use the web interface to:
   - Upload files
   - Adjust layout settings
   - Choose output options
   - Download the converted PDF

### Command Line

```bash
py -3.9 -m venv venv 
# Convert a single file
python main.py input.pdf output.pdf --slides_per_row 3

# Convert all files in a directory
python main.py input_folder output_folder --slides_per_row 3

# Combine multiple files into one PDF
python main.py input_folder combined.pdf --single_file --slides_per_row 3

# Combine files with each PDF starting on a new page
python main.py input_folder combined.pdf --single_file --slides_per_row 3 --no_new_page
```

## Options

- `--slides_per_row`: Number of slides per row (default: 2)
- `--gap`: Space between slides in points (default: 10)
- `--margin`: Margin on sides and bottom in points (default: 20)
- `--top_margin`: Margin at the top in points (default: 0)
- `--single_file`: Combine all slides into a single PDF
- `--no_new_page`: Disable starting each PDF's slides on a new page (only applies when --single_file is used)
- `--rtl`: Enable right-to-left layout for RTL languages
- `--ocr`: Add searchable text layer to the PDF (enabled by default)

## Notes

- When combining files into a single PDF, each PDF's slides will start on a new page by default
- The layout page at the beginning of the combined PDF shows the order of PDFs
- For best results, use similar-sized slides
- Adjust the margins and gaps to optimize the layout
- The "Slides per Row" setting affects the size of each slide

## License

MIT License