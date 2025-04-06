# PPT/PDF to Multi-Slide PDF Converter

A tool to convert PowerPoint presentations and PDFs into multi-slide PDFs with customizable layouts.

## Features

- Convert PowerPoint (.ppt, .pptx) and PDF files to multi-slide PDFs
- Customize layout with adjustable slides per row, gaps, and margins
- Combine multiple files into a single PDF
- Option to start each PDF's slides on a new page
- Modern web interface with drag-and-drop support
- Layout page showing the order of PDFs in the combined output

## Requirements

- Python 3.7+
- Poppler (for PDF conversion)
- Microsoft PowerPoint (for PPT/PPTX conversion on Windows)

## Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Install Poppler:
   - Windows: Download from [poppler-windows releases](https://github.com/oschwartz10612/poppler-windows/releases/) and add to PATH
   - Linux: `sudo apt-get install poppler-utils`
   - macOS: `brew install poppler`

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

## Notes

- When combining files into a single PDF, each PDF's slides will start on a new page by default
- The layout page at the beginning of the combined PDF shows the order of PDFs
- For best results, use similar-sized slides
- Adjust the margins and gaps to optimize the layout
- The "Slides per Row" setting affects the size of each slide

## License

MIT License