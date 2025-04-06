# PPT/PDF to Multi-Slide PDF Converter

A tool to convert PowerPoint presentations and PDFs into PDFs with multiple slides per page. This tool is particularly useful for creating handouts or study materials from presentations.

## Features

- Convert PowerPoint (.ppt, .pptx) and PDF files to multi-slide PDFs
- Customize layout with adjustable slides per row, gaps, and margins
- Process multiple files at once
- Option to combine all slides into a single PDF file
- Layout page showing the order of PDFs in the combined output
- Option to start each PDF's slides on a new page
- Web-based interface with drag-and-drop support
- Command-line interface for batch processing
- Support for non-ASCII characters in file paths

## Requirements

- Python 3.6 or higher
- Microsoft PowerPoint (for PPT/PPTX conversion on Windows)
- Poppler (for PDF conversion)
- Required Python packages:
  - flask
  - reportlab
  - pdf2image
  - comtypes (Windows only)
  - pillow

## Installation

1. Install Python dependencies:
```bash
pip install flask reportlab pdf2image pillow
```

2. On Windows, install comtypes for PowerPoint conversion:
```bash
pip install comtypes
```

3. Install Poppler:
   - Windows: Download and add to PATH
   - Linux: `sudo apt-get install poppler-utils`
   - macOS: `brew install poppler`

## Usage

### Web Interface

1. Run the Flask application:
```bash
python app.py
```

2. Open your browser and navigate to `http://localhost:5000`

3. Features:
   - Drag and drop files or click to browse
   - Adjust layout settings:
     - Slides per row
     - Gap between slides
     - Page margins
     - Top margin
   - Choose between:
     - Separate PDFs (zipped together)
     - Single combined PDF (all slides in one file)
   - Option to start each PDF on a new page
   - Layout page showing PDF order
   - View progress during conversion
   - Automatic download of results

### Command Line

```bash
python main.py input_path output_path [options]
```

Options:
- `--slides_per_row`: Number of slides per row (default: 2)
- `--gap`: Space between slides in points (default: 10)
- `--margin`: Margin on sides and bottom in points (default: 20)
- `--top_margin`: Margin at the top in points (default: 0)
- `--single_file`: Combine all slides into a single PDF file
- `--new_page_per_pdf`: Start each PDF's slides on a new page (only applies in single_file mode)

Examples:
```bash
# Convert a single file
python main.py presentation.pptx output.pdf --slides_per_row 3

# Process a directory of files
python main.py input_folder output_folder --slides_per_row 2 --gap 15

# Combine multiple files into a single PDF
python main.py input_folder combined.pdf --single_file --slides_per_row 3

# Combine files with each PDF starting on a new page
python main.py input_folder combined.pdf --single_file --new_page_per_pdf --slides_per_row 3
```

## Notes

- When using the single file option, a layout page is automatically added showing the order of PDFs
- The `new_page_per_pdf` option ensures each PDF's slides start on a new page
- The tool maintains aspect ratios and optimizes slide sizes for readability
- For best results with PowerPoint files, ensure Microsoft PowerPoint is installed
- The web interface supports multiple file uploads and provides visual feedback during processing

## License

MIT License