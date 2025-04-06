import os
import sys
import math
import tempfile
import shutil
import argparse
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from PIL import Image
import pathlib
import glob


def convert_ppt_to_images(ppt_file, output_dir):
    """
    Convert each slide of the given PowerPoint file to an image.
    This function uses the COM interface (via comtypes) to work on Windows
    with Microsoft PowerPoint installed.
    """
    try:
        import comtypes.client
    except ImportError:
        raise Exception("The 'comtypes' package is required on Windows to convert PPT slides to images.")

    # Convert paths to absolute paths using pathlib for better Unicode handling
    ppt_file = str(pathlib.Path(ppt_file).resolve())
    output_dir = str(pathlib.Path(output_dir).resolve())

    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(ppt_file, WithWindow=False)
    slide_count = len(presentation.Slides)
    image_paths = []

    for slide in presentation.Slides:
        img_path = os.path.join(output_dir, f"slide_{slide.SlideIndex}.png")
        slide.Export(img_path, "PNG")
        image_paths.append(img_path)

    presentation.Close()
    powerpoint.Quit()
    return image_paths


def convert_pdf_to_images(pdf_file, output_dir):
    """
    Convert each page of the given PDF file (assumed to be a PowerPoint export)
    to an image using pdf2image. Poppler must be installed for this to work.
    """
    try:
        from pdf2image import convert_from_path
    except ImportError:
        raise Exception("The 'pdf2image' package is required to convert PDF pages to images.")

    # Convert paths to absolute paths using pathlib for better Unicode handling
    pdf_file = str(pathlib.Path(pdf_file).resolve())
    output_dir = str(pathlib.Path(output_dir).resolve())

    # Verify the file exists and is accessible
    if not os.path.exists(pdf_file):
        raise FileNotFoundError(f"PDF file not found: {pdf_file}")
    if not os.access(pdf_file, os.R_OK):
        raise PermissionError(f"Cannot read PDF file: {pdf_file}")

    try:
        # Convert all pages in the PDF to PIL Image objects
        pages = convert_from_path(pdf_file)
    except Exception as e:
        raise Exception(f"Failed to convert PDF to images: {str(e)}")

    image_paths = []
    for i, page in enumerate(pages):
        img_path = os.path.join(output_dir, f"slide_{i + 1}.png")
        page.save(img_path, 'PNG')
        image_paths.append(img_path)
    return image_paths


def convert_file_to_images(file_path, output_dir):
    """
    Determine the file type (PPT/PPTX or PDF) and convert its slides/pages to images.
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.ppt', '.pptx']:
        return convert_ppt_to_images(file_path, output_dir)
    elif ext == '.pdf':
        return convert_pdf_to_images(file_path, output_dir)
    else:
        raise Exception("Unsupported file type. Only .ppt, .pptx, and .pdf are supported.")


def create_pdf_from_images(image_paths, output_pdf, slides_per_row=2, gap=10, margin=20, top_margin=0):
    """
    Arrange slide images onto A4 pages, maximizing the use of space while ensuring readability.
    • slides_per_row: Number of slides in each row.
    • gap: Space (in points) between slides.
    • margin: Margin (in points) on the sides and bottom of the page.
    • top_margin: Margin (in points) at the top of the page.
    """
    a4_width, a4_height = A4

    if not image_paths:
        raise Exception("No slide images found to create the PDF.")

    # Use the first image to determine the original slide dimensions.
    with Image.open(image_paths[0]) as im:
        orig_w, orig_h = im.size
    aspect_ratio = orig_w / orig_h  # Maintain width-to-height ratio

    # Calculate available space for slides
    available_width = a4_width - 2 * margin - (slides_per_row - 1) * gap
    available_height = a4_height - top_margin - margin  # Only bottom margin

    # Calculate maximum possible slide dimensions
    # First try fitting based on width
    slide_width = available_width / slides_per_row
    slide_height = slide_width / aspect_ratio

    # Calculate how many rows we can fit
    rows_fit = math.floor((available_height + gap) / (slide_height + gap))
    
    # If we can't fit at least one row, adjust based on height
    if rows_fit < 1:
        rows_fit = 1
        slide_height = available_height
        slide_width = slide_height * aspect_ratio
        # Check if this fits within width constraints
        if slide_width * slides_per_row + (slides_per_row - 1) * gap > available_width:
            # If not, scale down to fit width
            slide_width = available_width / slides_per_row
            slide_height = slide_width / aspect_ratio
    else:
        # Try to fit one more row if possible, but only if the content remains readable
        # Calculate the size if we add one more row
        potential_rows = rows_fit + 1
        potential_slide_height = (available_height - (potential_rows - 1) * gap) / potential_rows
        potential_slide_width = potential_slide_height * aspect_ratio
        
        # Only add the extra row if the slides don't become too small
        # We consider slides too small if they're less than 60% of the original size
        if potential_slide_width >= slide_width * 0.6:
            rows_fit = potential_rows
            slide_height = potential_slide_height
            slide_width = potential_slide_width

    # Calculate actual margins to center the content horizontally
    total_slide_width = slide_width * slides_per_row + (slides_per_row - 1) * gap
    x_margin = (a4_width - total_slide_width) / 2

    c = canvas.Canvas(output_pdf, pagesize=A4)
    total_slides = len(image_paths)
    max_slides_per_page = slides_per_row * rows_fit

    # Process slides in batches for each page
    for page_start in range(0, total_slides, max_slides_per_page):
        page_slides = image_paths[page_start:page_start + max_slides_per_page]
        rows = math.ceil(len(page_slides) / slides_per_row)

        for idx, image_path in enumerate(page_slides):
            col = idx % slides_per_row
            row = idx // slides_per_row
            x = x_margin + col * (slide_width + gap)
            # y is calculated from the top of the page (ReportLab's origin is bottom-left)
            y = a4_height - top_margin - (row + 1) * slide_height - row * gap
            c.drawImage(image_path, x, y, width=slide_width, height=slide_height)
        c.showPage()  # End the current page
    c.save()


def process_file(input_path, output_path, slides_per_row=2, gap=10, margin=20, top_margin=0):
    """
    Process a single file and convert it to PDF with the specified layout.
    """
    temp_dir = tempfile.mkdtemp()
    try:
        image_paths = convert_file_to_images(input_path, temp_dir)
        create_pdf_from_images(image_paths, output_path, slides_per_row, gap, margin, top_margin)
    finally:
        shutil.rmtree(temp_dir)


def process_directory(input_dir, output_dir, slides_per_row=2, gap=10, margin=20, top_margin=0):
    """
    Process all PDF and PowerPoint files in the input directory.
    Creates corresponding output files in the output directory.
    """
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Get all supported files in the input directory
    supported_extensions = ['.pdf', '.ppt', '.pptx']
    input_files = []
    for ext in supported_extensions:
        input_files.extend(glob.glob(os.path.join(input_dir, f'*{ext}')))
    
    if not input_files:
        print(f"No supported files found in {input_dir}")
        return
    
    # Process each file
    for input_file in input_files:
        # Create output filename by replacing extension with .pdf
        filename = os.path.basename(input_file)
        output_file = os.path.join(output_dir, os.path.splitext(filename)[0] + '.pdf')
        
        try:
            print(f"Processing {filename}...")
            process_file(input_file, output_file, slides_per_row, gap, margin, top_margin)
            print(f"Created {output_file}")
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}", file=sys.stderr)


def process_files(input_paths, output_path, slides_per_row=2, gap=10, margin=20, top_margin=0, single_file=False):
    """
    Process multiple files and convert them to PDF(s) with the specified layout.
    
    Args:
        input_paths: List of input file paths
        output_path: Output file path or directory
        slides_per_row: Number of slides per row
        gap: Space between slides
        margin: Margin on sides and bottom
        top_margin: Margin at the top
        single_file: If True, combine all slides into a single PDF file
    """
    if not input_paths:
        raise Exception("No input files provided")

    if single_file:
        # For single file output, process all files and combine their slides
        all_image_paths = []
        temp_dirs = []
        
        try:
            # Process each file and collect all image paths
            for input_path in input_paths:
                temp_dir = tempfile.mkdtemp()
                temp_dirs.append(temp_dir)
                image_paths = convert_file_to_images(input_path, temp_dir)
                all_image_paths.extend(image_paths)
            
            # Create single PDF with all slides
            create_pdf_from_images(all_image_paths, output_path, slides_per_row, gap, margin, top_margin)
        finally:
            # Clean up temporary directories
            for temp_dir in temp_dirs:
                shutil.rmtree(temp_dir)
    else:
        # Process each file separately
        if os.path.isdir(output_path):
            # Output is a directory
            for input_path in input_paths:
                filename = os.path.basename(input_path)
                file_output_path = os.path.join(output_path, os.path.splitext(filename)[0] + '.pdf')
                process_file(input_path, file_output_path, slides_per_row, gap, margin, top_margin)
        else:
            # Output is a single file path
            if len(input_paths) > 1:
                raise Exception("Multiple input files require an output directory when single_file is False")
            process_file(input_paths[0], output_path, slides_per_row, gap, margin, top_margin)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert PowerPoint presentations or PDFs to PDFs with multiple slides per page."
    )
    parser.add_argument("input", help="Path to input file or directory")
    parser.add_argument("output", help="Path for output file or directory")
    parser.add_argument("--slides_per_row", type=int, default=2, help="Number of slides per row (default: 2)")
    parser.add_argument("--gap", type=int, default=10, help="Gap (in points) between slides (default: 10)")
    parser.add_argument("--margin", type=int, default=20, help="Margin (in points) on the sides and bottom (default: 20)")
    parser.add_argument("--top_margin", type=int, default=0, help="Margin (in points) at the top of the page (default: 0)")
    parser.add_argument("--single_file", action="store_true", help="Combine all slides into a single PDF file")
    args = parser.parse_args()

    try:
        # Get list of input files
        if os.path.isdir(args.input):
            # Get all supported files in the input directory
            supported_extensions = ['.pdf', '.ppt', '.pptx']
            input_files = []
            for ext in supported_extensions:
                input_files.extend(glob.glob(os.path.join(args.input, f'*{ext}')))
        else:
            input_files = [args.input]

        process_files(
            input_files,
            args.output,
            args.slides_per_row,
            args.gap,
            args.margin,
            args.top_margin,
            args.single_file
        )
        print(f"Successfully created PDF(s) in: {args.output}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
