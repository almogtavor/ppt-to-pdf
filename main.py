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
    Convert each slide of the given PowerPoint file to a JPEG image.
    Uses the COM interface (via comtypes) on Windows with Microsoft PowerPoint installed.
    The exported PNG is converted to a lower-quality JPEG to reduce file size.
    """
    try:
        import comtypes.client
    except ImportError:
        raise Exception("The 'comtypes' package is required on Windows to convert PPT slides to images.")

    ppt_file = str(pathlib.Path(ppt_file).resolve())
    output_dir = str(pathlib.Path(output_dir).resolve())

    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(ppt_file, WithWindow=False)
    image_paths = []

    for slide in presentation.Slides:
        # Export slide as PNG first
        png_path = os.path.join(output_dir, f"slide_{slide.SlideIndex}.png")
        slide.Export(png_path, "PNG")
        # Convert PNG to JPEG with compression
        jpg_path = os.path.join(output_dir, f"slide_{slide.SlideIndex}.jpg")
        with Image.open(png_path) as img:
            rgb_img = img.convert("RGB")
            rgb_img.save(jpg_path, "JPEG", quality=60, optimize=True)
        os.remove(png_path)
        image_paths.append(jpg_path)

    presentation.Close()
    powerpoint.Quit()
    return image_paths

def convert_pdf_to_images(pdf_file, output_dir):
    """
    Convert PDF pages to JPEG images using PyMuPDF.
    The JPEG quality is set to 60 to reduce file size.
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise Exception("The 'PyMuPDF' package is required to convert PDF pages to images. Install it with: pip install pymupdf")

    pdf_file = str(pathlib.Path(pdf_file).resolve())
    output_dir = str(pathlib.Path(output_dir).resolve())

    print(f"\nConverting PDF: {pdf_file}")
    print(f"Output directory: {output_dir}")

    if not os.path.exists(pdf_file):
        raise FileNotFoundError(f"PDF file not found: {pdf_file}")
    if not os.access(pdf_file, os.R_OK):
        raise PermissionError(f"Cannot read PDF file: {pdf_file}")

    try:
        doc = fitz.open(pdf_file)
        if len(doc) == 0:
            raise Exception("PDF file is empty")
        image_paths = []
        for i, page in enumerate(doc):
            # Use a 2x zoom for better quality before compression
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_path = os.path.join(output_dir, f"slide_{i + 1}.jpg")
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.save(img_path, 'JPEG', quality=60, optimize=True)
            image_paths.append(img_path)
        print(f"Successfully converted and saved {len(image_paths)} pages")
        return image_paths
    except Exception as e:
        print(f"Error during PDF conversion: {str(e)}")
        raise Exception(f"Failed to convert PDF to images: {str(e)}")

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

def add_images_to_canvas(c, image_paths, slides_per_row=2, gap=10, margin=20, top_margin=0):
    """
    Draws the provided slide images onto the given ReportLab canvas 'c'
    using a grid layout on as many A4 pages as needed.
    Returns the number of pages used.
    """
    a4_width, a4_height = A4

    if not image_paths:
        return 0

    with Image.open(image_paths[0]) as im:
        orig_w, orig_h = im.size
    aspect_ratio = orig_w / orig_h

    available_width = a4_width - 2 * margin - (slides_per_row - 1) * gap
    available_height = a4_height - top_margin - margin
    slide_width = available_width / slides_per_row
    slide_height = slide_width / aspect_ratio

    rows_fit = math.floor((available_height + gap) / (slide_height + gap))
    if rows_fit < 1:
        rows_fit = 1
        slide_height = available_height
        slide_width = slide_height * aspect_ratio
        if slide_width * slides_per_row + (slides_per_row - 1) * gap > available_width:
            slide_width = available_width / slides_per_row
            slide_height = slide_width / aspect_ratio
    else:
        potential_rows = rows_fit + 1
        potential_slide_height = (available_height - (potential_rows - 1) * gap) / potential_rows
        potential_slide_width = potential_slide_height * aspect_ratio
        if potential_slide_width >= slide_width * 0.6:
            rows_fit = potential_rows
            slide_height = potential_slide_height
            slide_width = potential_slide_width

    total_slide_width = slide_width * slides_per_row + (slides_per_row - 1) * gap
    x_margin = (a4_width - total_slide_width) / 2

    total_slides = len(image_paths)
    max_slides_per_page = slides_per_row * rows_fit
    pages_used = 0

    for page_start in range(0, total_slides, max_slides_per_page):
        page_slides = image_paths[page_start:page_start + max_slides_per_page]
        for idx, image_path in enumerate(page_slides):
            col = idx % slides_per_row
            row = idx // slides_per_row
            x = x_margin + col * (slide_width + gap)
            y = a4_height - top_margin - (row + 1) * slide_height - row * gap
            c.drawImage(image_path, x, y, width=slide_width, height=slide_height,
                        preserveAspectRatio=True, mask='auto')
        c.showPage()
        pages_used += 1

    return pages_used

def create_pdf_from_images(image_paths, output_pdf, slides_per_row=2, gap=10, margin=20, top_margin=0, pdf_names=None):
    """
    Create a PDF from a list of slide images using grid layout.
    """
    c = canvas.Canvas(output_pdf, pagesize=A4)
    c.setPageCompression(1)  # Enable page compression
    add_images_to_canvas(c, image_paths, slides_per_row, gap, margin, top_margin)
    c.save()

def process_file(input_path, output_path, slides_per_row=2, gap=10, margin=20, top_margin=0):
    """
    Process a single file and convert it to a PDF with the specified layout.
    """
    temp_dir = tempfile.mkdtemp()
    try:
        image_paths = convert_file_to_images(input_path, temp_dir)
        if not image_paths:
            raise Exception("No images were generated from the input file.")
        create_pdf_from_images(image_paths, output_path, slides_per_row, gap, margin, top_margin)
    finally:
        shutil.rmtree(temp_dir)

def process_directory(input_dir, output_dir, slides_per_row=2, gap=10, margin=20, top_margin=0):
    """
    Process all PDF and PowerPoint files in the input directory.
    Creates corresponding output files in the output directory.
    """
    os.makedirs(output_dir, exist_ok=True)
    supported_extensions = ['.pdf', '.ppt', '.pptx']
    input_files = []
    for ext in supported_extensions:
        input_files.extend(glob.glob(os.path.join(input_dir, f'*{ext}')))
    
    if not input_files:
        print(f"No supported files found in {input_dir}")
        return
    
    for input_file in input_files:
        filename = os.path.basename(input_file)
        output_file = os.path.join(output_dir, os.path.splitext(filename)[0] + '.pdf')
        try:
            print(f"Processing {filename}...")
            process_file(input_file, output_file, slides_per_row, gap, margin, top_margin)
            print(f"Created {output_file}")
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}", file=sys.stderr)

def process_files(input_paths, output_path, slides_per_row=2, gap=10, margin=20, top_margin=0,
                  single_file=False, new_page_per_pdf=False):
    """
    Convert multiple files to PDF(s) with the specified layout.
    
    Parameters:
        input_paths: List of input file paths.
        output_path: Output file path or directory.
        slides_per_row: Number of slides per row.
        gap: Space between slides.
        margin: Margin on sides and bottom.
        top_margin: Margin at the top.
        single_file: If True, combine all slides into a single PDF.
        new_page_per_pdf: If True (and single_file is True), ensure each input file's slides start on new pages.
    """
    if not input_paths:
        raise Exception("No input files provided")

    if single_file:
        if new_page_per_pdf:
            c = canvas.Canvas(output_path, pagesize=A4)
            c.setPageCompression(1)
            current_page = 1
            for input_path in input_paths:
                print(f"\nProcessing file: {input_path}")
                temp_dir = tempfile.mkdtemp(prefix='ppt_to_pdf_')
                try:
                    image_paths = convert_file_to_images(input_path, temp_dir)
                    if not image_paths:
                        raise Exception("No images were generated for file: " + input_path)
                    bookmark_name = os.path.splitext(os.path.basename(input_path))[0]
                    c.bookmarkPage(f"page_{current_page}")
                    c.addOutlineEntry(bookmark_name, f"page_{current_page}", 0)
                    pages_used = add_images_to_canvas(c, image_paths, slides_per_row, gap, margin, top_margin)
                    current_page += pages_used
                finally:
                    shutil.rmtree(temp_dir)
            c.save()
        else:
            all_image_paths = []
            temp_dirs = []
            pdf_names = []
            try:
                for input_path in input_paths:
                    print(f"\nProcessing file: {input_path}")
                    temp_dir = tempfile.mkdtemp(prefix='ppt_to_pdf_')
                    temp_dirs.append(temp_dir)
                    image_paths = convert_file_to_images(input_path, temp_dir)
                    if not image_paths:
                        raise Exception("No images for file: " + input_path)
                    pdf_names.append(os.path.splitext(os.path.basename(input_path))[0])
                    all_image_paths.extend(image_paths)
                create_pdf_from_images(all_image_paths, output_path, slides_per_row, gap, margin, top_margin, pdf_names)
            finally:
                for d in temp_dirs:
                    shutil.rmtree(d)
    else:
        if os.path.isdir(output_path):
            for input_path in input_paths:
                filename = os.path.basename(input_path)
                file_output_path = os.path.join(output_path, os.path.splitext(filename)[0] + '.pdf')
                process_file(input_path, file_output_path, slides_per_row, gap, margin, top_margin)
        else:
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
    parser.add_argument("--no_new_page", action="store_true", help="Disable forcing each PDF's slides on a new page (only applies when --single_file is used)")
    args = parser.parse_args()

    try:
        if os.path.isdir(args.input):
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
            args.single_file,
            new_page_per_pdf=not args.no_new_page
        )
        print(f"Successfully created PDF(s) in: {args.output}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
