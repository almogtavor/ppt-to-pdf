import os
import sys
import math
import tempfile
import shutil
import argparse
import io
import subprocess  # New import for calling OCRmyPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from PIL import Image
import pathlib
import glob
import numpy as np
import cv2 
from skimage.metrics import structural_similarity as compare_ssim

# Increased scale factor for improved resolution; adjust as needed
COMPOSITE_SCALE = 3

def convert_ppt_to_images(ppt_file, output_dir):
    """
    Convert each slide of the given PowerPoint file to a JPEG image.
    Uses the COM interface (via comtypes) on Windows with Microsoft PowerPoint installed.
    Slides are exported as PNG and then converted to high-quality JPEG.
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
        png_path = os.path.join(output_dir, f"slide_{slide.SlideIndex}.png")
        slide.Export(png_path, "PNG")
        with Image.open(png_path) as img:
            rgb_img = img.convert("RGB")
            jpg_path = os.path.join(output_dir, f"slide_{slide.SlideIndex}.jpg")
            rgb_img.save(jpg_path, "JPEG", quality=150, optimize=True, progressive=True)
        os.remove(png_path)
        image_paths.append(jpg_path)

    presentation.Close()
    powerpoint.Quit()
    return image_paths

def convert_pdf_to_images(pdf_file, output_dir):
    """
    Convert PDF pages to JPEG images using PyMuPDF.
    The JPEG quality is set to 95.
    The zoom factor remains unchanged.
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
        zoom = 2  # Keep zoom factor unchanged
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            img_path = os.path.join(output_dir, f"slide_{i + 1}.jpg")
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img.save(img_path, 'JPEG', quality=150, optimize=True, progressive=True)
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

import pytesseract

def extract_text(image_path, lang="eng+heb"):
    """Extract visible text from image using Tesseract OCR."""
    img = cv2.imread(image_path)
    return pytesseract.image_to_string(img, lang=lang)

def filter_progressive_slides(
    image_paths,
    ssim_threshold: float = 0.96,
    subset_ratio_threshold: float = 0.92,
    removed_ratio_threshold: float = 0.05,
    shrink: float = 0.35,
    dilate_iter: int = 1,
    ocr_lang: str = "eng+heb"
):
    """
    Improved filtering:
    • Detects chains of progressive slides.
    • Uses OCR to compare slide text; keeps last in each chain.
    • Always keeps final slide.
    """
    if len(image_paths) <= 1:
        return image_paths

    kernel = np.ones((3, 3), np.uint8)
    keep = []
    prev_g, prev_b = None, None
    prev_idx = 0
    chain_detected = False  # tracks a sequence of progressive slides

    def _prep(p):
        g = cv2.imread(p, cv2.IMREAD_GRAYSCALE)
        g = cv2.resize(g, (0, 0), fx=shrink, fy=shrink)
        _, b = cv2.threshold(g, 230, 255, cv2.THRESH_BINARY_INV)
        b = cv2.dilate(b, kernel, iterations=dilate_iter)
        return g, b

    prev_g, prev_b = _prep(image_paths[0])

    for idx in range(1, len(image_paths)):
        cur_g, cur_b = _prep(image_paths[idx])
        ssim = compare_ssim(prev_g, cur_g)

        overlap = cv2.bitwise_and(prev_b, cur_b)
        overlap_ratio = np.sum(overlap) / max(np.sum(prev_b), 1)

        removed = cv2.bitwise_and(prev_b, cv2.bitwise_not(cur_b))
        removed_ratio = np.sum(removed) / max(np.sum(prev_b), 1)

        is_progressive = (
            (ssim > ssim_threshold and overlap_ratio > subset_ratio_threshold)
            and removed_ratio < removed_ratio_threshold
        )

        # # 👇 OCR-based override if textual content differs
        # prev_text = extract_text(image_paths[prev_idx], lang=ocr_lang).strip()
        # cur_text = extract_text(image_paths[idx], lang=ocr_lang).strip()
        # # Normalize
        # prev_clean = prev_text.replace('\n', ' ').strip()
        # cur_clean = cur_text.replace('\n', ' ').strip()

        # # If previous content is NOT a subset of current, treat as unique
        # if prev_clean not in cur_clean:
        #     is_progressive = False

        if is_progressive:
            print(f"Slide {prev_idx + 1} is progressive of {idx + 1} → skipped")
            chain_detected = True
        else:
            if chain_detected:
                keep.append(image_paths[prev_idx])  # keep last from previous chain
                chain_detected = False
            else:
                keep.append(image_paths[prev_idx])
        prev_g, prev_b, prev_idx = cur_g, cur_b, idx

    keep.append(image_paths[prev_idx])
    return keep

def composite_page(page_images, slides_per_row, gap, margin, top_margin, a4_size, scale=COMPOSITE_SCALE, rtl=False):
    a4_w, a4_h = a4_size
    comp_w, comp_h = a4_w * scale, a4_h * scale
    composite = Image.new("RGB", (int(comp_w), int(comp_h)), "white")

    margin_scaled = margin * scale
    top_margin_scaled = top_margin * scale
    gap_scaled = gap * scale

    with Image.open(page_images[0]) as im:
        aspect_ratio = im.width / im.height

    avail_w = comp_w - 2 * margin_scaled
    avail_h = comp_h - top_margin_scaled - margin_scaled
    slide_width = avail_w / slides_per_row
    slide_height = slide_width / aspect_ratio

    for idx, image_path in enumerate(page_images):
        row = idx // slides_per_row
        col = idx % slides_per_row
        if rtl:
            col = slides_per_row - 1 - col

        x = margin_scaled + col * (slide_width + gap_scaled)
        y = top_margin_scaled + row * (slide_height + gap_scaled)
        try:
            with Image.open(image_path) as slide_img:
                slide_resized = slide_img.resize((int(slide_width), int(slide_height)), Image.LANCZOS)
                composite.paste(slide_resized, (int(x), int(y)))
        except Exception as e:
            print(f"Error processing {image_path}: {e}")
    return composite

def add_images_to_canvas(c, image_paths, slides_per_row=2, gap=10, margin=20, top_margin=0, rtl=False):
    """
    Groups slide images into pages, composites each page into a high-resolution image,
    and draws the composite image onto the canvas.
    Returns the number of pages used.
    """
    a4_w, a4_h = A4  # in points (72 dpi)
    pages_used = 0

    avail_w = a4_w - 2 * margin
    with Image.open(image_paths[0]) as im:
        aspect_ratio = im.width / im.height
    slide_width = avail_w / slides_per_row
    slide_height = slide_width / aspect_ratio
    avail_h = a4_h - top_margin - margin
    rows_fit = math.floor((avail_h + gap) / (slide_height + gap))
    max_slides_per_page = slides_per_row * rows_fit

    for i in range(0, len(image_paths), max_slides_per_page):
        page_group = image_paths[i:i + max_slides_per_page]
        composite_img = composite_page(page_group, slides_per_row, gap, margin, top_margin, (a4_w, a4_h), scale=COMPOSITE_SCALE, rtl=rtl)
        img_buffer = io.BytesIO()
        composite_img.save(img_buffer, format="JPEG", quality=150, optimize=True, progressive=True)
        img_buffer.seek(0)
        img_reader = ImageReader(img_buffer)
        c.drawImage(img_reader, 0, 0, width=a4_w, height=a4_h)
        c.showPage()
        pages_used += 1
    return pages_used

def create_pdf_from_images(image_paths, output_pdf, slides_per_row=2, gap=10, margin=20, top_margin=0, pdf_names=None, rtl=False):
    """
    Create a PDF from a list of slide images by compositing each page into a high-resolution image.
    """
    c = canvas.Canvas(output_pdf, pagesize=A4)
    c.setPageCompression(1)
    add_images_to_canvas(c, image_paths, slides_per_row, gap, margin, top_margin, rtl)
    c.save()

def process_file(input_path, output_path, slides_per_row=2, gap=10, margin=20, top_margin=0, rtl=False, filter_progressive=False):
    """
    Process a single file and convert it to a PDF with the specified layout.
    """
    temp_dir = tempfile.mkdtemp()
    try:
        image_paths = convert_file_to_images(input_path, temp_dir)
        if filter_progressive:
            image_paths = filter_progressive_slides(image_paths)
        if not image_paths:
            raise Exception("No images were generated from the input file.")
        create_pdf_from_images(image_paths, output_path, slides_per_row, gap, margin, top_margin, rtl=rtl)
    finally:
        shutil.rmtree(temp_dir)

def process_directory(input_dir, output_dir, slides_per_row=2, gap=10, margin=20, top_margin=0):
    """
    Process all PDF and PowerPoint files in the input directory.
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
                  single_file=False, new_page_per_pdf=False, rtl=False, filter_progressive=False):
    """
    Convert multiple files to PDF(s) with the specified layout.
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
                    if filter_progressive:
                        image_paths = filter_progressive_slides(image_paths)
                    if not image_paths:
                        raise Exception("No images were generated for file: " + input_path)
                    bookmark_name = os.path.splitext(os.path.basename(input_path))[0]
                    c.bookmarkPage(f"page_{current_page}")
                    c.addOutlineEntry(bookmark_name, f"page_{current_page}", 0)
                    pages_used = add_images_to_canvas(c, image_paths, slides_per_row, gap, margin, top_margin, rtl)
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
                    if filter_progressive:
                        image_paths = filter_progressive_slides(image_paths)
                    if not image_paths:
                        raise Exception("No images for file: " + input_path)
                    pdf_names.append(os.path.splitext(os.path.basename(input_path))[0])
                    all_image_paths.extend(image_paths)
                create_pdf_from_images(all_image_paths, output_path, slides_per_row, gap, margin, top_margin, pdf_names, rtl)
            finally:
                for d in temp_dirs:
                    shutil.rmtree(d)
    else:
        if os.path.isdir(output_path):
            for input_path in input_paths:
                filename = os.path.basename(input_path)
                file_output_path = os.path.join(output_path, os.path.splitext(filename)[0] + '.pdf')
                process_file(input_path, file_output_path, slides_per_row, gap, margin, top_margin, rtl, filter_progressive)
        else:
            if len(input_paths) > 1:
                raise Exception("Multiple input files require an output directory when single_file is False")
            process_file(input_paths[0], output_path, slides_per_row, gap, margin, top_margin, rtl, filter_progressive)

def run_ocr_on_pdf(pdf_path, ocr_lang="eng+heb"):
    """
    Run OCRmyPDF on a single PDF file to add a hidden, searchable text layer.
    The output will overwrite the original file.
    """
    temp_output = pdf_path + ".ocr.pdf"
    print(f"Running OCR on {pdf_path}...")
    try:
        subprocess.run(["ocrmypdf", "-l", ocr_lang, pdf_path, temp_output], check=True)
        shutil.move(temp_output, pdf_path)
        print("OCR applied successfully.")
    except Exception as e:
        print(f"OCR failed: {e}", file=sys.stderr)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Convert PowerPoint presentations or PDFs to PDFs with multiple slides per page by compositing a high-resolution screenshot of each page."
    )
    parser.add_argument("input", help="Path to input file or directory")
    parser.add_argument("output", help="Path for output file or directory")
    parser.add_argument("--slides_per_row", type=int, default=2, help="Number of slides per row (default: 2)")
    parser.add_argument("--gap", type=int, default=10, help="Gap (in points) between slides (default: 10)")
    parser.add_argument("--margin", type=int, default=20, help="Margin (in points) on the sides and bottom (default: 20)")
    parser.add_argument("--top_margin", type=int, default=0, help="Margin (in points) at the top of the page (default: 0)")
    parser.add_argument("--single_file", action="store_true", help="Combine all slides into a single PDF file")
    parser.add_argument("--no_new_page", action="store_true", help="Disable forcing each PDF's slides on a new page (only applies when --single_file is used)")
    parser.add_argument("--ocr", action="store_true", default=True, help="Run OCR on the generated PDF(s) to add a searchable text layer (requires OCRmyPDF)")
    parser.add_argument("--ocr-lang", default="eng+heb",
                    help="Tesseract languages to use for OCR (e.g. 'eng', 'heb', or 'eng+heb')")
    parser.add_argument("--rtl", action="store_true", help="Enable right-to-left layout")
    parser.add_argument("--filter-progressive", action="store_true", default=True, help="Filter out progressive slides (slides that are just builds of the next one)")
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
            new_page_per_pdf=not args.no_new_page,
            rtl=args.rtl,
            filter_progressive=args.filter_progressive
        )
        print(f"Successfully created PDF(s) in: {args.output}")

        if args.ocr:
            if os.path.isdir(args.output):
                pdf_files = glob.glob(os.path.join(args.output, "*.pdf"))
                for pdf in pdf_files:
                    run_ocr_on_pdf(pdf, args.ocr_lang)
            else:
                run_ocr_on_pdf(args.output, args.ocr_lang)

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
