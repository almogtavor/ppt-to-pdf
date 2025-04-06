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
    Convert PDF pages to images using PyMuPDF.
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise Exception("The 'PyMuPDF' package is required to convert PDF pages to images. Install it with: pip install pymupdf")

    # Use absolute paths for better compatibility
    pdf_file = str(pathlib.Path(pdf_file).resolve())
    output_dir = str(pathlib.Path(output_dir).resolve())

    print(f"\nConverting PDF: {pdf_file}")
    print(f"Output directory: {output_dir}")

    # Check file access
    if not os.path.exists(pdf_file):
        raise FileNotFoundError(f"PDF file not found: {pdf_file}")
    if not os.access(pdf_file, os.R_OK):
        raise PermissionError(f"Cannot read PDF file: {pdf_file}")

    try:
        print("Starting PDF conversion...")
        # Open the PDF
        doc = fitz.open(pdf_file)
        print(f"PDF has {len(doc)} pages")
        
        if len(doc) == 0:
            raise Exception("PDF file is empty")

        image_paths = []
        for i, page in enumerate(doc):
            # Get the page as a pixmap
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x zoom for better quality
            
            # Convert to PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # Save as JPEG with compression
            img_path = os.path.join(output_dir, f"slide_{i + 1}.jpg")
            img.save(img_path, 'JPEG', quality=75, optimize=True)
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


def create_pdf_from_images(image_paths, output_pdf, slides_per_row=2, gap=10, margin=20, top_margin=0, pdf_names=None):
    """
    Create a PDF from slide images with the specified layout.
    """
    a4_width, a4_height = A4

    if not image_paths:
        raise Exception("No slide images found to create the PDF.")

    # Create PDF with outlines enabled
    c = canvas.Canvas(output_pdf, pagesize=A4)
    
    # Get slide dimensions from first image
    with Image.open(image_paths[0]) as im:
        orig_w, orig_h = im.size
    aspect_ratio = orig_w / orig_h

    # Calculate available space
    available_width = a4_width - 2 * margin - (slides_per_row - 1) * gap
    available_height = a4_height - top_margin - margin

    # Calculate slide dimensions
    slide_width = available_width / slides_per_row
    slide_height = slide_width / aspect_ratio

    # Calculate number of rows
    rows_fit = math.floor((available_height + gap) / (slide_height + gap))
    
    # Adjust for minimum size
    if rows_fit < 1:
        rows_fit = 1
        slide_height = available_height
        slide_width = slide_height * aspect_ratio
        if slide_width * slides_per_row + (slides_per_row - 1) * gap > available_width:
            slide_width = available_width / slides_per_row
            slide_height = slide_width / aspect_ratio
    else:
        # Try to fit one more row if possible
        potential_rows = rows_fit + 1
        potential_slide_height = (available_height - (potential_rows - 1) * gap) / potential_rows
        potential_slide_width = potential_slide_height * aspect_ratio
        
        # Only add row if slides remain readable
        if potential_slide_width >= slide_width * 0.6:
            rows_fit = potential_rows
            slide_height = potential_slide_height
            slide_width = potential_slide_width

    # Center content horizontally
    total_slide_width = slide_width * slides_per_row + (slides_per_row - 1) * gap
    x_margin = (a4_width - total_slide_width) / 2

    total_slides = len(image_paths)
    max_slides_per_page = slides_per_row * rows_fit
    
    # Track current PDF and its starting page
    current_pdf_name = None
    current_pdf_start_page = 1
    page_num = 1

    # Add slides to pages
    for page_start in range(0, total_slides, max_slides_per_page):
        page_slides = image_paths[page_start:page_start + max_slides_per_page]
        
        # Check if this is the start of a new PDF by looking at the first image path
        if pdf_names and page_start < len(image_paths):
            first_slide_name = os.path.basename(image_paths[page_start])
            pdf_index = int(first_slide_name.split('_')[1].split('.')[0]) - 1
            if pdf_index < len(pdf_names):
                new_pdf_name = pdf_names[pdf_index]
                if new_pdf_name != current_pdf_name:
                    # Add outline entry for the new PDF
                    c.bookmarkPage(f"page_{page_num}")
                    c.addOutlineEntry(new_pdf_name, f"page_{page_num}", 0)
                    current_pdf_name = new_pdf_name
                    current_pdf_start_page = page_num

        for idx, image_path in enumerate(page_slides):
            col = idx % slides_per_row
            row = idx // slides_per_row
            x = x_margin + col * (slide_width + gap)
            y = a4_height - top_margin - (row + 1) * slide_height - row * gap
            
            # Draw image
            c.drawImage(
                image_path,
                x, y,
                width=slide_width,
                height=slide_height,
                preserveAspectRatio=True,
                mask='auto'
            )
        c.showPage()
        page_num += 1
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


def process_files(input_paths, output_path, slides_per_row=2, gap=10, margin=20, top_margin=0, single_file=False, new_page_per_pdf=None):
    """
    Convert multiple files to PDF(s) with the specified layout.
    
    Parameters:
        input_paths: List of input file paths
        output_path: Output file path or directory
        slides_per_row: Number of slides per row
        gap: Space between slides
        margin: Margin on sides and bottom
        top_margin: Margin at the top
        single_file: If True, combine all slides into a single PDF file
        new_page_per_pdf: If True, start each PDF's slides on a new page. Defaults to True when single_file is True.
    """
    if not input_paths:
        raise Exception("No input files provided")

    # Set new_page_per_pdf to True by default when single_file is True
    if new_page_per_pdf is None:
        new_page_per_pdf = single_file

    print(f"\nProcessing {len(input_paths)} files")
    print(f"Output path: {output_path}")
    print(f"Single file mode: {single_file}")
    print(f"New page per PDF: {new_page_per_pdf}")

    if single_file:
        # Combine all slides into one PDF
        all_image_paths = []
        temp_dirs = []
        pdf_names = []
        
        try:
            # Convert each file to images and collect all paths
            for input_path in input_paths:
                print(f"\nProcessing file: {input_path}")
                # Create a unique temp directory for each file
                temp_dir = tempfile.mkdtemp(prefix='ppt_to_pdf_')
                print(f"Created temporary directory: {temp_dir}")
                temp_dirs.append(temp_dir)
                
                try:
                    # Clean up any existing files in the temp directory
                    for file in os.listdir(temp_dir):
                        try:
                            os.remove(os.path.join(temp_dir, file))
                        except:
                            pass
                    
                    image_paths = convert_file_to_images(input_path, temp_dir)
                    print(f"Converted to {len(image_paths)} images")
                    
                    # Add blank page if needed
                    if new_page_per_pdf and all_image_paths:
                        # Create a white image for the blank page
                        blank_page = Image.new('RGB', (100, 100), (255, 255, 255))
                        blank_path = os.path.join(temp_dir, 'blank.jpg')
                        blank_page.save(blank_path, 'JPEG', quality=75, optimize=True)
                        all_image_paths.append(blank_path)
                    
                    all_image_paths.extend(image_paths)
                except Exception as e:
                    print(f"Error processing {input_path}: {str(e)}")
                    raise
                
                # Get PDF name without extension
                pdf_name = os.path.splitext(os.path.basename(input_path))[0]
                pdf_names.append(pdf_name)
            
            if not all_image_paths:
                raise Exception("No images were generated from the input files")
                
            print(f"\nTotal images to process: {len(all_image_paths)}")
            print(f"Creating combined PDF: {output_path}")
            
            # Create the combined PDF
            create_pdf_from_images(all_image_paths, output_path, slides_per_row, gap, margin, top_margin, pdf_names)
            print("PDF creation completed successfully")
        finally:
            # Clean up temporary files
            print("\nCleaning up temporary files...")
            for temp_dir in temp_dirs:
                try:
                    # Remove all files in the temp directory
                    for file in os.listdir(temp_dir):
                        try:
                            os.remove(os.path.join(temp_dir, file))
                        except:
                            pass
                    # Remove the temp directory itself
                    os.rmdir(temp_dir)
                except Exception as e:
                    print(f"Warning: Could not clean up temporary directory {temp_dir}: {str(e)}")
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
    parser.add_argument("--no_new_page", action="store_true", help="Disable starting each PDF's slides on a new page (only applies when --single_file is used)")
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
            args.single_file,
            not args.no_new_page  # Invert the flag since we want new_page_per_pdf to be True by default
        )
        print(f"Successfully created PDF(s) in: {args.output}")
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)
