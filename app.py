import streamlit as st
import os
import tempfile
from main import process_files
import glob

st.set_page_config(
    page_title="PPT/PDF to Multi-Slide PDF Converter",
    page_icon="📄",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .stApp {
        max-width: 1200px;
        margin: 0 auto;
    }
    .upload-area {
        border: 2px dashed #ccc;
        border-radius: 5px;
        padding: 20px;
        text-align: center;
        margin: 20px 0;
    }
    .file-list {
        margin: 20px 0;
    }
    .settings {
        background-color: #f0f2f6;
        padding: 20px;
        border-radius: 5px;
        margin: 20px 0;
    }
    </style>
""", unsafe_allow_html=True)

st.title("📄 PPT/PDF to Multi-Slide PDF Converter")
st.markdown("""
    Convert PowerPoint presentations and PDFs into multi-slide PDFs with customizable layouts.
    Upload your files and adjust the settings to create the perfect layout.
""")

# File upload
st.subheader("Upload Files")
uploaded_files = st.file_uploader(
    "Choose files",
    type=['pdf', 'ppt', 'pptx'],
    accept_multiple_files=True
)

# Settings
st.subheader("Layout Settings")
col1, col2 = st.columns(2)

with col1:
    slides_per_row = st.slider(
        "Slides per Row",
        min_value=1,
        max_value=6,
        value=2,
        help="Number of slides to display in each row"
    )
    gap = st.slider(
        "Gap between Slides",
        min_value=0,
        max_value=50,
        value=10,
        help="Space between slides in points"
    )

with col2:
    margin = st.slider(
        "Margin",
        min_value=0,
        max_value=50,
        value=20,
        help="Margin on sides and bottom in points"
    )
    top_margin = st.slider(
        "Top Margin",
        min_value=0,
        max_value=50,
        value=0,
        help="Margin at the top in points"
    )

# Additional options
st.subheader("Output Options")
single_file = st.checkbox(
    "Combine all slides into a single PDF",
    value=True,
    help="Create one PDF with all slides combined"
)
new_page_per_pdf = st.checkbox(
    "Start each PDF's slides on a new page",
    value=True,
    help="Add a blank page between different PDFs"
)

# Process button
if st.button("Convert to PDF", type="primary"):
    if not uploaded_files:
        st.error("Please upload at least one file")
    else:
        with st.spinner("Processing files..."):
            try:
                # Create temporary directory for uploaded files
                with tempfile.TemporaryDirectory() as temp_dir:
                    # Save uploaded files
                    input_paths = []
                    for uploaded_file in uploaded_files:
                        file_path = os.path.join(temp_dir, uploaded_file.name)
                        with open(file_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        input_paths.append(file_path)

                    # Create output file
                    output_file = os.path.join(temp_dir, "output.pdf")

                    # Process files
                    process_files(
                        input_paths,
                        output_file,
                        slides_per_row=slides_per_row,
                        gap=gap,
                        margin=margin,
                        top_margin=top_margin,
                        single_file=single_file,
                        new_page_per_pdf=new_page_per_pdf
                    )

                    # Read the output file
                    with open(output_file, "rb") as f:
                        pdf_bytes = f.read()

                    # Download button
                    st.download_button(
                        label="Download PDF",
                        data=pdf_bytes,
                        file_name="output.pdf",
                        mime="application/pdf"
                    )
                    st.success("Conversion completed successfully!")

            except Exception as e:
                st.error(f"An error occurred: {str(e)}")

# Instructions
st.markdown("""
    ### Instructions
    1. Upload your PowerPoint (.ppt, .pptx) or PDF files
    2. Adjust the layout settings to your preference
    3. Choose whether to combine all slides into a single PDF
    4. Click "Convert to PDF" to process the files
    5. Download the resulting PDF

    ### Tips
    - For best results, use similar-sized slides
    - Adjust the margins and gaps to optimize the layout
    - The "Slides per Row" setting affects the size of each slide
""") 