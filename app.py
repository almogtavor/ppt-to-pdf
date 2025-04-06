from flask import Flask, request, jsonify, send_file
import os
import tempfile
import shutil
from main import process_file, process_directory
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

ALLOWED_EXTENSIONS = {'pdf', 'ppt', 'pptx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return app.send_static_file('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'files[]' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    files = request.files.getlist('files[]')
    if not files or files[0].filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    # Get parameters
    slides_per_row = int(request.form.get('slides_per_row', 3))
    gap = int(request.form.get('gap', 10))
    margin = int(request.form.get('margin', 20))
    top_margin = int(request.form.get('top_margin', 0))
    single_file = request.form.get('single_file', 'false').lower() == 'true'
    
    # Create output directory
    output_dir = os.path.join(app.config['UPLOAD_FOLDER'], 'output')
    os.makedirs(output_dir, exist_ok=True)
    
    input_paths = []
    try:
        # Save all uploaded files
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(input_path)
                input_paths.append(input_path)
        
        if not input_paths:
            return jsonify({'error': 'No valid files processed'}), 400
        
        # Process files
        if single_file:
            # Create a single output file
            output_path = os.path.join(output_dir, 'combined.pdf')
            process_files(input_paths, output_path, slides_per_row, gap, margin, top_margin, single_file=True)
            return send_file(output_path, as_attachment=True, download_name='combined.pdf')
        else:
            # Process each file separately
            for input_path in input_paths:
                filename = os.path.basename(input_path)
                output_path = os.path.join(output_dir, os.path.splitext(filename)[0] + '.pdf')
                process_files([input_path], output_path, slides_per_row, gap, margin, top_margin)
            
            # Create zip file of results
            zip_path = os.path.join(app.config['UPLOAD_FOLDER'], 'converted.zip')
            shutil.make_archive(zip_path[:-4], 'zip', output_dir)
            return send_file(zip_path, as_attachment=True, download_name='converted.zip')
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        # Clean up input files
        for input_path in input_paths:
            if os.path.exists(input_path):
                os.remove(input_path)

if __name__ == '__main__':
    app.run(debug=True) 