import os
import subprocess
import tempfile
from flask import Flask, request, render_template, send_file, flash, redirect, url_for

app = Flask(__name__)
app.secret_key = 'super_secret_key' # Needed for flashing messages

# Set max upload size to 50MB (adjust if needed)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files[]' not in request.files:
        flash('No file part')
        return redirect(url_for('index'))
    
    files = request.files.getlist('files[]')
    if not files or files[0].filename == '':
        flash('No selected files')
        return redirect(url_for('index'))
    
    # Use the local input and output directories
    base_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir = os.path.join(base_dir, "input")
    output_dir = os.path.join(base_dir, "output")
    
    # Ensure directories exist
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    # Clear old files in input directory
    for filename in os.listdir(input_dir):
        fpath = os.path.join(input_dir, filename)
        if os.path.isfile(fpath):
            os.remove(fpath)
            
    try:
        for file in files:
            file_path = os.path.join(input_dir, file.filename)
            file.save(file_path)
            
        output_excel = os.path.join(output_dir, "Astrya_Invoices.xlsx")
        
        # Run the POC script
        script_path = os.path.join(base_dir, "extract_invoices_to_excel (1).py")
        
        import sys
        
        # Subprocess to run the extraction script
        print(f"Running POC script on directory: {input_dir}")
        result = subprocess.run(
            [sys.executable, script_path, "--pdf_dir", input_dir, "--output", output_excel],
            capture_output=True,
            text=True
        )
        
        print("Script stdout:", result.stdout)
        print("Script stderr:", result.stderr)
        
        if result.returncode != 0:
            flash(f'Error processing files. Script output: {result.stderr}')
            return redirect(url_for('index'))
            
        if not os.path.exists(output_excel):
            flash('Output Excel file was not generated.')
            return redirect(url_for('index'))
            
        return send_file(
            output_excel, 
            as_attachment=True, 
            download_name="Extracted_Invoices.xlsx"
        )
        
    except Exception as e:
        flash(f'An error occurred: {str(e)}')
        return redirect(url_for('index'))
        
if __name__ == '__main__':
    app.run(debug=True, port=5000)
