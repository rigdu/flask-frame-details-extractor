import os
import pandas as pd
from flask import Flask, request, render_template, send_file, redirect, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Material mapping function
def extract_material(detail):
    if not isinstance(detail, str):
        return None
    detail = detail.upper()
    if "TITAN" in detail or "TITA" in detail:
        return "TITANIUM"
    elif "METAL" in detail or detail.endswith("M") or detail.endswith("ME") or detail.endswith("META"):
        return "METAL"
    elif "SHELL" in detail or "PLASTIC" in detail:
        return "SHELL/PLASTICS"
    return None

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            # Process Excel file
            df = pd.read_excel(filepath)
            df['MATERIAL'] = df['Details'].apply(extract_material)

            output_filename = f"output_{filename}"
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            df.to_excel(output_path, index=False)

            return redirect(url_for('download_file', filename=output_filename))

    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
