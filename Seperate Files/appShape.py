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

def extract_shape(detail):
    if not isinstance(detail, str):
        return None

    detail = detail.upper()

    # Shape keyword mapping
    shape_map = {
        "Star": ["STAR"],
        "Round": ["ROUND", "RND", "ROU"],
        "Clubmaster": ["CLUBMASTER", "CLUB"],
        "Hexa": ["HEXA", "HEX"],
        "Cat Eye": ["CAT EYE", "CAT"],
        "Wayfarer": ["WAYFARER", "WAY"],
        "Butterfly": ["BUTTERFLY", "BUTTR", "BUTT"],
        "Aviator": ["AVIATOR", "AVI"],
        "Pillow": ["PILLOW", "PILL"],
        "Square": ["SQUARE", "SQR", "SQ"],
        "Pilot": ["PILOT", "PIL"],
        "Oval": ["OVAL", "OV", "OVA"],
        "Rectangle": ["RECTANGLE", "RECTANGL", "RECT", "REC", "RE"]
    }

    for shape, keywords in shape_map.items():
        for keyword in keywords:
            if keyword in detail:
                return shape

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
            df['SHAPE'] = df['Details'].apply(extract_shape)

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
