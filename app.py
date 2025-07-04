import os
import re
import pandas as pd
from flask import Flask, request, render_template, send_file, redirect, url_for
from werkzeug.utils import secure_filename

# Create Flask app instance
app = Flask(__name__)

# Define upload and output directories
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# ---------- Extraction Functions for Each Attribute ----------

def extract_gender(text):
    """
    Extracts gender ('Ladies', 'Gents', 'Unisex') from a detail string.
    Looks for L, G, or U as the third or later token (case-insensitive).
    """
    if not isinstance(text, str):
        return ''
    tokens = text.upper().split()[2:]
    for token in reversed(tokens):
        clean_token = token.rstrip('.')
        if clean_token in ['L', 'G', 'U']:
            return {'L': 'Ladies', 'G': 'Gents', 'U': 'Unisex'}[clean_token]
    return ''

def extract_material(detail):
    """
    Extracts material type (e.g., 'TITANIUM', 'METAL', 'PLASTICS') from the detail string.
    Identifies material based on predefined keywords.
    """
    if not isinstance(detail, str):
        return None
    detail = detail.upper()
    if "TITAN" in detail or "TITA" in detail:
        return "TITANIUM"
    elif "METAL" in detail or detail.endswith(("M", "ME", "META")):
        return "METAL"
    elif "SHELL" in detail or "PLASTIC" in detail:
        return "PLASTICS"
    return None

def extract_shape(detail):
    """
    Extracts frame shape (e.g., 'Round', 'Rectangle', 'Aviator') from keywords in the detail string.
    Uses a comprehensive mapping of shape names to possible keywords.
    """
    if not isinstance(detail, str):
        return None
    detail = detail.upper()
    shape_map = {
        "Star": ["STAR"], "Round": ["ROUND", "RND", "ROU"],
        "Clubmaster": ["CLUBMASTER", "CLUB"], "Hexa": ["HEXA", "HEX"],
        "Cat Eye": ["CAT EYE", "CAT"], "Wayfarer": ["WAYFARER", "WAY"],
        "Butterfly": ["BUTTERFLY", "BUTTR", "BUTT"], "Aviator": ["AVIATOR", "AVI"],
        "Pillow": ["PILLOW", "PILL"], "Square": ["SQUARE", "SQR", "SQ"],
        "Pilot": ["PILOT", "PIL"], "Oval": ["OVAL", "OV", "OVA"],
        "Rectangle": ["RECTANGLE", "RECTANGL", "RECT", "REC", "RE"],
        "Irregular": ["IRREGULAR", "IRREGU", "IRR", "IRREG", "IRRE"]
    }
    for shape, keywords in shape_map.items():
        for keyword in keywords:
            if keyword in detail:
                return shape
    return None

def extract_style(detail):
    """
    Extracts style (e.g., 'Full Rim', 'Rimless', 'Supra') from keywords in the detail string.
    """
    if not isinstance(detail, str):
        return None
    detail = detail.upper()
    if "FULL" in detail or "FUL" in detail or detail.endswith("F") or "WAYFARER" in detail:
        return "Full Rim"
    elif "R/L" in detail or "R/" in detail or "WAS" in detail or "WASHER" in detail or "WASH" in detail:
        return "Rimless"
    elif "SUPRA" in detail or "SU" in detail or "SUP" in detail:
        return "Supra"
    return None

def extract_frame_size(detail):
    """
    Extracts frame size from patterns like '49-18-135' or standalone numbers (46-70).
    Returns the first valid match as an integer.
    """
    if not isinstance(detail, str):
        return None

    detail = detail.upper()

    # Pattern: e.g. '49-18-135' (eye size, bridge size, temple length)
    match_full = re.search(r'\b([4-6][0-9]|70)-\d{1,2}-\d{2,3}\b', detail)
    if match_full:
        return int(match_full.group(1))

    # Standalone number typically representing eye size (between 46 and 70)
    match_number = re.findall(r'\b([4-6][0-9]|70)\b', detail)
    if match_number:
        return int(match_number[0])  # Return the first valid match

    return None

def extract_color(detail):
    """
    Extracts color code from the detail string using a prioritized set of rules:
    1. 'C-XXX' pattern (e.g., C-1234)
    2. 'C' followed immediately by digits (e.g., C123)
    3. Second token if it's alphanumeric and not an ignored word.
    4. Token immediately preceding a detected frame size pattern.
    5. Middle part of a hyphenated number sequence (e.g., 6193-2502-51 -> 2502).
    6. Third-last token if the second-last token is a frame size pattern.
    Ignores common material/style words to prevent misidentification.
    """
    if not isinstance(detail, str):
        return None

    detail = detail.upper()
    tokens = detail.strip().split()

    # Words to ignore as invalid color values (e.g., materials, styles)
    ignore_words = {'METAL', 'SUPRA', 'R/L', 'SPG', 'META', 'RIM', 'WASHER', 'TITAN', 'PLASTIC', 'SHELL'}

    # Priority 1: C-XXX pattern
    for token in tokens:
        if token.startswith("C-"):
            result = token[2:].split("-")[0]
            if result not in ignore_words:
                return result

    # Priority 2: C followed by digits (e.g., C123)
    for token in tokens:
        if re.match(r'^C\d+$', token):
            result = token[1:]
            if result not in ignore_words:
                return result

    # Priority 3: Second token if alphanumeric and not an ignored word
    if len(tokens) >= 2:
        second = tokens[1]
        if re.match(r'^[A-Z0-9]{2,6}$', second) and second not in ignore_words:
            return second

    # Priority 4: Token before frame size (e.g., "BLUE 49-18-135")
    for i in range(1, len(tokens)):
        if re.match(r'\d{2}-\d{2}-\d{3}', tokens[i]):
            before = tokens[i-1]
            if re.match(r'^[A-Z0-9]{2,6}$', before) and before not in ignore_words:
                return before

    # Priority 5: Middle part of hyphenated numbers (e.g., 6193-2502-51 -> 2502)
    match = re.search(r'\b\d{3,5}-(\d{2,5})-\d{2,5}\b', detail)
    if match:
        result = match.group(1)
        if result not in ignore_words:
            return result

    # Priority 6: Third-last token if the second-last is a frame size pattern
    if len(tokens) >= 3 and re.match(r'\d{2}-\d{2}-\d{3}', tokens[-2]):
        result = tokens[-3]
        if result not in ignore_words:
            return result

    return None

# ---------- Flask Routes ----------

@app.route('/', methods=['GET', 'POST'])
def index():
    """
    Main page handler:
    - On GET request, renders the upload form.
    - On POST request, handles file upload, extracts selected attributes from
      the 'Details' column of an Excel file, and provides a download link.
    """
    if request.method == 'POST':
        # Get selected attributes from the form (e.g., ['gender', 'material'])
        options = request.form.getlist('options')
        file = request.files.get('file')

        # Redirect if no file is provided
        if not file or file.filename == '':
            return redirect(request.url)

        # Securely save the uploaded file
        filename = secure_filename(file.filename)
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)

        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(filepath)

        # Check if 'Details' column exists in the DataFrame
        if 'Details' not in df.columns:
            return "⚠️ Error: 'Details' column not found in uploaded Excel file. Please ensure the header is correct."

        # DEBUG: Print DataFrame columns and first few rows of 'Details' column for verification
        print("Columns in uploaded Excel:", df.columns.tolist())
        if 'Details' in df.columns:
            print("First 5 rows of 'Details' column:\n", df[['Details']].head())
        else:
            print("❌ 'Details' column was NOT FOUND after reading the file.")

        # Apply extraction functions based on selected options
        # Each 'if' block checks if the option was selected and if 'Details' column exists
        if 'gender' in options and 'Details' in df.columns:
            df['GENDER'] = df['Details'].apply(extract_gender)
        if 'material' in options and 'Details' in df.columns:
            df['MATERIAL'] = df['Details'].apply(extract_material)
        if 'shape' in options and 'Details' in df.columns:
            df['SHAPE'] = df['Details'].apply(extract_shape)
        if 'style' in options and 'Details' in df.columns:
            df['STYLE'] = df['Details'].apply(extract_style)
        if 'size' in options and 'Details' in df.columns:
            df['SIZE'] = df['Details'].apply(extract_frame_size)
        if 'color' in options and 'Details' in df.columns:
            df['COLOR'] = df['Details'].apply(extract_color)

        # Save the processed DataFrame to a new Excel file in the outputs directory
        output_filename = f"processed_{filename}"
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        df.to_excel(output_path, index=False)

        # Redirect the user to the download page for the newly created file
        return redirect(url_for('download_file', filename=output_filename))

    # For GET requests, render the initial upload form
    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    """
    Provides a mechanism to download the processed output file.
    """
    return send_file(os.path.join(OUTPUT_FOLDER, filename), as_attachment=True)

# ---------- Main Entry Point ----------
if __name__ == '__main__':
    # Run the Flask application in debug mode for development
    app.run(debug=True)
