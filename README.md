# Flask Frame Details Extractor

A Flask web app for extracting gender, material, shape, style, size, and color from eyewear "Details" in Excel files.

## Features

- Upload an Excel file with a `Details` column
- Choose which attributes to extract (gender, material, shape, style, size, color)
- Download the processed file with new columns for selected attributes

## Requirements

- Python 3.8+
- Flask
- pandas
- openpyxl

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
python app.py
```

Go to [http://localhost:5000](http://localhost:5000) in your browser.

## Project Structure

- `app.py` — Main Flask app
- `templates/index.html` — Upload and options form
- `uploads/`, `outputs/` — Created automatically for file handling

## Notes

- Only Excel files (`.xlsx`) are supported.
- Make sure your file contains a `Details` column.

---
File Structure
flask-frame-details-extractor/
│
├── app.py
├── uploads/         # (auto-created, for uploads)
├── outputs/         # (auto-created, for outputs)
├── templates/
│   └── index.html   # HTML form for upload and options
