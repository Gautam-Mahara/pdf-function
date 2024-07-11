from flask import Flask, request, send_file, Response
from flask_cors import CORS
from pdf2docx import Converter
import os

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = './uploads'
CONVERTED_FOLDER = './converted'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

@app.route('/pdf-to-word', methods=['POST'])
def convert_pdf_to_word():
    if 'file' not in request.files:
        return "No file part", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    # Save the uploaded PDF file
    pdf_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(pdf_path)

    # Convert PDF to Word
    word_filename = f"{os.path.splitext(file.filename)[0]}.docx"
    word_path = os.path.join(CONVERTED_FOLDER, word_filename)

    try:
        converter = Converter(pdf_path)
        converter.convert(word_path)
    except Exception as e:
        print(f"Conversion failed: {str(e)}")
        return f"Conversion failed: {str(e)}", 500

    # Prepare response with the converted Word file
    with open(word_path, 'rb') as f:
        response = Response(f.read(), mimetype='application/msword')
        response.headers['Content-Disposition'] = f'attachment; filename={word_filename}'
    
    return response

@app.route('/', methods=['GET'])
def home():
    return "Welcome to PDF to Word converter API"

if __name__ == '__main__':
    app.run(debug=True)
