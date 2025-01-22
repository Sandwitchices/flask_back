import os
from flask import Flask, request, jsonify
from flask_cors import CORS
from werkzeug.utils import secure_filename
from pptx import Presentation

app = Flask(__name__)

# Enable CORS for all domains
CORS(app)

# Configure upload folder
app.config['UPLOAD_FOLDER'] = 'uploads'

# Ensure the upload folder exists
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/')
def home():
    return "Welcome to the PPT Parser API!"

@app.route('/parse', methods=['POST'])
def parse_ppt():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400

    if file and file.filename.endswith('.pptx'):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        try:
            # Parse the PowerPoint file
            presentation = Presentation(file_path)
            slides_content = []
            for i, slide in enumerate(presentation.slides):
                slide_text = []
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        slide_text.append(shape.text.strip())
                slides_content.append({"slide": i + 1, "text": slide_text})

            return jsonify({"slides": slides_content}), 200

        except Exception as e:
            return jsonify({"error": f"Error parsing PPTX file: {str(e)}"}), 500

    return jsonify({"error": "Invalid file format. Only .pptx files are allowed"}), 400

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))  # Dynamically set port
    app.run(host='0.0.0.0', port=port)
