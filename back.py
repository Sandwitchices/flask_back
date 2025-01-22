from flask import Flask, request, jsonify
from pptx import Presentation
import os

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


@app.route('/')
def home():
    return jsonify({"message": "Welcome to the PPT Parser API!"})


@app.route('/parse', methods=['POST'])
def parse_ppt():
    if 'file' not in request.files:
        return jsonify({"error": "No file part in the request"}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No file selected for uploading"}), 400
    
    if not file.filename.endswith('.pptx'):
        return jsonify({"error": "Unsupported file format. Please upload a .pptx file."}), 400
    
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)
    
    try:
        # Parse the PPTX file
        presentation = Presentation(file_path)
        content = []
        for slide in presentation.slides:
            slide_content = []
            for shape in slide.shapes:
                if shape.has_text_frame:
                    slide_content.append(shape.text)
            content.append({"slide": len(content) + 1, "text": slide_content})
        
        return jsonify({"slides": content}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        # Clean up: Remove the uploaded file
        if os.path.exists(file_path):
            os.remove(file_path)


if __name__ == '__main__':
    app.run(debug=True)
