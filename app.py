from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
import os

# Import the format_docx functions from both formatting styles
# from format_style_1 import format_docx as format_docx_style_1
# from format_style_2 import format_docx as format_docx_style_2
# from format_style_3 import format_docx as format_docx_style_3
# from format_style_4 import format_docx as format_docx_style_4
from format_style_5 import format_docx as format_docx_style_5
# from format_style_6 import format_docx as format_docx_style_6


app = Flask(__name__)

# Configure upload folder
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Define route to render the index.html page
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/about')
def about():
    return render_template('about.html')

# Handle form submission and processing
@app.route('/process', methods=['POST'])
def process():
    if 'docx_file' not in request.files:
        return "No file part in the request"

    docx_file = request.files['docx_file']
    if docx_file.filename == '':
        return "No file selected"

    formatting_style = request.form.get('formatting_style')

    # Save uploaded file
    docx_filename = secure_filename(docx_file.filename)
    docx_path = os.path.join(app.config['UPLOAD_FOLDER'], docx_filename)
    docx_file.save(docx_path)

    formatted_file = None

    if formatting_style == 'style_1':
        # Handle format style 1
        formatted_file = format_docx_style_1(docx_path)

    elif formatting_style == 'style_2':
        # Handle format style 2
        formatted_file = format_docx_style_2(docx_path)
    elif formatting_style == 'style_3':
        # Handle format style 3
        formatted_file = format_docx_style_3(docx_path)
    elif formatting_style == 'style_4':
        # Handle format style 4
        formatted_file = format_docx_style_4(docx_path)
    elif formatting_style == 'style_5':
        # Handle format style 5
        formatted_file = format_docx_style_5(docx_path)
    elif formatting_style == 'style_6':
        # Handle format style 6
        formatted_file = format_docx_style_6(docx_path)

    else:
        return "Invalid formatting style selected"

    # Serve the formatted file for download
    return send_file(formatted_file, as_attachment=True)

# if __name__ == '__main__':
#     app.run(debug=True)
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Use dynamic port
    app.run(host="0.0.0.0", port=port)