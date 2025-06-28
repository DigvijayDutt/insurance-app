from flask import Flask, render_template, request, send_from_directory
import pandas as pd
from docx import Document
from docx.shared import Inches
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)
        
        # Process the Excel and generate Word files
        generate_docs(filepath)
        return "Files generated! Check the 'outputs' folder."

    return render_template('index.html')

def generate_docs(excel_path):
    df = pd.read_excel(excel_path)
    columns = df.columns
    rooms = ["LIVING ROOM", "BEDROOM", "KITCHEN", "STORAGE"]
    data = df.values

    for j in range(len(data)):
        doc = Document()
        doc.add_picture("./static/logo.jpg", width=Inches(1.25))
        doc.add_paragraph("www.trinitycontents.com")
        doc.add_heading("FIRST INSPECTION REPORT", level=1)
        doc.add_picture("./static/home.jpg", width=Inches(2.5))

        for i in range(len(columns)):
            doc.add_heading(f"{columns[i]}", level=2)
            doc.add_paragraph(str(data[j][i]))

        doc.add_heading("PHOTOGRAPHS", level=2)
        for room in rooms:
            doc.add_heading(room, level=3)
            for l in range(1, 5):
                img_path = f"./static/images/{room}/{l}.jpg"
                if os.path.exists(img_path):
                    doc.add_picture(img_path)

        output_path = os.path.join(OUTPUT_FOLDER, f"claim_{j}.docx")
        doc.save(output_path)

@app.route('/outputs/<filename>')
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
