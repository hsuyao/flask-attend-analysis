from flask import Flask, request, send_file
from aspose.cells import Workbook
import os

app = Flask(__name__)

# Define upload folder
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure upload folder exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if file is present
        if 'file' not in request.files:
            return 'No file uploaded', 400
        file = request.files['file']
        if file.filename == '':
            return 'No file selected', 400
        if file and file.filename.endswith(('.xls', '.xlsx', '.ods')):
            # Save the uploaded file
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)
            # Repair the file using Aspose.Cells
            repaired_filepath = repair_excel(filepath)
            # Send the repaired file for download
            return send_file(repaired_filepath, as_attachment=True)
        else:
            return 'Invalid file type. Please upload an Excel file (.xls, .xlsx, .ods)', 400
    return '''
    <!doctype html>
    <title>Upload Excel File</title>
    <h1>Upload an Excel File</h1>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file>
      <input type=submit value=Upload>
    </form>
    '''

def repair_excel(filepath):
    # Load the workbook with repair options
    workbook = Workbook(filepath)
    # Save the repaired workbook
    repaired_filepath = os.path.join(app.config['UPLOAD_FOLDER'], 'repaired_' + os.path.basename(filepath))
    workbook.save(repaired_filepath)
    return repaired_filepath

if __name__ == '__main__':
    app.run(debug=True)
