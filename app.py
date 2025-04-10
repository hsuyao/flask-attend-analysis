# app.py
from flask import Flask, request, jsonify, send_file, redirect, url_for, session
from flask_session import Session
import uuid
import os
import traceback
from config import logger
from excel_handler import process_excel
from render_table import render_combined_table

app = Flask(__name__)

# Configure Flask-Session
app.config['SESSION_TYPE'] = 'filesystem'  # Store session data on the filesystem
app.config['SECRET_KEY'] = 'your-secret-key-here'  # Replace with a secure key
Session(app)

# CSS styles for both pages
CSS_STYLES = """
<style>
    .table-wrapper {
        overflow-x: auto;
        margin: 20px auto;
    }
    .excel-table {
        border-collapse: collapse;
        font-family: Arial, sans-serif;
        white-space: nowrap;
    }
    .excel-table th, .excel-table td {
        border: 1px solid #000;
        padding: 2px;
        text-align: left;
        vertical-align: top;
        min-width: 70px;
        line-height: 1.2;
    }
    .excel-table .separator {
        min-width: 10px;
        width: 10px;
    }
    .excel-table .title-row th {
        background-color: #005566;
        color: white;
        text-align: center;
        font-weight: bold;
        padding: 2px;
        line-height: 1.2;
    }
    .excel-table .header th {
        background-color: #107C10;
        color: white;
        padding: 2px;
        line-height: 1.2;
    }
    .excel-table .subheader th {
        background-color: #5DBB63;
        color: white;
        padding: 2px;
        line-height: 1.2;
    }
    .excel-table tr.even {
        background-color: #F3F2F1;
        color: black;
    }
    .excel-table tr.odd {
        background-color: #FFFFFF;
        color: black;
    }
    .excel-table .sub-row {
        background-color: #E1DFDD;
        font-size: 0.85em;
    }
    .excel-table .district-header {
        background-color: #107C10;
        color: white;
        text-align: center;
        font-weight: bold;
        padding: 2px;
        line-height: 1.2;
    }
    .excel-table .highlight-green {
        background-color: #90EE90;
    }
    .excel-table .highlight-red {
        background-color: #FFB6C1;
    }
    .button {
        background-color: #005566;
        color: white;
        padding: 8px 16px;
        border: none;
        cursor: pointer;
        margin: 10px 5px;
        display: inline-block;
        text-decoration: none;
    }
    .button:hover {
        background-color: #003f4c;
    }
    .button-container {
        text-align: center;
        margin-top: 10px;
    }
</style>
"""

@app.route('/')
def index():
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        {CSS_STYLES}
    </head>
    <body>
        <h2>Upload Excel File for Analysis</h2>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xls,.xlsx">
            <input type="submit" value="Upload and Analyze" class="button">
        </form>
    </body>
    </html>
    """

@app.route('/upload', methods=['POST'])
def upload_file():
    logger.info("Received upload request")
    if 'file' not in request.files:
        logger.error("No file uploaded")
        return jsonify({"error": "No file uploaded"}), 400
    
    file = request.files['file']
    if not file or file.filename == '':
        logger.error("No file selected")
        return jsonify({"error": "No file selected"}), 400

    filename = file.filename.lower()
    logger.debug(f"Uploaded file: {filename}")
    if not (filename.endswith('.xls') or filename.endswith('.xlsx')):
        logger.error("Invalid file format")
        return jsonify({"error": "Only .xls and .xlsx files are supported"}), 400
    
    file_extension = '.xls' if filename.endswith('.xls') else '.xlsx'
    
    try:
        result = process_excel(file.stream, file_extension)
        
        # Store results in session
        session['latest_analytic_date'] = result['latest_analytic_date']
        session['latest_attendance_data'] = result['latest_attendance_data']
        session['latest_week_display'] = result['latest_week_display']
        session['latest_district_counts'] = result['latest_district_counts']
        session['latest_main_district'] = result['latest_main_district']
        session['all_attendance_data'] = result['all_attendance_data']
        
        # Store the file stream in session (as bytes, since BytesIO is not JSON-serializable)
        result['output_stream'].seek(0)
        session['latest_file_stream'] = result['output_stream'].read()
        
        return redirect(url_for('result'))
    except Exception as e:
        logger.error(f"Processing error: {str(e)}")
        logger.debug(f"Full traceback: {traceback.format_exc()}")
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

@app.route('/result')
def result():
    # Retrieve data from session
    latest_attendance_data = session.get('latest_attendance_data')
    latest_week_display = session.get('latest_week_display', "No week data available yet")
    latest_district_counts = session.get('latest_district_counts')
    latest_main_district = session.get('latest_main_district')
    all_attendance_data = session.get('all_attendance_data', [])
    
    combined_table_html = render_combined_table(
        latest_week_display,
        latest_attendance_data,
        latest_district_counts,
        latest_main_district,
        all_attendance_data
    )
    download_button = '<a href="/download" class="button">Download Excel</a>' if 'latest_file_stream' in session else ''
    
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        {CSS_STYLES}
    </head>
    <body>
        <div class="table-wrapper">
            {combined_table_html}
        </div>
        <div class="button-container">
            <a href="/" class="button">Back to Upload Page</a>
            {download_button}
        </div>
    </body>
    </html>
    """

@app.route('/download', methods=['GET'])
def download_file():
    if 'latest_file_stream' not in session:
        return jsonify({"error": "No processed file available"}), 404
    file_stream = BytesIO(session['latest_file_stream'])
    file_stream.seek(0)
    return send_file(
        file_stream,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"analyzed_{uuid.uuid4().hex}.xlsx"
    )

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    logger.info(f"Starting server on port {port}")
    app.run(debug=False, host='0.0.0.0', port=port)
