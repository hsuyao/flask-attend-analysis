# app.py
from flask import Flask, request, jsonify, send_file, redirect, url_for
import uuid
import os
import traceback
from config import logger, state
from excel_handler import process_excel
from render_table import render_combined_table

app = Flask(__name__)

@app.route('/')
def index():
    latest_date_display = state.latest_analytic_date if state.latest_analytic_date else "No analytics available yet"
    week_display = state.latest_week_display if state.latest_week_display else "No week data available yet"
    
    combined_table_html = render_combined_table(week_display)
    download_button = '<form action="/download" method="get"><input type="submit" value="Download Processed XLS" class="button"></form>' if state.latest_file_stream else ''
    
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            .table-wrapper {{
                overflow-x: auto;
                margin: 20px auto;
            }}
            .excel-table {{
                border-collapse: collapse;
                font-family: Arial, sans-serif;
                white-space: nowrap;
            }}
            .excel-table th, .excel-table td {{
                border: 1px solid #000;
                padding: 2px;
                text-align: left;
                vertical-align: top;
                min-width: 70px;
                line-height: 1.2;
            }}
            .excel-table .separator {{
                min-width: 10px;
                width: 10px;
            }}
            .excel-table .title-row th {{
                background-color: #005566;
                color: white;
                text-align: center;
                font-weight: bold;
                padding: 2px;
                line-height: 1.2;
            }}
            .excel-table .header th {{
                background-color: #107C10;
                color: white;
                padding: 2px;
                line-height: 1.2;
            }}
            .excel-table .subheader th {{
                background-color: #5DBB63;
                color: white;
                padding: 2px;
                line-height: 1.2;
            }}
            .excel-table tr.even {{
                background-color: #F3F2F1;
                color: black;
            }}
            .excel-table tr.odd {{
                background-color: #FFFFFF;
                color: black;
            }}
            .excel-table .sub-row {{
                background-color: #E1DFDD;
                font-size: 0.85em;
            }}
            .excel-table .district-header {{
                background-color: #107C10;
                color: white;
                text-align: center;
                font-weight: bold;
                padding: 2px;
                line-height: 1.2;
            }}
            .excel-table .highlight-green {{
                background-color: #90EE90;
            }}
            .excel-table .highlight-red {{
                background-color: #FFB6C1;
            }}
            .button {{
                background-color: #005566;
                color: white;
                padding: 8px 16px;
                border: none;
                cursor: pointer;
                margin-top: 10px;
            }}
            .button:hover {{
                background-color: #003f4c;
            }}
        </style>
    </head>
    <body>
        <h2>Upload Excel File for Analysis</h2>
        <p>Latest Analytic Date: {latest_date_display}</p>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xls,.xlsx">
            <input type="submit" value="Upload and Analyze" class="button">
        </form>
        {download_button}
        <h3>Latest Attendance Data</h3>
        <div class="table-wrapper">
            {combined_table_html}
        </div>
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
        output_stream = process_excel(file.stream, file_extension)
        return redirect(url_for('index'))
    except Exception as e:
        logger.error(f"Processing error: {str(e)}")
        logger.debug(f"Full traceback: {traceback.format_exc()}")
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

@app.route('/download', methods=['GET'])
def download_file():
    if state.latest_file_stream is None:
        return jsonify({"error": "No processed file available"}), 404
    state.latest_file_stream.seek(0)
    return send_file(
        state.latest_file_stream,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"analyzed_{uuid.uuid4().hex}.xlsx"
    )

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    logger.info(f"Starting server on port {port}")
    app.run(debug=False, host='0.0.0.0', port=port)
