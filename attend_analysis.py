from flask import Flask, request, Response, jsonify, g, send_file, redirect, url_for
import aspose.cells as ac
from io import BytesIO
import uuid
import argparse
import os
import logging
import traceback
from datetime import datetime

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)

START_COLUMN = 8

# Global variables to store the latest analytic data and file
latest_analytic_date = None
latest_attendance_data = None  # {'attended': {}, 'not_attended': {}}
latest_file_stream = None

def chinese_to_int(chinese_num):
    """Convert Chinese numerals to Arabic integers."""
    numeral_map = {
        '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
        '六': 6, '七': 7, '八': 8, '九': 9, '十': 10
    }
    return numeral_map.get(chinese_num, 0)

def repair_xls(file_stream):
    file_stream.seek(0)
    file_content = file_stream.read()
    input_stream = BytesIO(file_content)
    try:
        logger.debug("Attempting to repair .xls file with Aspose.Cells")
        workbook = ac.Workbook(input_stream, ac.LoadOptions(ac.LoadFormat.EXCEL_97_TO_2003))
        output_stream = BytesIO()
        workbook.save(output_stream, ac.SaveFormat.XLSX)
        output_stream.seek(0)
        logger.info("Successfully repaired .xls file into .xlsx")
        return output_stream
    except Exception as e:
        logger.error(f"Failed to repair .xls file: {str(e)}")
        logger.debug(f"Repair traceback: {traceback.format_exc()}")
        return None

def classify_attendance(sheet, week_col):
    logger.debug(f"Classifying attendance for week column: {week_col}")
    attended = {}
    not_attended = {}
    max_row = sheet.cells.max_row
    
    for row in range(2, max_row + 1):
        district = f"{sheet.cells.get(row, 0).value}{sheet.cells.get(row, 1).value}"
        name = sheet.cells.get(row, 3).value
        if not name or not district.startswith("二大區"):
            continue
        attendance = sheet.cells.get(row, week_col).value
        if attendance == 1:
            if district not in attended:
                attended[district] = []
            attended[district].append(name)
        else:
            if district not in not_attended:
                not_attended[district] = []
            not_attended[district].append(name)
    return attended, not_attended

def write_summary(sheet, attended, not_attended):
    logger.debug(f"Writing summary with attended: {attended}, not_attended: {not_attended}")
    districts = sorted(set(attended.keys()).union(not_attended.keys()))
    row = 0

    for i, district in enumerate(districts):
        sheet.cells.get(row, i * 2).value = district
        sheet.cells.get(row, i * 2 + 1).value = district
        sheet.cells.get(row + 1, i * 2).value = "本週到會"  # "Attended this week"
        sheet.cells.get(row + 1, i * 2 + 1).value = "未到會"  # "Not attended"

    max_len = max(max(len(attended.get(d, [])), len(not_attended.get(d, []))) for d in districts)
    for r in range(max_len):
        for i, district in enumerate(districts):
            attended_list = attended.get(district, [])
            not_attended_list = not_attended.get(district, [])
            if r < len(attended_list):
                sheet.cells.get(r + 2, i * 2).value = attended_list[r]
            if r < len(not_attended_list):
                sheet.cells.get(r + 2, i * 2 + 1).value = not_attended_list[r]

    logger.debug("Summary written successfully")

def process_excel(file_stream, file_extension):
    global latest_analytic_date, latest_attendance_data, latest_file_stream
    file_stream.seek(0)
    file_content = file_stream.read()
    buffered_stream = BytesIO(file_content)
    logger.info(f"Processing file with extension: {file_extension}, Size: {len(file_content)} bytes")

    if file_extension == '.xls':
        repaired_stream = repair_xls(file_stream)
        if repaired_stream:
            buffered_stream = repaired_stream
            file_extension = '.xlsx'
            logger.debug("Using repaired .xlsx stream for analysis")
        else:
            logger.warning("Repair failed; attempting to process original .xls file")

    try:
        if file_extension == '.xls':
            workbook = ac.Workbook(buffered_stream, ac.LoadOptions(ac.LoadFormat.EXCEL_97_TO_2003))
        elif file_extension == '.xlsx':
            workbook = ac.Workbook(buffered_stream, ac.LoadOptions(ac.LoadFormat.XLSX))
        else:
            raise ValueError("Unsupported file format")
    except Exception as e:
        logger.error(f"Failed to load workbook: {str(e)}")
        raise

    input_sheet = workbook.worksheets[0]
    logger.debug(f"Loaded sheet: {input_sheet.name}, Rows: {input_sheet.cells.max_row}, Columns: {input_sheet.cells.max_column}")

    week_cols = []
    current_month = "2025年1月"
    for col in range(START_COLUMN, input_sheet.cells.max_column + 1):
        month_header = str(input_sheet.cells.get(0, col).value or "")
        week_header = str(input_sheet.cells.get(1, col).value or "")
        if "2025年" in month_header:
            current_month = month_header.strip()
        if "週" in week_header:
            week_cols.append((col, week_header, current_month))

    logger.info(f"Detected week columns with months: {week_cols}")

    if not week_cols:
        logger.warning("No week columns detected; output will lack analytic sheets")

    latest_date = None
    latest_attended = None
    latest_not_attended = None
    for col, week_name, month_prefix in week_cols:
        logger.info(f"Processing week: {week_name} in {month_prefix}")
        attended, not_attended = classify_attendance(input_sheet, col)

        if not any(attended.values()):
            logger.info(f"No attendees for {week_name} in {month_prefix}, skipping sheet creation")
            continue

        new_sheet_name = f"{month_prefix}{week_name} 主日"
        week_str = week_name.replace("第", "").replace("週", "")
        week_num = chinese_to_int(week_str)
        month_num = int(month_prefix.split("年")[1].replace("月", ""))
        current_date = datetime(2025, month_num, min(week_num * 7, 28))
        if latest_date is None or current_date > latest_date:
            latest_date = current_date
            latest_attended = attended
            latest_not_attended = not_attended

        existing_names = [sheet.name for sheet in workbook.worksheets]
        if new_sheet_name in existing_names:
            logger.error(f"Duplicate sheet name detected: {new_sheet_name}")
            raise ValueError(f"Sheet name '{new_sheet_name}' already exists")

        new_sheet = workbook.worksheets.add(new_sheet_name)
        logger.debug(f"Created new sheet: {new_sheet_name}")
        write_summary(new_sheet, attended, not_attended)

    for sheet in workbook.worksheets:
        if sheet.name == "Evaluation Warning":
            workbook.worksheets.remove(sheet)
            logger.debug("Removed 'Evaluation Warning' sheet")
            break

    if latest_date:
        latest_analytic_date = latest_date.strftime("%Y年%m月%d日")
        latest_attendance_data = {'attended': latest_attended, 'not_attended': latest_not_attended}
        logger.debug(f"Updated latest_analytic_date to: {latest_analytic_date}")

    output_stream = BytesIO()
    workbook.save(output_stream, ac.SaveFormat.XLSX)
    output_stream.seek(0)
    latest_file_stream = BytesIO(output_stream.read())
    logger.info("File processing completed successfully")
    return output_stream

@app.route('/')
def index():
    global latest_analytic_date, latest_attendance_data
    latest_date_display = latest_analytic_date if latest_analytic_date else "No analytics available yet"
    
    # Build Excel-style table HTML with districts parallel
    table_html = ""
    if latest_attendance_data:
        districts = sorted(set(latest_attendance_data['attended'].keys()).union(latest_attendance_data['not_attended'].keys()))
        max_len = max(max(len(latest_attendance_data['attended'].get(d, [])), len(latest_attendance_data['not_attended'].get(d, []))) for d in districts)
        
        table_html = """
        <div class="table-wrapper">
            <table class="excel-table">
                <tr class="header">
        """
        # District headers spanning two columns
        for district in districts:
            table_html += f'<th colspan="2">{district}</th>'
        table_html += """
                </tr>
                <tr class="subheader">
        """
        # Subheaders for Attended and Not Attended under each district
        for _ in districts:
            table_html += '<th>本週到會</th><th>未到會</th>'
        table_html += "</tr>"

        # Data rows
        for r in range(max_len):
            row_class = "even" if r % 2 == 0 else "odd"
            table_html += f'<tr class="{row_class}">'
            for district in districts:
                attended_list = latest_attendance_data['attended'].get(district, [])
                not_attended_list = latest_attendance_data['not_attended'].get(district, [])
                attended = attended_list[r] if r < len(attended_list) else ''
                not_attended = not_attended_list[r] if r < len(not_attended_list) else ''
                table_html += f'<td>{attended}</td><td>{not_attended}</td>'
            table_html += '</tr>'
        table_html += "</table></div>"

    download_button = '<form action="/download" method="get"><input type="submit" value="Download Processed XLS" class="button"></form>' if latest_file_stream else ''
    
    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            .table-wrapper {{
                overflow-x: auto;
                max-width: 100%;
                margin: 20px auto;
            }}
            .excel-table {{
                border-collapse: collapse;
                font-family: Arial, sans-serif;
                white-space: nowrap;
            }}
            .excel-table th, .excel-table td {{
                border: 1px solid #000;
                padding: 8px;
                text-align: left;
                vertical-align: top;
                min-width: 100px;
            }}
            .excel-table th {{
                background-color: #4CAF50;
                color: white;
            }}
            .excel-table .subheader th {{
                background-color: #66BB6A;
                color: white;
            }}
            .excel-table tr.even {{
                background-color: #f2f2f2;
            }}
            .excel-table tr.odd {{
                background-color: #ffffff;
            }}
            .button {{
                background-color: #008CBA;
                color: white;
                padding: 10px 20px;
                border: none;
                cursor: pointer;
                margin-top: 10px;
            }}
            .button:hover {{
                background-color: #006d9e;
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
        {table_html}
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
    filename = file.filename.lower()
    logger.debug(f"Uploaded file: {filename}")
    if not (filename.endswith('.xls') or filename.endswith('.xlsx')):
        logger.error("Invalid file format")
        return jsonify({"error": "Only .xls and .xlsx files are supported"}), 400
    
    file_extension = '.xls' if filename.endswith('.xls') else '.xlsx'
    
    try:
        process_excel(file.stream, file_extension)
        return redirect(url_for('index'))
    except Exception as e:
        logger.error(f"Processing error: {str(e)}")
        logger.debug(f"Full traceback: {traceback.format_exc()}")
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

@app.route('/download', methods=['GET'])
def download_file():
    global latest_file_stream
    if latest_file_stream is None:
        return jsonify({"error": "No processed file available"}), 404
    latest_file_stream.seek(0)
    return send_file(
        latest_file_stream,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"analyzed_{uuid.uuid4().hex}.xlsx"
    )

def get_port():
    parser = argparse.ArgumentParser(description="Flask web service for Excel analysis")
    parser.add_argument('--port', type=int, default=os.getenv('PORT', 5000), help='Port to run the server on')
    args = parser.parse_args()
    return args.port

if __name__ == '__main__':
    port = get_port()
    logger.info(f"Starting server on port {port}")
    app.run(debug=True, host='0.0.0.0', port=port)
