from flask import Flask, request, send_file, jsonify
from flask import Response
import aspose.cells as ac
from io import BytesIO
import uuid
import argparse
import os
import logging
import traceback

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)

START_COLUMN = 8

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
    
    for row in range(2, max_row + 1):  # Start from row 2 (data)
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

    # Write district headers
    for i, district in enumerate(districts):
        sheet.cells.get(row, i * 2).value = district
        sheet.cells.get(row, i * 2 + 1).value = district
        sheet.cells.get(row + 1, i * 2).value = "本週到會"  # "Attended this week"
        sheet.cells.get(row + 1, i * 2 + 1).value = "未到會"  # "Not attended"

    # Determine the maximum number of rows needed for names
    max_len = max(max(len(attended.get(d, [])), len(not_attended.get(d, []))) for d in districts)

    # Write attendance lists
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

    # Detect month transitions and week columns
    week_cols = []
    current_month = "2025年1月"  # Default to January
    for col in range(START_COLUMN, input_sheet.cells.max_column + 1):
        month_header = str(input_sheet.cells.get(0, col).value or "")
        week_header = str(input_sheet.cells.get(1, col).value or "")
        if "2025年" in month_header:
            current_month = month_header.strip()  # Update month prefix
        if "週" in week_header:
            week_cols.append((col, week_header, current_month))

    logger.info(f"Detected week columns with months: {week_cols}")

    if not week_cols:
        logger.warning("No week columns detected; output will lack analytic sheets")

    for col, week_name, month_prefix in week_cols:
        logger.info(f"Processing week: {week_name} in {month_prefix}")
        attended, not_attended = classify_attendance(input_sheet, col)

        # Skip sheet creation if no one attended
        if not any(attended.values()):
            logger.info(f"No attendees for {week_name} in {month_prefix}, skipping sheet creation")
            continue

        new_sheet_name = f"{month_prefix}{week_name} 主日"

        # Check for duplicate sheet names
        existing_names = [sheet.name for sheet in workbook.worksheets]
        if new_sheet_name in existing_names:
            logger.error(f"Duplicate sheet name detected: {new_sheet_name}")
            raise ValueError(f"Sheet name '{new_sheet_name}' already exists")

        new_sheet = workbook.worksheets.add(new_sheet_name)
        logger.debug(f"Created new sheet: {new_sheet_name}")
        write_summary(new_sheet, attended, not_attended)

    # Remove "Evaluation Warning" sheet if it exists
    for sheet in workbook.worksheets:
        if sheet.name == "Evaluation Warning":
            workbook.worksheets.remove(sheet)
            logger.debug("Removed 'Evaluation Warning' sheet")
            break

    output_stream = BytesIO()
    workbook.save(output_stream, ac.SaveFormat.XLSX)
    output_stream.seek(0)
    logger.info("File processing completed successfully")
    return output_stream

@app.route('/')
def index():
    return """
    <!DOCTYPE html>
    <html>
    <body>
        <h2>Upload Excel File for Analysis</h2>
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xls,.xlsx">
            <input type="submit" value="Upload and Analyze">
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
    filename = file.filename.lower()
    logger.debug(f"Uploaded file: {filename}")
    if not (filename.endswith('.xls') or filename.endswith('.xlsx')):
        logger.error("Invalid file format")
        return jsonify({"error": "Only .xls and .xlsx files are supported"}), 400
    
    file_extension = '.xls' if filename.endswith('.xls') else '.xlsx'
    
    try:
        output_stream = process_excel(file.stream, file_extension)
        logger.info("Sending analyzed file for download")
        output_stream.seek(0)
        response = Response(output_stream.read(), mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response.headers["Content-Disposition"] = f"attachment; filename=analyzed_{uuid.uuid4().hex}.xlsx"
        return response
    except Exception as e:
        logger.error(f"Processing error: {str(e)}")
        logger.debug(f"Full traceback: {traceback.format_exc()}")
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

def get_port():
    parser = argparse.ArgumentParser(description="Flask web service for Excel analysis")
    parser.add_argument('--port', type=int, default=os.getenv('PORT', 5000), help='Port to run the server on')
    args = parser.parse_args()
    return args.port

if __name__ == '__main__':
    port = get_port()
    logger.info(f"Starting server on port {port}")
    app.run(debug=True, host='0.0.0.0', port=port)
