from flask import Flask, request, jsonify, send_file, redirect, url_for, session, render_template
from flask_session import Session
from io import BytesIO
import uuid
import os
import traceback
from config import logger
from excel_handler import process_excel
from render_table import render_attendance_table

app = Flask(__name__)

app.config['SESSION_TYPE'] = 'filesystem'
app.config['SECRET_KEY'] = 'your-secret-key-here'
Session(app)

@app.route('/')
def index():
    return render_template('index.html')

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
        
        if not result['all_attendance_data']:
            return render_template('index.html', error="上傳的文件中無任何出席紀錄，請檢查數據後重新上傳。")

        session['latest_analytic_date'] = result['latest_analytic_date']
        session['latest_attendance_data'] = result['latest_attendance_data']
        session['latest_week_display'] = result['latest_week_display']
        session['latest_district_counts'] = result['latest_district_counts']
        session['latest_main_district'] = result['latest_main_district']
        session['latest_main_district_counts'] = result['latest_main_district_counts']
        session['all_attendance_data'] = result['all_attendance_data']
        
        result['output_stream'].seek(0)
        session['latest_file_stream'] = result['output_stream'].read()
        
        return redirect(url_for('result'))
    except Exception as e:
        logger.error(f"Processing error: {str(e)}")
        logger.debug(f"Full traceback: {traceback.format_exc()}")
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

@app.route('/result')
def result():
    latest_attendance_data = session.get('latest_attendance_data')
    latest_week_display = session.get('latest_week_display', "No week data available yet")
    latest_district_counts = session.get('latest_district_counts')
    latest_main_district_counts = session.get('latest_main_district_counts')
    all_attendance_data = session.get('all_attendance_data', [])
    
    if not latest_attendance_data or not latest_attendance_data.get('attended'):
        return render_template('index.html', error="最新週無有效出席數據，請檢查文件內容。")

    all_attendance_data.sort(key=lambda x: x[0])
    
    attendance_table_html = render_attendance_table(
        latest_week_display,
        latest_attendance_data,
        all_attendance_data,
        latest_district_counts,
        latest_main_district_counts
    )
    
    week_options = [(week_name, idx) for idx, (_, _, week_name) in enumerate(all_attendance_data)]
    
    return render_template(
        'result.html',
        attendance_table_html=attendance_table_html,
        stats_table_html="",
        has_file_stream='latest_file_stream' in session,
        week_options=week_options,
        selected_week_idx=len(all_attendance_data) - 1 if all_attendance_data else 0
    )

@app.route('/get_week_data/<int:week_idx>')
def get_week_data(week_idx):
    all_attendance_data = session.get('all_attendance_data', [])
    latest_district_counts = session.get('latest_district_counts')
    latest_main_district_counts = session.get('latest_main_district_counts')
    
    if not all_attendance_data or week_idx < 0 or week_idx >= len(all_attendance_data):
        return jsonify({
            'attendance_table': '<div class="district-section"><table class="excel-table"><tr class="title-row"><th>無資料</th></tr></table></div>',
            'stats_table': ''
        }), 400
    
    _, attendance_data, week_name = all_attendance_data[week_idx]
    
    attendance_table_html = render_attendance_table(
        week_name,
        attendance_data,
        all_attendance_data,
        latest_district_counts,
        latest_main_district_counts
    )
    
    return jsonify({
        'attendance_table': attendance_table_html,
        'stats_table': ''
    })

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
