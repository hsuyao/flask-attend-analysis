from flask import Flask, request, jsonify, send_file, redirect, url_for, session, render_template
from flask_session import Session
from io import BytesIO
import uuid
import os
from celery import Celery
from config import db, COLLECTION_NAME, DB_TYPE
from excel_handler import process_excel
from render_table import render_attendance_table
from database import init_database, get_six_month_averages
from utils import parse_district
from datetime import datetime
import logging

app = Flask(__name__)
app.config['SESSION_TYPE'] = 'filesystem'
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'your-secret-key-here')
app.config['CELERY_BROKER_URL'] = os.getenv('CELERY_BROKER_URL', 'redis://localhost:6379/0')
app.config['CELERY_RESULT_BACKEND'] = os.getenv('CELERY_RESULT_BACKEND', 'redis://localhost:6379/0')
Session(app)

# Configure Celery
celery = Celery(app.name, broker=app.config['CELERY_BROKER_URL'])
celery.conf.update(app.config)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def get_version_info():
    try:
        with open('/app/version_info.txt', 'r') as f:
            return f.read().strip()
    except Exception:
        return "Unknown-Unknown"

@celery.task(bind=True)
def process_excel_task(self, file_content, file_extension):
    """Background task to process Excel file"""
    logger.info("Starting Excel processing task")
    buffered_stream = BytesIO(file_content)
    try:
        self.update_state(state='PROGRESS', meta={'stage': 'Parsing Excel', 'progress': 20})
        result = process_excel(buffered_stream, file_extension)
        self.update_state(state='PROGRESS', meta={'stage': 'Writing to database', 'progress': 60})
        # Database writing is handled within process_excel
        self.update_state(state='PROGRESS', meta={'stage': 'Finalizing', 'progress': 90})
        logger.info("Excel processing task completed successfully")
        return result
    except Exception as e:
        logger.error(f"Task failed: {str(e)}")
        self.update_state(state='FAILURE', meta={'stage': 'Error', 'error': str(e)})
        raise

@app.route('/')
def index():
    return render_template('index.html', version=get_version_info())

@app.route('/upload', methods=['POST'])
def upload_file():
    logger.info("Received upload request")
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files['file']
    if not file or file.filename == '':
        return jsonify({"error": "No file selected"}), 400

    filename = file.filename.lower()
    if not filename.endswith(('.xls', '.xlsx')):
        return jsonify({"error": "Only .xls and .xlsx files supported"}), 400

    file_extension = '.xls' if filename.endswith('.xls') else '.xlsx'
    file_content = file.stream.read()

    try:
        task = process_excel_task.apply_async(args=[file_content, file_extension])
        logger.info(f"Started Celery task: {task.id}")
        return jsonify({"task_id": task.id}), 202
    except Exception as e:
        logger.error(f"Failed to start task: {str(e)}")
        return jsonify({"error": f"Task initiation failed: {str(e)}"}), 500

@app.route('/task_status/<task_id>')
def task_status(task_id):
    task = process_excel_task.AsyncResult(task_id)
    if task.state == 'PENDING':
        response = {'state': task.state, 'stage': 'Waiting', 'progress': 0}
    elif task.state == 'PROGRESS':
        response = {'state': task.state, 'stage': task.info.get('stage', 'Processing'), 'progress': task.info.get('progress', 0)}
    elif task.state == 'SUCCESS':
        response = {'state': task.state, 'stage': 'Completed', 'progress': 100, 'result': task.get()}
    elif task.state == 'FAILURE':
        response = {'state': task.state, 'stage': 'Failed', 'progress': 0, 'error': task.info.get('error', 'Unknown error')}
    else:
        response = {'state': task.state, 'stage': 'Unknown', 'progress': 0}
    logger.debug(f"Task status: {task_id} - {response}")
    return jsonify(response)

@app.route('/result')
def result():
    latest_attendance_data = session.get('latest_attendance_data')
    latest_week_display = session.get('latest_week_display', "No week data")
    latest_district_counts = session.get('latest_district_counts')
    latest_main_district_counts = session.get('latest_main_district_counts')
    all_attendance_data = session.get('all_attendance_data', [])

    if not latest_attendance_data or not latest_attendance_data.get('attended'):
        return render_template('index.html', error="No valid attendance data", version=get_version_info())

    all_attendance_data.sort(key=lambda x: x[0])
    latest_date = all_attendance_data[-1][0]

    all_names = set()
    for district in latest_attendance_data['attended']:
        all_names.update(latest_attendance_data['attended'][district])
    for district in latest_attendance_data['not_attended']:
        all_names.update(latest_attendance_data['not_attended'][district])
    avg_attendance_rates = session.get('avg_attendance_rates')
    if not avg_attendance_rates:
        avg_attendance_rates = get_six_month_averages(list(all_names), latest_date)
        session['avg_attendance_rates'] = avg_attendance_rates

    attendance_table_html = render_attendance_table(
        latest_week_display, latest_attendance_data, all_attendance_data,
        latest_district_counts, latest_main_district_counts, avg_attendance_rates
    )

    week_options = [(week_name, idx) for idx, (_, _, week_name) in enumerate(all_attendance_data)]
    return render_template(
        'result.html',
        attendance_table_html=attendance_table_html,
        stats_table_html="",
        has_file_stream=True,
        week_options=week_options,
        selected_week_idx=len(all_attendance_data) - 1 if all_attendance_data else 0,
        version=get_version_info()
    )

@app.route('/get_week_data/<int:week_idx>')
def get_week_data(week_idx):
    all_attendance_data = session.get('all_attendance_data', [])
    if not all_attendance_data or week_idx < 0 or week_idx >= len(all_attendance_data):
        return jsonify({
            'attendance_table': '<div class="district-section"><table class="excel-table"><tr class="title-row"><th>No data</th></tr></table></div>'
        }), 400

    date, attendance_data, week_name = all_attendance_data[week_idx]
    latest_main_district = session.get('latest_main_district', '')
    _, district_counts, _, main_district, main_district_counts = classify_attendance_for_week(all_attendance_data[week_idx])
    if not main_district:
        main_district = latest_main_district

    all_names = set()
    for district in attendance_data['attended']:
        all_names.update(attendance_data['attended'][district])
    for district in attendance_data['not_attended']:
        all_names.update(attendance_data['not_attended'][district])
    avg_attendance_rates = session.get('avg_attendance_rates')
    if not avg_attendance_rates:
        avg_attendance_rates = get_six_month_averages(list(all_names), date)
        session['avg_attendance_rates'] = avg_attendance_rates

    attendance_table_html = render_attendance_table(
        week_name, attendance_data, all_attendance_data,
        district_counts, main_district_counts, avg_attendance_rates
    )

    return jsonify({'attendance_table': attendance_table_html})

def classify_attendance_for_week(week_data):
    date, data, _ = week_data
    attended = data['attended']
    not_attended = data['not_attended']
    main_district = None
    district_counts = {}
    main_district_counts = {}
    age_categories = ['青職以上', '大專', '中學', '大學', '小學', '學齡前']

    if DB_TYPE == "mongodb":
        records = db[COLLECTION_NAME].find({"date": date.strftime('%Y-%m-%d')})
        age_mapping = {(record["district"], record["name"]): record["age_group"] for record in records}
    elif DB_TYPE == "sqlite":
        cursor = db.cursor()
        query = """
            SELECT district, name, age_group
            FROM attendance_records
            WHERE date = ?
        """
        cursor.execute(query, (date.strftime('%Y-%m-%d'),))
        age_mapping = {
            (row["district"], row["name"]): row["age_group"]
            for row in cursor.fetchall()
        }

    for district in set(attended.keys()).union(not_attended.keys()):
        main_district_value = parse_district(district)[0]
        if not main_district:
            main_district = main_district_value
        district_counts[district] = {'total': 0, 'ages': {age: 0 for age in age_categories}}
        main_district_counts[main_district_value] = {'total': 0, 'ages': {age: 0 for age in age_categories}}

    for district, names in attended.items():
        main_district_value = parse_district(district)[0]
        for name in names:
            effective_age = age_mapping.get((district, name), '青職以上')
            district_counts[district]['total'] += 1
            main_district_counts[main_district_value]['total'] += 1
            district_counts[district]['ages'][effective_age] += 1
            main_district_counts[main_district_value]['ages'][effective_age] += 1

    total_attendance = sum(d['total'] for d in district_counts.values())
    district_counts['總計'] = total_attendance
    return attended, district_counts, not_attended, main_district, main_district_counts

@app.route('/download', methods=['GET'])
def download_file():
    logger.info("Received download request")
    all_attendance_data = session.get('all_attendance_data', [])
    if not all_attendance_data:
        return jsonify({"error": "No data available"}), 404

    try:
        file_stream = generate_excel(all_attendance_data)
        file_stream.seek(0)
        download_name = f"analyzed_{uuid.uuid4().hex}.xlsx"
        return send_file(
            file_stream,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=download_name
        )
    except Exception as e:
        logger.error(f"Failed to generate Excel: {str(e)}")
        return jsonify({"error": f"Download failed: {str(e)}"}), 500

if __name__ == '__main__':
    init_database()
    port = int(os.getenv('PORT', 5000))
    logger.info(f"Starting server on port {port}")
    app.run(debug=False, host='0.0.0.0', port=port)
