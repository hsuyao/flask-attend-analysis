from flask import Flask, request, jsonify, send_file, redirect, url_for, session, render_template
from flask_session import Session
from io import BytesIO
import uuid
import os
from celery import Celery
from config import db, COLLECTION_NAME, DB_TYPE
from excel_handler import process_excel
from render_table import render_attendance_table, render_stats_table
from database import init_database, get_six_month_averages
from utils import parse_district, chinese_to_int
from datetime import datetime, timedelta
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

@app.route('/task_result/<task_id>')
def task_result(task_id):
    task = process_excel_task.AsyncResult(task_id)
    if task.state == 'SUCCESS':
        result = task.get()
        logger.info(f"Task {task_id} result: {result}")
        session['latest_attendance_data'] = result.get('latest_attendance_data')
        session['latest_week_display'] = result.get('latest_week_display')
        session['latest_district_counts'] = result.get('latest_district_counts')
        session['latest_main_district_counts'] = result.get('latest_main_district_counts')
        session['all_attendance_data'] = result.get('all_attendance_data')
        session['latest_main_district'] = result.get('latest_main_district')
        logger.info(f"Session updated with task result: {session.get('latest_attendance_data')}")
        return jsonify({"status": "success", "redirect": url_for('result')})
    elif task.state == 'FAILURE':
        logger.error(f"Task {task_id} failed: {task.info.get('error')}")
        return jsonify({"status": "error", "error": task.info.get('error')}), 500
    else:
        return jsonify({"status": "pending"}), 202

@app.route('/result')
def result():
    logger.info(f"Session contents: {session}")
    latest_attendance_data = session.get('latest_attendance_data')
    logger.info(f"latest_attendance_data: {latest_attendance_data}")
    latest_week_display = session.get('latest_week_display', "No week data")
    latest_district_counts = session.get('latest_district_counts')
    latest_main_district_counts = session.get('latest_main_district_counts')
    all_attendance_data = session.get('all_attendance_data', [])
    logger.info(f"all_attendance_data length: {len(all_attendance_data)}")

    if not latest_attendance_data or not latest_attendance_data.get('attended'):
        logger.error("No valid attendance data found in session")
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

    records = db[COLLECTION_NAME].find({"date": date.strftime('%Y-%m-%d')})
    age_mapping = {(record["district"], record["name"]): record["age_group"] for record in records}

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

@app.route('/history')
def history():
    """Render history page with available main districts"""
    try:
        # Get distinct main districts from database
        districts = db[COLLECTION_NAME].distinct("district")
        main_districts = sorted(
            set(parse_district(d)[0] for d in districts if parse_district(d)[0]),
            key=lambda x: chinese_to_int(x[0])
        )
        logger.info(f"Loaded main districts: {main_districts}")
        return render_template(
            'history.html',
            main_districts=main_districts,
            version=get_version_info()
        )
    except Exception as e:
        logger.error(f"Failed to load history page: {str(e)}")
        return render_template(
            'index.html',
            error="無法載入歷史紀錄頁面",
            version=get_version_info()
        )

@app.route('/get_weeks_for_district/<district>')
def get_weeks_for_district(district):
    """Get available weeks for a given main district with at least one attendance"""
    try:
        # Find dates with at least one attended record for the main district
        pipeline = [
            {"$match": {
                "district": {"$regex": f"^{district}"},
                "attended": 1
            }},
            {"$group": {
                "_id": "$date"
            }},
            {"$sort": {"_id": -1}}
        ]
        dates = [doc["_id"] for doc in db[COLLECTION_NAME].aggregate(pipeline)]
        logger.debug(f"Found {len(dates)} dates with attendance for district {district}")

        # Convert dates to display format
        weeks = []
        for date_str in dates:
            date = datetime.strptime(date_str, '%Y-%m-%d')
            month = date.strftime('%Y年%m月')
            week_num = (date.day - 1) // 7 + 1
            week_display = f"{month}第{chinese_to_int_reverse(week_num)}週"
            weeks.append({"date": date_str, "display": week_display})
        
        logger.info(f"Loaded {len(weeks)} weeks for district {district}")
        return jsonify({"weeks": weeks})
    except Exception as e:
        logger.error(f"Failed to get weeks for district {district}: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/get_history_data/<district>/<week_date>')
def get_history_data(district, week_date):
    """Get attendance data for a specific district and week"""
    try:
        # Convert week_date to datetime
        date = datetime.strptime(week_date, '%Y-%m-%d')
        month = date.strftime('%Y年%m月')
        week_num = (date.day - 1) // 7 + 1
        week_display = f"{month}第{chinese_to_int_reverse(week_num)}週"

        # Fetch attendance data from database
        records = db[COLLECTION_NAME].find({
            "district": {"$regex": f"^{district}"},
            "date": week_date
        })
        records_list = list(records)
        logger.debug(f"Fetched {len(records_list)} records for {district} on {week_date}")

        # Organize data
        attended = {}
        not_attended = {}
        district_counts = {}
        main_district_counts = {}
        age_categories = ['青職以上', '大專', '中學', '大學', '小學', '學齡前']
        age_mapping = {}

        for record in records_list:
            sub_district = record["district"]
            name = record["name"]
            age_mapping[(sub_district, name)] = record.get("age_group", "青職以上")
            # Handle potential string or integer values for 'attended'
            attended_status = record.get("attended")
            if isinstance(attended_status, str):
                attended_status = int(attended_status)
            logger.debug(f"Record: {name}, district: {sub_district}, attended: {attended_status}")
            if attended_status == 1:
                attended.setdefault(sub_district, []).append(name)
            else:
                not_attended.setdefault(sub_district, []).append(name)

        # Log attendance data
        logger.debug(f"Attended: {attended}")
        logger.debug(f"Not attended: {not_attended}")

        # Calculate district counts based only on attended
        all_districts = set(attended.keys()).union(not_attended.keys())
        logger.debug(f"All districts: {all_districts}")
        for sub_district in all_districts:
            district_counts[sub_district] = {'total': 0, 'ages': {age: 0 for age in age_categories}}
            main_district_counts[district] = {'total': 0, 'ages': {age: 0 for age in age_categories}}

        # Count only attended
        for sub_district, names in attended.items():
            for name in names:
                effective_age = age_mapping.get((sub_district, name), '青職以上')
                district_counts[sub_district]['total'] += 1
                main_district_counts[district]['total'] += 1
                district_counts[sub_district]['ages'][effective_age] += 1
                main_district_counts[district]['ages'][effective_age] += 1

        total_attendance = sum(d['total'] for d in district_counts.values())
        district_counts['總計'] = total_attendance
        logger.debug(f"District counts: {district_counts}")
        logger.debug(f"Main district counts: {main_district_counts}")

        # Get previous week's data for highlight comparison
        # Find the most recent date before the current week_date
        prev_records = db[COLLECTION_NAME].aggregate([
            {"$match": {
                "district": {"$regex": f"^{district}"},
                "date": {"$lt": week_date}
            }},
            {"$group": {
                "_id": "$date"
            }},
            {"$sort": {"_id": -1}},
            {"$limit": 1}
        ])
        prev_date = None
        for doc in prev_records:
            prev_date = doc["_id"]
            break
        logger.debug(f"Previous week date for {week_date}: {prev_date}")

        prev_attended = set()
        prev_not_attended = set()
        if prev_date:
            prev_records = db[COLLECTION_NAME].find({
                "district": {"$regex": f"^{district}"},
                "date": prev_date
            })
            for record in prev_records:
                attended_status = record.get("attended")
                if isinstance(attended_status, str):
                    attended_status = int(attended_status)
                if attended_status == 1:
                    prev_attended.add((record["district"], record["name"]))
                else:
                    prev_not_attended.add((record["district"], record["name"]))
            logger.debug(f"Previous week attended: {prev_attended}")
            logger.debug(f"Previous week not attended: {prev_not_attended}")

        # Prepare all_attendance_data for rendering
        attendance_data = {'attended': attended, 'not_attended': not_attended}
        all_attendance_data = [(date, attendance_data, week_display)]
        prev_attendance_data = {
            'attended': {k: [n for d, n in prev_attended if d == k] for k in all_districts},
            'not_attended': {k: [n for d, n in prev_not_attended if d == k] for k in all_districts}
        }
        if prev_date and (prev_attended or prev_not_attended):
            all_attendance_data.insert(0, (datetime.strptime(prev_date, '%Y-%m-%d'), prev_attendance_data, ""))

        # Calculate attendance rates
        all_names = set()
        for sub_district in attended:
            all_names.update(attended[sub_district])
        for sub_district in not_attended:
            all_names.update(not_attended[sub_district])
        avg_attendance_rates = get_six_month_averages(list(all_names), date)
        logger.debug(f"Average attendance rates: {avg_attendance_rates}")

        # Render table
        attendance_table_html = render_attendance_table(
            week_display, attendance_data, all_attendance_data,
            district_counts, main_district_counts, avg_attendance_rates
        )

        logger.info(f"Rendered history data for {district} on {week_date}")
        return jsonify({'attendance_table': attendance_table_html})
    except Exception as e:
        logger.error(f"Failed to get history data for {district} on {week_date}: {str(e)}")
        return jsonify({"error": str(e)}), 500

def chinese_to_int_reverse(num):
    """Convert integer to Chinese numeral"""
    numeral_map = {1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八', 9: '九', 10: '十'}
    return numeral_map.get(num, str(num))

if __name__ == '__main__':
    init_database()
    port = int(os.getenv('PORT', 5000))
    logger.info(f"Starting server on port {port}")
    app.run(debug=False, host='0.0.0.0', port=port)
