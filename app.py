from flask import Flask, request, jsonify, send_file, redirect, url_for, session, render_template
from flask_session import Session
from io import BytesIO
import uuid
import os
from concurrent.futures import ThreadPoolExecutor
from excel_handler import process_excel, generate_excel
from render_table import render_attendance_table
from database import init_database, get_six_month_averages, get_event_name
from config import db, COLLECTION_NAME
from utils import parse_district, chinese_to_int, parse_week_display
from datetime import datetime
import logging

app = Flask(__name__)
app.config['SESSION_TYPE'] = 'filesystem'
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'your-secret-key-here')
Session(app)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Log application startup
logger.info("Starting Flask application")

# Initialize database
try:
    init_database()
    db.command("ping")  # Test MongoDB connection
    logger.info("Successfully connected to MongoDB")
except Exception as e:
    logger.error(f"Failed to initialize database: {str(e)}")
    raise

# Thread pool for background tasks
executor = ThreadPoolExecutor(max_workers=2)
tasks = {}  # Store task status and results

def get_version_info():
    # Read version info from file
    try:
        with open('/app/version_info.txt', 'r') as f:
            return f.read().strip()
    except Exception:
        return "Unknown-Unknown"

def process_excel_task(file_content, file_extension, task_id):
    # Process Excel file in background thread
    logger.info(f"Starting Excel processing task {task_id}")
    buffered_stream = BytesIO(file_content)
    try:
        tasks[task_id] = {'state': 'PROGRESS', 'stage': 'Parsing Excel', 'progress': 20}
        result = process_excel(buffered_stream, file_extension)
        tasks[task_id] = {
            'state': 'SUCCESS', 
            'stage': 'Completed', 
            'progress': 100, 
            'result': result
        }
        logger.info(f"Excel processing task {task_id} completed successfully")
    except Exception as e:
        logger.error(f"Task {task_id} failed: {str(e)}")
        tasks[task_id] = {
            'state': 'FAILURE', 
            'stage': 'Error', 
            'progress': 0, 
            'error': str(e)
        }

@app.route('/')
def index():
    return render_template('index.html', version=get_version_info())

@app.route('/upload', methods=['POST'])
def upload_file():
    # Handle file upload and start background task
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
    task_id = str(uuid.uuid4())

    try:
        tasks[task_id] = {'state': 'PENDING', 'stage': 'Waiting', 'progress': 0}
        executor.submit(process_excel_task, file_content, file_extension, task_id)
        logger.info(f"Started background task: {task_id}")
        return jsonify({"task_id": task_id}), 202
    except Exception as e:
        logger.error(f"Failed to start task: {str(e)}")
        return jsonify({"error": f"Task initiation failed: {str(e)}"}), 500

@app.route('/task_status/<task_id>')
def task_status(task_id):
    # Check status of background task
    task = tasks.get(task_id, {'state': 'PENDING', 'stage': 'Not Found', 'progress': 0})
    logger.debug(f"Task status: {task_id} - {task}")
    return jsonify(task)

@app.route('/task_result/<task_id>')
def task_result(task_id):
    # Retrieve task result and store in session
    task = tasks.get(task_id)
    if not task:
        return jsonify({"status": "error", "error": "Task not found"}), 404
    if task['state'] == 'SUCCESS':
        result = task['result']
        logger.info(f"Task {task_id} result: {result}")
        session['latest_attendance_data'] = result.get('latest_attendance_data')
        session['latest_week_display'] = result.get('latest_week_display')
        session['latest_district_counts'] = result.get('latest_district_counts')
        session['latest_main_district_counts'] = result.get('latest_main_district_counts')
        session['all_attendance_data'] = result.get('all_attendance_data')
        session['latest_main_district'] = result.get('latest_main_district')
        session['event_name'] = result.get('event_name')
        logger.info(f"Session updated with task result")
        return jsonify({"status": "success", "redirect": url_for('result')})
    elif task['state'] == 'FAILURE':
        logger.error(f"Task {task_id} failed: {task['error']}")
        return jsonify({"status": "error", "error": task['error']}), 500
    else:
        return jsonify({"status": "pending"}), 202

@app.route('/result')
def result():
    # Render result page with attendance data
    logger.info(f"Session contents: {session}")
    latest_attendance_data = session.get('latest_attendance_data')
    latest_week_display = session.get('latest_week_display', "No week data")
    latest_district_counts = session.get('latest_district_counts')
    latest_main_district_counts = session.get('latest_main_district_counts')
    all_attendance_data = session.get('all_attendance_data', [])
    event_name = session.get('event_name', "未指定活動")

    if not latest_attendance_data or not latest_attendance_data.get('attended'):
        logger.error("No valid attendance data found in session")
        return render_template('index.html', error="No valid attendance data", version=get_version_info())

    all_attendance_data.sort(key=lambda x: parse_week_display(x[2]))  # Sort by parsed week_display
    placeholder_date = datetime.now()

    all_names = set()
    for district in latest_attendance_data['attended']:
        all_names.update(latest_attendance_data['attended'][district])
    for district in latest_attendance_data['not_attended']:
        all_names.update(latest_attendance_data['not_attended'][district])
    avg_attendance_rates = session.get('avg_attendance_rates')
    if not avg_attendance_rates:
        avg_attendance_rates = get_six_month_averages(list(all_names), placeholder_date)
        session['avg_attendance_rates'] = avg_attendance_rates

    attendance_table_html = render_attendance_table(
        latest_week_display, latest_attendance_data, all_attendance_data,
        latest_district_counts, latest_main_district_counts, avg_attendance_rates,
        event_name=event_name
    )

    week_options = [(week_name, idx) for idx, (_, _, week_name, _) in enumerate(all_attendance_data)]
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
    # Fetch attendance data for a specific week
    all_attendance_data = session.get('all_attendance_data', [])
    if not all_attendance_data or week_idx < 0 or week_idx >= len(all_attendance_data):
        return jsonify({
            'attendance_table': '<div class="district-section"><table class="excel-table"><tr class="title-row"><th>No data</th></tr></table></div>'
        }), 400

    all_attendance_data.sort(key=lambda x: parse_week_display(x[2]))  # Sort by parsed week_display
    date, attendance_data, week_name, event_name = all_attendance_data[week_idx]
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
        avg_attendance_rates = get_six_month_averages(list(all_names), datetime.now())
        session['avg_attendance_rates'] = avg_attendance_rates

    attendance_table_html = render_attendance_table(
        week_name, attendance_data, all_attendance_data,
        district_counts, main_district_counts, avg_attendance_rates,
        event_name=event_name
    )

    return jsonify({'attendance_table': attendance_table_html})

def classify_attendance_for_week(week_data):
    # Classify attendance data for a specific week
    date, data, week_display, event_name = week_data
    attended = data['attended']
    not_attended = data['not_attended']
    main_district = None
    district_counts = {}
    main_district_counts = {}
    age_categories = ['青職以上', '大專', '中學', '大學', '小學', '學齡前']

    records = db[COLLECTION_NAME].find({"week_display": week_display})
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
    # Download processed Excel file
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
    # Render history page with available main districts
    try:
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
    # Get available weeks for a given main district
    try:
        pipeline = [
            {"$match": {
                "district": {"$regex": f"^{district}"},
                "attended": 1
            }},
            {"$group": {
                "_id": "$week_display"
            }}
        ]
        weeks = [doc["_id"] for doc in db[COLLECTION_NAME].aggregate(pipeline)]
        weeks.sort(key=parse_week_display)  # Sort by parsed week_display
        logger.info(f"Loaded {len(weeks)} weeks for district {district}")
        return jsonify({"weeks": [{"date": week, "display": week} for week in weeks]})
    except Exception as e:
        logger.error(f"Failed to get weeks for district {district}: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/get_history_data/<district>/<path:week_display>')
def get_history_data(district, week_display):
    # Get attendance data for a specific district and week
    try:
        records = db[COLLECTION_NAME].find({
            "district": {"$regex": f"^{district}"},
            "week_display": week_display
        })
        records_list = list(records)
        logger.debug(f"Fetched {len(records_list)} records for {district} on {week_display}")

        if not records_list:
            logger.warning(f"No records found for {district} on {week_display}")
            return jsonify({
                'attendance_table': '<div class="district-section"><table class="excel-table"><tr class="title-row"><th>無資料</th></tr></table></div>'
            }), 404

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
            attended_status = record.get("attended")
            if isinstance(attended_status, str):
                attended_status = int(attended_status)
            logger.debug(f"Record: {name}, district: {sub_district}, attended: {attended_status}")
            if attended_status == 1:
                attended.setdefault(sub_district, []).append(name)
            else:
                not_attended.setdefault(sub_district, []).append(name)

        for sub_district in set(attended.keys()).union(not_attended.keys()):
            district_counts[sub_district] = {'total': 0, 'ages': {age: 0 for age in age_categories}}
            main_district_counts[district] = {'total': 0, 'ages': {age: 0 for age in age_categories}}

        for sub_district, names in attended.items():
            for name in names:
                effective_age = age_mapping.get((sub_district, name), '青職以上')
                district_counts[sub_district]['total'] += 1
                main_district_counts[district]['total'] += 1
                district_counts[sub_district]['ages'][effective_age] += 1
                main_district_counts[district]['ages'][effective_age] += 1

        total_attendance = sum(d['total'] for d in district_counts.values())
        district_counts['總計'] = total_attendance

        # Get all week_displays before the current one, sorted by parsed week_display
        all_weeks = db[COLLECTION_NAME].distinct("week_display", {
            "district": {"$regex": f"^{district}"}
        })
        all_weeks = [w for w in all_weeks if parse_week_display(w) < parse_week_display(week_display)]
        all_weeks.sort(key=parse_week_display, reverse=True)
        prev_week = all_weeks[0] if all_weeks else None
        logger.debug(f"Previous week for {week_display}: {prev_week}, available weeks: {all_weeks}")

        prev_attended = {}
        prev_not_attended = {}
        if prev_week:
            prev_records = db[COLLECTION_NAME].find({
                "district": {"$regex": f"^{district}"},
                "week_display": prev_week
            })
            for record in prev_records:
                sub_district = record["district"]
                name = record["name"]
                attended_status = record.get("attended")
                if isinstance(attended_status, str):
                    attended_status = int(attended_status)
                if attended_status == 1:
                    prev_attended.setdefault(sub_district, []).append(name)
                else:
                    prev_not_attended.setdefault(sub_district, []).append(name)
            logger.debug(f"Previous week {prev_week} data: attended={prev_attended}, not_attended={prev_not_attended}")

        attendance_data = {'attended': attended, 'not_attended': not_attended}
        all_attendance_data = [(datetime.now(), attendance_data, week_display, None)]
        prev_attendance_data = {
            'attended': prev_attended,
            'not_attended': prev_not_attended
        }
        if prev_week and (prev_attended or prev_not_attended):
            all_attendance_data.insert(0, (datetime.now(), prev_attendance_data, prev_week, None))

        all_names = set()
        for sub_district in attended:
            all_names.update(attended[sub_district])
        for sub_district in not_attended:
            all_names.update(not_attended[sub_district])
        avg_attendance_rates = get_six_month_averages(list(all_names), datetime.now())

        # Get event name from database
        event_name = get_event_name(week_display)
        logger.debug(f"Event name for {week_display} in district {district}: {event_name}")

        attendance_table_html = render_attendance_table(
            week_display, attendance_data, all_attendance_data,
            district_counts, main_district_counts, avg_attendance_rates,
            event_name=event_name
        )

        logger.info(f"Rendered history data for {district} on {week_display}")
        return jsonify({'attendance_table': attendance_table_html})
    except Exception as e:
        logger.error(f"Failed to get history data for {district} on {week_display}: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    logger.info(f"Starting server on port {port}")
    app.run(debug=False, host='0.0.0.0', port=port)
