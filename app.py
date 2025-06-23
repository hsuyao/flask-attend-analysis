from flask import Flask, request, jsonify, send_file, redirect, url_for, session, render_template
from flask_session import Session
from io import BytesIO
import uuid
import os
from concurrent.futures import ThreadPoolExecutor
from excel_handler import process_excel, generate_excel
from render_table import render_attendance_table, render_average_attendance_table
from database import init_database, get_six_month_averages, get_event_name, get_event_totals, get_six_month_trimmed_mean, get_six_month_trimmed_mean_by_event
from user import init_users_collection, create_user, verify_user, create_admin_if_not_exists
from config import db, DB_OFFLINE, COLLECTION_NAME
from utils import parse_district, chinese_to_int, parse_week_display
from datetime import datetime
from urllib.parse import urlencode
import logging
from admin_routes import admin_bp

app = Flask(__name__)
app.config['SESSION_TYPE'] = 'filesystem'
app.config['SECRET_KEY'] = os.getenv('SECRET_KEY', 'your-secret-key-here')
Session(app)
app.config['DB_OFFLINE'] = DB_OFFLINE          # 供 template 使用
app.register_blueprint(admin_bp)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Log application startup
logger.info("Starting Flask application")

# ---------------- 初始化 ----------------
if not DB_OFFLINE:
    try:
        init_database()
        init_users_collection()
        db.command("ping")
        admin_username = os.getenv('ADMIN_ACCOUNT')
        admin_password = os.getenv('ADMIN_PASSWORD')
        if not (admin_username and admin_password):
            raise ValueError("Admin credentials not set")
        create_admin_if_not_exists(admin_username, admin_password)
        logger.info("資料庫初始化完成")
    except Exception as e:
        logger.error(f"DB init error: {e}")
        raise
else:
    logger.warning("⚠️  離線模式：跳過資料庫 / 使用者初始化")

# Thread pool for background tasks
executor = ThreadPoolExecutor(max_workers=2)
tasks = {}  # Store task status and results

def get_version_info():
    try:
        with open('/app/version_info.txt', 'r') as f:
            return f.read().strip()
    except Exception:
        return "Unknown-Unknown"

def is_authenticated():
    return 'user' in session and session['user'].get('username')

def is_admin():
    return is_authenticated() and session['user'].get('role') == 'admin'

def process_excel_task(file_contents, file_extensions, task_id, save_to_db=True):
    logger.info(f"Starting Excel processing task {task_id}, save_to_db={save_to_db}")
    combined_result = {
        'latest_analytic_date': None,
        'latest_attendance_data': None,
        'latest_week_display': None,
        'latest_district_counts': None,
        'latest_main_district': None,
        'latest_main_district_counts': None,
        'all_attendance_data': [],
        'event_name': "未指定活動",
        'records_written': 0,
        'week_avg_rates': {}  # Store per-week attendance rates
    }
    
    try:
        for idx, (file_content, file_extension) in enumerate(zip(file_contents, file_extensions)):
            buffered_stream = BytesIO(file_content)
            tasks[task_id] = {'state': 'PROGRESS', 'stage': f'Parsing Excel File {idx + 1}', 'progress': 20 + (idx * 20)}
            result = process_excel(buffered_stream, file_extension, save_to_db=save_to_db)
            
            # Log processing result for this file
            logger.info(f"Processed file {idx + 1}: {result.get('latest_week_display', 'No week display')}, "
                       f"records to write: {len(result.get('all_records', []))}")
            
            if result['latest_attendance_data']:
                if not combined_result['latest_attendance_data'] or parse_week_display(result['latest_week_display']) > parse_week_display(combined_result['latest_week_display'] or ''):
                    combined_result['latest_analytic_date'] = result['latest_analytic_date']
                    combined_result['latest_attendance_data'] = result['latest_attendance_data']
                    combined_result['latest_week_display'] = result['latest_week_display']
                    combined_result['latest_district_counts'] = result['latest_district_counts']
                    combined_result['latest_main_district'] = result['latest_main_district']
                    combined_result['latest_main_district_counts'] = result['latest_main_district_counts']
                    combined_result['event_name'] = result['event_name']
            
            combined_result['all_attendance_data'].extend(result['all_attendance_data'])
            combined_result['records_written'] += result.get('records_written', 0)
            combined_result['week_avg_rates'].update(result.get('week_avg_rates', {}))
        
        seen_weeks = set()
        unique_attendance_data = []
        for item in sorted(combined_result['all_attendance_data'], key=lambda x: parse_week_display(x[2])):
            if item[2] not in seen_weeks:
                seen_weeks.add(item[2])
                unique_attendance_data.append(item)
        combined_result['all_attendance_data'] = unique_attendance_data

        tasks[task_id] = {
            'state': 'SUCCESS', 
            'stage': 'Completed', 
            'progress': 100, 
            'result': combined_result
        }
        logger.info(f"Excel processing task {task_id} completed successfully, {combined_result['records_written']} records written")
    except Exception as e:
        tasks[task_id] = {
            'state': 'FAILURE',
            'stage': 'Error',
            'progress': 0,
            'error': str(e)
        }
        logger.error(f"Excel processing task {task_id} failed: {str(e)}")
        raise

@app.route('/')
def index():
    if not is_authenticated():
        return redirect(url_for('login'))
    is_anonymous = session.get('user', {}).get('role') == 'anonymous'
    return render_template('index.html', 
                           version=get_version_info(), 
                           is_anonymous=is_anonymous,
                           is_admin=is_admin)

@app.route('/login', methods=['GET', 'POST'])
def login():
    # 離線模式：只能匿名
    if DB_OFFLINE:
        if request.method == 'POST':
            # POST 直接回覆錯誤訊息
            return render_template(
                'login.html',
                error="資料庫離線，暫時僅支援匿名登入",
                version=get_version_info()
            )
        # GET 顯示同樣提示
        return render_template(
            'login.html',
            error="資料庫離線，請點選『匿名登入』",
            version=get_version_info()
        )

    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        user = verify_user(username, password)
        if user:
            session['user'] = {
                'username': user['username'],
                'role': user['role']
            }
            logger.info(f"User {username} logged in successfully")
            return redirect(url_for('index'))
        return render_template('login.html', error="無效的使用者名稱或密碼", version=get_version_info())
    return render_template('login.html', version=get_version_info())

@app.route('/anonymous_login')
def anonymous_login():
    session['user'] = {
        'username': 'anonymous',
        'role': 'anonymous'
    }
    logger.info("Anonymous user logged in")
    return redirect(url_for('index'))

@app.route('/register', methods=['GET', 'POST'])
def register():
    # 離線模式停用註冊
    if DB_OFFLINE:
        return render_template(
            'register.html',
            error="資料庫離線，無法註冊新帳號",
            version=get_version_info()
        )

    if request.method == 'POST':
        username = request.form.get('username')
        email = request.form.get('email')
        password = request.form.get('password')
        if create_user(username, email, password):
            logger.info(f"User {username} registered successfully")
            return redirect(url_for('login'))
        return render_template('register.html', error="使用者名稱已存在", version=get_version_info())
    return render_template('register.html', version=get_version_info())

@app.route('/logout')
def logout():
    session.pop('user', None)
    logger.info("User logged out")
    return redirect(url_for('login'))

@app.route('/upload', methods=['POST'])
def upload_file():
    if not is_authenticated():
        return jsonify({"error": "請先登入"}), 401
    
    is_anonymous = session.get('user', {}).get('role') == 'anonymous'
    file_keys = ['file1', 'file2', 'file3', 'file4']
    files = []
    file_extensions = []
    
    for key in file_keys:
        if key in request.files and request.files[key].filename:
            file = request.files[key]
            filename = file.filename.lower()
            if not filename.endswith(('.xls', '.xlsx')):
                return jsonify({"error": f"File {filename}: Only .xls and .xlsx files supported"}), 400
            files.append(file.stream.read())
            file_extensions.append('.xls' if filename.endswith('.xls') else '.xlsx')
    
    if not files:
        return jsonify({"error": "No files selected"}), 400

    task_id = str(uuid.uuid4())
    try:
        tasks[task_id] = {'state': 'PENDING', 'stage': 'Waiting', 'progress': 0}
        executor.submit(
            process_excel_task,
            files,
            file_extensions,
            task_id,
            save_to_db=(not is_anonymous and not DB_OFFLINE)
        )
        logger.info(f"Started background task: {task_id}")
        return jsonify({"task_id": task_id}), 202
    except Exception as e:
        logger.error(f"Failed to start task: {str(e)}")
        return jsonify({"error": f"Task initiation failed: {str(e)}"}), 500

@app.route('/task_status/<task_id>')
def task_status(task_id):
    if not is_authenticated():
        return jsonify({"error": "請先登入"}), 401
    task = tasks.get(task_id, {'state': 'PENDING', 'stage': 'Not Found', 'progress': 0})
    logger.debug(f"Task status: {task_id} - {task}")
    return jsonify(task)

@app.route('/task_result/<task_id>')
def task_result(task_id):
    if not is_authenticated():
        return jsonify({"error": "請先登入"}), 401
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
        session['week_avg_rates'] = result.get('week_avg_rates', {})
        logger.info(f"Session updated with task result, {result.get('records_written', 0)} records written")
        query = urlencode({
            'district': result.get('latest_main_district', ''),
            'week': result.get('latest_week_display', '')
        })
        redirect_url = url_for('history') + ('?' + query if query else '')
        return jsonify({"status": "success", "redirect": redirect_url})
    elif task['state'] == 'FAILURE':
        logger.error(f"Task {task_id} failed: {task['error']}")
        return jsonify({"status": "error", "error": task['error']}), 500
    else:
        return jsonify({"status": "pending"}), 202

@app.route('/result')
def result():
    if not is_authenticated():
        return redirect(url_for('login'))
    
    is_anonymous = session.get('user', {}).get('role') == 'anonymous'
    latest_attendance_data = session.get('latest_attendance_data')
    latest_week_display = session.get('latest_week_display', "No week data")
    latest_district_counts = session.get('latest_district_counts')
    latest_main_district_counts = session.get('latest_main_district_counts')
    all_attendance_data = session.get('all_attendance_data', [])
    event_name = session.get('event_name', "未指定活動")
    latest_main_district = session.get('latest_main_district', '')
    week_avg_rates = session.get('week_avg_rates', {})

    if not latest_attendance_data or not latest_attendance_data.get('attended'):
        logger.error("No valid attendance data found in session")
        return render_template('index.html', error="No valid attendance data", version=get_version_info(), is_anonymous=is_anonymous)

    sorted_attendance_data = sorted(all_attendance_data, key=lambda x: parse_week_display(x[2]), reverse=True)
    placeholder_date = datetime.now()

    sunday_data = None
    for idx, (date, data, week_name, evt_name) in enumerate(sorted_attendance_data):
        if evt_name == "主日":
            sunday_data = (idx, date, data, week_name, evt_name)
            break
    if not sunday_data and sorted_attendance_data:
        sunday_data = (0, sorted_attendance_data[0][0], sorted_attendance_data[0][1], sorted_attendance_data[0][2], sorted_attendance_data[0][3])

    if not sunday_data:
        logger.error("No 主日 or fallback data found")
        return render_template('index.html', error="No 主日 or fallback data available", version=get_version_info(), is_anonymous=is_anonymous)

    sunday_idx, sunday_date, sunday_attendance_data, sunday_week_display, sunday_event_name = sunday_data
    sunday_district_counts = latest_district_counts
    sunday_main_district_counts = latest_main_district_counts
    sunday_main_district = latest_main_district

    _, sunday_district_counts, _, sunday_main_district, sunday_main_district_counts = classify_attendance_for_week(sorted_attendance_data[sunday_idx])

    all_names = set()
    for district in sunday_attendance_data['attended']:
        all_names.update(sunday_attendance_data['attended'][district])
    for district in sunday_attendance_data['not_attended']:
        all_names.update(sunday_attendance_data['not_attended'][district])
    avg_attendance_rates = week_avg_rates.get(sunday_week_display)
    if not avg_attendance_rates and not is_anonymous:
        avg_attendance_rates = get_six_month_averages(list(all_names), sunday_week_display)
        week_avg_rates[sunday_week_display] = avg_attendance_rates
        session['week_avg_rates'] = week_avg_rates

    event_totals = get_event_totals(sunday_week_display, sunday_main_district) if not is_anonymous else {}

    attendance_table_html = render_attendance_table(
        sunday_week_display, sunday_attendance_data, sorted_attendance_data,
        sunday_district_counts, sunday_main_district_counts, avg_attendance_rates,
        event_name=sunday_event_name,
        is_history_page=True,
        event_totals=event_totals
    )
    week_options = [
        (idx, week_name) for idx, (_, _, week_name, _) in enumerate(sorted_attendance_data)
    ]
    logger.info(f"Generated week_options: {week_options}")
    return render_template(
        'result.html',
        attendance_table_html=attendance_table_html,
        has_file_stream=True,
        week_options=week_options,
        selected_week_display=sunday_week_display,   # 新增
        version=get_version_info(),
        is_anonymous=is_anonymous
)

@app.route('/get_week_data_by_name/<path:week_display>')
def get_week_data_by_name(week_display):
    if not is_authenticated():
        return jsonify({"error":"請先登入"}), 401

    all_data = session.get('all_attendance_data', [])
    # 先依週次字串找出索引
    sorted_data = sorted(all_data, key=lambda x: parse_week_display(x[2]), reverse=True)
    for idx, item in enumerate(sorted_data):
        if item[2] == week_display:
            # 直接呼叫既有函式重用邏輯
            return get_week_data(idx)
    return jsonify({"error":"週次不存在"}), 404

@app.route('/get_week_data/<int:week_idx>')
def get_week_data(week_idx):
    if not is_authenticated():
        return jsonify({"error": "請先登入"}), 401
    
    is_anonymous = session.get('user', {}).get('role') == 'anonymous'
    all_attendance_data = session.get('all_attendance_data', [])
    week_avg_rates = session.get('week_avg_rates', {})
    if not all_attendance_data or week_idx < 0 or week_idx >= len(all_attendance_data):
        return jsonify({
            'attendance_table': '<div class="district-section"><table class="excel-table"><tr class="title-row"><th>No data</th></tr></table></div>'
        }), 400

    sorted_attendance_data = sorted(all_attendance_data, key=lambda x: parse_week_display(x[2]), reverse=True)
    date, attendance_data, week_name, event_name = sorted_attendance_data[week_idx]
    logger.info(f"Fetching data for week_idx={week_idx}, week_name={week_name}")
    latest_main_district = session.get('latest_main_district', '')
    _, district_counts, _, main_district, main_district_counts = classify_attendance_for_week(sorted_attendance_data[week_idx])
    if not main_district:
        main_district = latest_main_district

    all_names = set()
    for district in attendance_data['attended']:
        all_names.update(attendance_data['attended'][district])
    for district in attendance_data['not_attended']:
        all_names.update(attendance_data['not_attended'][district])
    avg_attendance_rates = week_avg_rates.get(week_name)
    if not avg_attendance_rates and not is_anonymous:
        avg_attendance_rates = get_six_month_averages(list(all_names), week_name)
        week_avg_rates[week_name] = avg_attendance_rates
        session['week_avg_rates'] = week_avg_rates

    event_totals = get_event_totals(week_name, main_district) if not is_anonymous else {}

    attendance_table_html = render_attendance_table(
        week_name, attendance_data, sorted_attendance_data,
        district_counts, main_district_counts, avg_attendance_rates,
        event_name=event_name,
        is_history_page=True,
        event_totals=event_totals
    )

    return jsonify({'attendance_table': attendance_table_html})

def classify_attendance_for_week(week_data):
    date, data, week_display, event_name = week_data
    attended = data['attended']
    not_attended = data['not_attended']
    main_district = None
    district_counts = {}
    main_district_counts = {}
    age_categories = ['青職以上', '大專', '中學', '小學', '學齡前']

    is_anonymous = session.get('user', {}).get('role') == 'anonymous'
    if not is_anonymous:
        records = db[COLLECTION_NAME].find({"week_display": week_display, "event_name": event_name})
        age_mapping = {(record["district"], record["name"]): record["age_group"] for record in records}
    else:
        age_mapping = {}

    for district in set(attended.keys()).union(not_attended.keys()):
        main_district_value = parse_district(district)[0]
        if not main_district:
            main_district = main_district_value
        district_counts[district] = {'total': 0, 'ages': {age: 0 for age in age_categories}}
        main_district_counts[main_district_value] = {'total': 0, 'ages': {age: 0 for age in age_categories}}

    for district, names in attended.items():
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
    if not is_authenticated():
        return jsonify({"error": "請先登入"}), 401
    
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
    if not is_authenticated():
        return redirect(url_for('login'))
    
    if DB_OFFLINE:
        return render_template('index.html',
                               error="資料庫離線，無法查看歷史紀錄",
                               version=get_version_info(), is_anonymous=True)
    is_anonymous = session.get('user', {}).get('role') == 'anonymous'
    if is_anonymous:
        return render_template('index.html', error="匿名使用者無法查看歷史紀錄", version=get_version_info(), is_anonymous=True)
    
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
            version=get_version_info(),
            is_anonymous=is_anonymous
        )

@app.route('/average_attendance')
def average_attendance():
    if not is_authenticated():
        return redirect(url_for('login'))
    
    if DB_OFFLINE:
        return render_template('index.html',
                               error="資料庫離線，無法查看半年平均出席",
                               version=get_version_info(), is_anonymous=True)
    is_anonymous = session.get('user', {}).get('role') == 'anonymous'
    if is_anonymous:
        return render_template('index.html', error="匿名使用者無法查看半年平均出席", version=get_version_info(), is_anonymous=True)
    
    try:
        districts = db[COLLECTION_NAME].distinct("district")
        main_districts = sorted(
            set(parse_district(d)[0] for d in districts if parse_district(d)[0]),
            key=lambda x: chinese_to_int(x[0])
        )
        logger.info(f"Loaded main districts for average attendance: {main_districts}")
        return render_template(
            'average_attendance.html',
            main_districts=main_districts,
            version=get_version_info()
        )
    except Exception as e:
        logger.error(f"Failed to load average attendance page: {str(e)}")
        return render_template(
            'index.html',
            error="無法載入半年平均出席頁面",
            version=get_version_info(),
            is_anonymous=is_anonymous
        )

@app.route('/get_weeks_for_district/<district>')
def get_weeks_for_district(district):
    if not is_authenticated():
        return jsonify({"error": "請先登入"}), 401
    
    is_anonymous = session.get('user', {}).get('role') == 'anonymous'
    if is_anonymous:
        return jsonify({"error": "匿名使用者無法訪問此資源"}), 403
    
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
        weeks.sort(key=parse_week_display, reverse=True)
        logger.info(f"Loaded {len(weeks)} weeks for district {district}")
        return jsonify({"weeks": [{"date": week, "display": week} for week in weeks]})
    except Exception as e:
        logger.error(f"Failed to get weeks for district {district}: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/get_history_data/<district>/<path:week_display>')
def get_history_data(district, week_display):
    if not is_authenticated():
        return jsonify({"error": "請先登入"}), 401
    
    is_anonymous = session.get('user', {}).get('role') == 'anonymous'
    if is_anonymous:
        return jsonify({"error": "匿名使用者無法訪問此資源"}), 403
    
    try:
        records = db[COLLECTION_NAME].find({
            "district": {"$regex": f"^{district}"},
            "week_display": week_display,
            "event_name": "主日"
        })
        records_list = list(records)
        logger.debug(f"Fetched {len(records_list)} records for {district} on {week_display} with event_name: 主日")

        # ─────────────────────────────────────────
        # Return empty table even the lord day is missing
        # ─────────────────────────────────────────
        missing_sunday = False
        if not records_list:
            missing_sunday = True         # missing sunday
        attended = {}
        not_attended = {}
        district_counts = {}
        main_district_counts = {}
        age_categories = ['青職以上', '大專', '中學', '小學', '學齡前']
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

        # if missing sunday, give zeros  
        if missing_sunday:
            age_categories = ['青職以上', '大專', '中學', '小學', '學齡前']
            district_counts = { '總計': 0 }
            main_district_counts = { district: {'total':0, 'ages':{age:0 for age in age_categories}} }
            attendance_data = {'attended': {}, 'not_attended': {}}
            all_attendance_data = [(datetime.now(), attendance_data, week_display, None)]

        all_weeks = db[COLLECTION_NAME].distinct("week_display", {
            "district": {"$regex": f"^{district}"},
            "event_name": "主日"
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
                "week_display": prev_week,
                "event_name": "主日"
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
        week_avg_rates = session.get('week_avg_rates', {})
        avg_attendance_rates = week_avg_rates.get(week_display)
        if not avg_attendance_rates and not is_anonymous:
            avg_attendance_rates = get_six_month_averages(list(all_names), week_display)
            week_avg_rates[week_display] = avg_attendance_rates
            session['week_avg_rates'] = week_avg_rates

        event_name = "缺少主日數據" if missing_sunday else get_event_name(week_display)
        event_totals = {} if missing_sunday else get_event_totals(week_display, district)


        logger.debug(f"Event name for {week_display} in district {district}: {event_name}")
        logger.info(f"Passing event_totals to render for {week_display}, {district}: {event_totals}")

        attendance_table_html = render_attendance_table(
            week_display, attendance_data, all_attendance_data,
            district_counts, main_district_counts, avg_attendance_rates,
            event_name=event_name,
            is_history_page=True,
            event_totals=event_totals
        )

        logger.info(f"Rendered history data for {district} on {week_display}")
        return jsonify({'attendance_table': attendance_table_html})
    except Exception as e:
        logger.error(f"Failed to get history data for {district} on {week_display}: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/get_average_attendance_data/<district>/<date>')
def get_average_attendance_data(district, date):
    if not is_authenticated():
        return jsonify({"error": "請先登入"}), 401
    
    is_anonymous = session.get('user', {}).get('role') == 'anonymous'
    if is_anonymous:
        return jsonify({"error": "匿名使用者無法訪問此資源"}), 403
    
    try:
        end_date = datetime.strptime(date, '%Y-%m-%d')
        # Get all unique event names from the database
        event_names = db[COLLECTION_NAME].distinct("event_name")
        logger.debug(f"Retrieved event names: {event_names}")
        
        # Get trimmed mean data for all events
        trimmed_mean_data_list = []
        for event in event_names:
            if event == "未指定活動":
                continue  # Skip default event name
            data = get_six_month_trimmed_mean_by_event(district, end_date, event)
            if any(data["districts"].values()) or data["counts"].get(district):
                trimmed_mean_data_list.append(data)
        
        if not trimmed_mean_data_list:
            logger.warning(f"No average attendance data for {district} up to {date}")
            return jsonify({
                'attendance_table': '<div class="district-section"><table class="excel-table"><tr class="title-row"><th>無資料</th></tr></table></div>'
            }), 404

        # Sort by event name to ensure consistent order
        trimmed_mean_data_list.sort(key=lambda x: x["event_name"])
        attendance_table_html = render_average_attendance_table(district, end_date, trimmed_mean_data_list)
        logger.info(f"Rendered average attendance data for {district} up to {date}")
        return jsonify({'attendance_table': attendance_table_html})
    except ValueError:
        logger.error(f"Invalid date format: {date}")
        return jsonify({"error": "無效的日期格式"}), 400
    except Exception as e:
        logger.error(f"Failed to get average attendance data for {district} on {date}: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
