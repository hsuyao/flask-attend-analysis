from flask import Flask, request, jsonify, send_file, redirect, url_for, session, render_template, Response
from flask_session import Session
from io import BytesIO
import uuid
import os
import traceback
import sqlite3
from config import logger
from excel_handler import process_excel, generate_excel
from render_table import render_attendance_table, render_stats_table
from database import init_database, get_six_month_average, DATABASE_PATH
from utils import parse_district  # 新增導入

app = Flask(__name__)

app.config['SESSION_TYPE'] = 'filesystem'
app.config['SECRET_KEY'] = 'your-secret-key-here'
Session(app)

def get_version_info():
    try:
        with open('/app/version_info.txt', 'r') as f:
            version = f.read().strip()
        return version
    except Exception as e:
        logger.error(f"讀取版本資訊時發生錯誤: {str(e)}")
        return "Unknown-Unknown"

@app.route('/')
def index():
    version = get_version_info()
    return render_template('index.html', version=version)

@app.route('/upload', methods=['POST'])
def upload_file():
    logger.info("收到上傳請求")
    if 'file' not in request.files:
        logger.error("未上傳檔案")
        return jsonify({"error": "未上傳檔案"}), 400
    
    file = request.files['file']
    if not file or file.filename == '':
        logger.error("未選擇檔案")
        return jsonify({"error": "未選擇檔案"}), 400

    filename = file.filename.lower()
    logger.debug(f"已上傳檔案: {filename}")
    if not (filename.endswith('.xls') or filename.endswith('.xlsx')):
        logger.error("無效的檔案格式")
        return jsonify({"error": "僅支援 .xls 和 .xlsx 檔案"}), 400
    
    file_extension = '.xls' if filename.endswith('.xls') else '.xlsx'
    
    try:
        result = process_excel(file.stream, file_extension)
        
        if not result['all_attendance_data']:
            version = get_version_info()
            return render_template('index.html', error="上傳的檔案中無任何出勤紀錄，請檢查資料後重新上傳。", version=version)

        session['latest_analytic_date'] = result['latest_analytic_date']
        session['latest_attendance_data'] = result['latest_attendance_data']
        session['latest_week_display'] = result['latest_week_display']
        session['latest_district_counts'] = result['latest_district_counts']
        session['latest_main_district'] = result['latest_main_district']
        session['latest_main_district_counts'] = result['latest_main_district_counts']
        session['all_attendance_data'] = result['all_attendance_data']
        
        return redirect(url_for('result'))
    except Exception as e:
        logger.error(f"處理錯誤: {str(e)}")
        logger.debug(f"完整錯誤追蹤: {traceback.format_exc()}")
        return jsonify({"error": f"處理失敗: {str(e)}"}), 500

@app.route('/result')
def result():
    latest_attendance_data = session.get('latest_attendance_data')
    latest_week_display = session.get('latest_week_display', "尚未有週次資料")
    latest_district_counts = session.get('latest_district_counts')
    latest_main_district_counts = session.get('latest_main_district_counts')
    all_attendance_data = session.get('all_attendance_data', [])
    
    if not latest_attendance_data or not latest_attendance_data.get('attended'):
        version = get_version_info()
        return render_template('index.html', error="最新週無有效出勤資料，請檢查檔案內容。", version=version)

    all_attendance_data.sort(key=lambda x: x[0])
    
    # 計算每人的半年平均出勤率
    latest_date = all_attendance_data[-1][0]
    avg_attendance_rates = {}
    for district, names in latest_attendance_data['attended'].items():
        for name in names:
            avg_attendance_rates[name] = get_six_month_average(name, latest_date)
    for district, names in latest_attendance_data['not_attended'].items():
        for name in names:
            avg_attendance_rates[name] = get_six_month_average(name, latest_date)

    attendance_table_html = render_attendance_table(
        latest_week_display,
        latest_attendance_data,
        all_attendance_data,
        latest_district_counts,
        latest_main_district_counts,
        avg_attendance_rates
    )
    
    week_options = [(week_name, idx) for idx, (_, _, week_name) in enumerate(all_attendance_data)]
    
    version = get_version_info()
    return render_template(
        'result.html',
        attendance_table_html=attendance_table_html,
        stats_table_html="",  # 初始時不渲染統計表，交由 AJAX 處理
        has_file_stream=True,
        week_options=week_options,
        selected_week_idx=len(all_attendance_data) - 1 if all_attendance_data else 0,
        version=version
    )

@app.route('/get_week_data/<int:week_idx>')
def get_week_data(week_idx):
    all_attendance_data = session.get('all_attendance_data', [])
    
    if not all_attendance_data or week_idx < 0 or week_idx >= len(all_attendance_data):
        return jsonify({
            'attendance_table': '<div class="district-section"><table class="excel-table"><tr class="title-row"><th>無資料</th></tr></table></div>',
            'stats_table': ''
        }), 400
    
    date, attendance_data, week_name = all_attendance_data[week_idx]
    
    # 計算對應週次的統計資料
    latest_main_district = session.get('latest_main_district', '')
    _, district_counts, _, main_district, main_district_counts = classify_attendance_for_week(all_attendance_data[week_idx])
    if not main_district:
        main_district = latest_main_district

    avg_attendance_rates = {}
    for district, names in attendance_data['attended'].items():
        for name in names:
            avg_attendance_rates[name] = get_six_month_average(name, date)
    for district, names in attendance_data['not_attended'].items():
        for name in names:
            avg_attendance_rates[name] = get_six_month_average(name, date)
    
    attendance_table_html = render_attendance_table(
        week_name,
        attendance_data,
        all_attendance_data,
        district_counts,
        main_district_counts,
        avg_attendance_rates
    )
    
    # 生成統計表
    stats_table_html = ""
    if main_district and district_counts and main_district_counts:
        stats_table_html = render_stats_table(main_district, district_counts, main_district_counts)
    
    return jsonify({
        'attendance_table': attendance_table_html,
        'stats_table': stats_table_html
    })

def classify_attendance_for_week(week_data):
    """為特定週次重新計算出勤和統計資料"""
    date, data, _ = week_data
    attended = data['attended']
    not_attended = data['not_attended']
    
    main_district = None
    district_counts = {}
    main_district_counts = {}
    age_categories = ['青職以上', '大專', '中學', '大學', '小學', '學齡前']
    
    # 從資料庫中獲取該週次所有人的年齡分組
    conn = sqlite3.connect(DATABASE_PATH)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT district, name, age_group FROM attendance_records
        WHERE date = ?
    ''', (date.strftime('%Y-%m-%d'),))
    records = cursor.fetchall()
    conn.close()
    
    age_mapping = {(district, name): age_group for district, name, age_group in records}
    
    # 初始化統計結構
    for district in set(attended.keys()).union(not_attended.keys()):
        main_district_value = parse_district(district)[0]
        if not main_district:
            main_district = main_district_value
        if district not in district_counts:
            district_counts[district] = {'total': 0, 'ages': {age: 0 for age in age_categories}}
        if main_district_value not in main_district_counts:
            main_district_counts[main_district_value] = {'total': 0, 'ages': {age: 0 for age in age_categories}}
    
    # 計算出勤統計
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

@app.route('/download', methods=['GET]'])
def download_file():
    all_attendance_data = session.get('all_attendance_data', [])
    if not all_attendance_data:
        return jsonify({"error": "無可用的處理檔案"}), 404
    
    file_stream = generate_excel(all_attendance_data)
    file_stream.seek(0)
    return send_file(
        file_stream,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"analyzed_{uuid.uuid4().hex}.xlsx"
    )

if __name__ == '__main__':
    init_database()
    port = int(os.getenv('PORT', 5000))
    logger.info(f"正在啟動伺服器，端口 {port}")
    app.run(debug=False, host='0.0.0.0', port=port)
