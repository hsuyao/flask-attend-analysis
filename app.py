from flask import Flask, request, jsonify, send_file, redirect, url_for, session, render_template
from flask_session import Session
from io import BytesIO
import uuid
import os
from config import logger, db, COLLECTION_NAME
from excel_handler import process_excel, generate_excel
from render_table import render_attendance_table
from database import init_database, get_six_month_averages
from utils import parse_district
from datetime import datetime

app = Flask(__name__)
app.config['SESSION_TYPE'] = 'filesystem'
app.config['SECRET_KEY'] = 'your-secret-key-here'
Session(app)

def get_version_info():
    try:
        with open('/app/version_info.txt', 'r') as f:
            return f.read().strip()
    except Exception:
        return "Unknown-Unknown"

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

    try:
        result = process_excel(file.stream, file_extension)
        if not result['all_attendance_data']:
            return render_template('index.html', error="No attendance records found", version=get_version_info())

        session['latest_analytic_date'] = result['latest_analytic_date']
        session['latest_attendance_data'] = result['latest_attendance_data']
        session['latest_week_display'] = result['latest_week_display']
        session['latest_district_counts'] = result['latest_district_counts']
        session['latest_main_district'] = result['latest_main_district']
        session['latest_main_district_counts'] = result['latest_main_district_counts']
        session['all_attendance_data'] = result['all_attendance_data']
        return redirect(url_for('result'))
    except Exception as e:
        logger.error(f"Processing error: {str(e)}")
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

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

    # Cache attendance rates
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

    records = db[COLLECTION_NAME].find({"date": date.strftime('%Y-%m-%d')}, hint="date_name_attended_idx")
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

if __name__ == '__main__':
    init_database()
    port = int(os.getenv('PORT', 5000))
    logger.info(f"Starting server on port {port}")
    app.run(debug=False, host='0.0.0.0', port=port)
